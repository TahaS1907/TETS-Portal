"""Canli Qlik Sense baglantisi (Playwright ile).

Akis:
  1) Tarayici acilir, Qlik hub'a gidilir, kullanici adi/sifre girilir
     (form girisi, Windows/NTLM girisi veya manuel giris).
  2) Her sorgu icin Qlik "App Integration" URL'i ile ilgili sheet secimli acilir:
     /sense/app/<app>/sheet/<sheet>/state/analysis/options/clearselections/select/<alan>/<deger>
  3) Tablo nesnesine sag tik -> Indir -> Veri ile Excel export alinir, pandas ile okunur.

Sirket portalinin sayfa yapisi bu kodu yazan tarafindan gorulmedigi icin tum
secici/etiketler config.yaml'dan degistirilebilir. Ayar bulmak icin
"Kalibrasyon" modu tarayiciyi Playwright Inspector ile acar.
"""
from __future__ import annotations

import os
import re
import tempfile
import time
from pathlib import Path
from urllib.parse import quote

import pandas as pd

from ..columns import normalize_frame
from ..config import save_learned
from ..normalize import key
from .base import DataSource, FrameSource
from .excel_source import _fix_header


SHEET_RE = re.compile(r"/sense/app/([^/?#]+)/sheet/([^/?#]+)")
TABLE_SEL = ".qv-object-table, .qv-object-sn-table, .qv-object-pivot-table, .qv-object-sn-pivot-table, .qv-object"


MORE_SEL = ("button[aria-label*='More' i], button[title*='More' i], button[aria-label*='Daha' i], "
            "button[title*='Daha' i], [data-testid*='more' i], [tid='nav-menu-more'], button[aria-haspopup='menu']")
SCREEN_JS = "\n".join(line for line in (Path(__file__).with_name("qlik_screen.js")).read_text(encoding="utf-8")
                      .splitlines() if not line.startswith("//")).strip()
ENGINE_JS = "\n".join(line for line in (Path(__file__).with_name("qlik_engine.js")).read_text(encoding="utf-8")
                      .splitlines() if not line.startswith("//")).strip()


def _unique(headers: list[str]) -> list[str]:
    seen, out = {}, []
    for h in headers:
        n = seen.get(h, 0)
        out.append(h if n == 0 else f"{h}_{n}")
        seen[h] = n + 1
    return out


TABLE_KINDS = [".qv-object-sn-table", ".qv-object-table", ".qv-object-pivot-table", ".qv-object-sn-pivot-table",
               ".qv-object"]
# Sayfadaki EN BUYUK tabloyu sec (filtre kutulari/KPI'lar secilmesin)
DETECT_JS = """(sels) => {
  for (const sel of sels) {
    const els = [...document.querySelectorAll(sel)].filter(e => e.offsetWidth && e.offsetHeight);
    if (!els.length) continue;
    els.sort((a, b) => b.offsetWidth * b.offsetHeight - a.offsetWidth * a.offsetHeight);
    const t = els[0].closest('[tid]') || els[0].querySelector('[tid]');
    if (t) return t.getAttribute('tid');
  }
  return null;
}"""


# Tablo hazir mi? ok: basliklar gorunuyor | nodata | limit: 'maksimum kayit asildi' mesaji | bekle
CHECK_JS = """({ oid, tableSel, known }) => {
  const root = (oid && document.querySelector(`[tid="${oid}"], [data-qid="${oid}"]`)) || document.querySelector(tableSel);
  if (!root) return 'bekle';
  const t = (root.innerText || '');
  if (/(No data|Veri yok|Gösterilecek veri yok)/i.test(t)) return 'nodata';
  if (/(Maksimum kay|maximum number of|exceed|calculation condition|hesaplama ko)/i.test(t)) return 'limit';
  if (root.querySelectorAll('th, [role=columnheader]').length >= 2) return 'ok';
  const norm = s => String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const K = new Set(known.map(norm));
  let n = 0;
  for (const e of root.querySelectorAll('*')) {
    if ([...e.children].some(c => (c.innerText || '').trim())) continue;
    if (K.has(norm(e.innerText))) n++;
    if (n >= 3) return 'ok';
  }
  return 'bekle';
}"""


class QlikError(RuntimeError):
    pass


class QlikSession:
    def __init__(self, cfg: dict, username: str = "", password: str = "", log=None):
        self.cfg = cfg["qlik"]
        self.cfg_root = cfg
        self.username, self.password = username, password
        self.log = log or (lambda m: None)
        self.timeout = int(self.cfg["zaman_asimi_sn"]) * 1000
        self.download_dir = Path(tempfile.mkdtemp(prefix="tets_qlik_"))
        self._pw = self.browser = self.context = self.page = None
        self.ws_urls: list[str] = []  # Qlik sayfasinin kendi actigi motor baglantilari
        self.good_method = ""         # bu oturumda ise yarayan okuma yontemi
        self.dead_methods: set[str] = set()

    # ------------------------------------------------------------------ yasam dongusu
    def __enter__(self):
        self.start()
        return self

    def __exit__(self, *exc):
        self.close()

    def start(self):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError as e:
            raise QlikError("Playwright kurulu degil. KURULUM.bat calistirin.") from e
        if not self.cfg["base_url"]:
            raise QlikError("config.yaml > qlik > base_url bos. Qlik Sense adresini girin.")
        self._pw = sync_playwright().start()
        channel = self.cfg.get("tarayici") or None
        opts = dict(headless=bool(self.cfg.get("gorunmez")), accept_downloads=True, ignore_https_errors=True,
                    no_viewport=True, args=["--start-maximized"])
        if self.cfg["giris"]["mod"] == "windows" and self.username:
            opts["http_credentials"] = {"username": self.username, "password": self.password}
        if self.cfg.get("tarayici_yolu"):
            opts["executable_path"] = self.cfg["tarayici_yolu"]
            channel = "chromium"
        # Kalici profil: oturum cerezleri saklanir, cogu zaman sifre tekrar sorulmaz.
        profile = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "TETS-Portal" / f"profil_{channel or 'chromium'}"
        profile.mkdir(parents=True, exist_ok=True)
        try:
            self.context = self._pw.chromium.launch_persistent_context(
                str(profile), channel=None if channel == "chromium" else channel, **opts)
        except Exception:
            self.log(f"'{channel}' acilamadi, Playwright Chromium deneniyor")
            self.context = self._pw.chromium.launch_persistent_context(str(profile), **opts)
        self.context.set_default_timeout(self.timeout)
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        self.page.on("websocket", lambda w: self.ws_urls.append(w.url))
        self.login()

    @property
    def alive(self) -> bool:
        try:
            return self._pw is not None and self.page is not None and not self.page.is_closed()
        except Exception:
            return False

    def close(self):
        try:
            self.context and self.context.close()
        except Exception:
            pass
        if self._pw:
            try:
                self._pw.stop()
            except Exception:
                pass
        self._pw = self.context = self.page = None

    # ------------------------------------------------------------------ giris
    def _root(self) -> str:
        base = self.cfg["base_url"].rstrip("/")
        proxy = self.cfg.get("sanal_proxy", "").strip("/")
        return f"{base}/{proxy}" if proxy else base

    def _logged_in(self) -> bool:
        url = self.page.url
        return url.startswith(self._root()) and ("/hub" in url or "/sense/app/" in url) \
            and "login" not in url.lower()

    def _wait(self, cond, what: str):
        """Kullanici tarayicida bir sey yaparken bekler (tarayici kapatilirsa hata verir)."""
        deadline = time.time() + int(self.cfg.get("bekleme_sn", 600))
        while time.time() < deadline:
            if not self.alive:
                raise QlikError("Tarayici kapatildi.")
            try:
                if cond():
                    return
                self.page.wait_for_timeout(700)
            except QlikError:
                raise
            except Exception:
                time.sleep(0.7)  # sayfa gecisi sirasinda olabilir
        raise QlikError(f"Zaman asimi: {what}")

    def login(self):
        g = self.cfg["giris"]
        page = self.page
        self.log("» Qlik Sense aciliyor...")
        page.goto(self._root() + "/hub", wait_until="domcontentloaded")
        page.wait_for_timeout(1500)
        if self._logged_in():
            self.log("» Oturum zaten acik.")
            return
        if g["mod"] == "form" and self.username and self.password:
            try:
                user = page.locator(g["kullanici_selector"]).first
                user.wait_for(state="visible", timeout=15000)
                user.fill(self.username)
                pw = page.locator(g["sifre_selector"]).first
                if not pw.is_visible():  # iki adimli (once kullanici, sonra sifre) ekranlar
                    page.locator(g["gonder_selector"]).first.click()
                    pw.wait_for(state="visible", timeout=15000)
                pw.fill(self.password)
                page.locator(g["gonder_selector"]).first.click()
            except Exception:
                self.log("Giris formu otomatik doldurulamadi, tarayicidan giris yapin.")
        self.log("» Tarayicida Qlik Sense'e giris yapin (sifrenizi tarayiciya girin)...")
        self._wait(self._logged_in, "Qlik girisi yapilmadi")
        self.log("» Giris basarili.")

    def detect_table(self) -> str | None:
        """Acik sayfadaki en buyuk tablonun nesne kimligini bulur."""
        try:
            self.page.locator(TABLE_SEL).first.wait_for(state="visible", timeout=20000)
            self.page.wait_for_timeout(1000)
            return self.page.evaluate(DETECT_JS, TABLE_KINDS)
        except Exception:
            return None

    def resolve_sheet(self, app: dict) -> dict:
        """Uygulamanin tablo sayfasi bilinmiyorsa kullanicinin tarayicida acmasini bekler, sonra hatirlar."""
        if app.get("app_id") and app.get("sheet_id"):
            return app
        name = app.get("ad") or "ilgili portal"
        page, want = self.page, app.get("app_id", "")

        def on_sheet():
            m = SHEET_RE.search(page.url)
            return bool(m) and (not want or m.group(1) == want)

        if not on_sheet():
            if want and f"/app/{want}" not in page.url:
                page.goto(f"{self._root()}/sense/app/{want}/overview", wait_until="domcontentloaded")
            self.log(f"» Tarayicida '{name}' icinde tablonun bulundugu sayfayi acin...")
            self._wait(on_sheet, f"'{name}' sayfasi acilmadi")
        m = SHEET_RE.search(page.url)
        learned = {"app_id": m.group(1), "sheet_id": m.group(2)}
        if not app.get("object_id"):
            try:
                tid = self.detect_table()
                if tid:
                    learned["object_id"] = tid
            except Exception:
                self.log("Tablo nesnesi otomatik bulunamadi; sayfadaki ilk tablo kullanilacak.")
        app.update(learned)
        if app.get("_bolum"):
            save_learned(app["_bolum"], learned)
        self.log(f"» '{name}' sayfasi algilandi ve kaydedildi. Otomasyon basliyor.")
        return app

    # ------------------------------------------------------------------ export
    def sheet_url(self, app: dict, selections: dict[str, str | list]) -> str:
        url = f"{self._root()}/sense/app/{app['app_id']}"
        if app.get("sheet_id"):
            url += f"/sheet/{app['sheet_id']}"
        url += "/state/analysis/options/clearselections"
        for field, val in selections.items():
            vals = val if isinstance(val, (list, tuple)) else [val]
            vals = [str(v) for v in vals if str(v).strip()]
            if not field or not vals:
                continue
            joined = ";".join(f"[{v}]" for v in vals)
            url += f"/select/{quote(field, safe='')}/{quote(joined, safe='')}"
        return url

    def _click_text(self, labels: list[str], timeout=4000, required=False) -> bool:
        page = self.page
        end = time.time() + timeout / 1000
        while time.time() < end:
            for t in labels:
                # "Download as..." gibi devami olan menu yazilarini da yakalar
                rx = re.compile(rf"^\s*{re.escape(t)}(?![A-Za-zÇĞİÖŞÜçğıöşü]).*$", re.I)
                for loc in (page.get_by_role("menuitem", name=rx),
                            page.get_by_role("button", name=rx),
                            page.get_by_text(rx)):
                    try:
                        if loc.count() and loc.first.is_visible():
                            loc.first.click()
                            return True
                    except Exception:
                        continue
            page.wait_for_timeout(250)
        if required:
            raise QlikError(f"Menude su seceneklerden biri bulunamadi: {labels}")
        return False

    METHODS = {"sayfa": "sayfanin kendi Qlik baglantisi", "ekran": "filtre + ekrandan okuma (indirmeden)",
               "menu": "sag tik > Download as > Apply", "motor": "motor (WebSocket) baglantisi"}
    # FA takibinde indirme YOK: otomatik = sayfa baglantisi, olmazsa ekrandan (kaydirarak) okuma.
    CHOICES = {"otomatik": ["sayfa", "ekran"], "ekran": ["ekran"], "menu": ["menu"],
               "sayfa": ["sayfa"], "motor": ["motor"]}

    def export(self, app: dict, selections: dict, must_have: list[str] | None = None,
               validate=None, methods: list[str] | None = None) -> pd.DataFrame:
        """Filtreleri uygulayip tablo verisini alir.

        Yontemler sirayla denenir; ise yarayan yontem oturum boyunca hatirlanir ve hep o kullanilir.
        """
        self.resolve_sheet(app)
        chosen = self.cfg.get("export_yontemi", "otomatik")
        order = list(methods or self.CHOICES.get(chosen, self.CHOICES["otomatik"]))
        order = [m for m in order if m not in self.dead_methods] or order
        if self.good_method in order and not methods:
            order.remove(self.good_method)
            order.insert(0, self.good_method)
        errors, redetected = [], False
        i = 0
        while i < len(order):
            m = order[i]
            i += 1
            try:
                if m == "sayfa":
                    df = self.export_engine(app, selections, must_have or [], "client")
                elif m == "motor":
                    df = self.export_engine(app, selections, must_have or [], "ws")
                elif m == "ekran":
                    df = self.export_screen(app, selections)
                else:
                    df = self.export_menu(app, selections)
                # Yanlis nesnenin dosyasi indiyse (orn. sadece 'AC' kolonu) kullanma, siradaki yonteme gec
                miss = []
                if len(df.columns) and validate:
                    miss = validate(df)        # orn. tarih/hareket/PN/SN kolonlari eksik mi
                elif len(df.columns) and must_have:
                    have = {re.sub(r"[^A-Z0-9]", "", str(c).upper()) for c in df.columns}
                    miss = [c for c in must_have if re.sub(r"[^A-Z0-9]", "", c.upper()) not in have]
                if miss:
                    if not redetected:
                        # kayitli nesne yanlis (filtre kutusu vb.): unut, en buyuk tabloyu bul, tekrar dene
                        redetected = True
                        tid = self.detect_table()
                        self.log(f"  okunan tabloda {miss} yok; tablo yeniden bulundu: {tid}")
                        if tid and tid != app.get("object_id"):
                            app["object_id"] = tid
                            if app.get("_bolum"):
                                save_learned(app["_bolum"], {"object_id": tid})
                            i -= 1
                            continue
                    raise QlikError(f"okunan tabloda {miss} kolonu yok (gelen: {list(df.columns)[:6]})")
            except Exception as e:
                msg = str(e).split("\n")[0].replace("Page.evaluate: ", "")
                errors.append(f"{self.METHODS[m]}: {msg}")
                self.log(f"  {self.METHODS[m]} olmadi -> {msg}")
                if m in ("sayfa", "motor"):
                    self.dead_methods.add(m)  # bu ortamda calismiyor: oturum boyunca tekrar deneme
                continue
            if self.good_method != m:
                self.log(f"» Veri '{self.METHODS[m]}' yontemiyle okunuyor.")
                self.good_method = m
            return df
        raise QlikError("Qlik tablosu okunamadi. Denenenler: " + " | ".join(errors))

    def export_screen(self, app: dict, selections: dict) -> pd.DataFrame:
        """Filtreleri Qlik'te uygular (secim cubugunda gorunur) ve tabloyu ekrandan okur."""
        page = self.page
        sel = {f: v for f, v in selections.items() if f and str(v).strip()}
        self.log(f"Qlik filtresi: {sel}")
        self._widen()
        page.goto(self.sheet_url(app, sel), wait_until="domcontentloaded")
        oid = app.get("object_id", "")
        oid = oid or self.detect_table() or ""
        target = page.locator(f"[tid='{oid}'], [data-qid='{oid}']").first if oid else page.locator(TABLE_SEL).first
        target.wait_for(state="visible", timeout=self.timeout)
        warmed = getattr(self, "_warmed", False)
        state = self._wait_table(oid, None if warmed else int(self.cfg.get("ucak_secme_sn", 12)))
        if state == "limit" and not warmed:
            # Kullanicinin elle yaptigi gibi: once listeden bir ucak sec (tablo acilsin), sonra ucagi
            # kaldirip PN/SN filtresiyle tekrar ac.
            self._warmed = True
            ac = self._first_aircraft()
            field = (app.get("alanlar") or {}).get("aircraft", "AC")
            if ac:
                self.log(f"» Tablo acilmadi; once ucak seciliyor ({ac}), sonra filtreler uygulanacak")
                page.goto(self.sheet_url(app, {field: ac}), wait_until="domcontentloaded")
                self._wait_table(oid)
                page.goto(self.sheet_url(app, sel), wait_until="domcontentloaded")
                state = self._wait_table(oid)
        if state == "nodata":
            return pd.DataFrame()
        if state != "ok":
            bar = self._selection_bar_text()
            raise QlikError("Tablo ekranda acilmadi"
                            + (" (Qlik: 'Maksimum kayit sayisi asildi' - filtre uygulanmamis)" if state == "limit" else "")
                            + (f"; secim cubugu: {bar}" if bar else ""))
        res = page.evaluate(SCREEN_JS, {"oid": oid, "tableSel": TABLE_SEL,
                                        "maxRows": int(self.cfg.get("ekran_max_satir", 5000)),
                                        "waitMs": int(self.cfg.get("ekran_bekleme_ms", 450)),
                                        "knownHeaders": self._known_headers()})
        self.log(f"  ekrandan {len(res['rows'])} satir, {len(res['headers'])} kolon okundu"
                 + (" (yazilardan tanindi)" if res.get("mode") == "goz" else "")
                 + (" (saga kaydirildi)" if res.get("hscroll") else ""))
        if res.get("truncated"):
            self.log("  UYARI: ekrandan okunan satir limiti doldu")
        return pd.DataFrame(res["rows"], columns=_unique(res["headers"]), dtype=object)

    def engine_urls(self, app_id: str) -> list[str]:
        """Denenecek motor adresleri: once Qlik sayfasinin kendi kullandigi adres."""
        out = [u for u in reversed(self.ws_urls) if f"/app/{app_id}" in u or quote(app_id, safe="") in u]
        if self.cfg.get("motor_url"):
            out.append(self.cfg["motor_url"] + quote(app_id, safe=""))
        out.append(self._root().replace("https://", "wss://", 1).replace("http://", "ws://", 1)
                   + "/app/" + quote(app_id, safe=""))
        return list(dict.fromkeys(out))

    def _goto_app(self, app: dict):
        """Sayfa dogru uygulamada degilse uygulamanin tablo sayfasini acar ve Qlik'in yuklenmesini bekler."""
        page = self.page
        if f"/app/{app['app_id']}" not in page.url:
            page.goto(f"{self._root()}/sense/app/{app['app_id']}/sheet/{app['sheet_id']}/state/analysis",
                      wait_until="domcontentloaded")
        try:
            page.locator(TABLE_SEL).first.wait_for(state="visible", timeout=30000)
        except Exception:
            pass

    def _wait_table(self, oid: str, limit_sn: int | None = None) -> str:
        """Filtreler uygulanip tablo cizilene kadar bekler.
        Doner: ok | nodata | limit (Qlik 'maksimum kayit' gosteriyor) | bekle (zaman asimi)."""
        deadline = time.time() + int(self.cfg.get("ekran_yukleme_sn", 40))
        state, said = "bekle", False
        known = self._known_headers()
        limit_since = None
        while time.time() < deadline:
            try:
                state = self.page.evaluate(CHECK_JS, {"oid": oid, "tableSel": TABLE_SEL, "known": known})
            except Exception:
                state = "bekle"  # sayfa gecisi
            if state in ("ok", "nodata"):
                self.page.wait_for_timeout(600)  # son satirlarin da cizilmesi icin
                return state
            if state == "limit":
                limit_since = limit_since or time.time()
                if limit_sn and time.time() - limit_since > limit_sn:
                    return "limit"  # filtre bu surede gelmediyse ucak adimina gec
            else:
                limit_since = None
            if state == "limit" and not said:
                self.log("  filtrelerin uygulanmasi bekleniyor (tablo 'maksimum kayit' gosteriyor)")
                said = True
            self.page.wait_for_timeout(500)
        return state

    def _first_aircraft(self) -> str:
        """Sayfadaki AC listesinden ilk ucak tescilini alir (orn. 2-ADAE, TC-JFV)."""
        try:
            return self.page.evaluate(r"""() => {
              const rx = /^[A-Z0-9]{1,2}-[A-Z0-9]{3,5}$/;
              for (const e of document.querySelectorAll('*')) {
                if ([...e.children].some(c => (c.innerText || '').trim())) continue;
                const t = (e.innerText || '').trim();
                if (rx.test(t)) return t;
              }
              return '';
            }""") or ""
        except Exception:
            return ""

    def _selection_bar_text(self) -> str:
        try:
            loc = self.page.get_by_text(re.compile(r"No selections applied|Seçim uygulanmadı", re.I))
            return "secim yok (No selections applied)" if loc.count() else ""
        except Exception:
            return ""

    def _known_headers(self) -> list[str]:
        """Ekranda baslik olarak aranacak kolon adlari (config + bilinen varyasyonlar)."""
        from ..columns import ALIASES
        names = [a for al in ALIASES.values() for a in al]
        for sec in ("fa_hareket", "fa_listesi", "envanter"):
            names += [v for v in (self.cfg.get(sec, {}).get("alanlar") or {}).values() if v]
        names += [v for v in (self.cfg_root.get("kolonlar") or {}).values() if v]
        e = self.cfg.get("envanter", {})
        names += [e.get("notes_kolonu", ""), e.get("durum_alani", "")]
        return [n for n in names if n]

    def _widen(self):
        """Okumadan once sayfa alanini genis/uzun yapar: Qlik tabloyu buyuk cizer, daha cok kolon ve satir
        ayni anda gorunur, kaydirma azalir. (config: ekran_genislik / ekran_yukseklik, 0 = kapali)"""
        w, h = int(self.cfg.get("ekran_genislik", 2560)), int(self.cfg.get("ekran_yukseklik", 1600))
        if w and h and not getattr(self, "_widened", False):
            try:
                self.page.set_viewport_size({"width": w, "height": h})
                self._widened = True
            except Exception:
                pass

    def export_engine(self, app: dict, selections: dict, must_have: list[str], mode: str) -> pd.DataFrame:
        sel = {}
        for f, v in selections.items():
            vals = [str(x) for x in (v if isinstance(v, (list, tuple)) else [v]) if str(x).strip()]
            if f and vals:
                sel[f] = vals
        self.log(f"Qlik sorgusu: {sel}")
        self._goto_app(app)
        args = {"appId": app["app_id"], "objectId": app.get("object_id", ""), "sheetId": app.get("sheet_id", ""),
                "selections": sel, "mustHave": must_have,
                "maxRows": int(self.cfg.get("max_satir", 200000)), "timeoutMs": self.timeout}
        attempts = [("client", "")] if mode == "client" else [("ws", u) for u in self.engine_urls(app["app_id"])]
        errors = []
        for mode, url in attempts:
            try:
                res = self.page.evaluate(ENGINE_JS, {**args, "mode": mode, "wsUrl": url})
                break
            except Exception as e:
                msg = str(e).split("\n")[0].replace("Page.evaluate: ", "")
                errors.append(f"{'sayfa baglantisi' if mode == 'client' else url}: {msg}")
                self.log(f"  okunamadi ({errors[-1]})")
        else:
            raise QlikError("Qlik verisi okunamadi -> " + " | ".join(errors))
        self.log(f"  veri {'sayfanin kendi baglantisiyla' if mode == 'client' else 'motor baglantisiyla'} okundu")
        for line in res.get("log", []):
            self.log("  " + line)
        if res.get("objectId") and res["objectId"] != app.get("object_id"):
            app["object_id"] = res["objectId"]
            if app.get("_bolum"):
                save_learned(app["_bolum"], {"object_id": res["objectId"]})
        return pd.DataFrame(res["rows"], columns=_unique(res["headers"]), dtype=object)

    def export_menu(self, app: dict, selections: dict) -> pd.DataFrame:
        page, ex = self.page, self.cfg["export"]
        url = self.sheet_url(app, selections)
        self.log(f"Qlik sorgusu: {selections}")
        page.goto(url, wait_until="domcontentloaded")
        oid = app.get("object_id", "")
        oid = oid or self.detect_table() or ""
        target = page.locator(f"[tid='{oid}'], [data-qid='{oid}']").first if oid else page.locator(TABLE_SEL).first
        target.wait_for(state="visible", timeout=self.timeout)
        page.wait_for_timeout(1500)  # secimlerin uygulanmasi icin
        if page.get_by_text(re.compile(r"(No data|Veri yok|Gösterilecek veri yok)", re.I)).count():
            return pd.DataFrame()
        downloads = []
        def handler(d):
            downloads.append(d)
        page.on("download", handler)
        try:
            self._open_download_menu(target, ex)                      # sag tik > Download as...
            if not downloads:
                self._click_text(ex["veri_secenegi"], timeout=2500)   # (varsa) Excel / Data
            if not downloads:
                self._click_text(ex["onay_butonu"], timeout=8000)     # acilan pencerede Apply
            # Apply'dan sonra dosya ya dogrudan iner ya da "indirmek icin tiklayin" linki cikar
            end = time.time() + self.timeout / 1000
            clicked = False
            while not downloads and time.time() < end:
                link = page.locator(ex["link_selector"]).first
                if not clicked and link.count() and link.is_visible():
                    link.click()
                    clicked = True
                page.wait_for_timeout(300)
            if not downloads:
                raise QlikError("Apply'dan sonra dosya inmedi")
        finally:
            page.remove_listener("download", handler)
        dl = downloads[0]
        path = self.download_dir / (dl.suggested_filename or "export.xlsx")
        dl.save_as(path)
        page.keyboard.press("Escape")
        raw = pd.read_excel(path, dtype=object) if path.suffix.lower() != ".csv" else pd.read_csv(path, dtype=str)
        return _fix_header(raw)

    def _open_download_menu(self, target, ex):
        """Indirme menusunu acar: once sag tik, olmazsa nesnenin '...' (daha fazla) butonu."""
        page = self.page
        box = target.bounding_box() or {"width": 200, "height": 200}
        attempts = [
            lambda: target.click(button="right", position={"x": box["width"] / 2, "y": min(box["height"] / 2, 120)}),
            lambda: (target.hover(), target.locator(MORE_SEL).first.click(timeout=3000)),
            lambda: (target.hover(), page.locator(MORE_SEL).first.click(timeout=3000)),
        ]
        for attempt in attempts:
            try:
                attempt()
            except Exception:
                continue
            if self._click_text(ex["sag_tik_menu"], timeout=3000):
                return
            page.keyboard.press("Escape")
        raise QlikError(f"Indirme menusu bulunamadi (aranan: {ex['sag_tik_menu']}). "
                        "Ayarlar > Kalibrasyon ile menu adlarini kontrol edin.")

    def calibrate(self, app: dict | None = None):
        """Playwright Inspector'i acar: secicileri ve export adimlarini elle bulmak icin."""
        if app and app.get("app_id") and app.get("sheet_id"):
            self.page.goto(self.sheet_url(app, {}))
        self.page.pause()


class QlikSource(DataSource):
    """FA motoru icin canli Qlik veri kaynagi. Her sorgu sonucu oturum boyunca onbellekte tutulur."""

    def __init__(self, session: QlikSession, cfg: dict, owns_session: bool = True):
        self.s, self.owns_session = session, owns_session
        self.app = cfg["qlik"]["fa_hareket"]
        self.f = self.app.get("alanlar", {})
        self.fa_app = cfg["qlik"].get("fa_listesi", {})
        self.overrides = cfg.get("kolonlar", {})
        self._cache: dict[tuple, FrameSource] = {}

    def _missing_columns(self, df: pd.DataFrame) -> list[str]:
        """FA motorunun ihtiyac duydugu kolonlardan eksik olanlar (ekrandan eksik okuma vb.)."""
        from ..columns import map_columns
        found = map_columns(df, self.overrides)
        need = self.s.cfg_root["fa"].get("zorunlu_kolonlar") or ["date", "tx_type", "pn", "sn"]
        return [c for c in need if c not in found]

    def _frame(self, selections: dict) -> FrameSource:
        sel = {self.f.get(k, k): v for k, v in selections.items() if v}
        ck = tuple(sorted((k, key(v)) for k, v in sel.items()))
        if ck not in self._cache:
            raw = self.s.export(self.app, sel, [self.f.get("pn", "PN"), self.f.get("sn", "SN")],
                                validate=self._missing_columns)
            df = normalize_frame(raw, self.overrides, source="qlik") if not raw.empty else _empty()
            self._cache[ck] = FrameSource(df)
        return self._cache[ck]

    def part_history(self, pn="", sn="", grb=""):
        parts = []
        if key(pn) and key(sn):
            parts.append(self._frame({"pn": pn, "sn": sn}).part_history(pn, sn, ""))
        if key(grb):
            parts.append(self._frame({"grb": grb}).part_history("", "", grb))
        if not parts:
            return _empty()
        df = pd.concat(parts, ignore_index=True).drop_duplicates(
            subset=["date", "tx_raw", "pn_k", "sn_k", "aircraft_k", "position", "wo_k"])
        return df.sort_values(["date", "seq"]).reset_index(drop=True)

    def query_levels(self, inst):
        levels = [{"pn": inst.get("pn", "")}]
        if inst.get("pn_group"):
            levels.append({"pn_group": inst["pn_group"]})
        levels.append({})
        return levels

    def aircraft_events(self, aircraft, start, end, **hint):
        sel = {"aircraft": aircraft, **{k: v for k, v in hint.items() if v}}
        return self._frame(sel).aircraft_events(aircraft, start, end)

    def nha_events(self, nha_pn, nha_sn, start, end):
        return self._frame({"nha_pn": nha_pn, "nha_sn": nha_sn}).nha_events(nha_pn, nha_sn, start, end)

    def fa_lookup(self, fa_no):
        if not self.fa_app.get("app_id"):
            return None
        from .excel_source import load_fa_list
        raw = self.s.export(self.fa_app, {self.fa_app.get("alanlar", {}).get("fa_no", "FA No"): fa_no})
        if raw.empty:
            return None
        tmp = self.s.download_dir / "fa_lookup.xlsx"
        raw.to_excel(tmp, index=False)
        return FrameSource(_empty(), load_fa_list(tmp)).fa_lookup(fa_no)

    def close(self):
        if self.owns_session:
            self.s.close()


def _empty() -> pd.DataFrame:
    from ..columns import CANONICAL
    cols = CANONICAL + ["tx_raw", "source", "seq"] + [c + "_k" for c in
                                                    ("pn", "sn", "grb", "aircraft", "nha_pn", "nha_sn", "wo", "new_pn", "new_sn")]
    df = pd.DataFrame(columns=cols)
    df["date"] = pd.to_datetime(df["date"])
    return df
