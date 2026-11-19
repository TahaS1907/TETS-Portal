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
ENGINE_JS = "\n".join(line for line in (Path(__file__).with_name("qlik_engine.js")).read_text(encoding="utf-8")
                      .splitlines() if not line.startswith("//")).strip()


def _unique(headers: list[str]) -> list[str]:
    seen, out = {}, []
    for h in headers:
        n = seen.get(h, 0)
        out.append(h if n == 0 else f"{h}_{n}")
        seen[h] = n + 1
    return out


class QlikError(RuntimeError):
    pass


class QlikSession:
    def __init__(self, cfg: dict, username: str = "", password: str = "", log=None):
        self.cfg = cfg["qlik"]
        self.username, self.password = username, password
        self.log = log or (lambda m: None)
        self.timeout = int(self.cfg["zaman_asimi_sn"]) * 1000
        self.download_dir = Path(tempfile.mkdtemp(prefix="tets_qlik_"))
        self._pw = self.browser = self.context = self.page = None

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
                page.locator(TABLE_SEL).first.wait_for(state="visible", timeout=20000)
                tid = page.locator(TABLE_SEL).first.evaluate(
                    "e => { const t = e.closest('[tid]') || e.querySelector('[tid]'); return t ? t.getAttribute('tid') : null }")
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
                rx = re.compile(rf"^\s*{re.escape(t)}\s*(\.\.\.|…|>|›)?\s*$", re.I)
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

    def export(self, app: dict, selections: dict, must_have: list[str] | None = None) -> pd.DataFrame:
        """Tablo verisini alir. Once Qlik motorundan (Engine API) dogrudan, olmazsa menuden indirerek."""
        self.resolve_sheet(app)
        method = self.cfg.get("export_yontemi", "otomatik")
        if method in ("otomatik", "motor"):
            try:
                return self.export_engine(app, selections, must_have or [])
            except Exception as e:
                if method == "motor":
                    raise
                self.log(f"Motor ile okunamadi ({e}); menu ile indirme deneniyor")
        return self.export_menu(app, selections)

    def engine_url(self, app_id: str) -> str:
        base = self.cfg.get("motor_url") or (
            self._root().replace("https://", "wss://", 1).replace("http://", "ws://", 1) + "/app/")
        return base + quote(app_id, safe="")

    def export_engine(self, app: dict, selections: dict, must_have: list[str]) -> pd.DataFrame:
        sel = {}
        for f, v in selections.items():
            vals = [str(x) for x in (v if isinstance(v, (list, tuple)) else [v]) if str(x).strip()]
            if f and vals:
                sel[f] = vals
        self.log(f"Qlik sorgusu (motor): {sel}")
        res = self.page.evaluate(ENGINE_JS, {
            "wsUrl": self.engine_url(app["app_id"]), "appId": app["app_id"],
            "objectId": app.get("object_id", ""), "sheetId": app.get("sheet_id", ""),
            "selections": sel, "mustHave": must_have,
            "maxRows": int(self.cfg.get("max_satir", 200000)), "timeoutMs": self.timeout,
        })
        for line in res.get("log", []):
            self.log("  " + line)
        if res.get("objectId") and res["objectId"] != app.get("object_id"):
            app["object_id"] = res["objectId"]
            if app.get("_bolum"):
                save_learned(app["_bolum"], {"object_id": res["objectId"]})
        headers = _unique(res["headers"])
        return pd.DataFrame(res["rows"], columns=headers, dtype=object)

    def export_menu(self, app: dict, selections: dict) -> pd.DataFrame:
        page, ex = self.page, self.cfg["export"]
        url = self.sheet_url(app, selections)
        self.log(f"Qlik sorgusu: {selections}")
        page.goto(url, wait_until="domcontentloaded")
        oid = app.get("object_id", "")
        target = page.locator(f"[tid='{oid}'], [data-qid='{oid}']").first if oid else page.locator(TABLE_SEL).first
        target.wait_for(state="visible", timeout=self.timeout)
        page.wait_for_timeout(1500)  # secimlerin uygulanmasi icin
        if page.get_by_text(re.compile(r"(No data|Veri yok|Gösterilecek veri yok)", re.I)).count():
            return pd.DataFrame()
        self._open_download_menu(target, ex)
        self._click_text(ex["veri_secenegi"], timeout=3000)
        with page.expect_download(timeout=self.timeout) as dl:
            self._click_text(ex["onay_butonu"], timeout=3000)
            link = page.locator(ex["link_selector"]).first
            link.wait_for(state="visible", timeout=self.timeout)
            link.click()
        path = self.download_dir / (dl.value.suggested_filename or "export.xlsx")
        dl.value.save_as(path)
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

    def _frame(self, selections: dict) -> FrameSource:
        sel = {self.f.get(k, k): v for k, v in selections.items() if v}
        ck = tuple(sorted((k, key(v)) for k, v in sel.items()))
        if ck not in self._cache:
            raw = self.s.export(self.app, sel, [self.f.get("pn", "PN"), self.f.get("sn", "SN")])
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

    def aircraft_events(self, aircraft, start, end):
        return self._frame({"aircraft": aircraft}).aircraft_events(aircraft, start, end)

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
