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

import re
import tempfile
import time
from pathlib import Path
from urllib.parse import quote

import pandas as pd

from ..columns import normalize_frame
from ..normalize import key
from .base import DataSource, FrameSource
from .excel_source import _fix_header


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
        launch = dict(headless=bool(self.cfg.get("gorunmez")))
        if self.cfg.get("tarayici_yolu"):
            launch["executable_path"] = self.cfg["tarayici_yolu"]
            channel = "chromium"
        try:
            self.browser = self._pw.chromium.launch(channel=None if channel == "chromium" else channel, **launch)
        except Exception:
            self.log(f"'{channel}' acilamadi, Playwright Chromium deneniyor")
            self.browser = self._pw.chromium.launch(**launch)
        ctx = dict(accept_downloads=True, ignore_https_errors=True)
        if self.cfg["giris"]["mod"] == "windows" and self.username:
            ctx["http_credentials"] = {"username": self.username, "password": self.password}
        self.context = self.browser.new_context(**ctx)
        self.context.set_default_timeout(self.timeout)
        self.page = self.context.new_page()
        self.login()

    def close(self):
        for obj in (self.context, self.browser):
            try:
                obj and obj.close()
            except Exception:
                pass
        if self._pw:
            self._pw.stop()
        self._pw = None

    # ------------------------------------------------------------------ giris
    def _root(self) -> str:
        base = self.cfg["base_url"].rstrip("/")
        proxy = self.cfg.get("sanal_proxy", "").strip("/")
        return f"{base}/{proxy}" if proxy else base

    def login(self):
        g = self.cfg["giris"]
        page = self.page
        self.log("Qlik Sense aciliyor...")
        page.goto(self._root() + "/hub", wait_until="domcontentloaded")
        ok = g.get("basarili_url_icerir", "/hub")
        if g["mod"] == "form":
            user = page.locator(g["kullanici_selector"]).first
            try:
                user.wait_for(state="visible", timeout=15000)
            except Exception:
                user = None  # SSO zaten acik olabilir
            if user is not None:
                user.fill(self.username)
                pw = page.locator(g["sifre_selector"]).first
                if not pw.is_visible():  # iki adimli (once kullanici, sonra sifre) ekranlar
                    page.locator(g["gonder_selector"]).first.click()
                    pw.wait_for(state="visible", timeout=15000)
                pw.fill(self.password)
                page.locator(g["gonder_selector"]).first.click()
        elif g["mod"] == "manuel":
            self.log("Tarayicida giris yapin; hub acilinca devam edilecek (5 dk).")
        deadline = time.time() + (300 if g["mod"] == "manuel" else self.timeout / 1000)
        while time.time() < deadline:
            if ok in page.url and "login" not in page.url.lower():
                self.log("Giris basarili.")
                return
            page.wait_for_timeout(500)
        raise QlikError(f"Giris dogrulanamadi (su anki adres: {page.url}). Sifreyi ve config ayarlarini kontrol edin.")

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
                for loc in (page.get_by_role("menuitem", name=t, exact=True),
                            page.get_by_role("button", name=t, exact=True),
                            page.get_by_text(t, exact=True)):
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

    def export(self, app: dict, selections: dict) -> pd.DataFrame:
        if not app.get("app_id"):
            raise QlikError("config.yaml icinde app_id tanimli degil.")
        page, ex = self.page, self.cfg["export"]
        url = self.sheet_url(app, selections)
        self.log(f"Qlik sorgusu: {selections}")
        page.goto(url, wait_until="domcontentloaded")
        oid = app.get("object_id", "")
        target = page.locator(f"[tid='{oid}'], [data-qid='{oid}']").first if oid else \
            page.locator(".qv-object-table, .qv-object-pivot-table, .qv-object-sn-table, .qv-object").first
        target.wait_for(state="visible", timeout=self.timeout)
        page.wait_for_timeout(1500)  # secimlerin uygulanmasi icin
        if page.get_by_text(re.compile(r"(No data|Veri yok|Gösterilecek veri yok)", re.I)).count():
            return pd.DataFrame()
        target.click(button="right")
        self._click_text(ex["sag_tik_menu"], required=True)
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

    def calibrate(self, app: dict | None = None):
        """Playwright Inspector'i acar: secicileri ve export adimlarini elle bulmak icin."""
        if app and app.get("app_id"):
            self.page.goto(self.sheet_url(app, {}))
        self.page.pause()


class QlikSource(DataSource):
    """FA motoru icin canli Qlik veri kaynagi. Her sorgu sonucu oturum boyunca onbellekte tutulur."""

    def __init__(self, session: QlikSession, cfg: dict):
        self.s = session
        self.app = cfg["qlik"]["fa_hareket"]
        self.f = self.app.get("alanlar", {})
        self.fa_app = cfg["qlik"].get("fa_listesi", {})
        self.overrides = cfg.get("kolonlar", {})
        self._cache: dict[tuple, FrameSource] = {}

    def _frame(self, selections: dict) -> FrameSource:
        sel = {self.f.get(k, k): v for k, v in selections.items() if v}
        ck = tuple(sorted((k, key(v)) for k, v in sel.items()))
        if ck not in self._cache:
            raw = self.s.export(self.app, sel)
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
        self.s.close()


def _empty() -> pd.DataFrame:
    from ..columns import CANONICAL
    cols = CANONICAL + ["tx_raw", "source", "seq"] + [c + "_k" for c in
                                                    ("pn", "sn", "grb", "aircraft", "nha_pn", "nha_sn", "wo", "new_pn", "new_sn")]
    df = pd.DataFrame(columns=cols)
    df["date"] = pd.to_datetime(df["date"])
    return df
