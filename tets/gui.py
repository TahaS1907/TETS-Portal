"""TETS Portal masaustu arayuzu (Tkinter - Python ile birlikte gelir, ek kurulum gerekmez).

Sekmeler: FA Takip | Envanter Takip | Ayarlar
Qlik islemleri tek bir arka plan is parcaciginda calisir; boylece bir kez giris
yapildiktan sonra ayni oturumla art arda sorgu yapilabilir ve arayuz donmaz.
"""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
from tkinter.scrolledtext import ScrolledText

from . import __version__
from .config import CONFIG_PATH, ROOT, ensure_user_config, load_config

RED = "#C8102E"


def open_path(p):
    p = str(p)
    if sys.platform.startswith("win"):
        os.startfile(p)  # noqa
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", p])


METHOD_LABELS = {
    "otomatik": "Otomatik - once ekrandan oku (indirmeden), olmazsa Download as ile indir",
    "ekran": "Sadece ekrandan oku (Download as kullanma)",
    "menu": "Sadece Download as ile indir",
}


class Worker(threading.Thread):
    """Sirali is kuyrugu. Playwright nesneleri yalnizca bu thread'de yasar."""

    def __init__(self, ui_queue):
        super().__init__(daemon=True)
        self.jobs: queue.Queue = queue.Queue()
        self.ui = ui_queue
        self.source = None
        self.source_key = None
        self.session = None  # acik Qlik tarayici oturumu (tum sekmeler ortak kullanir)

    def submit(self, fn, *args):
        self.jobs.put((fn, args))

    def run(self):
        while True:
            fn, args = self.jobs.get()
            try:
                fn(self, *args)
            except Exception as e:
                self.ui.put(("log", f"HATA: {e}"))
                self.ui.put(("log", traceback.format_exc()))
                self.ui.put(("error", str(e)))
            finally:
                self.ui.put(("idle", None))

    def reset_source(self):
        if self.source is not None:
            try:
                self.source.close()
            except Exception:
                pass
        self.source, self.source_key = None, None

    def reset_session(self):
        self.reset_source()
        if self.session is not None:
            self.session.close()
        self.session = None


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.cfg = load_config()
        self.title(f"TETS Portal {__version__}")
        self.geometry("1100x760")
        self.minsize(900, 600)
        self.ui_q: queue.Queue = queue.Queue()
        self.worker = Worker(self.ui_q)
        self.worker.start()
        self.last_report = None
        self._style()
        self._header()
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(0, 6))
        self._fa_tab(nb)
        self._env_tab(nb)
        self._settings_tab(nb)
        self.status = tk.StringVar(value="Hazir")
        bar = ttk.Frame(self)
        bar.pack(fill="x", padx=10, pady=(0, 8))
        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=180)
        self.progress.pack(side="right")
        ttk.Label(bar, textvariable=self.status).pack(side="left")
        self.after(100, self._poll)
        self.protocol("WM_DELETE_WINDOW", self._quit)

    # ------------------------------------------------------------------ gorunum
    def _style(self):
        st = ttk.Style(self)
        try:
            st.theme_use("vista" if sys.platform.startswith("win") else "clam")
        except tk.TclError:
            pass
        st.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        st.configure("H.TLabel", font=("Segoe UI", 16, "bold"), foreground=RED)

    def _header(self):
        f = ttk.Frame(self)
        f.pack(fill="x", padx=10, pady=8)
        logo = ROOT / "assets" / "logo.png"
        if logo.exists():
            try:
                self._logo = tk.PhotoImage(file=str(logo))
                h = self._logo.height()
                if h > 60:
                    self._logo = self._logo.subsample(max(1, h // 60))
                ttk.Label(f, image=self._logo).pack(side="left", padx=(0, 10))
            except tk.TclError:
                pass
        ttk.Label(f, text="TETS Portal", style="H.TLabel").pack(side="left")
        cred = ttk.LabelFrame(f, text="Qlik Sense")
        cred.pack(side="right")
        self.user = tk.StringVar()
        self.pw = tk.StringVar()
        col = 0
        if self.cfg["qlik"]["giris"]["mod"] != "manuel":
            ttk.Label(cred, text="Kullanici").grid(row=0, column=0, padx=4, pady=3)
            ttk.Entry(cred, textvariable=self.user, width=16).grid(row=0, column=1, padx=4)
            ttk.Label(cred, text="Sifre").grid(row=0, column=2, padx=4)
            ttk.Entry(cred, textvariable=self.pw, show="•", width=16).grid(row=0, column=3, padx=4)
            col = 4
        self.qlik_state = tk.StringVar(value="kapali")
        ttk.Button(cred, text="Qlik'i Ac", style="Accent.TButton", command=self._open_qlik).grid(
            row=0, column=col, padx=4, pady=3)
        ttk.Label(cred, textvariable=self.qlik_state, foreground="#555").grid(row=0, column=col + 1, padx=4)

    def _fa_tab(self, nb):
        tab = ttk.Frame(nb, padding=8)
        nb.add(tab, text="  FA Takip  ")
        src = ttk.LabelFrame(tab, text="Veri kaynagi", padding=6)
        src.pack(fill="x")
        self.mode = tk.StringVar(value="qlik" if self.cfg["qlik"]["base_url"] else "excel")
        ttk.Radiobutton(src, text="Canli Qlik Sense", variable=self.mode, value="qlik").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(src, text="Excel export dosyalari", variable=self.mode, value="excel").grid(row=0, column=1, sticky="w", padx=10)
        self.excel_files: list[str] = []
        self.excel_lbl = tk.StringVar(value="(dosya secilmedi)")
        ttk.Button(src, text="Hareket dosyalari...", command=self._pick_excel).grid(row=0, column=2, padx=6)
        ttk.Label(src, textvariable=self.excel_lbl, foreground="#555").grid(row=0, column=3, sticky="w")

        q = ttk.LabelFrame(tab, text="FA sorgusu", padding=6)
        q.pack(fill="x", pady=6)
        self.v = {k: tk.StringVar() for k in ("fa", "pn", "sn", "grb", "tarih")}
        labels = [("FA No", "fa"), ("PN", "pn"), ("SN", "sn"), ("GRB", "grb"), ("FA tarihi (ops.)", "tarih")]
        for c, (lbl, k) in enumerate(labels):
            ttk.Label(q, text=lbl).grid(row=0, column=c, sticky="w", padx=4)
            e = ttk.Entry(q, textvariable=self.v[k], width=18)
            e.grid(row=1, column=c, padx=4, pady=2)
            e.bind("<Return>", lambda _e: self._run_single())
        btns = ttk.Frame(q)
        btns.grid(row=1, column=len(labels), padx=8)
        self.btn_find = ttk.Button(btns, text="FA'i Bul", style="Accent.TButton", command=self._run_single)
        self.btn_find.pack(side="left", padx=2)
        self.btn_batch = ttk.Button(btns, text="Toplu (Excel listesi)...", command=self._run_batch)
        self.btn_batch.pack(side="left", padx=2)
        ttk.Button(btns, text="Son raporu ac", command=self._open_report).pack(side="left", padx=2)

        pan = ttk.PanedWindow(tab, orient="vertical")
        pan.pack(fill="both", expand=True)
        self.out = ScrolledText(pan, font=("Consolas", 10), wrap="none", height=20)
        self.logbox = ScrolledText(pan, font=("Consolas", 9), height=6, foreground="#555")
        pan.add(self.out, weight=4)
        pan.add(self.logbox, weight=1)
        for tag, color in (("FA_ATA", "#0a7a2f"), ("FA_SIL", RED), ("PNSN_CHG", "#b26b00"),
                           ("DEPODA", "#1f5fa8"), ("FA_DOGRU", "#0a7a2f"), ("WARN", RED)):
            self.out.tag_config(tag, foreground=color, font=("Consolas", 10, "bold"))

    def _env_tab(self, nb):
        tab = ttk.Frame(nb, padding=12)
        nb.add(tab, text="  Envanter Takip  ")
        ttk.Label(tab, text="Teknik Envanter Takibi > Null Veri Duzenleme Raporu\n"
                            "Durum = false secilir ve veri Excel olarak indirilir.").pack(anchor="w")
        self.env_null = tk.BooleanVar(value=True)
        ttk.Radiobutton(tab, text="Sadece Notes bos olanlar", variable=self.env_null, value=True).pack(anchor="w", pady=(10, 0))
        ttk.Radiobutton(tab, text="Tamami", variable=self.env_null, value=False).pack(anchor="w")
        ttk.Button(tab, text="Raporu indir", style="Accent.TButton", command=self._run_env).pack(anchor="w", pady=10)

    def _settings_tab(self, nb):
        tab = ttk.Frame(nb, padding=12)
        nb.add(tab, text="  Ayarlar  ")
        ttk.Label(tab, text=f"Ayar dosyasi: {CONFIG_PATH}").pack(anchor="w")
        ttk.Label(tab, text=f"Rapor klasoru: {self.cfg['cikti_klasoru']}").pack(anchor="w", pady=(0, 10))
        ttk.Label(tab, text="Qlik'ten veri okuma yontemi:").pack(anchor="w")
        cur = self.cfg["qlik"].get("export_yontemi", "otomatik")
        self.method_cb = ttk.Combobox(tab, values=list(METHOD_LABELS.values()), state="readonly", width=70)
        self.method_cb.set(METHOD_LABELS.get(cur, METHOD_LABELS["otomatik"]))
        self.method_cb.pack(anchor="w", pady=(2, 12))
        self.method_cb.bind("<<ComboboxSelected>>", self._set_method)
        for text, cmd in (("config.yaml'i ac", lambda: open_path(ensure_user_config())),
                          ("Ayarlari yeniden yukle", self._reload),
                          ("Rapor klasorunu ac", self._open_out_dir),
                          ("Kayitli Qlik sayfalarini unut (yeniden sec)", self._forget),
                          ("Kalibrasyon (Qlik'i Inspector ile ac)", self._calibrate),
                          ("Ornek veri ile dene", self._demo)):
            ttk.Button(tab, text=text, command=cmd, width=40).pack(anchor="w", pady=3)
        ttk.Label(tab, foreground="#555", justify="left", text=(
            "Kalibrasyon: Qlik giris yapilmis tarayiciyi Playwright Inspector ile acar.\n"
            "Inspector'daki 'Pick locator' ile tablo nesnesini / menu adimlarini secip\n"
            "config.yaml > qlik bolumune yazabilirsiniz.")).pack(anchor="w", pady=10)

    # ------------------------------------------------------------------ olaylar
    def _pick_excel(self):
        files = filedialog.askopenfilenames(title="Qlik hareket export dosyalari",
                                            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("Tum", "*.*")])
        if files:
            self.excel_files = list(files)
            self.excel_lbl.set(", ".join(Path(f).name for f in files))
            self.mode.set("excel")
            self.worker.submit(lambda w: w.reset_source())

    def _snapshot(self) -> dict:
        """Tk degiskenleri sadece ana thread'de okunur; is parcacigina kopya gider."""
        return {"mode": self.mode.get(), "files": list(self.excel_files),
                "user": self.user.get().strip(), "pw": self.pw.get()}

    def _ensure_session(self, w: Worker, snap: dict):
        """Tarayiciyi acar; kullanici sifresini tarayiciya girene kadar bekler. Oturum acik kalir."""
        from .sources.qlik_source import QlikSession
        if w.session is None or not w.session.alive:
            w.reset_session()
            self.ui_q.put(("qlik", "giris bekleniyor"))
            s = QlikSession(self.cfg, snap["user"], snap["pw"], self._log)
            try:
                s.start()
            except Exception:
                s.close()
                self.ui_q.put(("qlik", "kapali"))
                raise
            w.session = s
            self.ui_q.put(("qlik", "bagli ✓"))
        return w.session

    def _ensure_source(self, w: Worker, snap: dict):
        from .fa.runner import open_source
        k = (snap["mode"], tuple(snap["files"]))
        if snap["mode"] == "qlik":
            from .sources.qlik_source import QlikSource
            sess = self._ensure_session(w, snap)
            if not isinstance(w.source, QlikSource) or w.source.s is not sess:
                w.reset_source()
                w.source, w.source_key = QlikSource(sess, self.cfg, owns_session=False), k
            return w.source
        if w.source is None or w.source_key != k:
            w.reset_source()
            w.source = open_source(self.cfg, "excel", snap["files"], None, log=self._log)
            w.source_key = k
        return w.source

    def _open_qlik(self):
        snap = self._snapshot()
        self._busy("Tarayicida Qlik'e giris yapin...")

        def job(w):
            self._ensure_session(w, snap)
            self._log("» Giris tamam. Simdi FA sorgusu yapabilir veya envanter indirebilirsiniz.")
        self.worker.submit(job)

    def _busy(self, text):
        self.status.set(text)
        self.progress.start(12)
        self.btn_find.state(["disabled"])
        self.btn_batch.state(["disabled"])

    def _run_single(self):
        vals = {k: v.get().strip() for k, v in self.v.items()}
        if not ((vals["pn"] and vals["sn"]) or vals["grb"] or vals["fa"]):
            messagebox.showwarning("Eksik bilgi", "PN + SN, GRB veya FA No girin.")
            return
        snap = self._snapshot()
        self._busy("FA dongusu takip ediliyor...")

        def job(w):
            import pandas as pd
            from .fa.engine import FAEngine
            from .fa.report import write_report
            src = self._ensure_source(w, snap)
            fa_date = pd.to_datetime(vals["tarih"], dayfirst=True, errors="coerce") if vals["tarih"] else None
            res = FAEngine(src, self.cfg, self._log).trace(vals["pn"], vals["sn"], vals["grb"], vals["fa"], fa_date)
            # Tekli sorguda Excel rapor yazilmaz; sonuc sadece ekrandaki aciklamada gorunur.
            # (Istenirse config.yaml > fa > tekli_excel_rapor: true ile geri acilir.)
            path = write_report([res], self.cfg["cikti_klasoru"]) if self.cfg["fa"].get("tekli_excel_rapor") else None
            self.ui_q.put(("result", (res, path)))
        self.worker.submit(job)

    def _run_batch(self):
        path = filedialog.askopenfilename(title="FA listesi (FA No, PN, SN, GRB kolonlari)",
                                          filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if not path:
            return
        snap = self._snapshot()
        self._busy("Toplu sorgu calisiyor...")

        def job(w):
            from .fa.runner import run_batch
            src = self._ensure_source(w, snap)
            prog = lambda n, t, name: self.ui_q.put(("status", f"Toplu sorgu {n}/{t}: {name}"))
            results, rpt = run_batch(src, self.cfg, path, self._log, prog)
            self.ui_q.put(("batch", (results, rpt)))
        self.worker.submit(job)

    def _run_env(self):
        self._busy("Envanter raporu indiriliyor...")
        only_null, snap = self.env_null.get(), self._snapshot()

        def job(w):
            from .envanter import run_envanter
            p = run_envanter(self.cfg, self._ensure_session(w, snap), only_null, self._log)
            self.ui_q.put(("file", p))
        self.worker.submit(job)

    def _calibrate(self):
        snap = self._snapshot()
        self._busy("Kalibrasyon: Inspector penceresini kapatinca biter")

        def job(w):
            self._ensure_session(w, snap).calibrate(self.cfg["qlik"]["fa_hareket"])
        self.worker.submit(job)

    def _demo(self):
        sys.path.insert(0, str(ROOT))
        from tests.sample_data import write
        h, f = write(ROOT / "ornek_veri")
        self.excel_files = [str(h)]
        self.excel_lbl.set(h.name + "  (ornek)")
        self.mode.set("excel")
        self.v["pn"].set("CFM56-7B26/3")
        self.v["sn"].set("890111")
        self.v["fa"].set("FA-1004")
        self.worker.submit(lambda w: w.reset_source())
        messagebox.showinfo("Ornek veri", f"Ornek veri yuklendi. 'FA'i Bul' ile deneyin.\n"
                                          f"Toplu deneme icin liste: {f}")

    def _set_method(self, _e=None):
        from .config import save_preference
        label = self.method_cb.get()
        value = next(k for k, v in METHOD_LABELS.items() if v == label)
        self.cfg["qlik"]["export_yontemi"] = value
        save_preference("export_yontemi", value)

        def job(w):
            if w.session is not None:
                w.session.cfg["export_yontemi"] = value
                w.session.good_method = ""
        self.worker.submit(job)
        self.status.set(f"Okuma yontemi: {label}")

    def _forget(self):
        from .config import forget_learned
        forget_learned()
        self._reload()
        messagebox.showinfo("Qlik sayfalari", "Bir sonraki islemde portal sayfasini tarayicida tekrar secmeniz istenecek.")

    def _reload(self):
        self.cfg = load_config()
        self.worker.submit(lambda w: w.reset_source())
        self.status.set("Ayarlar yeniden yuklendi")

    def _open_report(self):
        if self.last_report:
            open_path(self.last_report)

    def _open_out_dir(self):
        Path(self.cfg["cikti_klasoru"]).mkdir(parents=True, exist_ok=True)
        open_path(self.cfg["cikti_klasoru"])

    def _log(self, msg):
        self.ui_q.put(("log", msg))
        if msg.startswith("»"):  # kullaniciya yonelik onemli adimlar durum cubugunda da gorunur
            self.ui_q.put(("status", msg.lstrip("» ")))

    # ------------------------------------------------------------------ UI kuyrugu
    def _show_result(self, res):
        self.out.delete("1.0", "end")
        for line in res.text().split("\n"):
            tag = ()
            if line.startswith("SONUÇ"):
                tag = (res.status,)
            elif line.startswith(" !"):
                tag = ("WARN",)
            self.out.insert("end", line + "\n", tag)

    def _poll(self):
        try:
            while True:
                kind, data = self.ui_q.get_nowait()
                if kind == "log":
                    self.logbox.insert("end", data + "\n")
                    self.logbox.see("end")
                elif kind == "qlik":
                    self.qlik_state.set(data)
                elif kind == "status":
                    self.status.set(data)
                elif kind == "result":
                    res, path = data
                    self._show_result(res)
                    if path:
                        self.last_report = path
                        self.status.set(f"{res.status_text}  |  rapor: {Path(path).name}")
                    else:
                        self.status.set(res.status_text)
                elif kind == "batch":
                    results, self.last_report = data
                    self.out.delete("1.0", "end")
                    for r in results:
                        self.out.insert("end", r.text() + "\n" + "-" * 100 + "\n")
                    self.status.set(f"{len(results)} FA islendi | rapor: {Path(self.last_report).name}")
                    open_path(self.last_report)
                elif kind == "file":
                    self.last_report = data
                    self.status.set(f"Kaydedildi: {data}")
                    open_path(data)
                elif kind == "error":
                    messagebox.showerror("Hata", data)
                elif kind == "idle":
                    self.progress.stop()
                    self.btn_find.state(["!disabled"])
                    self.btn_batch.state(["!disabled"])
                    if self.status.get().endswith("..."):
                        self.status.set("Hazir")
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _quit(self):
        self.worker.submit(lambda w: w.reset_session())
        self.after(800, self.destroy)


def main():
    App().mainloop()
