"""Komut satiri.

Ornekler:
  python -m tets.cli fa --excel ornek_veri/ornek_hareketler.xlsx --pn 3214-01 --sn A100
  python -m tets.cli toplu --excel ornek_veri/ornek_hareketler.xlsx --liste ornek_veri/ornek_fa_listesi.xlsx
  python -m tets.cli fa --qlik --kullanici TK12345 --pn 3214-01 --sn A100     (sifre sorulur)
  python -m tets.cli envanter --kullanici TK12345 --null-notes
  python -m tets.cli kalibrasyon --kullanici TK12345
"""
from __future__ import annotations

import argparse
import getpass
import sys

from .config import load_config
from .fa.engine import FAEngine
from .fa.report import write_report
from .fa.runner import open_source, run_batch


def main(argv=None):
    p = argparse.ArgumentParser(prog="tets")
    p.add_argument("--config")
    sub = p.add_subparsers(dest="cmd", required=True)

    def src_args(sp):
        sp.add_argument("--excel", nargs="*", help="Hareket export dosyalari (Excel/CSV)")
        sp.add_argument("--qlik", action="store_true", help="Canli Qlik Sense'den cek")
        sp.add_argument("--kullanici", default="")

    fa = sub.add_parser("fa", help="Tek FA sorgusu")
    src_args(fa)
    for a in ("--pn", "--sn", "--grb", "--fa", "--fa-tarihi"):
        fa.add_argument(a, default="")
    fa.add_argument("--rapor", action="store_true", help="Excel rapor da yaz")
    tp = sub.add_parser("toplu", help="FA listesi Excel'i ile toplu sorgu")
    src_args(tp)
    tp.add_argument("--liste", required=True)
    env = sub.add_parser("envanter")
    env.add_argument("--kullanici", default="")
    env.add_argument("--null-notes", action="store_true")
    cal = sub.add_parser("kalibrasyon")
    cal.add_argument("--kullanici", default="")
    a = p.parse_args(argv)
    cfg = load_config(a.config)
    log = lambda m: print("  ·", m, file=sys.stderr)
    pw = ""
    if getattr(a, "qlik", False) or a.cmd in ("envanter", "kalibrasyon"):
        pw = getpass.getpass("Qlik sifresi: ")

    if a.cmd == "envanter":
        from .envanter import run_envanter
        print(run_envanter(cfg, a.kullanici, pw, a.null_notes, log))
        return
    if a.cmd == "kalibrasyon":
        from .sources.qlik_source import QlikSession
        with QlikSession(cfg, a.kullanici, pw, log) as s:
            s.calibrate(cfg["qlik"]["fa_hareket"])
        return
    mode = "qlik" if a.qlik else "excel"
    src = open_source(cfg, mode, a.excel, getattr(a, "liste", None), a.kullanici, pw, log)
    try:
        if a.cmd == "toplu":
            results, path = run_batch(src, cfg, a.liste, log)
            print(f"Rapor: {path}")
            return
        res = FAEngine(src, cfg, log).trace(a.pn, a.sn, a.grb, a.fa, a.fa_tarihi or None)
        print(res.text())
        if a.rapor:
            print(f"\nRapor: {write_report([res], cfg['cikti_klasoru'])}")
    finally:
        src.close()


if __name__ == "__main__":
    main()
