"""Arayuz ve komut satiri icin ortak calistirma katmani (tekli ve toplu FA sorgusu)."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from ..sources.base import DataSource
from ..sources.excel_source import ExcelSource, load_fa_list
from .engine import FAEngine, FAResult
from .report import write_report


def open_source(cfg: dict, mode: str, excel_paths=None, fa_list_path=None,
                username: str = "", password: str = "", log=None) -> DataSource:
    """mode: 'excel' (indirilmis export dosyalari) veya 'qlik' (canli portal)."""
    if mode == "excel":
        if not excel_paths:
            raise ValueError("Excel modu icin hareket dosyasi secin.")
        src = ExcelSource(excel_paths, cfg.get("kolonlar"), fa_list_path)
        if log:
            log(f"{len(src.df)} hareket satiri yuklendi.")
            for s in src.skipped:
                log(f"Atlanan sayfa: {s.splitlines()[0]}")
        return src
    from ..sources.qlik_source import QlikSession, QlikSource
    session = QlikSession(cfg, username, password, log)
    session.start()
    return QlikSource(session, cfg)


def run_batch(source: DataSource, cfg: dict, fa_list_path, log=None, progress=None) -> tuple[list[FAResult], Path]:
    items = load_fa_list(fa_list_path)
    engine = FAEngine(source, cfg, log)
    results = []
    for n, r in enumerate(items.itertuples(), 1):
        if progress:
            progress(n, len(items), r.fa_no or r.pn)
        res = engine.trace(r.pn, r.sn, r.grb, r.fa_no, None if pd.isna(r.date) else r.date)
        results.append(res)
        if log:
            log(f"[{n}/{len(items)}] {r.fa_no or r.pn}: {res.status_text}")
    return results, write_report(results, cfg["cikti_klasoru"], f"FA_Toplu_{Path(fa_list_path).stem}.xlsx")
