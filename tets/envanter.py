"""TETS Envanter Takip otomasyonu.

Teknik Envanter Takibi > Null Veri Duzenleme Raporu sayfasinda durum alani
'false' secilir, veri Excel olarak indirilir. Iki secenek:
  - tamami   : false olan tum satirlar
  - null_not : sadece Notes kolonu bos olanlar
Qlik'te null degerler URL ile secilemedigi icin Notes filtresi indirilen veride uygulanir.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from .columns import header_key
from .normalize import clean


def filter_null_notes(df: pd.DataFrame, notes_col: str) -> pd.DataFrame:
    col = next((c for c in df.columns if header_key(c) == header_key(notes_col)), None)
    if col is None:
        raise ValueError(f"'{notes_col}' kolonu bulunamadi. Kolonlar: {list(df.columns)}")
    return df[df[col].map(clean) == ""]


def run_envanter(cfg: dict, session, only_null_notes: bool, log=None) -> Path:
    """session: acik bir QlikSession (giris yapilmis tarayici)."""
    e = cfg["qlik"]["envanter"]
    # Notes kolonu olmayan bir dosya inerse (yanlis nesne) sistem tabloyu yeniden bulup tekrar dener
    # Envanterde amac zaten tablonun tamamini Excel olarak almak: "Download as" ile indirilir.
    df = session.export(e, {e["durum_alani"]: e["durum_degeri"]}, [e["notes_kolonu"]] if only_null_notes else None,
                        methods=["menu"])
    if only_null_notes:
        df = filter_null_notes(df, e["notes_kolonu"])
    out = Path(cfg["cikti_klasoru"])
    out.mkdir(parents=True, exist_ok=True)
    path = out / f"Envanter_{'NullNotes' if only_null_notes else 'Tamami'}_{datetime.now():%Y%m%d_%H%M}.xlsx"
    df.to_excel(path, index=False)
    if log:
        log(f"» {len(df)} satir kaydedildi: {path.name}")
    return path
