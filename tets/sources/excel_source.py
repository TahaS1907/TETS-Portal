"""Excel/CSV dosyalarindan (Qlik'ten elle indirilmis export'lar) veri kaynagi."""
from __future__ import annotations

from pathlib import Path

import pandas as pd

from ..columns import header_key, map_columns, normalize_frame
from ..normalize import clean, key
from .base import FrameSource


def read_table(path: str | Path) -> dict[str, pd.DataFrame]:
    path = Path(path)
    if path.suffix.lower() in (".csv", ".txt"):
        for enc in ("utf-8-sig", "cp1254", "latin-1"):
            try:
                return {path.name: pd.read_csv(path, sep=None, engine="python", encoding=enc, dtype=str)}
            except UnicodeDecodeError:
                continue
    sheets = pd.read_excel(path, sheet_name=None, dtype=object)
    return {f"{path.name}:{name}": _fix_header(df) for name, df in sheets.items()}


def _fix_header(df: pd.DataFrame) -> pd.DataFrame:
    """Qlik export'larinda bazen ustte baslik/filtre satirlari olur; gercek basligi bul."""
    if map_columns(df).keys() >= {"pn", "sn"}:
        return df
    for i in range(min(15, len(df))):
        row = df.iloc[i].tolist()
        probe = pd.DataFrame(columns=[str(x) for x in row])
        if map_columns(probe).keys() >= {"pn", "sn", "tx_type"}:
            new = df.iloc[i + 1:].copy()
            new.columns = [str(x) for x in row]
            return new.reset_index(drop=True)
    return df


class ExcelSource(FrameSource):
    def __init__(self, paths, column_overrides: dict | None = None, fa_list_path: str | None = None):
        if isinstance(paths, (str, Path)):
            paths = [paths]
        frames, self.skipped = [], []
        for p in paths:
            for name, raw in read_table(p).items():
                if raw.empty:
                    continue
                try:
                    frames.append(normalize_frame(raw, column_overrides, source=name))
                except ValueError as e:
                    self.skipped.append(str(e))
        if not frames:
            raise ValueError("Hareket tablosu okunamadi.\n" + "\n".join(self.skipped))
        df = pd.concat(frames, ignore_index=True)
        df = df.drop_duplicates(subset=["date", "tx_raw", "pn_k", "sn_k", "aircraft_k", "position", "wo_k"])
        df["seq"] = range(len(df))
        fa_list = load_fa_list(fa_list_path) if fa_list_path else None
        super().__init__(df.sort_values(["date", "seq"]).reset_index(drop=True), fa_list)


def load_fa_list(path) -> pd.DataFrame:
    """FA listesi / toplu sorgu dosyasi: FA No, PN, SN, GRB, (Tarih) kolonlari."""
    raw = next(iter(read_table(path).values()))
    m = map_columns(raw)
    out = pd.DataFrame(index=raw.index)
    for c in ("fa_no", "pn", "sn", "grb"):
        out[c] = raw[m[c]].map(clean) if c in m else ""
    out["date"] = pd.to_datetime(raw[m["date"]], errors="coerce", dayfirst=True) if "date" in m else pd.NaT
    out["fa_no_k"] = out.fa_no.map(key)
    return out[(out.fa_no != "") | (out.pn != "") | (out.grb != "")].reset_index(drop=True)


__all__ = ["ExcelSource", "load_fa_list", "read_table", "header_key"]
