"""Veri kaynagi arayuzu. FA motoru sadece bu arayuzu bilir; veri Excel'den
mi yoksa canli Qlik Sense'den mi geliyor umursamaz."""
from __future__ import annotations

from abc import ABC, abstractmethod
from datetime import datetime

import pandas as pd

from ..normalize import key


class DataSource(ABC):
    @abstractmethod
    def part_history(self, pn: str = "", sn: str = "", grb: str = "") -> pd.DataFrame:
        """Parcanin (PN+SN ve/veya GRB) tum hareketleri, tarih sirali."""

    @abstractmethod
    def aircraft_events(self, aircraft: str, start: datetime, end: datetime) -> pd.DataFrame:
        """Ucaktaki tum takma/sokme hareketleri (tarih araliginda)."""

    @abstractmethod
    def nha_events(self, nha_pn: str, nha_sn: str, start: datetime, end: datetime) -> pd.DataFrame:
        """Ust assy (NHA) uzerindeki takma/sokme hareketleri."""

    def fa_lookup(self, fa_no: str) -> dict | None:
        """FA numarasindan PN/SN/GRB bulur (FA listesi varsa)."""
        return None

    def close(self) -> None:
        pass


class FrameSource(DataSource):
    """Normalize edilmis tek bir DataFrame uzerinde calisan kaynak."""

    def __init__(self, frame: pd.DataFrame, fa_list: pd.DataFrame | None = None):
        self.df = frame
        self.fa_list = fa_list

    def part_history(self, pn="", sn="", grb=""):
        df = self.df
        pk, sk, gk = key(pn), key(sn), key(grb)
        mask = pd.Series(False, index=df.index)
        if pk and sk:
            mask |= (df.pn_k == pk) & (df.sn_k == sk)
            mask |= (df.new_pn_k == pk) & (df.new_sn_k == sk)  # PN SN CHG ile gelen kimlik
        elif sk and not pk:
            mask |= df.sn_k == sk
        if gk:
            mask |= df.grb_k == gk
        return df[mask].sort_values(["date", "seq"])

    def aircraft_events(self, aircraft, start, end):
        df = self.df
        m = (df.aircraft_k == key(aircraft)) & (df.date >= start) & (df.date <= end)
        return df[m].sort_values(["date", "seq"])

    def nha_events(self, nha_pn, nha_sn, start, end):
        df = self.df
        m = (df.date >= start) & (df.date <= end) & (df.nha_pn_k == key(nha_pn))
        if key(nha_sn):
            m &= df.nha_sn_k == key(nha_sn)
        return df[m].sort_values(["date", "seq"])

    def fa_lookup(self, fa_no):
        if self.fa_list is None or not key(fa_no):
            return None
        hit = self.fa_list[self.fa_list.fa_no_k == key(fa_no)]
        if hit.empty:
            return None
        r = hit.iloc[0]
        return {"pn": r.get("pn", ""), "sn": r.get("sn", ""), "grb": r.get("grb", ""),
                "fa_date": r.get("date")}
