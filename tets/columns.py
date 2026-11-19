"""Excel/Qlik export kolon basliklarini standart alan adlarina esler.

Basliklar sirketten sirkete, rapordan rapora degisir. Burada bilinen
Turkce/Ingilizce varyasyonlar var; eslesmeyen bir kolon icin config.yaml
icindeki `kolonlar:` bolumune elle ekleme yapilabilir.
"""
from __future__ import annotations

import re

import pandas as pd

from .normalize import clean

CANONICAL = [
    "date", "tx_type", "pn", "sn", "grb", "description", "pn_group", "aircraft",
    "position", "nla_position", "nha_pn", "nha_sn", "wo", "location", "new_pn", "new_sn", "fa_no",
]

ALIASES: dict[str, list[str]] = {
    "date": ["CREATEDDATE", "DATE", "TARIH", "TRANSACTIONDATE", "TXDATE", "ISLEMTARIHI", "HAREKETTARIHI",
             "MOVEMENTDATE", "ACTIONDATE", "INSTALLDATE", "REMOVALDATE", "BOOKINGDATE"],
    "tx_type": ["TRANSACTION", "TRANSACTIONTYPE", "TXTYPE", "TYPE", "HAREKET", "HAREKETTIPI",
                "HAREKETTURU", "ISLEM", "ISLEMTIPI", "MOVEMENT", "MOVEMENTTYPE", "ACTION", "EVENT"],
    "pn": ["PN", "PARTNO", "PARTNUMBER", "PARCANO", "PARCANUMARASI", "MATERIAL", "MALZEME"],
    "sn": ["SN", "SERIALNO", "SERIALNUMBER", "SERINO", "SERINUMARASI", "SERIAL"],
    "grb": ["GOODSRCVDBATCH", "GOODSRECEIVEDBATCH", "GRB", "GRBNO", "GRBNUMBER", "BATCH", "BATCHNO", "LABEL", "LABELNO", "ETIKET"],
    "description": ["PNDESCRIPTION", "DESCRIPTION", "DESC", "ACIKLAMA", "TANIM", "PARTDESCRIPTION", "PARCAADI",
                    "PARCATANIMI", "NOMENCLATURE"],
    "pn_group": ["PNGROUP", "PNGRUP", "PNGRUBU", "PARTGROUP", "GRUP", "GROUP", "INTERCHANGEGROUP",
                 "INTERCHANGEABILITY", "IPCGROUP", "ALTERNATEGROUP"],
    "aircraft": ["AC", "AIRCRAFT", "ACREG", "AIRCRAFTREG", "REGISTRATION", "TAIL", "TAILNO",
                 "TESCIL", "UCAK", "UCAKTESCIL"],
    "position": ["ACPOSITION", "POSITION", "POS", "POZISYON", "INSTALLPOSITION", "POSITIONCODE",
                 "LOCATIONONAC"],
    # ust assy icindeki pozisyon (ROT FA PN History: NLA_POSITION)
    "nla_position": ["NLAPOSITION", "NHAPOSITION", "ASSYPOSITION"],
    "nha_pn": ["NHAPN", "NHAPARTNO", "NHAPARTNUMBER", "PARENTPN", "USTASSYPN", "USTASSYPARCANO",
               "ASSYPN", "HIGHERASSYPN"],
    "nha_sn": ["NHASN", "NHASERIALNO", "NHASERIALNUMBER", "PARENTSN", "USTASSYSN", "ASSYSN",
               "HIGHERASSYSN"],
    "wo": ["WO", "WORKORDER", "WORKORDERNO", "ISEMRI", "ISEMRINO", "WP", "WORKPACKAGE", "TASK",
           "TASKNO", "EVENTNO"],
    "location": ["LOCATION", "LOKASYON", "STORE", "DEPO", "STATION", "ISTASYON", "WAREHOUSE"],
    "new_pn": ["NEWPN", "YENIPN", "TOPN", "PNTO", "NEWPARTNO"],
    "new_sn": ["NEWSN", "YENISN", "TOSN", "SNTO", "NEWSERIALNO"],
    "fa_no": ["FA", "FANO", "FANUMBER", "FANUMARASI"],
}


def header_key(h) -> str:
    return re.sub(r"[^A-Z0-9]", "", clean(h))


def map_columns(df: pd.DataFrame, overrides: dict | None = None) -> dict[str, str]:
    """{standart_ad: orijinal_kolon} sozlugu dondurur."""
    overrides = {k: v for k, v in (overrides or {}).items() if v}
    by_key = {header_key(c): c for c in df.columns}
    out: dict[str, str] = {}
    used: set[str] = set()
    for canon, actual in overrides.items():
        if actual in df.columns:
            out[canon] = actual
            used.add(actual)
        elif header_key(actual) in by_key:
            out[canon] = by_key[header_key(actual)]
            used.add(out[canon])
    for canon in CANONICAL:
        if canon in out:
            continue
        for alias in ALIASES.get(canon, []):
            col = by_key.get(alias)
            if col is not None and col not in used:
                out[canon] = col
                used.add(col)
                break
    return out


def normalize_frame(df: pd.DataFrame, overrides: dict | None = None, source: str = "") -> pd.DataFrame:
    """Ham tabloyu standart kolonlu, tipleri cozulmus bir tabloya cevirir."""
    from .normalize import key, tx_type

    mapping = map_columns(df, overrides)
    missing = [c for c in ("date", "tx_type", "pn", "sn") if c not in mapping]
    if missing:
        raise ValueError(
            f"{source or 'Tablo'}: zorunlu kolon(lar) bulunamadi: {missing}. "
            f"Mevcut kolonlar: {list(df.columns)}. config.yaml > kolonlar bolumunden esleyin."
        )
    out = pd.DataFrame(index=df.index)
    for canon in CANONICAL:
        out[canon] = df[mapping[canon]] if canon in mapping else ""
    raw_date = out["date"]
    out["date"] = pd.to_datetime(raw_date, errors="coerce", dayfirst=True)
    out["tx_raw"] = out["tx_type"].map(clean)
    out["tx_type"] = out["tx_type"].map(tx_type)
    for c in CANONICAL:
        if c not in ("date", "tx_type"):
            out[c] = out[c].map(clean)
    for c in ("pn", "sn", "grb", "aircraft", "nha_pn", "nha_sn", "wo", "new_pn", "new_sn"):
        out[c + "_k"] = out[c].map(key)
    # NHA hareketlerinde pozisyon = ust assy icindeki pozisyon (NLA_POSITION)
    nha_rows = out["tx_type"].isin(["NHA INSTALL", "NHA REMOVE"]) & (out["nla_position"] != "")
    out.loc[nha_rows, "position"] = out.loc[nha_rows, "nla_position"]
    no_pos = (out["position"] == "") & (out["nla_position"] != "")
    out.loc[no_pos, "position"] = out.loc[no_pos, "nla_position"]
    out["source"] = source
    out = out[out["date"].notna() & (out["pn_k"] != "")]
    out["seq"] = range(len(out))  # ayni gun icindeki orijinal sira
    return out.sort_values(["date", "seq"]).reset_index(drop=True)
