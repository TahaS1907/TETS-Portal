"""Metin, PN, pozisyon ve hareket tipi normalizasyonu + benzerlik fonksiyonlari.

Portal verisi tutarsiz yazilir ("LH", "LEFT", "SOL", "#1", "ENG 1", "NO.1" ...).
Bu modul hepsini karsilastirilabilir hale getirir.
"""
from __future__ import annotations

import math
import re
from difflib import SequenceMatcher

_TR = str.maketrans("İıŞşĞğÜüÖöÇç", "IISSGGUUOOCC")


def clean(value) -> str:
    """Bos/NaN degerleri '' yapar, Turkce karakterleri ASCII'ye cevirir, buyuk harf."""
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    s = str(value).translate(_TR).upper().strip()
    if s in {"NAN", "NONE", "NULL", "-", "NAT"}:
        return ""
    if s.endswith(".0") and s[:-2].isdigit():  # Excel'in sayiya cevirdigi SN'ler
        s = s[:-2]
    return re.sub(r"\s+", " ", s)


def key(value) -> str:
    """Sadece harf+rakam: 'P/N 123-45' -> 'PN12345'. Esitlik karsilastirmasi icin."""
    return re.sub(r"[^A-Z0-9]", "", clean(value))


# --------------------------------------------------------------------------- hareket tipleri
AC_INSTALL = "AC INSTALL"
AC_REMOVE = "AC REMOVE"
NHA_INSTALL = "NHA INSTALL"
NHA_REMOVE = "NHA REMOVE"
PNSN_CHG = "PN SN CHG"
STORE_IN = "STORE IN"
STORE_OUT = "STORE OUT"
SHIPPED = "SHIPPED"
SCRAP = "SCRAP"
OTHER = "OTHER"

INSTALLS = {AC_INSTALL, NHA_INSTALL}
REMOVES = {AC_REMOVE, NHA_REMOVE}

def _has(s: str, *patterns: str) -> bool:
    return any(re.search(p, s) for p in patterns)


def tx_type(raw) -> str:
    """Portaldaki serbest hareket metnini standart tipe cevirir."""
    s = clean(raw).replace("/", " ").replace("_", " ").replace("-", " ")
    s = re.sub(r"\bA C\b", "AC", s)
    if not s:
        return OTHER
    if _has(s, r"\b(PN|P N|PART ?NO|SN|S N|SERIAL)\b.*(CHG|CHANGE|DEGIS)", r"(CHG|CHANGE|DEGIS).*\b(PN|SN|SERIAL)\b"):
        return PNSN_CHG
    removal = _has(s, r"UNINST", r"REMOV", r"\bREM\b", r"SOKUM", r"SOKUL", r"\bSOK\b", r"\bOFF\b")
    install = not removal and _has(s, r"INST", r"TAKIL", r"TAKMA", r"FITTED", r"\bON\b")
    nha = _has(s, r"\bNHA\b", r"\bASSY\b", r"\bUST ?ASSY", r"PARENT", r"\bSUB\b", r"COMPONENT")
    if removal:
        return NHA_REMOVE if nha else AC_REMOVE
    if install:
        return NHA_INSTALL if nha else AC_INSTALL
    if _has(s, r"SCRAP", r"HURDA", r"\bBER\b"):
        return SCRAP
    if _has(s, r"SHIP", r"SEVK", r"REPAIR", r"TAMIR", r"VENDOR", r"OUTGOING"):
        return SHIPPED
    if _has(s, r"ISSUE", r"CIKIS", r"STORE OUT", r"STOCK OUT", r"PICK"):
        return STORE_OUT
    if _has(s, r"RECEI", r"GIRIS", r"RETURN", r"IADE", r"STORE IN", r"STOCK IN", r"\bGR\b", r"BOOK IN", r"DEPO"):
        return STORE_IN
    return OTHER


# --------------------------------------------------------------------------- PN
def pn_base(pn) -> str:
    """Dash/slash numarasiz ana PN: '123-01' -> '123', 'CFM56-7B26/3F' -> 'CFM56-7B26'."""
    s = clean(pn)
    for sep in ("/", "-", " "):
        if sep in s:
            head = s.rsplit(sep, 1)[0]
            if len(key(head)) >= 4:
                return key(head)
    return key(s)


def pn_prefix_ratio(a, b) -> float:
    """Ortak on-ek orani; motorlarda son kisim farkli olur (Senaryo 3)."""
    a, b = key(a), key(b)
    if not a or not b:
        return 0.0
    n = 0
    for x, y in zip(a, b):
        if x != y:
            break
        n += 1
    return n / max(len(a), len(b))


# --------------------------------------------------------------------------- aciklama
_DESC_STOP = {"ASSY", "ASSEMBLY", "THE", "OF", "AND", "WITH", "W", "PART", "UNIT", "KIT"}


def _tokens(s: str, stop: set[str]) -> list[str]:
    return [t for t in re.split(r"[^A-Z0-9]+", clean(s)) if t and t not in stop]


def text_similarity(a, b) -> float:
    """0..1 arasi aciklama benzerligi (token Jaccard + sirali token SequenceMatcher)."""
    ta, tb = _tokens(a, _DESC_STOP), _tokens(b, _DESC_STOP)
    if not ta or not tb:
        return 0.0
    sa, sb = set(ta), set(tb)
    jac = len(sa & sb) / len(sa | sb)
    seq = SequenceMatcher(None, " ".join(sorted(sa)), " ".join(sorted(sb))).ratio()
    return max(jac, seq * 0.9)


# --------------------------------------------------------------------------- pozisyon
_SIDE = {
    "L": "L", "LH": "L", "LHS": "L", "LEFT": "L", "SOL": "L", "PORT": "L",
    "R": "R", "RH": "R", "RHS": "R", "RIGHT": "R", "SAG": "R", "STBD": "R", "STARBOARD": "R",
    "C": "C", "CTR": "C", "CENTER": "C", "CENTRE": "C", "ORTA": "C", "MID": "C",
}
_POS_STOP = {"POS", "POSITION", "POZISYON", "NO", "NR", "NUM", "NUMBER", "THE", "OF", "AT", "ON"}
_POS_SYN = {
    "ENGINE": "ENG", "MOTOR": "ENG", "ENGINES": "ENG",
    "FORWARD": "FWD", "AFTER": "AFT", "ARKA": "AFT",
    "UPPER": "UPR", "UST": "UPR", "LOWER": "LWR", "ALT": "LWR",
    "INBOARD": "INBD", "IB": "INBD", "OUTBOARD": "OUTBD", "OB": "OUTBD",
    "MAIN": "MAIN", "NOSE": "NLG", "MLG": "MLG", "NLG": "NLG",
}


class Position:
    __slots__ = ("raw", "side", "numbers", "words")

    def __init__(self, raw):
        self.raw = clean(raw)
        s = self.raw.replace("#", " ").replace(".", " ").replace("/", " ")
        s = re.sub(r"(?<=[A-Z])(?=\d)|(?<=\d)(?=[A-Z])", " ", s)  # ENG1 -> ENG 1, 12L -> 12 L
        toks = [t for t in re.split(r"[^A-Z0-9]+", s) if t]
        sides, nums, words = set(), set(), set()
        for t in toks:
            if t in _SIDE:
                sides.add(_SIDE[t])
            elif t.isdigit():
                nums.add(str(int(t)))
            elif t not in _POS_STOP:
                words.add(_POS_SYN.get(t, t))
        self.side = sides.pop() if len(sides) == 1 else ("".join(sorted(sides)) if sides else "")
        self.numbers = frozenset(nums)
        self.words = frozenset(words)

    def __bool__(self):
        return bool(self.raw)


def position_compare(a, b) -> tuple[float, bool, bool]:
    """(benzerlik 0..1, taraf_celiskisi, numara_celiskisi).

    Taraf celiskisi: biri sol, digeri sag. Numara celiskisi: '#1' vs '#2'.
    Bilgi yoksa benzerlik 0.5 (notr) doner.
    """
    pa, pb = (a if isinstance(a, Position) else Position(a)), (b if isinstance(b, Position) else Position(b))
    if not pa or not pb:
        return 0.5, False, False
    parts: list[tuple[float, float]] = []
    side_conflict = num_conflict = False
    if pa.side and pb.side:
        same = pa.side == pb.side
        side_conflict = not same and set(pa.side).isdisjoint(pb.side)
        parts.append((0.4, 1.0 if same else (0.5 if not side_conflict else 0.0)))
    if pa.numbers and pb.numbers:
        if pa.numbers == pb.numbers:
            v = 1.0
        elif pa.numbers & pb.numbers:
            v = 0.5
        else:
            v, num_conflict = 0.0, True
        parts.append((0.35, v))
    if pa.words and pb.words:
        jac = len(pa.words & pb.words) / len(pa.words | pb.words)
        parts.append((0.25, jac))
    if not parts:
        return SequenceMatcher(None, pa.raw, pb.raw).ratio(), False, False
    tot = sum(w for w, _ in parts)
    return sum(w * v for w, v in parts) / tot, side_conflict, num_conflict
