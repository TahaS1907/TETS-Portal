"""FA dongu takip motoru.

Mantik (Fa portal Senaryolari.xlsx):
  S1  FA'deki parca takildiysa, yerine sokulen parca bulunur. Sokulen parca
      sonradan tekrar takildiysa onun yerine sokulen bulunur... Takilmayan ilk
      sokulen parcada dongu biter -> FA o parcaya atanir.
  S2  Takilan parcanin yerine sokulen yoksa "FA silinmeli". Eslesme sadece ayni
      PN ile degil PN grubu ve aciklama benzerligi ile de aranir (mod/EO).
  S3  Motorlarda uzun donguler; PN sonu farkli olabilir, aciklama ve pozisyon
      (ozellikle sag/sol, #1/#2) kesin uymali.
  S4  Ayni PN'den birden fazla takma/sokme: pozisyon benzerligi ile birebir
      (global) eslestirme yapilir.
  S5  NHA INSTALL/REMOVE: parca ust assy'e takildiysa ayni ust assy'den
      sokulen ayni PN/grup parcaya bakilir. Sokulen depoda ise FA gerekmez.
  S6  PN SN CHG: parca GRB ile de takip edilir; degisimden sonra hareket yoksa
      FA'e yeni PN/SN gelmeli.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd

from ..config import load_config
from ..normalize import (AC_INSTALL, AC_REMOVE, INSTALLS, NHA_INSTALL, NHA_REMOVE, PNSN_CHG, REMOVES,
                         SCRAP, SHIPPED, STORE_IN, STORE_OUT, key, pn_base, pn_prefix_ratio,
                         position_compare, text_similarity)
from ..sources.base import DataSource

STATUS_TEXT = {
    "FA_ATA": "FA bu parcaya atanmali",
    "FA_SIL": "FA silinmeli (yerine sokulen parca yok)",
    "PNSN_CHG": "PN/SN degisimi - FA'e yeni PN/SN gelmeli",
    "DEPODA": "Dongunun sonundaki parca depoda - FA kapatilabilir",
    "FA_DOGRU": "FA dogru parcada (parca takili degil, depoda degil)",
    "BULUNAMADI": "Parcaya ait hareket bulunamadi",
    "LIMIT": "Dongu limiti asildi - manuel kontrol gerekli",
    "HATA": "Hata",
}

SCENARIO_TEXT = {
    1: "S1 Dongu takibi",
    2: "S2 PN grubu / aciklama (mod) eslesmesi",
    3: "S3 Motor (taraf/pozisyon kesin)",
    4: "S4 Coklu ayni PN - pozisyon eslestirme",
    5: "S5 NHA (ust assy) hareketi",
    6: "S6 PN/SN degisimi (GRB takibi)",
}

_EV_FIELDS = ["date", "tx_type", "tx_raw", "pn", "sn", "grb", "description", "pn_group", "aircraft",
              "position", "nha_pn", "nha_sn", "wo", "location", "new_pn", "new_sn"]


def ev(row) -> dict:
    d = {f: row[f] for f in _EV_FIELDS if f in row}
    for k in ("pn", "sn", "grb", "aircraft", "nha_pn", "nha_sn", "wo"):
        d[k + "_k"] = key(d.get(k, ""))
    return d


def uid(e: dict) -> tuple:
    return (e["date"], e.get("tx_raw", ""), e["pn_k"], e["sn_k"], e["aircraft_k"], e["nha_sn_k"],
            e.get("position", ""), e["wo_k"])


def part_label(e: dict | None) -> str:
    if not e:
        return "-"
    s = f"PN {e.get('pn', '')} / SN {e.get('sn', '')}"
    if e.get("grb"):
        s += f" / GRB {e['grb']}"
    return s


def where(e: dict) -> str:
    if e["tx_type"] in (NHA_INSTALL, NHA_REMOVE):
        loc = f"NHA {e.get('nha_pn', '')} SN {e.get('nha_sn', '')}"
    else:
        loc = e.get("aircraft", "") or "?"
    if e.get("position"):
        loc += f" [{e['position']}]"
    return loc


@dataclass
class Candidate:
    event: dict
    score: float
    reasons: list[str]
    flags: set[int] = field(default_factory=set)


@dataclass
class Step:
    no: int
    installed: dict
    removed: dict | None
    score: float = 0.0
    reasons: list[str] = field(default_factory=list)
    scenarios: set[int] = field(default_factory=set)
    alternatives: list[Candidate] = field(default_factory=list)
    note: str = ""


@dataclass
class FAResult:
    query: dict
    status: str = "HATA"
    message: str = ""
    final: dict | None = None
    steps: list[Step] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    scenarios: set[int] = field(default_factory=set)
    confidence: float = 0.0
    changes: list[dict] = field(default_factory=list)

    @property
    def status_text(self) -> str:
        return STATUS_TEXT.get(self.status, self.status)

    def text(self) -> str:
        q = self.query
        L = [f"FA {q.get('fa_no') or '-'} | Sorgu: PN {q.get('pn') or '-'} / SN {q.get('sn') or '-'}"
             f" / GRB {q.get('grb') or '-'}", ""]
        L.append(f"SONUC : {self.status_text}")
        if self.final:
            L.append(f"PARCA : {part_label(self.final)}  ({self.final.get('description', '')})")
        if self.message:
            L.append(f"DETAY : {self.message}")
        if self.status == "HATA":
            L.append("NE YAPMALI: Bu bir veri alma hatasidir, FA sonucu degildir. Altta kayit kutusundaki "
                     "satirlarin ekran goruntusunu iletin; bu arada tabloyu Qlik'ten Excel'e indirip "
                     "'Excel export dosyalari' moduyla ayni sorguyu calistirabilirsiniz.")
        L.append(f"GUVEN : %{self.confidence:.0f}")
        if self.scenarios:
            L.append("SENARYO: " + ", ".join(SCENARIO_TEXT[s] for s in sorted(self.scenarios)))
        for c in self.changes:
            L.append(f"PN/SN CHG {c['date']:%d.%m.%Y}: {c['old']} -> {c['new']}")
        if self.steps:
            L += ["", "DONGU ZINCIRI:"]
        for s in self.steps:
            i = s.installed
            L.append(f" {s.no}. {i['date']:%d.%m.%Y} TAKILDI  {part_label(i)} -> {where(i)}"
                     + (f"  WO {i['wo']}" if i.get("wo") else ""))
            if s.removed:
                r = s.removed
                L.append(f"    yerine SOKULDU {part_label(r)} {where(r)} {r['date']:%d.%m.%Y}"
                         f"  (skor {s.score:.0f}: {'; '.join(s.reasons)})")
            if s.note:
                L.append(f"    -> {s.note}")
            for a in s.alternatives[:3]:
                L.append(f"       alternatif: {part_label(a.event)} {where(a.event)} skor {a.score:.0f}")
        if self.warnings:
            L += ["", "UYARILAR:"] + [f" ! {w}" for w in self.warnings]
        return "\n".join(L)


class FAEngine:
    def __init__(self, source: DataSource, cfg: dict | None = None, log=None):
        self.src = source
        self.c = (cfg or load_config())["fa"]
        self.log = log or (lambda m: None)
        self._engine_re = re.compile(self.c["motor_aciklama_regex"])

    # ------------------------------------------------------------------ yardimcilar
    def is_engine(self, e: dict) -> bool:
        pn = e.get("pn_k", "")
        if any(pn.startswith(key(p)) for p in self.c["motor_pn_onekleri"]):
            return True
        return bool(self._engine_re.search(e.get("description", "") or "")
                    or self._engine_re.search(e.get("pn_group", "") or ""))

    def _history(self, pn, sn, grb) -> pd.DataFrame:
        h = self.src.part_history(pn, sn, grb)
        return h.copy() if h is not None else pd.DataFrame()

    @staticmethod
    def _merge(a: pd.DataFrame, b: pd.DataFrame) -> pd.DataFrame:
        if a.empty:
            return b
        df = pd.concat([a, b], ignore_index=True)
        df = df.drop_duplicates(subset=["date", "tx_raw", "pn_k", "sn_k", "aircraft_k", "position", "wo_k"])
        return df.sort_values(["date", "seq"], kind="stable").reset_index(drop=True)

    # ------------------------------------------------------------------ Senaryo 6
    def resolve_identity(self, pn, sn, grb, hist: pd.DataFrame):
        """PN/SN degisimlerini (acik CHG hareketi veya GRB uzerinden) takip eder.

        Doner: (birlesik_gecmis, degisimler, kimlik_anahtarlari, guncel_kimlik_dict)
        """
        hist = hist.sort_values(["date", "seq"], kind="stable").reset_index(drop=True)
        if key(pn) and key(sn):
            ident = (key(pn), key(sn))
            disp = (pn, sn)
        else:
            first = hist.iloc[0]
            ident, disp = (first.pn_k, first.sn_k), (first.pn, first.sn)
        if not key(grb):
            g = hist[(hist.pn_k == ident[0]) & (hist.sn_k == ident[1]) & (hist.grb_k != "")]
            grb = g.grb.iloc[-1] if not g.empty else ""
        seen, changes, since = {ident}, [], pd.Timestamp.min
        for _ in range(6):
            mine = hist[(hist.pn_k == ident[0]) & (hist.sn_k == ident[1]) & (hist.date >= since)]
            chg = mine[mine.tx_type == PNSN_CHG]
            new = None
            if not chg.empty:
                r = chg.iloc[0]
                if r.new_pn_k or r.new_sn_k:
                    new = (r.new_pn_k or ident[0], r.new_sn_k or ident[1])
                    ndisp, d, how = (r.new_pn or r.pn, r.new_sn or r.sn), r.date, "PN SN CHG hareketi"
                else:
                    later = hist[(hist.date >= r.date) & ~hist.apply(lambda x: (x.pn_k, x.sn_k) in seen, axis=1)]
                    if not later.empty:
                        x = later.iloc[0]
                        new, ndisp, d, how = (x.pn_k, x.sn_k), (x.pn, x.sn), r.date, "PN SN CHG + GRB"
            elif key(grb):
                later = hist[(hist.grb_k == key(grb)) & (hist.date >= since)
                             & ~hist.apply(lambda x: (x.pn_k, x.sn_k) in seen, axis=1)]
                if not later.empty:
                    x = later.iloc[0]
                    new, ndisp, d, how = (x.pn_k, x.sn_k), (x.pn, x.sn), x.date, "ayni GRB farkli PN/SN"
            if new is None or new in seen:
                break
            changes.append({"date": d, "old": f"{disp[0]} / {disp[1]}", "new": f"{ndisp[0]} / {ndisp[1]}",
                            "new_pn": ndisp[0], "new_sn": ndisp[1], "how": how})
            seen.add(new)
            ident, disp, since = new, ndisp, d
            hist = self._merge(hist, self._history(ndisp[0], ndisp[1], grb))
        own = hist[hist.apply(lambda x: (x.pn_k, x.sn_k) in seen or (key(grb) and x.grb_k == key(grb)), axis=1)]
        current = {"pn": disp[0], "sn": disp[1], "grb": grb}
        return own.reset_index(drop=True), changes, seen, current

    # ------------------------------------------------------------------ eslesme skoru
    def score(self, i: dict, r: dict, engine: bool, win_days: int) -> Candidate | None:
        reasons, flags = [], set()
        # 1) kimlik: PN / motor PN koku / PN grubu / dash farki / aciklama
        if i["pn_k"] == r["pn_k"]:
            id_pts = 35
            reasons.append("ayni PN")
        elif engine and (pn_base(i["pn"]) == pn_base(r["pn"]) or pn_prefix_ratio(i["pn"], r["pn"]) >= 0.6):
            id_pts = 28
            reasons.append("motor PN koku ayni")
            flags.add(3)
        elif i.get("pn_group") and key(i["pn_group"]) == key(r.get("pn_group", "")):
            id_pts = 28
            reasons.append(f"ayni PN grubu ({i['pn_group']})")
            flags.add(2)
        elif pn_base(i["pn"]) == pn_base(r["pn"]):
            id_pts = 24
            reasons.append("PN koku ayni (dash farki)")
        else:
            id_pts = 0
        desc = text_similarity(i.get("description", ""), r.get("description", ""))
        if id_pts == 0:
            if desc < self.c["aciklama_esik"]:
                return None
            flags.add(2)
            reasons.append(f"sadece aciklama benzerligi %{desc * 100:.0f} (mod?)")
        desc_pts = 20 * desc if (i.get("description") and r.get("description")) else 10
        if id_pts and desc >= 0.5 and "aciklama" not in " ".join(reasons):
            reasons.append(f"aciklama %{desc * 100:.0f}")
        # 2) zaman / is emri
        t_pts = 0.0
        if i["wo_k"] and i["wo_k"] == r["wo_k"]:
            t_pts += 20
            reasons.append("ayni is emri")
        days = abs((r["date"] - i["date"]).total_seconds()) / 86400
        t_pts += 12 if days < 1 else max(0.0, 12 * (1 - days / (win_days + 1)))
        if days >= 1:
            reasons.append(f"{days:.0f} gun fark")
        if r["date"] > i["date"] + pd.Timedelta(days=1):
            t_pts -= 3  # sokum takimdan sonra kaydedilmis
        # 3) pozisyon
        sim, side_conf, num_conf = position_compare(i.get("position", ""), r.get("position", ""))
        if engine and (side_conf or num_conf):
            return None  # S3: saga takilan motorun yerine soldan sokulen olamaz
        pos_pts = 25 * sim
        if i.get("position") and r.get("position"):
            if side_conf:
                pos_pts -= 25
                reasons.append("taraf uyusmuyor!")
            elif num_conf:
                pos_pts -= 10
                reasons.append("pozisyon no farkli")
            elif sim >= 0.7:
                reasons.append(f"pozisyon uyumlu ({r['position']})")
        total = id_pts + desc_pts + t_pts + pos_pts
        return Candidate(r, max(0.0, min(100.0, 100 * total / 112)), reasons, flags)

    # ------------------------------------------------------------------ yerine sokuleni bul
    def find_removal(self, inst: dict, used: set):
        """Doner: (secilen Candidate|None, inst icin tum adaylar, senaryo bayraklari, not)."""
        engine = self.is_engine(inst)
        win_days = self.c["motor_pencere_gun"] if engine else self.c["pencere_gun"]
        win = pd.Timedelta(days=win_days)
        d = inst["date"]
        flags = {3} if engine else set()
        if inst["tx_type"] == NHA_INSTALL:
            flags.add(5)
            if not inst["nha_pn_k"]:
                return None, [], flags, "NHA INSTALL hareketinde ust assy PN bilgisi yok"
            evs = self.src.nha_events(inst["nha_pn"], inst["nha_sn"], d - win, d + win)
            rem_t, ins_t = NHA_REMOVE, NHA_INSTALL
        else:
            if not inst["aircraft_k"]:
                return None, [], flags, "Takma hareketinde ucak bilgisi yok"
            evs = self.src.aircraft_events(inst["aircraft"], d - win, d + win)
            rem_t, ins_t = AC_REMOVE, AC_INSTALL
        evs = [ev(r) for _, r in evs.iterrows()] if evs is not None and len(evs) else []
        same_part = lambda e: e["pn_k"] == inst["pn_k"] and e["sn_k"] == inst["sn_k"]
        rems = [e for e in evs if e["tx_type"] == rem_t and uid(e) not in used and not same_part(e)]
        ins = {uid(e): e for e in evs if e["tx_type"] == ins_t}
        ins[uid(inst)] = inst
        me = uid(inst)

        # tum takilan x sokulen skor matrisi
        matrix: dict[tuple, Candidate] = {}
        for iu, ie in ins.items():
            ie_engine = engine if iu == me else self.is_engine(ie)
            for r in rems:
                c = self.score(ie, r, ie_engine, win_days)
                if c:
                    matrix[(iu, uid(r))] = c
        mine = sorted((c for (iu, _), c in matrix.items() if iu == me), key=lambda c: -c.score)
        # S4: ayni removal icin yarisan baska takilanlar var mi?
        my_rem = {uid(c.event) for c in mine}
        rivals = {iu for (iu, ru) in matrix if iu != me and ru in my_rem}
        if rivals:
            flags.add(4)
        # global greedy birebir atama
        taken_i, taken_r, assign = set(), set(), {}
        for (iu, ru), c in sorted(matrix.items(), key=lambda kv: -kv[1].score):
            if c.score < self.c["min_skor"]:
                break
            if iu in taken_i or ru in taken_r:
                continue
            taken_i.add(iu)
            taken_r.add(ru)
            assign[iu] = c
        best = assign.get(me)
        note = ""
        if best is None and mine and mine[0].score >= self.c["min_skor"]:
            note = "Uygun sokulen parcalar ayni is/gun icindeki diger takilan parcalarla eslesti"
        if best is not None and rivals:
            best.reasons.append(f"{len(rivals) + 1} takilan arasindan pozisyona gore secildi")
        return best, mine, flags, note

    # ------------------------------------------------------------------ durum
    @staticmethod
    def state_after(hist: pd.DataFrame, after: dict) -> tuple[str, dict | None]:
        """Bir hareketten sonra parcanin son durumu: TAKILI / DEPODA / SEVK / HURDA / ACIKTA."""
        rows = [ev(r) for _, r in hist[hist.date >= after["date"]].iterrows()]
        idx = next((k for k, e in enumerate(rows) if uid(e) == uid(after)), None)
        if idx is None:
            idx = next((k for k, e in enumerate(rows)
                        if e["tx_type"] == after["tx_type"] and e["date"] == after["date"]), None)
        rest = rows[idx + 1:] if idx is not None else [e for e in rows if e["date"] > after["date"]]
        state, last = "ACIKTA", None
        for e in rest:
            t = e["tx_type"]
            if t in INSTALLS:
                return "TAKILI", e
            if t == STORE_IN:
                state, last = "DEPODA", e
            elif t == STORE_OUT:
                state, last = "DEPODAN_CIKTI", e
            elif t == SHIPPED:
                state, last = "SEVK", e
            elif t == SCRAP:
                state, last = "HURDA", e
            elif t in REMOVES:
                state, last = "ACIKTA", e
        return state, last

    # ------------------------------------------------------------------ ana akis
    def trace(self, pn: str = "", sn: str = "", grb: str = "", fa_no: str = "", fa_date=None) -> FAResult:
        q = {"pn": pn, "sn": sn, "grb": grb, "fa_no": fa_no, "fa_date": fa_date}
        res = FAResult(q)
        try:
            return self._trace(res)
        except Exception as e:  # GUI/toplu modda tek kayit hatasi tum isi durdurmasin
            res.status, res.message = "HATA", f"{type(e).__name__}: {e}"
            return res

    def _trace(self, res: FAResult) -> FAResult:
        q = res.query
        if q["fa_no"] and not (key(q["pn"]) or key(q["grb"])):
            hit = self.src.fa_lookup(q["fa_no"])
            if hit:
                q.update({k: v for k, v in hit.items() if v is not None and v == v and k in q and not q.get(k)})
                q["fa_date"] = q.get("fa_date") or hit.get("fa_date")
        if not (key(q["pn"]) and key(q["sn"])) and not key(q["grb"]):
            res.status, res.message = "HATA", "PN+SN veya GRB girilmeli"
            return res
        self.log(f"Parca gecmisi aliniyor: {q['pn']} {q['sn']} {q['grb']}")
        hist = self._history(q["pn"], q["sn"], q["grb"])
        if hist.empty:
            res.status = "BULUNAMADI"
            return res
        hist, changes, ids, current = self.resolve_identity(q["pn"], q["sn"], q["grb"], hist)
        res.changes = changes
        moves = hist[hist.tx_type.isin(INSTALLS | REMOVES)]
        if changes:
            res.scenarios.add(6)
            last_chg = changes[-1]["date"]
            if moves[moves.date > last_chg].empty:
                res.status = "PNSN_CHG"
                res.final = {"pn": current["pn"], "sn": current["sn"], "grb": current["grb"],
                             "description": hist.description.iloc[-1], "date": last_chg}
                res.message = (f"Parca yerde PN/SN degisimi gecirmis ({changes[-1]['how']}). Sonrasinda "
                               f"takma/sokme yok -> FA {changes[-1]['new']} olarak guncellenmeli.")
                res.confidence = 90 if changes[-1]["how"].startswith("PN SN CHG") else 75
                return res

        start = self._start_install(moves, q.get("fa_date"))
        if start is None:
            if moves.empty:
                state, last = self.state_after(hist, ev(hist.iloc[0])) if len(hist) > 1 else ("ACIKTA", None)
            else:
                state, last = self.state_after(hist, ev(moves.iloc[-1]))
            base = ev(moves.iloc[-1]) if not moves.empty else ev(hist.iloc[-1])
            res.final = base
            if state == "DEPODA":
                res.status, res.message = "DEPODA", "FA'deki parcanin kendisi depoya girmis."
            else:
                res.status = "FA_DOGRU"
                res.message = ("Parcanin son takma/sokme hareketi sokum; sonrasinda takilmamis. "
                               f"Son durum: {state}.")
            res.confidence = 80
            return res
        return self._chain(res, start)

    def _start_install(self, moves: pd.DataFrame, fa_date) -> dict | None:
        if moves.empty:
            return None
        if fa_date is not None and not pd.isna(fa_date):
            fd = pd.Timestamp(fa_date) - pd.Timedelta(days=self.c["pencere_gun"])
            cand = moves[(moves.date >= fd) & moves.tx_type.isin(INSTALLS)]
            if not cand.empty:
                return ev(cand.iloc[0])
        last = moves.iloc[-1]
        return ev(last) if last.tx_type in INSTALLS else None

    def _chain(self, res: FAResult, inst: dict) -> FAResult:
        used, visited = set(), set()
        limit = self.c["motor_max_adim"] if self.is_engine(inst) else self.c["max_adim"]
        scores: list[float] = []
        for n in range(1, limit + 1):
            if uid(inst) in visited:
                res.status, res.message = "LIMIT", "Dongu kendi kendine kapandi (ayni takma tekrarlandi)."
                return res
            visited.add(uid(inst))
            self.log(f"Adim {n}: {part_label(inst)} {inst['date']:%d.%m.%Y} {where(inst)}")
            best, cands, flags, note = self.find_removal(inst, used)
            step = Step(n, inst, None, scenarios={1} | flags)
            res.steps.append(step)
            res.scenarios |= step.scenarios
            if best is None:
                step.alternatives = cands[:5]
                step.scenarios.add(2)
                res.scenarios.add(2)
                res.status, res.final = "FA_SIL", inst
                step.note = note or "Yerine sokulen parca bulunamadi"
                res.message = (f"{part_label(inst)} {inst['date']:%d.%m.%Y} tarihinde {where(inst)} "
                               "takilmis, ancak yerine sokulen bir parca (PN/PN grubu/aciklama bazinda) yok.")
                if inst["pn"].startswith("EO"):
                    res.message += " EO'lu kalem: mod kapsaminda ekleme olabilir."
                if cands:
                    res.warnings.append("Esik altinda olasi eslesmeler var, alternatifleri kontrol edin.")
                res.confidence = 70 if not cands else 55
                if scores:
                    res.confidence = min(res.confidence, min(scores))
                return res
            r = best.event
            used.add(uid(r))
            step.removed, step.score, step.reasons = r, best.score, best.reasons
            step.scenarios |= best.flags
            res.scenarios |= step.scenarios
            step.alternatives = [c for c in cands if uid(c.event) != uid(r)][:3]
            conf = best.score
            if step.alternatives and best.score - step.alternatives[0].score < self.c["belirsizlik_farki"]:
                conf -= 15
                res.warnings.append(f"Adim {n}: sokulen parca secimi belirsiz "
                                    f"({part_label(r)} vs {part_label(step.alternatives[0].event)}).")
            scores.append(conf)

            # sokulen parcanin sonrasi
            r_hist = self._history(r["pn"], r["sn"], r["grb"])
            if not r_hist.empty:
                r_hist, r_chg, _, r_cur = self.resolve_identity(r["pn"], r["sn"], r["grb"], r_hist)
                after = [c for c in r_chg if c["date"] >= r["date"]]
                if after:
                    step.scenarios.add(6)
                    res.scenarios.add(6)
                    res.changes += after
                state, nxt = self.state_after(r_hist, r)
            else:
                state, nxt, after = "ACIKTA", None, []
            if state == "TAKILI" and nxt is not None:
                step.note = f"Sokulen parca {nxt['date']:%d.%m.%Y} tarihinde {where(nxt)} tekrar takilmis -> dongu devam"
                inst = nxt
                continue
            final = dict(r)
            if after:
                final.update({"pn": after[-1]["new_pn"], "sn": after[-1]["new_sn"]})
                step.note = f"Sokulen parca sonradan PN/SN degistirmis: {after[-1]['old']} -> {after[-1]['new']}"
            res.final = final
            res.confidence = min(scores)
            if state == "DEPODA":
                res.status = "DEPODA"
                res.message = (f"Dongu {n} adimda bitti. Son sokulen {part_label(final)} depoya girmis, "
                               "FA gerekmeyebilir.")
            else:
                res.status = "FA_ATA"
                res.message = f"Dongu {n} adimda bitti. Son sokulen parca bir daha takilmamis."
                if state in ("SEVK", "HURDA"):
                    res.warnings.append(f"Son parca durumu: {state} - FA atamadan once kontrol edin.")
            if n > 1:
                step.note = (step.note + " | " if step.note else "") + "Dongu sonu"
            return res
        res.status, res.message = "LIMIT", f"{limit} adimda dongu bitmedi."
        res.confidence = min(scores) if scores else 0
        return res
