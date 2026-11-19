"""FA sonuclarini bicimli Excel raporuna yazar."""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .engine import SCENARIO_TEXT, FAResult, part_label, where

STATUS_COLOR = {
    "FA_ATA": "C6EFCE", "FA_SIL": "FFC7CE", "PNSN_CHG": "FFEB9C", "DEPODA": "DDEBF7",
    "FA_DOGRU": "E2EFDA", "BULUNAMADI": "D9D9D9", "LIMIT": "F8CBAD", "HATA": "F8CBAD",
}


def _d(x):
    return x.strftime("%d.%m.%Y") if isinstance(x, (pd.Timestamp, datetime)) and not pd.isna(x) else ""


def results_to_frames(results: list[FAResult]):
    summary, chain, alts = [], [], []
    for r in results:
        q, f = r.query, r.final or {}
        summary.append({
            "FA No": q.get("fa_no", ""), "Sorgu PN": q.get("pn", ""), "Sorgu SN": q.get("sn", ""),
            "Sorgu GRB": q.get("grb", ""), "Sonuc": r.status_text, "Kod": r.status,
            "Hedef PN": f.get("pn", ""), "Hedef SN": f.get("sn", ""), "Hedef GRB": f.get("grb", ""),
            "Hedef Aciklama": f.get("description", ""), "Hedef Konum": where(f) if f.get("tx_type") else "",
            "Guven %": round(r.confidence), "Adim": len(r.steps),
            "Senaryolar": ", ".join(SCENARIO_TEXT[s] for s in sorted(r.scenarios)),
            "Detay": r.message, "Uyarilar": " | ".join(r.warnings),
        })
        for s in r.steps:
            i, m = s.installed, s.removed or {}
            chain.append({
                "FA No": q.get("fa_no", ""), "Adim": s.no, "Takma Tarihi": _d(i["date"]),
                "Takilan": part_label(i), "Takildigi Yer": where(i), "WO": i.get("wo", ""),
                "Sokulen": part_label(m) if m else "-", "Sokulen Yer": where(m) if m else "",
                "Sokum Tarihi": _d(m.get("date")), "Skor": round(s.score),
                "Gerekce": "; ".join(s.reasons), "Senaryo": ", ".join(f"S{x}" for x in sorted(s.scenarios)),
                "Not": s.note,
            })
            for a in s.alternatives:
                alts.append({"FA No": q.get("fa_no", ""), "Adim": s.no, "Aday": part_label(a.event),
                             "Yer": where(a.event), "Tarih": _d(a.event["date"]), "Skor": round(a.score),
                             "Gerekce": "; ".join(a.reasons)})
    return pd.DataFrame(summary), pd.DataFrame(chain), pd.DataFrame(alts)


def write_report(results: list[FAResult], out_dir: str | Path, name: str | None = None) -> Path:
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = out_dir / (name or f"FA_Rapor_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    summary, chain, alts = results_to_frames(results)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        summary.to_excel(xw, sheet_name="Ozet", index=False)
        (chain if not chain.empty else pd.DataFrame({"Bilgi": ["Zincir yok"]})).to_excel(xw, sheet_name="Dongu Zinciri", index=False)
        (alts if not alts.empty else pd.DataFrame({"Bilgi": ["Alternatif yok"]})).to_excel(xw, sheet_name="Alternatifler", index=False)
        pd.DataFrame({"Detay": "\n\n".join(r.text() for r in results).split("\n")}).to_excel(xw, sheet_name="Aciklama", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for c in ws[1]:
                c.font = Font(bold=True, color="FFFFFF")
                c.fill = PatternFill("solid", fgColor="C8102E")
            for col in ws.columns:
                width = min(60, max(10, *(len(str(c.value or "")) for c in col)) + 2)
                ws.column_dimensions[get_column_letter(col[0].column)].width = width
            ws.auto_filter.ref = ws.dimensions
        ws = xw.book["Ozet"]
        kod_col = list(summary.columns).index("Kod") + 1 if not summary.empty else None
        if kod_col:
            for row in ws.iter_rows(min_row=2):
                fill = STATUS_COLOR.get(row[kod_col - 1].value)
                if fill:
                    for c in row:
                        c.fill = PatternFill("solid", fgColor=fill)
                for c in row:
                    c.alignment = Alignment(vertical="top", wrap_text=False)
    return path
