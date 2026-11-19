"""Excel'deki 6 senaryonun hepsini iceren sentetik hareket verisi.

Calistirinca ornek_veri/ klasorune ornek_hareketler.xlsx ve ornek_fa_listesi.xlsx yazar;
uygulamayi canli portal olmadan denemek icin kullanilabilir.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

COLS = ["Tarih", "Hareket", "Part Number", "Serial No", "GRB", "Description", "PN Group", "A/C",
        "Position", "NHA PN", "NHA SN", "Work Order", "Location", "New PN", "New SN"]


def R(date, tx, pn, sn, grb="", desc="", grp="", ac="", pos="", nha_pn="", nha_sn="", wo="", loc="",
      new_pn="", new_sn=""):
    return [date, tx, pn, sn, grb, desc, grp, ac, pos, nha_pn, nha_sn, wo, loc, new_pn, new_sn]


def rows():
    return [
        # ---------------- S1: 3214-01 A100 -> B200 sokuldu -> B200 TC-JAB'ye takildi -> C300 sokuldu
        R("10.01.2025", "STORE RECEIPT", "3214-01", "A100", "G1", "ACTUATOR MLG", "MLGACT", loc="IST"),
        R("15.01.2025", "STORE ISSUE", "3214-01", "A100", "G1", "ACTUATOR MLG", "MLGACT", loc="IST"),
        R("01.03.2025", "AC REMOVE", "9999-1", "Q1", "", "BRAKE ASSY", "BRK", "TC-JAA", "RH MLG", wo="W1"),
        R("01.03.2025", "AC REMOVE", "3214-01", "B200", "G2", "ACTUATOR MLG", "MLGACT", "TC-JAA", "LH MLG", wo="W1"),
        R("01.03.2025", "AC INSTALL", "3214-01", "A100", "G1", "ACTUATOR MLG", "MLGACT", "TC-JAA", "LH MLG", wo="W1"),
        R("10.04.2025", "AC REMOVE", "3214-02", "C300", "G3", "ACTUATOR, MLG", "MLGACT", "TC-JAB", "MLG L", wo="W2"),
        R("10.04.2025", "AC INSTALL", "3214-01", "B200", "G2", "ACTUATOR MLG", "MLGACT", "TC-JAB", "L MLG", wo="W2"),
        # ---------------- S2a: EO kalem takilmis, yerine sokulen yok
        R("01.05.2025", "AC INSTALL", "EO-5566", "E1", "G4", "BRACKET MOD KIT", "", "TC-JAC", "FWD CARGO", wo="W3"),
        # ---------------- S2b: farkli PN ama ayni PN grubu (mod degisimi)
        R("20.05.2025", "AC REMOVE", "3290011-2", "V9", "G5", "BLEED AIR VALVE", "BLEEDVLV", "TC-JAC", "ENG 1 BLEED", wo="W4"),
        R("20.05.2025", "AC INSTALL", "EO-7788", "E2", "G6", "VALVE BLEED AIR", "BLEEDVLV", "TC-JAC", "ENG 1 BLEED", wo="W4"),
        R("25.05.2025", "REPAIR SHIPMENT", "3290011-2", "V9", "G5", "BLEED AIR VALVE", "BLEEDVLV", loc="VENDOR"),
        # ---------------- S3: motor; sol motor tuzak, sag motor eslesmeli, dongu ikinci ucakta devam
        R("01.06.2025", "AC REMOVE", "CFM56-7B26/F", "890222", "G7", "ENGINE ASSY", "ENG", "TC-JAD", "ENG #1 LH", wo="W5"),
        R("01.06.2025", "AC REMOVE", "CFM56-7B26/3F", "890444", "G8", "ENGINE ASSY", "ENG", "TC-JAD", "#2 RIGHT", wo="W5"),
        R("01.06.2025", "AC INSTALL", "CFM56-7B26", "890333", "G9", "ENGINE ASSY", "ENG", "TC-JAD", "ENG #1 LH", wo="W5"),
        R("01.06.2025", "AC INSTALL", "CFM56-7B26/3", "890111", "G10", "ENGINE ASSY", "ENG", "TC-JAD", "ENG #2 RH", wo="W5"),
        R("12.07.2025", "AC REMOVE", "CFM56-7B26E", "890555", "G11", "ENGINE", "ENG", "TC-JAE", "ENG 2", wo="W6"),
        R("12.07.2025", "AC REMOVE", "CFM56-7B26E", "890666", "G12", "ENGINE", "ENG", "TC-JAE", "ENG 1", wo="W6b"),
        R("15.07.2025", "AC INSTALL", "CFM56-7B26/3F", "890444", "G8", "ENGINE ASSY", "ENG", "TC-JAE", "NO 2 ENG", wo="W6"),
        # ---------------- S4: ayni PN'den iki takma iki sokme, pozisyonla eslesme
        R("05.08.2025", "AC REMOVE", "LAMP-100", "L8", "", "READING LIGHT", "", "TC-JAA", "ROW 14 R", wo="W7"),
        R("05.08.2025", "AC REMOVE", "LAMP-100", "L9", "", "READING LIGHT", "", "TC-JAA", "ROW12 L", wo="W7"),
        R("05.08.2025", "AC INSTALL", "LAMP-100", "L1", "", "READING LIGHT", "", "TC-JAA", "SEAT ROW 12 LH", wo="W7"),
        R("05.08.2025", "AC INSTALL", "LAMP-100", "L2", "", "READING LIGHT", "", "TC-JAA", "SEAT ROW 14 RH", wo="W7"),
        # ---------------- S5: NHA (ust assy) hareketleri
        R("01.09.2025", "NHA REMOVE", "BRG-55", "N7", "", "BEARING", "BRG", nha_pn="WHL-1", nha_sn="W1", wo="W8"),
        R("01.09.2025", "NHA INSTALL", "BRG-55", "N1", "", "BEARING", "BRG", nha_pn="WHL-1", nha_sn="W1", wo="W8"),
        R("05.09.2025", "NHA REMOVE", "BRG-55", "N8", "", "BEARING", "BRG", nha_pn="WHL-1", nha_sn="W2", wo="W9"),
        R("05.09.2025", "NHA INSTALL", "BRG-55", "N2", "", "BEARING", "BRG", nha_pn="WHL-1", nha_sn="W2", wo="W9"),
        R("10.09.2025", "STORE RECEIPT", "BRG-55", "N8", "", "BEARING", "BRG", loc="IST"),
        # ---------------- S6a: acik PN SN CHG, sonrasinda hareket yok
        R("01.02.2025", "STORE RECEIPT", "4455-A", "X1", "G600", "PUMP HYD", "HYDP", loc="IST"),
        R("20.02.2025", "PN SN CHG", "4455-A", "X1", "G600", "PUMP HYD", "HYDP", loc="IST", new_pn="4455-B", new_sn="X1A"),
        # ---------------- S6b: CHG satirinda yeni PN yok, GRB uzerinden bulunur
        R("01.02.2025", "STORE RECEIPT", "6677-1", "Y1", "G700", "SENSOR TEMP", "", loc="IST"),
        R("03.03.2025", "P/N CHANGE", "6677-1", "Y1", "G700", "SENSOR TEMP", "", loc="IST"),
        R("03.03.2025", "STORE RECEIPT", "6677-2", "Y1", "G700", "SENSOR TEMP", "", loc="IST"),
        # ---------------- S6c: CHG sonrasi takilmis -> dongu yeni kimlikle devam
        R("01.09.2025", "PN SN CHG", "7001", "Z1", "G800", "FAN COOLING", "", loc="IST", new_pn="7001-B", new_sn="Z1B"),
        R("01.10.2025", "AC REMOVE", "7001", "Z5", "G801", "FAN COOLING", "", "TC-JAB", "AVIONICS FAN", wo="W10"),
        R("01.10.2025", "AC INSTALL", "7001-B", "Z1B", "G800", "FAN COOLING", "", "TC-JAB", "AVIONICS FAN", wo="W10"),
        # ---------------- FA dogru: parca sokulmus, bir daha takilmamis
        R("11.11.2025", "AC REMOVE", "8800", "K1", "", "COMPUTER FMC", "", "TC-JAA", "E/E BAY", wo="W11"),
    ]


def frame() -> pd.DataFrame:
    return pd.DataFrame(rows(), columns=COLS)


FA_LIST = [
    ("FA-1001", "3214-01", "A100", ""),
    ("FA-1002", "EO-5566", "E1", ""),
    ("FA-1003", "EO-7788", "E2", ""),
    ("FA-1004", "CFM56-7B26/3", "890111", ""),
    ("FA-1005", "LAMP-100", "L1", ""),
    ("FA-1006", "BRG-55", "N1", ""),
    ("FA-1007", "BRG-55", "N2", ""),
    ("FA-1008", "", "", "G600"),
    ("FA-1009", "6677-1", "Y1", "G700"),
    ("FA-1010", "7001", "Z1", "G800"),
    ("FA-1011", "8800", "K1", ""),
]


def write(out_dir: Path) -> tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    h = out_dir / "ornek_hareketler.xlsx"
    f = out_dir / "ornek_fa_listesi.xlsx"
    frame().to_excel(h, index=False)
    pd.DataFrame(FA_LIST, columns=["FA No", "PN", "SN", "GRB"]).to_excel(f, index=False)
    return h, f


if __name__ == "__main__":
    print(*write(Path(__file__).resolve().parent.parent / "ornek_veri"), sep="\n")
