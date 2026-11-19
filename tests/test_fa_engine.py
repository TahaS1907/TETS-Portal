"""Excel'deki senaryolarin her biri icin motor testleri (sentetik veri)."""
import pandas as pd
import pytest

from tets.config import load_config
from tets.fa.engine import FAEngine
from tets.fa.report import write_report
from tets.normalize import position_compare, tx_type
from tets.sources.excel_source import ExcelSource
from tests.sample_data import write


@pytest.fixture(scope="module")
def files(tmp_path_factory):
    return write(tmp_path_factory.mktemp("veri"))


@pytest.fixture(scope="module")
def engine(files):
    return FAEngine(ExcelSource(files[0], fa_list_path=files[1]), load_config())


def test_s1_dongu_ikinci_ucakta_biter(engine):
    r = engine.trace("3214-01", "A100")
    assert r.status == "FA_ATA"
    assert (r.final["pn"], r.final["sn"]) == ("3214-02", "C300")
    assert len(r.steps) == 2 and 1 in r.scenarios


def test_s2_sokulen_yok_fa_silinmeli(engine):
    r = engine.trace("EO-5566", "E1")
    assert r.status == "FA_SIL"


def test_s2_pn_grubu_ile_mod_eslesmesi(engine):
    r = engine.trace("EO-7788", "E2")
    assert r.status == "FA_ATA" and r.final["sn"] == "V9" and 2 in r.scenarios
    assert any("SEVK" in w for w in r.warnings)


def test_s3_motor_taraf_kesin(engine):
    r = engine.trace("CFM56-7B26/3", "890111")
    assert r.status == "FA_ATA"
    assert r.steps[0].removed["sn"] == "890444"      # soldaki 890222 secilmemeli
    assert r.final["sn"] == "890555"                 # #1 motor 890666 secilmemeli
    assert 3 in r.scenarios


@pytest.mark.parametrize("sn,expected", [("L1", "L9"), ("L2", "L8")])
def test_s4_pozisyon_eslestirme(engine, sn, expected):
    r = engine.trace("LAMP-100", sn)
    assert r.final["sn"] == expected and 4 in r.scenarios


def test_s5_nha(engine):
    r = engine.trace("BRG-55", "N1")
    assert r.status == "FA_ATA" and r.final["sn"] == "N7" and 5 in r.scenarios


def test_s5_nha_sokulen_depoda(engine):
    assert engine.trace("BRG-55", "N2").status == "DEPODA"


def test_s6_pnsn_chg_grb_ile(engine):
    r = engine.trace(grb="G600")
    assert r.status == "PNSN_CHG" and (r.final["pn"], r.final["sn"]) == ("4455-B", "X1A")


def test_s6_chg_satirinda_yeni_pn_yok(engine):
    r = engine.trace("6677-1", "Y1", "G700")
    assert r.status == "PNSN_CHG" and r.final["pn"] == "6677-2"


def test_s6_chg_sonrasi_dongu(engine):
    r = engine.trace("7001", "Z1", "G800")
    assert r.status == "FA_ATA" and r.final["sn"] == "Z5" and {1, 6} <= r.scenarios


def test_fa_dogru(engine):
    assert engine.trace("8800", "K1").status == "FA_DOGRU"


def test_fa_no_ile_sorgu(engine):
    r = engine.trace(fa_no="FA-1001")
    assert r.final["sn"] == "C300"


def test_bulunamadi(engine):
    assert engine.trace("YOK", "YOK").status == "BULUNAMADI"


def test_rapor(engine, tmp_path):
    rs = [engine.trace("3214-01", "A100"), engine.trace("EO-5566", "E1")]
    p = write_report(rs, tmp_path)
    xl = pd.read_excel(p, sheet_name=None)
    assert {"Ozet", "Dongu Zinciri", "Alternatifler", "Aciklama"} <= set(xl)
    assert list(xl["Ozet"]["Hedef SN"]) == ["C300", "E1"]


def test_toplu(files):
    from tets.fa.runner import open_source, run_batch
    cfg = load_config()
    src = open_source(cfg, "excel", [files[0]])
    cfg = {**cfg, "cikti_klasoru": str(files[0].parent)}
    results, path = run_batch(src, cfg, files[1])
    assert len(results) == 12 and path.exists()
    assert all(r.status != "HATA" for r in results)


def test_normalizasyon():
    assert tx_type("A/C INSTALL") == "AC INSTALL"
    assert tx_type("REMOVE FROM NHA") == "NHA REMOVE"
    assert tx_type("Sökül") == "AC REMOVE"
    sim, side, num = position_compare("ENG #2 RH", "#2 RIGHT")
    assert sim > 0.9 and not side and not num
    assert position_compare("LH MLG", "RH MLG")[1]


def test_envanter_null_filtre():
    from tets.envanter import filter_null_notes
    df = pd.DataFrame({"PN": [1, 2, 3], "NOTES": [None, "x", "-"]})
    assert list(filter_null_notes(df, "Notes").PN) == [1, 3]


def test_qlik_url():
    from tets.sources.qlik_source import QlikSession
    cfg = load_config()
    cfg["qlik"]["base_url"] = "https://qlik.local"
    s = QlikSession(cfg)
    url = s.sheet_url({"app_id": "A1", "sheet_id": "S1"}, {"PN": "CFM56-7B26/3", "SN": ""})
    assert url == ("https://qlik.local/sense/app/A1/sheet/S1/state/analysis/options/clearselections"
                   "/select/PN/%5BCFM56-7B26%2F3%5D")


def test_rot_fa_pn_history_kolonlari():
    """Qlik'teki ROT FA PN History tablosunun gercek basliklari taninmali."""
    from tets.columns import map_columns
    cols = ["PN", "SN", "PN_DESCRIPTION", "NHA_PN", "NHA_SN", "NLA_POSITION", "PN_GROUP", "GOODS_RCVD_BATCH",
            "BATCH", "TRANSACTION_TYPE", "AC", "AC_POSITION", "LOCATION", "CREATED_DATE", "CREATED_BY", "WO",
            "TASK_CARD", "ORDER_TYPE", "ORDER_NO", "ORDER_LINE", "UNIT_COST", "Count(AC)"]
    m = map_columns(pd.DataFrame(columns=cols))
    assert m == {"date": "CREATED_DATE", "tx_type": "TRANSACTION_TYPE", "pn": "PN", "sn": "SN",
                 "grb": "GOODS_RCVD_BATCH", "description": "PN_DESCRIPTION", "pn_group": "PN_GROUP",
                 "aircraft": "AC", "position": "AC_POSITION", "nla_position": "NLA_POSITION",
                 "nha_pn": "NHA_PN", "nha_sn": "NHA_SN", "wo": "WO", "location": "LOCATION"}


def test_nha_pozisyonu_nla_kolonundan(engine):
    r = engine.trace("BRG-55", "N1")
    assert r.steps[0].removed["position"] == "INBD"


def test_s6_son_hareket_pnsn_chg_grb_ile_devam(engine):
    """Sadece PN+SN girildi; son hareket PN SN CHG. Yeni PN/SN GRB'den bulunup dongu devam etmeli."""
    r = engine.trace("5500-1", "P1")
    assert r.status == "FA_ATA" and r.final["sn"] == "P7"
    assert r.changes and r.changes[0]["new"] == "5500-2 / P1X" and 6 in r.scenarios


def test_ayni_grb_paralel_parca_degisim_sayilmaz():
    """GRB'si ayni ama eski parca da hareket etmeye devam ediyorsa PN/SN degisimi sayilmamali."""
    import pandas as pd
    from tests.sample_data import COLS, R
    from tets.columns import normalize_frame
    from tets.sources.base import FrameSource
    rows = [R("01.01.2025", "STORE RECEIPT", "A-1", "S1", "GX", "X"),
            R("02.01.2025", "STORE RECEIPT", "A-1", "S2", "GX", "X"),
            R("05.01.2025", "AC REMOVE", "A-1", "S1", "GX", "X", ac="TC-X", wo="W")]
    eng = FAEngine(FrameSource(normalize_frame(pd.DataFrame(rows, columns=COLS))), load_config())
    r = eng.trace("A-1", "S1")
    assert not r.changes and r.status == "FA_DOGRU"


def test_s6_degisim_sonrasi_depoda_fa_pnsn_degismeli(engine):
    r = engine.trace("7B26E-X", "Q9")
    assert r.status == "PNSN_CHG" and r.final["pn"] == "7B27E-X"
    assert "değiştirilmeli" in r.text() and "depoya girmiş" in r.text()
