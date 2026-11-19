"""config.yaml okuma. Dosyada olmayan her ayar icin buradaki varsayilan kullanilir."""
from __future__ import annotations

import copy
from pathlib import Path

import yaml

ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = ROOT / "config.yaml"

DEFAULTS: dict = {
    "fa": {
        "pencere_gun": 3,            # takilan/sokulen eslesmesi icin +/- gun
        "motor_pencere_gun": 10,     # motor degisimleri daha uzun surebilir
        "max_adim": 15,              # dongu derinlik limiti
        "motor_max_adim": 40,        # motorlarda uzun donguler (Senaryo 3)
        "min_skor": 45,              # bu skorun altindaki aday eslesme sayilmaz (0-100)
        "belirsizlik_farki": 6,      # ilk iki aday arasi fark bundan azsa uyar
        "aciklama_esik": 0.55,       # PN/grup tutmazsa aciklama benzerligi en az bu olmali
        "motor_pn_onekleri": ["CFM", "LEAP", "V25", "PW", "GE90", "GENX", "TRENT", "CF6", "CF34", "PW1"],
        "motor_aciklama_regex": r"^(ENGINE|ENG|MOTOR|POWER ?PLANT)\b(?!.*\b(VALVE|SENSOR|PUMP|MOUNT|COWL|BRACKET|HOSE|DUCT|FILTER|SWITCH|HARNESS|SEAL|STARTER|GENERATOR|CONTROL|PROBE)\b)",
    },
    "kolonlar": {},  # ornek: {pn: "Part Number", grb: "GRB No"}
    "qlik": {
        "base_url": "https://sensedashboard.thy.com",
        "sanal_proxy": "",
        "giris": {
            "mod": "form",           # form | windows | manuel
            "kullanici_selector": "input[name='username'], #username-input, input[type='email'], input[name='loginfmt']",
            "sifre_selector": "input[type='password']",
            "gonder_selector": "button[type='submit'], #loginbtn, input[type='submit']",
            "basarili_url_icerir": "/hub",
        },
        "tarayici": "msedge",        # msedge | chrome | chromium
        "tarayici_yolu": "",         # ozel tarayici exe yolu (opsiyonel)
        "gorunmez": False,
        "zaman_asimi_sn": 90,
        "export": {
            "sag_tik_menu": ["Download", "İndir", "Indir", "Export", "Dışa aktar"],
            "veri_secenegi": ["Data", "Veri", "Export data", "Verileri dışa aktar"],
            "onay_butonu": ["Download", "İndir", "Export", "Dışa aktar"],
            "link_selector": "a.export-url, a[href*='TempContent'], a[download]",
        },
        "fa_hareket": {
            "app_id": "609d5232-db11-4195-90d1-483cffa420e1", "sheet_id": "", "object_id": "",
            "alanlar": {"pn": "PN", "sn": "SN", "grb": "GOODS_RCVD_BATCH", "aircraft": "AC",
                        "nha_pn": "NHA_PN", "nha_sn": "NHA_SN"},
        },
        "fa_listesi": {"app_id": "", "sheet_id": "", "object_id": "", "alanlar": {"fa_no": "FA No"}},
        "envanter": {
            "app_id": "410f6aca-0d09-4832-9fda-af9287922832", "sheet_id": "", "object_id": "",
            "durum_alani": "Durum", "durum_degeri": "false",
            "notes_kolonu": "Notes",
        },
    },
    "cikti_klasoru": "Raporlar",
}


def _merge(base: dict, extra: dict) -> dict:
    out = copy.deepcopy(base)
    for k, v in (extra or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config(path: str | Path | None = None) -> dict:
    p = Path(path) if path else CONFIG_PATH
    if not p.exists() and not path:
        p = ROOT / "config.ornek.yaml"
    user = {}
    if p.exists():
        with open(p, encoding="utf-8") as f:
            user = yaml.safe_load(f) or {}
    cfg = _merge(DEFAULTS, user)
    out_dir = Path(cfg["cikti_klasoru"])
    cfg["cikti_klasoru"] = str(out_dir if out_dir.is_absolute() else ROOT / out_dir)
    return cfg


def ensure_user_config() -> Path:
    """config.yaml yoksa ornekten olusturur ve yolunu dondurur."""
    if not CONFIG_PATH.exists():
        import shutil
        shutil.copy(ROOT / "config.ornek.yaml", CONFIG_PATH)
    return CONFIG_PATH
