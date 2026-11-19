# TETS-Portal

Şefliğimize ait portal işlerini tek çatı altında toplayan masaüstü uygulaması.
Kullanıcı Qlik Sense şifresini girer, uygulama portalı kendisi açar ve istenen otomasyonu yapar.

| Sekme | Ne yapar |
|---|---|
| **FA Takip** | PN / SN / GRB / FA No girilir, sistem FA döngüsünü hareket geçmişi üzerinden takip eder ve FA'in **hangi parçaya** atanması gerektiğini (ya da silinmesi/güncellenmesi gerektiğini) gerekçesiyle birlikte bulur. Tekli veya toplu (Excel listesi) çalışır. |
| **Envanter Takip** | Teknik Envanter Takibi > Null Veri Düzenleme Raporu'nda durum = false seçer, Excel'e indirir. "Tamamı" veya "Sadece Notes boş olanlar". |
| **Ayarlar** | config.yaml, kalibrasyon (Qlik sayfasını Playwright Inspector ile açar), örnek veriyle deneme. |

## Kurulum ve çalıştırma (Windows)

1. Python 3.10+ kurulu olmalı ("Add to PATH" işaretli).
2. `KURULUM.bat` → sanal ortamı ve paketleri kurar, `config.yaml` oluşturur.
3. `config.yaml` içinde `qlik.base_url`, `fa_hareket.app_id / sheet_id / object_id` ve alan adlarını doldurun.
4. `BASLAT.bat` → uygulama açılır.

Tarayıcı olarak varsayılan **Microsoft Edge** kullanılır (şirket bilgisayarında hazır, ek indirme gerekmez).
Portal olmadan denemek için: **Ayarlar > Örnek veri ile dene**.

Komut satırı da var:
```
python -m tets.cli fa --excel ornek_veri/ornek_hareketler.xlsx --pn 3214-01 --sn A100
python -m tets.cli toplu --excel hareketler.xlsx --liste fa_listesi.xlsx
python -m tets.cli fa --qlik --kullanici TK12345 --pn 3214-01 --sn A100   (şifre sorulur)
```

## FA motoru: senaryolar nasıl uygulanıyor

Kaynak: `Fa portal Senaryoları.xlsx` (dosyada 6 senaryo var).

| # | Senaryo | Uygulama |
|---|---|---|
| 1 | Takılan FA parçasının yerine sökülen bulunur; o da sonradan takıldıysa döngü devam eder | Zincir: *takıldı → yerine sökülen → o parça tekrar takıldı mı?* Takılmayan ilk sökülende biter → **FA buna atanmalı**. Döngü/sonsuz zincir koruması var. |
| 2 | Takılmış ama yerine sökülen yoksa FA silinmeli; PN grubu ve description ile mod eşleşmesi | Aday sökülenler sadece aynı PN değil, **PN grubu**, **dash farkı** ve **açıklama benzerliği** ile de aranır. Hiç aday yoksa **FA silinmeli**, eşik altı adaylar "alternatif" olarak raporlanır. |
| 3 | Motorlarda uzun döngü, PN sonu farklı, sağ/sol kesin | Motor tespit edilince (PN öneki / açıklama) daha geniş tarih penceresi ve adım limiti; **PN kökü** eşleşmesi; taraf (LH/RH) veya motor no (#1/#2) çelişirse aday **tamamen elenir**. |
| 4 | Aynı PN'den çok takma/sökme → pozisyonla eşleştirme | Pozisyonlar normalize edilir (`LH/LEFT/SOL/L`, `#2/NO 2/ENG2`, `ROW12 L`…). Aynı iş/gündeki **tüm takılan × sökülen** skorlanır ve birebir global atama yapılır. |
| 5 | NHA INSTALL/REMOVE, üst assy üzerinden takip | Parça üst assy'e takıldıysa aynı NHA PN/SN'den sökülenler aranır. Döngü sonundaki parça **depodaysa** "FA kapatılabilir" denir. |
| 6 | PN SN CHG, GRB üzerinden takip | Hem açık `PN SN CHG` hareketi (yeni PN/SN kolonlu veya kolonsuz) hem de **aynı GRB'de farklı PN/SN** görünmesi değişim sayılır. Sonrasında takma/sökme yoksa **"FA'e yeni PN/SN gelmeli"**; varsa döngü yeni kimlikle devam eder. |

Her eşleşme 0–100 arası skorlanır (PN/grup/açıklama + iş emri/tarih + pozisyon). Sonuçta güven yüzdesi,
uygulanan senaryolar, adım adım zincir ve alternatif adaylar gösterilir. İlk iki aday birbirine çok yakınsa
"belirsiz" uyarısı verilir. Tüm eşikler `config.yaml > fa` altından ayarlanabilir.

**Sonuç kodları:** `FA_ATA` (FA şu parçaya atanmalı) · `FA_SIL` (FA silinmeli) · `PNSN_CHG` (FA'e yeni PN/SN) ·
`DEPODA` (döngü sonundaki parça depoda) · `FA_DOGRU` (parça takılı değil, FA zaten doğru) · `BULUNAMADI` · `LIMIT`.

Excel raporu (`Raporlar/`): **Özet** (renkli), **Döngü Zinciri**, **Alternatifler**, **Açıklama** sayfaları.

## Veri kaynağı

- **Canlı Qlik Sense:** Playwright ile giriş yapılır; her sorgu Qlik'in uygulama entegrasyon URL'i
  (`/sense/app/<app>/sheet/<sheet>/state/analysis/options/clearselections/select/<alan>/<değer>`) ile seçimli açılır,
  tablo sağ tık > İndir > Veri ile Excel'e alınır. Oturum boyunca aynı sorgu tekrar çekilmez. Şifre sadece bellekte tutulur.
- **Excel modu:** Qlik'ten elle indirilen hareket export'ları seçilir. Kolon başlıkları (PN, P/N, Part Number, Parça No,
  Tarih, Hareket, A/C, Pozisyon, NHA PN…) otomatik tanınır; tanınmayan varsa `config.yaml > kolonlar` ile eşlenir.

Gereken hareket kolonları: tarih, hareket tipi, PN, SN (zorunlu); GRB, açıklama, PN grubu, uçak, pozisyon, NHA PN/SN,
iş emri, yeni PN/SN (ne kadar çok varsa o kadar isabetli).

### Kalibrasyon (ilk kurulumda bir kez)
Portal sayfa yapısı şirkete özel olduğu için seçiciler ayardan değiştirilebilir. **Ayarlar > Kalibrasyon** giriş yapılmış
tarayıcıyı Playwright Inspector ile açar; tablo nesnesinin `tid` değeri ve menü etiketleri buradan bulunup
`config.yaml > qlik` bölümüne yazılır.

## Proje yapısı
```
portal_giris.py          giriş noktası (BASLAT.bat çalıştırır)
tets/gui.py              3 sekmeli arayüz
tets/fa/engine.py        FA döngü motoru (senaryolar)
tets/fa/report.py        Excel rapor
tets/fa/runner.py        tekli/toplu çalıştırma
tets/sources/            Excel ve canlı Qlik veri kaynakları
tets/normalize.py        PN / pozisyon / hareket tipi normalizasyonu, benzerlik
tets/envanter.py         envanter otomasyonu
tests/                   senaryo testleri + örnek veri üretici (pytest)
```

Logo için `assets/logo.png` koymanız yeterli; varsa başlıkta gösterilir.
