# Logo Tiger 3 Enterprise — e-Fatura Transfer Süreci: Dersler ve Doğrular

**Konum**: RaporlamaPortali → Pages/EFatura.razor → "Logo'ya Aktar"
**Hedef**: Gelen e-Faturayı UBL XML'den parse edip Logo'da Satın Alma Faturası veya Alınan Hizmet Faturası olarak post etmek
**Logo sürümü**: Tiger 3 Enterprise v3.00.00.01 (Afyon Şeker — DOĞUŞ ÇAY GIDA, firma 211)
**Yazılım yığını**: .NET 8 Blazor Server + MudBlazor, COM late-binding (`UnityObjects.UnityApplication`), STA thread
**Belge tarihi**: 15 Mayıs 2026 — V17 itibariyle çalışan durum
**Kapsam**: Mayıs 2026 boyunca yaşanan keşif sürecinde nelerin doğru çıktığı, nerede yanıldık ve neden — gelecekte aynı tuzaklara düşmemek için.

> Bu belge yalnızca "ne yaptığımızı" değil, **nerede yanıldığımızı ve neden** belgelemek için yazıldı.
> Her bölümde önce **DOĞRU**, sonra **YANLIŞ**/yanıltıcı yollar.

---

## 0. Genel Akış

```
e-Fatura listesi (Pages/EFatura.razor)
     └─→ kullanıcı checkbox ile seçim + tip (Satın Alma / Alınan Hizmet)
            └─→ LogoyaAktar() → EFaturaLogoAktarimDialog
                   └─→ LogoAktarimService.HazirlaAsync (state hazırlığı)
                          ├─ EFaturaService.DetayGetirAsync (UBL parse)
                          ├─ LogoCariLookupService (VKN → cari adayları)
                          ├─ KDV/Tevkifat hesap kodu mapping (LogoHesapKoduMaps)
                          └─ Uyarılar listesi (eksik cari, eksik GL_CODE, vs.)
                   └─→ Dialog: kullanıcı eksikleri tamamlar (Master Kart, Birim, MasrafMerkezi, vs.)
                          └─→ "Logo'ya Aktar" → LogoAktarimService.AktarAsync
                                 └─→ LogoFaturaModel'e dönüşüm
                                 └─→ LogoUnityComService.AktarAsync
                                        └─→ STA thread başlat
                                        └─→ Login → her fatura için PostSingleInvoice
                                        └─→ COM: NewDataObject(18) → field set → Post()
                                 └─→ Sonuç UI'da (başarılı/hatalı + Logo NUMBER/REF)
```

---

## 1. DataObjectType — Hangi sayı doSlsInvoice, hangisi doPurchInvoice?

### DOĞRU (kesinleştirilmiş — KeyNet PMHS.dll v16.02.2026 enum decompile)

| Type | Adı | Açıklama |
|------|-----|----------|
| **18** | `doPurchInvoice` | **Satın Alma Faturası** (bu proje) |
| **19** | `doSalesInvoice` | Satış Faturası (SabnetFaturaTransfer ayrı projede) |
| 30 | `doAccountsRP` | Cari Hesap Kartı (CLCARD) |
| 23 | `doBankAccount` | Banka Hesabı |

### YANLIŞ yollar

| Versiyon | Sandığımız | Gerçek | Sebep |
|---|---|---|---|
| V7 (eski) | `doPurchInvoice = 23` | 23 = BANKACC | İlk geliştirme tahmini, field probe'la CARDTYPE/IBAN görünce yanlış anlaşıldı |
| V8 | Discovery probe ile 18 ve 19'un INVOICE tablosunu döndürdüğü görüldü, ama hangisinin Satın hangisinin Satış olduğu belirsiz kaldı | — | Sadece `TableName` ve field count'a bakmak yetersiz |
| V9–V10 (eski memory) | `doPurchInvoice = 19`, `doSalesInvoice = 18` (tam tersi) | Aslında 18 Purch, 19 Sales | KeyNet enum'unu okumadan empirik tahminle yazılmıştı |
| **V10 (14 May)** | KeyNet PMHS.dll decompile'ından kesinleştirildi → 18 ve 19 swap edildi | ✓ Doğru | SabnetFaturaTransfer'da 1159 hatası bu swap ile çözüldü |

**Ders**: Aynı tabloya yazan iki DataObjectType arasında ayrım için **field listesine bakmak yeterli değil** — DLL enum'undan kesin değeri al. Logo COM aynı INVOICE tablosuna iki farklı tip ile erişime izin verir ama post sırasında **modül kontrolü** yapar (TYPE=1/4 ile DataObjectType=18 uyumlu; TYPE=8 ile DataObjectType=19).

**Hata belirtisi**: "(1159) Fiş tip bilgisi fiş modülüne uygun değil"

---

## 2. Cari Karta GL_CODE Zorunluluğu

### DOĞRU
Fatura aktarımdan önce **cari kartında "Hesap Kodu" (GL_CODE) tanımlı olmalı**. RaporlamaPortali bunu dialog'da kontrol ediyor (`SeciliCari.GlKodVar`). Yoksa Logo aktarımı reddediyor.

### YANLIŞ yollar

| Sembol | Gerçek | Etki |
|---|---|---|
| `320.07.0000` | TELEFONLAR (Türk Telekom, Turkcell) hesabı | Müstahsil cariye yanlışlıkla atanmıştı (SabnetFaturaTransfer eski VB.NET kaynağında) → Logo "513: 320.07.0000 kodlu muhasebe hesabı bulunamadı" + "32558: alt hesaplar seçilmelidir" |
| `320.99.06.xxxx` veya `320.99.07.0000` | Doğru müstahsil cari hesabı | Bireysel cari için kullanıcının doğrulayacağı doğru kod |

**Hata belirtisi (zincirleme)**: Cari Post başarısızsa → ARP_CODE Logo'da yok → Fatura Post **"(1159) fiş tip uygun değil"** alıyor. Yani 1159 hatası **fatura tarafından değil**, eksik cari tarafından gelebilir.

**Ders**: 1159 hatası genelliği yüksek — kök sebep çoğu zaman fatura field'larında değil, **upstream'de eksik bir referans** (cari, hesap, malzeme kartı).

---

## 3. Header Zorunlu Field'ları

### DOĞRU — Logo Post için header'da olmazsa olmaz

```csharp
SetHeaderField(inv, "TYPE",         m.Type);        // 1=Satın Alma, 4=Alınan Hizmet
SetHeaderField(inv, "NUMBER",       m.FaturaNo);    // = e-Fatura no (FICHENO)
SetHeaderField(inv, "DATE",         m.Tarih);
SetHeaderField(inv, "DOC_NUMBER",   m.BelgeNo);     // = e-Fatura no (DOCODE)
SetHeaderField(inv, "DOC_DATE",     m.BelgeTarihi);
SetHeaderField(inv, "ARP_CODE",     m.CariKod);
SetHeaderField(inv, "GL_CODE",      m.CariGlKod);   // cari hesap kodu (320.99.x)
SetHeaderField(inv, "FACTORY",      m.Fabrika);     // 17 vs.
SetHeaderField(inv, "SOURCE_WH",    m.Ambar);       // 147
SetHeaderField(inv, "POST_FLAGS",   247);           // sabit
SetHeaderField(inv, "TOTAL_DISCOUNTED", ...);       // matrah
SetHeaderField(inv, "TOTAL_GROSS",      ...);       // matrah
SetHeaderField(inv, "TOTAL_NET",        ...);       // genel toplam
SetHeaderField(inv, "TC_NET",           ...);       // TL net (TL fatura için aynı)
SetHeaderField(inv, "ACCOUNTED_CNT", 1);            // DB'de çalışan faturalarda 1
SetHeaderField(inv, "GRPCODE",       1);            // Satın Alma istisnalı için 1
SetHeaderField(inv, "AFFECT_RISK",   1);
SetHeaderField(inv, "ESTATUS",       12);           // Kabul edildi
SetHeaderField(inv, "EBOOK_DOCTYPE", 99);           // e-fatura standart
SetHeaderField(inv, "EXIMVAT",       0);
SetHeaderField(inv, "EDURATION_TYPE",0);
SetHeaderField(inv, "PROFILE_ID",    m.ProfileId);  // 1=TEMEL, 2=TICARI
SetHeaderField(inv, "EINVOICE",      1);            // e-Fatura bayrağı (INT, DateTime DEĞİL!)
SetHeaderField(inv, "PROFILE_ID",    m.ProfileId);
// Raporlama Dövizi (USD)
SetHeaderField(inv, "EDTCURR_GLOBAL_CODE", "USD");
SetHeaderField(inv, "CURRSEL_TOTALS",      1);
SetHeaderField(inv, "RC_XRATE",            m.RaporlamaKuru);
SetHeaderField(inv, "REPORTRATE",          m.RaporlamaKuru);
SetHeaderField(inv, "RC_NET",              m.ToplamNet / m.RaporlamaKuru);
SetHeaderField(inv, "REPORTNET",           m.ToplamNet / m.RaporlamaKuru);
// V16: TIME alanları — auto-DISPATCH validation popup'ını engellemek için
SetHeaderField(inv, "TIME",      LogoTimeEncode(12, 0, 0));
SetHeaderField(inv, "DOC_TIME",  LogoTimeEncode(12, 0, 0));
SetHeaderField(inv, "SHIP_TIME", LogoTimeEncode(12, 1, 0));  // > TIME
```

### YANLIŞ — eksik kalan veya yanlış set edilen field'lar (geçmişte 1159 verdi)

| Field | Default | Yanlış senaryo | Doğru |
|---|---|---|---|
| `GRPCODE` | 0 | Satın Alma istisnalı için 0 → 1159 | 1 |
| `ACCOUNTED_CNT` | 0 | İlk Post için 0 → 1159 | 1 |
| `EBOOK_DOCTYPE` | 0 | e-Fatura kontrolünde yetersiz | 99 |
| `EINVOICE` | 0 | DateTime gönderilirse COM sessizce kabul etmiyor → header EINVOICETYP'le çelişiyor | `1` (int) |
| `EDTCURR_GLOBAL_CODE` | "TL" yazılırsa | Raporlama Dövizi boş kalıyor, "Dövizli Tutar" sütunu boş | Her zaman "USD" |
| `EINVOICE_TYPE` (header) | Set edilirse 2/4 | **V13'e kadar gerekli sandık, V15'te kaldırdık — Logo COM bunu DISPATCH'e cascade ediyor** (aşağıda detay) | Set ETME — line VATEXCEPT_CODE/CANDEDUCT Logo'ya istisna/tevkifat'ı yeterli söylüyor |
| `TOTAL_SERVICES` | Yok | Satın Alma'da set edilirse "fiş tipi modüle uygun değil" | Sadece **Type=4 (Hizmet)** için set et |

**Ders**: 1159 hatası bir kombinasyon hatası. Header'da bir alan eksikse veya yanlış kombinasyon varsa atıyor. Memory'deki "1159 = EINVOICE_TYPE eksik" diye eski not yanıltıcı; aslında **GRPCODE/ACCOUNTED_CNT/EBOOK_DOCTYPE/EINVOICE/ESTATUS** birlikte gerekiyordu.

---

## 4. Line (TRANSACTIONS) Zorunlu Field'ları

### DOĞRU
```csharp
SetField(line, "TYPE",        s.SatirTipi);   // Satın Alma → 0, Hizmet → 4 (NOT 1 = iade!)
SetField(line, "MASTER_CODE", s.MasterCode);
SetField(line, "SOURCEINDEX", m.Ambar);
SetField(line, "QUANTITY",    (double)s.Miktar);
SetField(line, "PRICE",       (double)s.BirimFiyat);
SetField(line, "TOTAL",       (double)s.ToplamNet);
SetField(line, "TOTAL_NET",   (double)s.ToplamNet);
SetField(line, "VAT_BASE",    (double)s.ToplamNet);
SetField(line, "UNIT_CODE",   s.Birim);
SetField(line, "UNIT_CONV1",  1);
SetField(line, "UNIT_CONV2",  1);
SetField(line, "VAT_RATE",    (double)s.KdvOran);
SetField(line, "VAT_AMOUNT",  (double)s.KdvTutar);
SetField(line, "BILLED",      1);
SetField(line, "FACTORY",     m.Fabrika);
SetField(line, "AFFECT_RISK", 1);
SetField(line, "EDT_CURR",    m.EdtCurr);
SetField(line, "EDTCURR_GLOBAL_CODE", "USD");   // Raporlama
SetField(line, "GL_CODE1",    malHesabi);       // Mal/Hizmet hesabı (150/153/740...)
SetField(line, "GL_CODE2",    kdvHesabi);       // İndirilecek KDV (191.01.x)
// İstisnalı (KDV=0)
SetField(line, "VATEXCEPT_CODE",   istisnaKodu);   // 327 vs.
SetField(line, "VATEXCEPT_REASON", istisnaSebep);
// Tevkifatlı
SetField(line, "CANDEDUCT",        1);
SetField(line, "DEDUCTION_PART1",  pay);
SetField(line, "DEDUCTION_PART2",  payda);
SetField(line, "DEDUCTION_TOT",    tevkifatTutar);
SetField(line, "DEDUCT_CODE",      tevkifatKodu);  // 3xx (alıcı tarafı, 6xx değil!)
SetField(line, "GL_CODE3",         tevkifatliKdvHesabi);  // 192.02.x
SetField(line, "GL_CODE4",         sorumluKdvHesabi);     // 360.10.x
```

### YANLIŞ yollar

| Konu | Yanlış | Doğru | Sebep |
|---|---|---|---|
| Line TYPE | 1 (iade satırı) | 0 (mal alım) veya 4 (hizmet) | İade modülüne uygun olmadığı için (1159) |
| GL_CODE1 vs GL_CODE2 sıralaması | Mal=GL_CODE2, KDV=GL_CODE1 | **Mal=GL_CODE1, KDV=GL_CODE2** | XML referansından doğrulandı; tersi (1159) |
| `VATEXCEPT_CODE` (alt çizgisiz: `VATEXCEPTCODE`) | "Bu isimde bir alan bulunamadı" | Alt çizgili: `VATEXCEPT_CODE` | API katmanı + DB tutarsızlığı |
| `TRCODE` / `IOCODE` line'da set | "Fiş tip uygun değil" | Line'da SET ETME | Logo bunları header'dan türetir |
| Tevkifat `DEDUCT_CODE`=6xx (satıcı) | Alıcı için yanlış | 6xx → 3xx çevir (`6` ile başlayan → `3` ile başlat) | Tevkifat kodu satıcı/alıcı yönüne göre değişir |

---

## 5. PAYMENT_LIST Zorunluluğu

### DOĞRU
```csharp
PAYMENT_LIST.AppendLine()
SetField(p, "DATE",             m.Tarih);
SetField(p, "PROCDATE",         m.Tarih);
SetField(p, "DISCOUNT_DUEDATE", m.Tarih);
SetField(p, "MODULENR", 4);            // Fatura modülü
SetField(p, "SIGN",     1);            // 1=Alacak (cari tarafından bakınca)
SetField(p, "TRCODE",   m.Type);       // header TYPE ile aynı (1 veya 4)
SetField(p, "TOTAL",    (double)m.ToplamNet);
SetField(p, "PAY_NO",   1);
SetField(p, "DISCTRDELLIST", 0);
```

Manuel fatura `DAYS=30` ekliyor (vade), bizim PESIN aktarımımız için gerekli değil.

---

## 6. DISPATCH Tuzağı — auto-create ve TIME validation

### Sorun
Logo Post() satın alma faturasında **otomatik internal DISPATCH** yaratıyor. Bu DISPATCH boş TIME (`00:00:00`) alanları içerdiği için Logo edit-validation:
**"e-İrsaliye için tüm bilgiler eksiksiz girilmelidir"** popup'ı + sürücü/plaka istiyor.

Manuel girilen fatura da DISPATCH yaratıyor ama TIME alanları doldurulduğu için tetiklenmiyor.

### DOĞRU çözüm (V16)
Logo TIME formatı: `HH × 16777216 + MM × 65536 + SS × 256`

```csharp
int LogoTimeEncode(int h, int m, int s) => h * 16777216 + m * 65536 + s * 256;
SetHeaderField(inv, "TIME",      LogoTimeEncode(12, 0, 0));   // 201326592
SetHeaderField(inv, "DOC_TIME",  LogoTimeEncode(12, 0, 0));
SetHeaderField(inv, "SHIP_TIME", LogoTimeEncode(12, 1, 0));   // 201392128
```

Logo Post sırasında bu değerleri auto-DISPATCH'e **cascade ediyor**. Kural: **SHIP_TIME > TIME** (Sevk Zamanı > Düzenleme Zamanı).

**Doğrulama**: Manuel fatura XML'inde `TIME=253769472 = 15*16777216 + 32*65536 + 55*256 = 15:32:55` — Bağlı İrsaliyeler ekranındaki Saat ile aynı.

### YANLIŞ yollar (V13–V15 — boşa çıkan hipotezler)

| V | Hipotez | Sonuç |
|---|---|---|
| V13 | DISPATCH override pre-Post (PAYMENT_LIST sonrası): `SetField(dispatchLine, "EINVOICE_TYPE", 0)` | DISPATCH collection Pre-Post boş → hiçbir şey yapmadı |
| V14 | DISPATCH override `Validation()` sonrası | Validation da DISPATCH yaratmıyor — diag log `DISPATCH-OVERRIDE count=0` → boş |
| V15 | Header'dan `EINVOICE_TYPE` set'ini tamamen kaldır (cascade'i suçladık) | Cascade kaldırıldı ama popup yine geldi → kök sebep cascade değildi |
| V16 | TIME/SHIP_TIME header'a set edip Logo'nun cascade etmesini kullan | ✓ Çalıştı |

**Ders**: Kullanıcının ipucu kritik oldu: "saatleri sadece girdiğimde sorun olmuyor sanki". Bu olmadan EINVOICE_TYPE cascade'i için çok zaman harcadık.

### Denenmemiş ama doğal olmayan yollar
- DISPATCH'i kullanıcı için silmek (DB read-only, sqluser yetkisi yok)
- DISPATCH'i POST sonrası COM ile editlemek (DataObject lifecycle olarak garantili değil)

---

## 7. Muhasebe Hesabı Eşlemesi — `LG_211_CRDACREF` (Card-Account Reference)

### DOĞRU (V17 — 15 May 2026)

Logo malzeme/hizmet kartının **"Muhasebe Hesapları" sekmesindeki kayıtlar `LG_{firma}_CRDACREF` tablosunda**:

```
CARDREF    → ITEMS.LOGICALREF (Satın Alma) veya SRVCARD.LOGICALREF (Hizmet)
TRCODE = 1 → malzeme/hizmet kartı bağlantıları
TYP = 1    → Alımlar Hesabı (GL_CODE1 için kullandığımız)
TYP = 3    → Satışlar Hesabı
TYP = 5    → Sarflar
TYP = 11   → Satınalma İadesi
TYP = 12   → Satış İadesi
TYP = 95/96 → Sayım Fazlası/Noksanı
ACCOUNTREF → EMUHACC.LOGICALREF (hesap)
CENTERREF  → opsiyonel masraf merkezi
```

DOGUSNDB'de 28,929 satır TYP=1 binding — yani hemen hemen tüm malzeme/hizmet kartları muhasebe hesabıyla eşli.

**Sorgu**:
```sql
SELECT TOP 1 e.CODE
FROM LG_211_CRDACREF c WITH(NOLOCK)
INNER JOIN LG_211_ITEMS k WITH(NOLOCK) ON k.LOGICALREF = c.CARDREF
INNER JOIN LG_211_EMUHACC e WITH(NOLOCK) ON e.LOGICALREF = c.ACCOUNTREF
WHERE k.CODE = @masterCode AND c.TRCODE = 1 AND c.TYP = 1
```

Hizmet için aynı sorguda `ITEMS` yerine `SRVCARD` kullan.

### YANLIŞ tablolar (önce gidilen yollar)

| Tablo | Neden başvurduk | Gerçek |
|---|---|---|
| `LG_211_ITEMS` | Direkt kart üzerinde olabilir | GL_CODE / ACCOUNTREF kolonu yok |
| `LG_211_EMUHACC` | Kullanıcının paylaştığı şema (`CARDREF/CARDTYPE/TYP/ACCOUNTREF`) bu tablo dedi | Bu Logo Tiger 3 sürümünde sadece **hesap kart tanımı** — paylaşılan şema farklı sürüm (muhtemelen Logo Wings) |
| `LG_211_ACCCODES` | Sistem İşletmeni "Muhasebe Bağlantı Kodları" prefix kuralları | 253 kayıt var; A.G prefix için kayıt YOK — kullanıcı bu yolla tanımlamamış; **kart-spesifik binding'i tutmuyor** |
| `LG_211_ACCOUNTTEMPLATES` | "Template" isminden | Boş — bu Logo'da kullanılmamış |
| `LG_211_EMUHACCSUBACCASGN` | ACCOUNTREF kolonu var | Sadece ana-alt hesap mapping (ilgisiz) |
| `LG_211_POACCREF` | "Purchase Order Account REF" | Sadece sipariş seviyesinde |

**Anahtar buluş**: ITEM 66022 + Hesap 76627'yi **aynı satırda** içeren tabloyu bulmak. INFORMATION_SCHEMA üzerinden kolon-bazlı arama ile `LG_211_CRDACREF` ortaya çıktı.

### Üç-katmanlı öncelik (V17)
[LogoAktarimService.cs](Services/Logo/LogoAktarimService.cs) GL_CODE1 için:
1. **Kullanıcı dialog'da manuel girdi** → onu kullan
2. **CRDACREF** (`LogoAccCodesLookupService.AlimHesabiAsync`) → kart-spesifik kayıt
3. **`LogoHesapKoduMaps` prefix mapping** (A.G.01.01→150.18.001 vs.) → CRDACREF boşsa fallback

---

## 8. Cari, Birim, Fabrika, Ambar — Eşleştirme Hizmetleri

### DOĞRU
- **Cari**: `LogoCariLookupService.BulAsync(vkn)` — VKN/TCKN ile arar, **çoklu cari** dönebilir (örn. aynı VKN'ye farklı cariler), dialog'da dropdown
- **Master kart**: `LogoMasterKartLookupService.MalzemeAraAsync()` (Satın Alma) ve `HizmetAraAsync()` (Hizmet) — autocomplete
- **Birim seti**: `LogoBirimSetiLookupService.KartBirimleriAsync(masterCode)` — kart seçildikten sonra Logo'daki birim seti yüklenir, UBL'deki birim ile en yakın eşleşme default seçilir
- **Fabrika/Ambar/MasrafMerkezi**: `LogoOrgLookupService` — Logo'dan organizasyon yapısı, dropdown

### YANLIŞ varsayım
- "Çoklu cari olmaz, tek dönecek" — yanlış, dialog mutlaka dropdown göstermeli

---

## 9. Raporlama Dövizi (Dövizli Tutar)

### DOĞRU
Afyon Şeker firma 211 USD raporlaması yapıyor. TL faturada bile:
- `EDTCURR_GLOBAL_CODE = "USD"` (header + line)
- `RC_XRATE = REPORTRATE = USD/TL TCMB satış kuru` (fatura tarihi)
- `RC_NET = REPORTNET = ToplamNet / RaporlamaKuru`

Kur kaynağı: `LG_211_EXCHANGE.RATES2` (CRTYPE=1=USD, EDATE<=fatura tarihi, ORDER BY EDATE DESC TOP 1).

### YANLIŞ
- TL faturada `EDTCURR_GLOBAL_CODE = "TL"` set ettik → Logo "Dövizli Tutar" sütununda boş gösteriyor

---

## 10. Fiş No (NUMBER) — e-Fatura numarası

### DOĞRU
- `FaturaNo` (NUMBER → FICHENO) ve `BelgeNo` (DOC_NUMBER → DOCODE) **ikisi de** e-Fatura numarası olsun
- Manuel fatura DB'de NUMBER = e-Fatura no, DOCODE = boş (Logo manuel formda ikincisini boş bırakıyor)
- Bizim için ikisini de dolu yapmak güvenli — kullanıcının istediği davranış

### YANLIŞ
- NUMBER'ı boş bırakırsak Logo `20260508000767` gibi tarih-bazlı otomatik kod üretir, kullanıcı kafa karışıyor — kağıt fatura sanıyor

---

## 11. E-Fatura Tipi (`EINVOICE` ve `EINVOICETYP`)

### DOĞRU
- `EINVOICE = 1` (INT, smallint flag — DateTime DEĞİL!)
- `EINVOICETYP` — header'da set ETME (V15'ten itibaren). DB'de manuel fatura için 2 (istisna) görünebilir ama bu Logo'nun internal davranışı. Bizim COM çağrımız sırasında set edersek auto-DISPATCH'e cascade ediyor.

### YANLIŞ
- `EINVOICE = DateTime` → COM sessizce kabul etmiyor, DB'de 0 kalıyor → header EINVOICETYP=2 ile çelişiyor → 1159
- V9–V14 boyunca `EINVOICE_TYPE` cascade'i 1159'a çare sandık ama aslında diğer alanlar (GRPCODE, ACCOUNTED_CNT, EBOOK_DOCTYPE) eklenince zaten gerek kalmamıştı

---

## 12. COM Late-Binding — .NET 8 Tuzakları

### DOĞRU
- `Type.GetTypeFromProgID("UnityObjects.UnityApplication")` + `Activator.CreateInstance`
- Tüm dispatch çağrıları `Type.InvokeMember` veya direkt IDispatch::Invoke ile
- STA thread (Blazor Server thread'leri MTA — ayrı STA thread spawn et)
- csproj: `<BuiltInComInteropSupport>true</BuiltInComInteropSupport>`, `<EnableComHosting>false</EnableComHosting>`

### YANLIŞ
- `dynamic` + RuntimeBinder → `"Cannot perform runtime binding on a null reference"` patlar
- `Marshal.GetObjectForNativeVariant` VT_DISPATCH için **null döner** — manuel VARIANT byte extraction gerekli (`IDispCall`, `ProbeRawVariant`)
- MTA thread'den COM çağrısı → CO_E_NOTINITIALIZED veya benzeri

### Bonus: ProgID kayıtsızsa
- `CO_E_CLASSSTRING (0x800401F3)` — "Invalid class string" / Logo Tiger yüklü değil
- **Yerel dev PC'de aktarım test edilemez**, sadece Logo'nun yüklü olduğu makinede

---

## 13. DISPATCH Lifecycle — Logo COM'un sınırı

Logo COM `Post()` sırasında auto-DISPATCH yaratır. Bu DISPATCH'e **pre-Post erişim imkansız**:
- `inv.DataFields.FieldByName("DISPATCHES").Lines.Count` Pre-Post = **0**
- `Validation()` çağrısı dahi DISPATCH'i pre-create etmiyor (V14 diag log doğruladı)

Bu yüzden:
- DISPATCH'in field'ları (TIME, EINVOICE_TYPE) → **header'a set et**, Logo Post sırasında cascade eder
- Veya **Post sonrası** DataObject'ı `Read(REF)` ile yeniden aç, modify, tekrar Post — (denemedik, riski yüksek)
- Veya **DB UPDATE** — `sqluser` read-only olduğu için yapılamadı

---

## 14. Tanılama (Diag) Stratejisi

### DOĞRU
- `[BUILD-V{N}]` etiketi her hatada yazılıyor → çalışan exe'nin doğru sürüm olduğunu doğrula
- `[V14] DISPATCH-OVERRIDE count={N}` gibi adımsal diag log
- Logo `ErrorCode`, `ErrorDesc`, `DBErrorDesc`, `ValidateErrors` toplu yakalama
- Header field enum'ı hata anında log dosyasına yazılıyor (`EnumerateFieldNamesAndLog`)
- Diag log: `C:\RaporlamaPortaliData\logo_unity_diag.txt` (paylaşılan AppData yolu, exe silinse de kalır)

### YANLIŞ
- Sadece UI'ya hata mesajı göstermek → uzak makinede debug için yetersiz, log dosyasına da yaz
- "Errors silently swallowed" → cari Post'unda yapmıştık, fatura silently fail oldu, sonra fatal exception fırlatma ile düzeltildi

---

## 15. Yerel/Uzak Çalışma Ayrımı

### DOĞRU
- Local geliştirme PC = `localhost:5050`, Logo COM bağlanmaz (ProgID yok)
- Uzak Logo bilgisayar = exe deploy edilmiş klasör (`C:\RaporPublish\_new\` veya `\RaporlamaPortali\`)
- DB sorguları (cari/master kart/kur arama) local'den DOGUSNDB'ye gidebiliyor (192.168.0.51\DOGUSLGSRV2) — bu işler
- Aktarım yalnızca Logo'nun yüklü olduğu uzak makinede mümkün

### YANLIŞ
- Local'den aktarım denedik → CO_E_CLASSSTRING → "kod bozuldu" sandık → değil, local PC'de Logo yok

---

## 16. Publish & Deploy

### DOĞRU
- `dotnet publish -c Release -r win-x64 --self-contained -o <hedef>` → 228 MB tek exe + WhatsApp/wwwroot/locale destek klasörleri
- csproj: `<PublishSingleFile>true</PublishSingleFile>`, `<IncludeNativeLibrariesForSelfExtract>true</IncludeNativeLibrariesForSelfExtract>`
- Hedef makinede .NET runtime YOK → self-contained zorunlu

### YANLIŞ tuzaklar
- Proje köküne yanlışlıkla kopyalanan `RaporPublish_new` klasörü → csproj exclude etmiyor → publish çıktısına gerçek olmayan DLL'ler kopyalandı (kullanıcıyla çok zaman kaybettirdi)
- `Remove-Item` PowerShell'de bazen blocked → `rm -rf` Bash üzerinden çalışıyor

---

## 17. Önemli Tablolar Hızlı Referansı

| Tablo | Görev |
|---|---|
| `LG_211_01_INVOICE` | Fatura header (FICHENO, DOCODE, EINVOICE, EINVOICETYP, REPORTRATE, REPORTNET) |
| `LG_211_01_STLINE` | Fatura kalemleri (transactions) |
| `LG_211_01_STFICHE` | DISPATCH (auto-yaratılan internal) — EDESPATCH/EDESPSTATUS bilgisi |
| `LG_211_CLCARD` | Cariler (kod 320.x, ARP_CODE referans) |
| `LG_211_ITEMS` | Malzeme kartları |
| `LG_211_SRVCARD` | Hizmet kartları |
| **`LG_211_CRDACREF`** | **Malzeme/Hizmet kartı muhasebe hesap eşlemeleri** ⭐ |
| `LG_211_EMUHACC` | Muhasebe hesap kart tanımları (kod → açıklama) |
| `LG_211_ACCCODES` | Sistem İşletmeni "Muhasebe Bağlantı Kodları" prefix kuralları (MODNR=1 mal alımı, 3 satış) |
| `LG_211_EXCHANGE` | Döviz kurları (CRTYPE=1=USD, RATES2=TCMB satış) |
| `LG_211_01_PAYTRANS` | Vade taksitleri |
| `LG_211_01_CLFLINE` | Cari hareket (ekstre) |

---

## 18. Sürüm Geçmişi (BUILD tag)

| Build | Tarih | Ana değişiklik |
|---|---|---|
| V7–V9 | Mayıs öncesi | DataObjectType keşfi, field enumeration, ilk header alanları |
| V10 | 14 May | `DO_PURCH_INVOICE = 19 → 18` (KeyNet enum); 1159 çözüldü |
| V11 | 14 May | Fiş No = e-Fatura no (NUMBER set), Raporlama Dövizi (RC_XRATE/REPORTRATE), EDTCURR_GLOBAL_CODE="USD" |
| V12 | 14 May | Prefix mapping (A.G.01.01.→150.18.001, A.G.02.→150.18.002) |
| V13 | 15 May | DISPATCH pre-Post override denemesi (boşa) — payment_list sonrası |
| V14 | 15 May | DISPATCH override Validation sonrasına alındı — yine boşa (count=0) |
| V15 | 15 May | Header'dan `EINVOICE_TYPE` set'i kaldırıldı (yanlış suçlama, ama zarar yok — istisna line'da) |
| **V16** | 15 May | **TIME/DOC_TIME/SHIP_TIME header → DISPATCH cascade → popup gitti** ✓ |
| **V17** | 15 May | **`LG_211_CRDACREF` üzerinden dinamik kart hesabı lookup'ı** ✓ |
| **V18** | 15 May | Dövizli fatura (EUR=20, TRCURR/TRRATE/TC_NET); EFatura Aktar menüsü + aktarım geçmişi tablosu; mükerrer kontrol + K-prefix; **CRDACREF.TRCODE hizmet için 3** (1 değil); Fiş No/Belge No ayrımı (K-prefix'te DOCODE orijinal kalır) ✓ |

---

## 19. Açık Konular / Gelecek İyileştirmeler

- **Hizmet için TYP belirleme**: Şu an Alınan Hizmet için de TYP=1 (Alımlar) çekiyoruz. Hizmet kartlarındaki "Sarflar" (TYP=5) veya "Diğer Giriş" (TYP=10) hesabı bazı senaryolarda daha doğru olabilir. Kullanıcının kullanım pattern'i gözlemlenmeli.
- **Vade desteği**: Manuel fatura'da `PAYMENT_LIST.DAYS=30` (30 günlük vade) var. Bizim aktarım PESIN. Cari kartından otomatik vade plan çekme (`LG_211_PAYPLANS`) iyileştirilebilir.
- **Tevkifat mapping genişletme**: Bazı tevkifat oranları henüz `LogoHesapKoduMaps.TevkifatSorumlulukKdv`'ye eklenmedi — kullanım çıktıkça ekle.
- **DB write yetkisi**: `sqluser` read-only. Logo destek bilgilendirilirse veya farklı kullanıcı sağlanırsa Post-Post DB UPDATE ile dispatch düzeltme imkanı doğar — şu an gerekli değil ama gelecekte aciden ihtiyaç olursa biliniyor.

---

## 20. Tek Cümle Özet

> Logo Tiger Unity COM ile satınalma e-faturası aktarımı için header'da **TYPE=1, GRPCODE=1, ACCOUNTED_CNT=1, EBOOK_DOCTYPE=99, EINVOICE=1, ESTATUS=12, TIME/SHIP_TIME** + `LG_211_CRDACREF.TYP=1`'den dinamik **GL_CODE1** + line'da **VATEXCEPT_CODE/REASON** ile fatura post edilir; auto-DISPATCH'e header alanları cascade eder, edit-validation popup'ları açılmaz.

---

*Sürüm: V17 — 15 Mayıs 2026. Sonraki güncelleme için BUILD tag'i ve bu dosyayı eş güncelle.*
