namespace RaporlamaPortali.Services.Logo;

// Logo Tiger Unity COM API'ye gönderilecek alış faturası modeli.
// Excel makrosundaki ProcessInvoice() ile aynı alanları temsil eder; sadece bizim
// kontrol etmemiz gereken alanlar burada — TIME, INTERNAL_REFERENCE, DATE_CREATED gibi
// alanları Logo'ya kendisi atatıyoruz (sabit veya boş bırakılıyor).
public class LogoFaturaModel
{
    // 4 = Alınan Hizmet Faturası, 1 = Satın Alma Faturası (mal alım)
    public int       Type        { get; set; } = 4;

    public string    FaturaNo    { get; set; } = "";            // NUMBER (boş = Logo otomatik)
    public DateTime  Tarih       { get; set; } = DateTime.Today;
    public string    BelgeNo     { get; set; } = "";            // DOC_NUMBER (e-fatura no)
    public DateTime  BelgeTarihi { get; set; } = DateTime.Today;// DOC_DATE

    // Cari (LG_211_CLCARD.CODE) ve cari'nin muhasebe hesap kodu (LG_211_EMUHACC.CODE)
    public string    CariKod     { get; set; } = "";            // ARP_CODE
    public string    CariGlKod   { get; set; } = "";            // GL_CODE — DB'den okunmalı

    public int       Fabrika     { get; set; } = 17;            // FACTORY
    public int       Ambar       { get; set; } = 147;           // SOURCE_WH
    public string    OdemeKodu   { get; set; } = "PESIN";       // PAYMENT_CODE
    public string    OzelKod     { get; set; } = "AFYON";       // TRADING_GRP

    // Döviz bilgisi (UBL XML'den)
    public string    Doviz       { get; set; } = "TL";          // İşlem dövizi kodu (TRCURR_GLOBAL_CODE)
    public decimal   DovizKuru   { get; set; } = 1m;
    public int       EdtCurr     { get; set; } = 0;             // 0 = TL, 20 = USD, 53 = EUR (Logo currency #)
    // Raporlama Dövizi (Afyon Şeker firma 211 → USD). TL fatura için bile Logo'nun "Dövizli Tutar"
    // kolonunda görünür. DB kolonu: REPORTRATE (RC_XRATE alanı XML/COM tarafında).
    // 0 ise Logo Raporlama Dövizi alanını boş bırakır.
    public decimal   RaporlamaKuru { get; set; }

    // Fatura toplamları (UBL XML'den hesaplanır)
    public decimal   ToplamMatrah     { get; set; }   // TOTAL_DISCOUNTED & TOTAL_GROSS & TOTAL_SERVICES
    public decimal   ToplamKdv        { get; set; }   // TOTAL_VAT
    public decimal   ToplamNet        { get; set; }   // TOTAL_NET & TC_NET (KDV dahil ödenecek)
    public decimal   ToplamTevkifat   { get; set; }   // hesaba katılır; alanı yok ama satırlardan toplanır

    public bool      Tevkifatli       { get; set; }   // PREACCLINES gerekiyor mu

    // E-Fatura bilgileri
    public int       ProfileId        { get; set; } = 2;          // PROFILE_ID (2 = TICARIFATURA, 1 = TEMELFATURA)
    public DateTime? EFaturaTarihi    { get; set; }                // EINVOICE

    // Satır listesi
    public List<LogoFaturaSatirModel> Satirlar { get; set; } = new();
}

// Tek satır (Logo TRANSACTIONS satırı).
public class LogoFaturaSatirModel
{
    public string    MasterCode    { get; set; } = "";   // MASTER_CODE — hizmet/malzeme kodu (kullanıcı seçer)
    // Satır tipi: hizmet faturasında 4, satın alma faturasında 0 (Logo varsayılan)
    public int       SatirTipi     { get; set; } = 4;

    public decimal   Miktar        { get; set; }
    public string    Birim         { get; set; } = "ADET";
    public decimal   BirimFiyat    { get; set; }
    public decimal   ToplamNet     { get; set; }         // TOTAL & TOTAL_NET (KDVsiz satır toplamı)
    public decimal   KdvOran       { get; set; }         // VAT_RATE (yüzde — 10 / 20 / 1)
    public decimal   KdvTutar      { get; set; }         // VAT_AMOUNT

    // Masraf merkezi (LG_211_EMCENTER.CODE) — varsayılan "7.04"
    public string    MasrafMerkezi { get; set; } = "7.04";

    // Muhasebe hesap kodları (yüzde göre seçilir, EFaturaService Kalem.KdvOran kullanılır)
    public string    HizmetMalHesabi    { get; set; } = "";   // GL_CODE1 — Hizmet/Malzeme/Stok/Gider (150/153/157/740/760.xx)
    public string    KdvHesabi          { get; set; } = "";   // GL_CODE2 — İndirilecek KDV (191.01.xx)
    public string    TevkifatKdvHesabi  { get; set; } = "";   // GL_CODE3 — Tevkifatlı İndirilecek KDV (192.02.xx) — sadece tevkifatlı
    public string    SorumluKdvHesabi   { get; set; } = "";   // GL_CODE4 — Sorumlu Sıfatı ile Ödenecek KDV (360.10.xx) — sadece tevkifatlı

    // İstisna alanları — KDV=0 ise Logo VATEXCEPT_CODE + VATEXCEPT_REASON dolu bekliyor (1159 hatasını önlemek için).
    public string    IstisnaKodu   { get; set; } = "";   // VATEXCEPT_CODE
    public string    IstisnaSebep  { get; set; } = "";   // VATEXCEPT_REASON

    // Tevkifat alanları (CANDEDUCT=1 + DEDUCTION_PARTx + DEDUCT_CODE)
    public bool      Tevkifatli    { get; set; }
    public decimal   TevkifatPay   { get; set; }         // DEDUCTION_PART1
    public decimal   TevkifatPayda { get; set; }         // DEDUCTION_PART2
    public decimal   TevkifatTutar { get; set; }         // DEDUCTION_TOT
    public string    TevkifatKodu  { get; set; } = "";   // DEDUCT_CODE (alıcı tarafı — 3xx)
}

// Toplu aktarım sonucu — her fatura için ayrı durum.
public class LogoAktarimSonuc
{
    public bool      LoginBasarili  { get; set; }
    public string    LoginHata      { get; set; } = "";
    public List<LogoFaturaAktarimSatir> Faturalar { get; set; } = new();

    public int       BasariliSayisi => Faturalar.Count(f => f.Basarili);
    public int       HataliSayisi   => Faturalar.Count(f => !f.Basarili);
}

public class LogoFaturaAktarimSatir
{
    public long      EFaturaId   { get; set; }   // EFatura.Id
    public string    FaturaNo    { get; set; } = "";
    public bool      Basarili    { get; set; }
    public string    LogoNo      { get; set; } = "";  // Logo'nun atadığı NUMBER (Post sonrası)
    public long      LogoRef     { get; set; }        // Logo INTERNAL_REFERENCE
    public string    Hata        { get; set; } = "";  // ErrorCode + ErrorDesc + DBErrorDesc
    public List<string> ValidateErrors { get; set; } = new();
}

public class LogoUnityCredentials
{
    public string  Kullanici { get; set; } = "";
    public string  Sifre     { get; set; } = "";
    public int     FirmaNo   { get; set; } = 211;
    public int     DonemNo   { get; set; } = 1;
}

// =====================================================================
// TOPTAN SATIŞ FATURASI (doSalesInvoice, DataObjectType=19)
// Referans: F:\30.12.2025 yedek\hakan\Genel Projeler\PMHS\PMHS\Raporlar\AyniAvansEntegrasyonu.vb
// KeyNet PMHS.dll enum: 18=Satın Alma, 19=Satış. SabnetFaturaTransfer 14 May'de doğruladı.
// =====================================================================

public class LogoSatisFaturaModel
{
    // Header
    public int       Tip          { get; set; } = 8;             // TYPE — XML'de 8, SabNet kodunda 9 → ilk denemede 8
    public string    FaturaNo     { get; set; } = "";            // NUMBER (örn. "B 1305260001")
    public DateTime  Tarih        { get; set; } = DateTime.Today;
    public string    CariKod      { get; set; } = "";            // ARP_CODE ("S" + TCKN)
    public string    CariUnvan    { get; set; } = "";            // doAccountsRP.TITLE
    public string    CariTcKimlik { get; set; } = "";            // doAccountsRP.TCKNO
    public string    CariEPosta   { get; set; } = "";            // doAccountsRP.E_MAIL
    public string    CariTelefon  { get; set; } = "";            // doAccountsRP.TELEPHONE1
    public string    TradingGrp   { get; set; } = "AFYON";       // TRADING_GRP
    public int       Fabrika      { get; set; } = 17;            // FACTORY
    public int       Ambar        { get; set; } = 147;           // SOURCE_WH
    public string    Isyeri       { get; set; } = "";            // DIVISION
    public string    Bolum        { get; set; } = "";            // DEPARTMENT
    public string    ProjeKodu    { get; set; } = "";            // PROJECT_CODE
    public int       PostFlags    { get; set; } = 247;
    public long      TimeKodu     { get; set; } = 138092067;     // Logo internal time sabit

    // Tek satır (her fatura tek malzeme — gübre satışı)
    public string    MasterCode   { get; set; } = "";            // MASTER_CODE (AvansStokKodu)
    public string    Birim        { get; set; } = "KG";          // UNIT_CODE
    public decimal   Miktar       { get; set; }                  // QUANTITY
    public decimal   BirimFiyat   { get; set; }                  // PRICE
    public decimal   KdvOran      { get; set; }                  // VAT_RATE
    public string    KdvDahilHaric{ get; set; } = "DAHIL";       // bilgi (PRICE ayarlamada)

    // İz sürme için kaynak satır kimliği (post sonrası SQL update için)
    public string    FormNo       { get; set; } = "";
    public string    HesapNo      { get; set; } = "";
}

public class LogoSatisFaturaResult
{
    public bool      Basarili    { get; set; }
    public string    FaturaNo    { get; set; } = "";
    public long      LogoRef     { get; set; }
    public string    Hata        { get; set; } = "";
    public List<string> ValidateErrors { get; set; } = new();
    public string    FormNo      { get; set; } = "";
    public string    HesapNo     { get; set; } = "";
}

public class LogoSatisAktarimSonuc
{
    public bool      LoginBasarili { get; set; }
    public string    LoginHata     { get; set; } = "";
    public List<LogoSatisFaturaResult> Faturalar { get; set; } = new();
    public int       BasariliSayisi => Faturalar.Count(f => f.Basarili);
    public int       HataliSayisi   => Faturalar.Count(f => !f.Basarili);
}
