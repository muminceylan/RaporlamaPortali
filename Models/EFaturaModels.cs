namespace RaporlamaPortali.Models;

public enum EFaturaYon { Gelen, Giden }

public class EFaturaFiltre
{
    public EFaturaYon Yon { get; set; } = EFaturaYon.Gelen;
    public DateTime Baslangic { get; set; } = DateTime.Today.AddDays(-30);
    public DateTime Bitis     { get; set; } = DateTime.Today;
    public string?  Tip       { get; set; }
    public string?  ProfileId { get; set; }
    public string?  KarsiUnvan { get; set; }
    public string?  KarsiVKN   { get; set; }
    public string?  FaturaNo   { get; set; }
    public int      MaxKayit   { get; set; } = 500;
    // GİB Sistem/Uygulama Yanıtı zarflarını listeden gizle (varsayılan: gizle)
    public bool     SistemYanitlariniGoster { get; set; } = false;
    // Logo cari kart kodu LIKE filtresi — örn. "320.06" yazılırsa "320.06%" ile başlayan
    // tüm cariler eşleşir. % işareti otomatik eklenir.
    public string?  LogoCariKodu { get; set; }

    // Çoklu fatura no filtresi — kullanıcı toplu fatura numarası listesi yapıştırır
    // (yeni satır / virgül / noktalı virgül / boşluk ile ayrılabilir). Dolu ise
    // diğer filtreler etkili olmaya devam eder ama sadece bu numaralardaki kayıtlar döner.
    public string?  FaturaNoListesiRaw { get; set; }

    public string[] FaturaNolariCoz()
    {
        if (string.IsNullOrWhiteSpace(FaturaNoListesiRaw)) return Array.Empty<string>();
        return FaturaNoListesiRaw
            .Split(new[] { '\n', '\r', ',', ';', '\t', ' ' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(s => s.Trim())
            .Where(s => s.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }
}

public class EFaturaListItem
{
    public long      Id              { get; set; }
    public DateTime  Tarih           { get; set; }
    public string    FaturaNo        { get; set; } = "";
    public string    Uuid            { get; set; } = "";
    public string    Yon             { get; set; } = "";
    public string    Tip             { get; set; } = "";
    public string    ProfileId       { get; set; } = "";
    public string    EnvelopeType    { get; set; } = "";
    public string    KarsiVKN        { get; set; } = "";
    public string    KarsiUnvan      { get; set; } = "";
    public string    Aciklama        { get; set; } = "";
    public string    DosyaAdi        { get; set; } = "";
    public long      DataSize        { get; set; }
    public string    Status          { get; set; } = "";
    // Logo Tiger LG_211_CLCARD lookup (VKN üzerinden)
    public string    LogoCariKod     { get; set; } = "";
    public string    LogoCariUnvan   { get; set; } = "";
    public string    LogoCariOzelKod { get; set; } = "";   // CLCARD.SPECODE (Özel Kod)
    // UBL XML'inden parse edilmiş fatura toplamları (listede gösterilir)
    public string    Doviz           { get; set; } = "";
    public decimal   Matrah          { get; set; }   // TaxExclusiveAmount
    public decimal   KdvTutar        { get; set; }   // sum of KDV (TaxTypeCode 0015)
    public decimal   ToplamTutar     { get; set; }   // PayableAmount (KDV dahil — varsa tevkifat sonrası ödenecek)
    public decimal   TevkifatTutar   { get; set; }   // WithholdingTaxTotal toplamı
    // Faturanın ilk kaleminin malzeme/ürün/hizmet adı (cac:InvoiceLine[1]/cac:Item/cbc:Name)
    // Name boşsa fallback olarak cbc:Note kullanılır
    public string    IlkKalem        { get; set; } = "";

    // Logo'ya işlendi mi kontrolü (sadece Gelen tarafı için doldurulur)
    // Kesin  : Aynı VKN'nin Logo carisinde FICHENO tam eşleşir
    // Modifiye: Aynı VKN, FICHENO'nun son 4 hanesi eşleşir + tutar eşit + tarih ±tolerans
    // Supheli : Aynı VKN, tutar eşit + tarih ±tolerans (numara hiç tutmadı)
    // Yok     : Logo'da bu faturaya işaret eden kayıt bulunamadı
    public LogoIslemDurumu LogoDurum     { get; set; } = LogoIslemDurumu.Bilinmiyor;
    public string          LogoFicheno   { get; set; } = "";  // eşleşen Logo fatura no
    public DateTime?       LogoTarih     { get; set; }        // Logo'daki işleme tarihi
    public decimal         LogoTutar     { get; set; }        // Logo GROSSTOTAL

    // Uygulama Yanıtı (POSTBOXENVELOPE) — RED yanıtı varsa fatura iptal/red sayılır,
    // Logo'ya işlenmemiş olması beklenir (uyarı listesinden düşmesi gerekir).
    public bool      RedEdildi    { get; set; }
    public DateTime? RedTarihi    { get; set; }
    public string    RedAciklama  { get; set; } = "";
}

public enum LogoIslemDurumu
{
    Bilinmiyor = 0,   // kontrol edilmedi
    Kesin      = 1,   // ✅ FICHENO tam eşleşti
    Modifiye   = 2,   // ✅ son 4 hane + tutar + tarih eşleşti
    Supheli    = 3,   // ⚠️ tutar + tarih eşleşti ama numara hiç tutmadı
    Yok        = 4    // ❌ hiç eşleşme yok
}

public class EFaturaKalem
{
    public int     Sira         { get; set; }
    public string  Aciklama     { get; set; } = "";   // cbc:Note (PDF'te görünen)
    public string  StokAdi      { get; set; } = "";   // cac:Item/cbc:Name
    public string  StokKodu     { get; set; } = "";   // SellersItemIdentification
    public decimal Miktar       { get; set; }
    public string  Birim        { get; set; } = "";
    public decimal BirimFiyat   { get; set; }
    public string  Doviz        { get; set; } = "";
    public decimal NetTutar     { get; set; }
    public decimal KdvOran      { get; set; }
    public decimal KdvTutar     { get; set; }
    public string  KdvAdi       { get; set; } = "";
    public string  IstisnaKodu  { get; set; } = "";
    public string  IstisnaSebep { get; set; } = "";
    public string  TevkifatKodu { get; set; } = "";
    public string  TevkifatAdi  { get; set; } = "";
    public decimal TevkifatOran { get; set; }
    public decimal TevkifatTutar{ get; set; }
}

public class EFaturaVergi
{
    public string  Adi          { get; set; } = "";
    public string  Kodu         { get; set; } = "";
    public decimal Oran         { get; set; }
    public decimal Matrah       { get; set; }
    public decimal Tutar        { get; set; }
    public string  IstisnaKodu  { get; set; } = "";
    public string  IstisnaSebep { get; set; } = "";
}

public class EFaturaDetay
{
    public long       Id              { get; set; }
    public string     FaturaNo        { get; set; } = "";
    public string     Uuid            { get; set; } = "";
    public DateTime?  Tarih           { get; set; }
    public string     Saat            { get; set; } = "";
    public string     Tip             { get; set; } = "";
    public string     ProfileId       { get; set; } = "";
    public string     Doviz           { get; set; } = "";
    public decimal    DovizKuru       { get; set; }

    public string     SaticiUnvan     { get; set; } = "";
    public string     SaticiVkn       { get; set; } = "";
    public string     SaticiVergiDairesi { get; set; } = "";
    public string     SaticiAdres     { get; set; } = "";
    public string     SaticiTelefon   { get; set; } = "";
    public string     SaticiEposta    { get; set; } = "";
    public string     SaticiWeb       { get; set; } = "";

    public string     AliciUnvan      { get; set; } = "";
    public string     AliciVkn        { get; set; } = "";
    public string     AliciVergiDairesi { get; set; } = "";
    public string     AliciAdres      { get; set; } = "";

    // Logo Tiger LG_211_CLCARD lookup (karşı taraf — bizim cari karttaki kod+ünvan)
    public string     LogoCariKod     { get; set; } = "";
    public string     LogoCariUnvan   { get; set; } = "";

    public List<EFaturaKalem>  Kalemler        { get; set; } = new();
    public List<EFaturaVergi>  VergiOzetleri   { get; set; } = new();
    public List<EFaturaVergi>  TevkifatOzetleri { get; set; } = new();
    public List<string>        FaturaNotlari   { get; set; } = new();

    public decimal    LineExtensionAmount  { get; set; }
    public decimal    TaxExclusiveAmount   { get; set; }
    public decimal    TaxInclusiveAmount   { get; set; }
    public decimal    AllowanceTotalAmount { get; set; }
    public decimal    ChargeTotalAmount    { get; set; }
    public decimal    PayableAmount        { get; set; }

    public string     XmlIcerik       { get; set; } = "";
    public string     XmlDosyaAdi     { get; set; } = "";
}
