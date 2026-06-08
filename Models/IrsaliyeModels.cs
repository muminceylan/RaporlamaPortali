namespace RaporlamaPortali.Models;

// Satınalma (TRCODE=1) ve Toptan Satış (TRCODE=8) irsaliyelerinin liste ve detay modelleri.
// Faturalanma durumu STFICHE.BILLED alanından gelir (0=faturalanmamış, 1=faturalanmış).

public class IrsaliyeOzet
{
    public int      LogicalRef     { get; set; }
    public string   FisNo          { get; set; } = "";
    public DateTime Tarih          { get; set; }
    public int      Trcode         { get; set; }
    public bool     Faturalandi    { get; set; }   // BILLED = 1
    public string   CariKodu       { get; set; } = "";
    public string   CariUnvani     { get; set; } = "";
    public int      AmbarNo        { get; set; }
    public string   AmbarAdi       { get; set; } = "";
    public decimal  NetTutar       { get; set; }
    public decimal  BrutTutar      { get; set; }
    public string?  Aciklama       { get; set; }
}

public class IrsaliyeDetay
{
    public int      LogicalRef       { get; set; }
    public string   FisNo            { get; set; } = "";
    public DateTime Tarih            { get; set; }
    public int      Trcode           { get; set; }
    public string   FisTuru          { get; set; } = "";
    public bool     Faturalandi      { get; set; }
    public string   CariKodu         { get; set; } = "";
    public string   CariUnvani       { get; set; } = "";
    public int      AmbarNo          { get; set; }
    public string   AmbarAdi         { get; set; } = "";
    public decimal  NetTutar         { get; set; }
    public decimal  BrutTutar        { get; set; }
    public decimal  ToplamIskonto    { get; set; }
    public decimal  ToplamKdv        { get; set; }
    public string?  Aciklama1        { get; set; }
    public string?  Aciklama2        { get; set; }
    public string?  Aciklama3        { get; set; }
    public string?  Aciklama4        { get; set; }
    public string?  Aciklama5        { get; set; }
    public string?  Aciklama6        { get; set; }
    public string?  MuafiyetKodu     { get; set; }   // STFICHE.VATEXCEPTCODE
    public string?  MuafiyetAciklama { get; set; }   // STFICHE.VATEXCEPTREASON

    // Kayıt bilgisi — Logo CAPIBLOCK_* alanları + L_CAPIUSER join
    public string?   KullaniciOlusturan  { get; set; }
    public DateTime? OlusturmaZamani     { get; set; }
    public string?   KullaniciDegistiren { get; set; }
    public DateTime? DegistirmeZamani    { get; set; }

    public List<IrsaliyeDetaySatir> Satirlar { get; set; } = new();
}

public class IrsaliyeDetaySatir
{
    public int     Sira             { get; set; }
    public string  MalzemeKodu      { get; set; } = "";
    public string  MalzemeAdi       { get; set; } = "";
    public decimal Miktar           { get; set; }
    public string  BirimAdi         { get; set; } = "";
    public decimal Fiyat            { get; set; }
    public decimal Tutar            { get; set; }   // LINENET
    public decimal KdvMatrah        { get; set; }   // VATMATRAH
    public int     AmbarNo          { get; set; }   // TRCODE'a göre DESTINDEX veya SOURCEINDEX
    public string? MuafiyetKodu     { get; set; }   // STLINE.VATEXCEPTCODE
    public string? MuafiyetAciklama { get; set; }   // STLINE.VATEXCEPTREASON
    public string? SatirAciklama    { get; set; }
}
