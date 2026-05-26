namespace RaporlamaPortali.Models;

/// <summary>
/// Finans raporundaki bir satıra çift tıklandığında Logo'dan çekilen
/// fiş + satır detayını taşır.
/// </summary>
public class FisDetay
{
    public string Modul        { get; set; } = "";   // BANKA / KASA / CARI / KREDİ KARTI
    public int    LogicalRef   { get; set; }         // tıklanan satırın logicalref'i
    public int    FicheRef     { get; set; }         // varsa fiş başlık logicalref'i (BNFICHE / CLFICHE)
    public string FisNo        { get; set; } = "";   // FICHENO
    public string FisTuru      { get; set; } = "";   // TRCODE açıklaması
    public int    TrCode       { get; set; }
    public DateTime Tarih      { get; set; }
    public string Aciklama1    { get; set; } = "";   // BNFICHE.GENEXP1 / CLFICHE.GENEXP1 vs.
    public string Aciklama2    { get; set; } = "";
    public string Aciklama3    { get; set; } = "";
    public decimal BorcToplam  { get; set; }
    public decimal AlacakToplam{ get; set; }
    public string KullaniciOlusturan  { get; set; } = "";
    public DateTime? OlusturmaZamani  { get; set; }
    public string KullaniciDegistiren { get; set; } = "";
    public DateTime? DegistirmeZamani { get; set; }
    public bool Iptal { get; set; }

    public List<FisDetaySatir> Satirlar { get; set; } = new();
}

/// <summary>
/// Fişin tek bir satırı — UI tablosunda bir kayıt.
/// </summary>
public class FisDetaySatir
{
    public int      LogicalRef       { get; set; }
    public bool     SecilenSatir     { get; set; }   // çift tıklanan satır mı (mavi vurgu için)
    public int      Sira             { get; set; }
    public string   Aciklama         { get; set; } = "";
    public string   IslemTipi        { get; set; } = "";    // Borç / Alacak (SIGN)
    public decimal  Tutar            { get; set; }
    public string   DovizKodu        { get; set; } = "";
    public decimal  DovizKuru        { get; set; }
    public decimal  DovizTutari      { get; set; }

    // Karşı hesap bilgileri (cari / banka / kasa / muh)
    public string   CariKodu         { get; set; } = "";
    public string   CariUnvani       { get; set; } = "";
    public string   BankaKodu        { get; set; } = "";
    public string   BankaAdi         { get; set; } = "";
    public string   KasaKodu         { get; set; } = "";
    public string   KasaAdi          { get; set; } = "";
    public string   MuhasebeKodu     { get; set; } = "";
    public string   MuhasebeAdi      { get; set; } = "";
    public string   ProjeKodu        { get; set; } = "";
    public string   ProjeAdi         { get; set; } = "";
    public string   OzelKod          { get; set; } = "";
    public string   IslemNo          { get; set; } = "";    // TRANNO
}
