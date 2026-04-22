namespace RaporlamaPortali.Models;

/// <summary>
/// INF_MD_FINANS_PROJE_RAPORU_211_YYYY view'üne ait tek satır.
/// Logo'da her yıl için ayrı view yaratılıyor; bu model view kolonlarını birebir taşır.
/// </summary>
public class FinansRaporSatiri
{
    public int      LogicalRef           { get; set; }
    public string   FisTuru              { get; set; } = "";
    public string   Firma                { get; set; } = "";
    public string   HareketTuru          { get; set; } = "";
    public string   Modul                { get; set; } = "";
    public string   ChKod                { get; set; } = "";
    public string   ChUnvani             { get; set; } = "";
    public string   BankaHesapKodu       { get; set; } = "";
    public string   BankaHesapAciklamasi { get; set; } = "";
    public string   ProjeKodu            { get; set; } = "";
    public string   ProjeAdi             { get; set; } = "";
    public DateTime Tarih                { get; set; }
    public int      Yil                  { get; set; }
    public int      Ay                   { get; set; }
    public int      Gun                  { get; set; }
    public string   HaraketOzelKodu      { get; set; } = "";
    // View'da ISLEM_NO bazı branch'larda string dönebiliyor — güvenli olsun diye string
    public string   IslemNo              { get; set; } = "";
    public decimal  Havale               { get; set; }
    public decimal  Cek                  { get; set; }
    public decimal  Devir                { get; set; }
    public decimal  Diger                { get; set; }
    public int      Sign                 { get; set; }
    public string   SpecOde              { get; set; } = "";
}
