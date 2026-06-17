namespace RaporlamaPortali.Models;

/// <summary>
/// Satış Kıyaslama Raporu — bir cariye ait iki dönemin malzeme bazında karşılaştırması.
/// </summary>
public class SatisKiyaslamaSonuc
{
    public CariBilgi Cari      { get; set; } = new();
    public DonemOzeti DonemA   { get; set; } = new();
    public DonemOzeti DonemB   { get; set; } = new();
    public List<SatisKiyaslamaSatiri> Satirlar { get; set; } = new();
}

public class CariBilgi
{
    public int    LogicalRef { get; set; }
    public string Kod        { get; set; } = "";
    public string Unvan      { get; set; } = "";
    public string Vkn        { get; set; } = "";
    public string Sehir      { get; set; } = "";
}

public class DonemOzeti
{
    public DateTime Baslangic    { get; set; }
    public DateTime Bitis        { get; set; }
    public int      FisSayisi    { get; set; }
    public decimal  ToplamMiktar { get; set; }
    public decimal  NetTutar     { get; set; }
    public decimal  KdvTutar     { get; set; }
    public decimal  BrutTutar    { get; set; }
    public decimal  IadeTutar    { get; set; }
    public string   Etiket       => $"{Baslangic:dd.MM.yyyy} - {Bitis:dd.MM.yyyy}";
}

public class SatisKiyaslamaSatiri
{
    public string  MalzemeKodu { get; set; } = "";
    public string  MalzemeAdi  { get; set; } = "";
    public string  Birim       { get; set; } = "";

    public decimal A_Miktar    { get; set; }
    public decimal A_NetTutar  { get; set; }
    public decimal A_OrtFiyat  => A_Miktar == 0 ? 0 : A_NetTutar / A_Miktar;

    public decimal B_Miktar    { get; set; }
    public decimal B_NetTutar  { get; set; }
    public decimal B_OrtFiyat  => B_Miktar == 0 ? 0 : B_NetTutar / B_Miktar;

    public decimal FarkMiktar  => A_Miktar - B_Miktar;
    public decimal FarkTutar   => A_NetTutar - B_NetTutar;

    /// <summary>(A - B) / B × 100. B sıfırsa A varsa +100, yoksa 0.</summary>
    public decimal FarkYuzdeTutar => B_NetTutar == 0
        ? (A_NetTutar == 0 ? 0 : 100m)
        : (A_NetTutar - B_NetTutar) / B_NetTutar * 100m;

    public decimal FarkYuzdeMiktar => B_Miktar == 0
        ? (A_Miktar == 0 ? 0 : 100m)
        : (A_Miktar - B_Miktar) / B_Miktar * 100m;
}
