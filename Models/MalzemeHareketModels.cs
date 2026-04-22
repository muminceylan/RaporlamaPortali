namespace RaporlamaPortali.Models;

/// <summary>
/// Logo STLINE tablosundan tek bir malzeme hareketi satırı.
/// INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_2025 view'ünün kolon yapısını birebir taşır.
/// </summary>
public class MalzemeHareketSatiri
{
    public int     Yil              { get; set; }
    public int     Ay               { get; set; }
    public DateTime Tarih           { get; set; }
    public string  FisTuru          { get; set; } = "";
    public string  FisNumarasi      { get; set; } = "";
    public string  CariHesapKodu    { get; set; } = "";
    public string  CariHesapUnvani  { get; set; } = "";
    public string  MalzemeKodu      { get; set; } = "";
    public string  MalzemeAciklamasi{ get; set; } = "";
    public decimal GirisMiktari     { get; set; }
    public decimal GirisFiyati      { get; set; }
    public decimal GirisTutari      { get; set; }
    public decimal CikisMiktari     { get; set; }
    public decimal CikisFiyati      { get; set; }
    public decimal CikisTutari      { get; set; }
}

/// <summary>
/// Kayıtlı malzeme listesi — kullanıcının isim verip kaydettiği malzeme kodu koleksiyonu.
/// JSON olarak <see cref="RaporlamaPortali.Services.AppDataPaths.MalzemeListeleriJson"/> dosyasında tutulur.
/// </summary>
public class KayitliMalzemeListesi
{
    public string       Ad             { get; set; } = "";
    public List<string> MalzemeKodlari { get; set; } = new();
    public DateTime     OlusturmaTarihi{ get; set; } = DateTime.Now;
    public DateTime     GuncellemeTarihi{ get; set; } = DateTime.Now;
}
