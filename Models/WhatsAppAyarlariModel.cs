namespace RaporlamaPortali.Models;

public class WhatsAppAyarlariModel
{
    public List<string> YetkiliNumaralar { get; set; } = new();
    public List<string> Tetikleyiciler   { get; set; } = new() { "tüm rapor", "tum rapor", "tumrapor", "tümrapor" };
    public string       RaporApiUrl      { get; set; } = "http://localhost:5050/api/rapor";
    public string       ExcelDosyasi     { get; set; } = "";
    public bool         OtomatikBaslat   { get; set; } = true;
}

public class WhatsAppLogKayit
{
    public DateTime Tarih  { get; set; }
    public string   Numara { get; set; } = "";
    public string   Mesaj  { get; set; } = "";
    public string   Sonuc  { get; set; } = "";
}

public class WhatsAppDurumModel
{
    /// <summary>BAGLI_DEGIL | QR_BEKLIYOR | BAGLI | HATA</summary>
    public string   Durum      { get; set; } = "BAGLI_DEGIL";
    /// <summary>Ham QR string — QRCoder ile image'a çevrilir</summary>
    public string   QrString   { get; set; } = "";
    public DateTime Guncelleme { get; set; } = DateTime.MinValue;
}
