namespace RaporlamaPortali.Models;

/// <summary>
/// Tarım Kredi Kooperatifi -> Bölge eşleşmesi (SQLite)
/// </summary>
public class TarimKrediBolgeEslesme
{
    public int Id { get; set; }
    public string FirmaAdi { get; set; } = "";
    public string Bolge { get; set; } = "";
    public string? CariKodu { get; set; }
    public DateTime OlusturmaTarihi { get; set; }
}

/// <summary>
/// Bölge tanımı — icmal mail adresi ve ilgili kişi bilgileri
/// </summary>
public class TarimKrediBolgeTanim
{
    public int Id { get; set; }
    public string Ad { get; set; } = "";
    public string? Email { get; set; }
    public string? IlgiliKisi { get; set; }
    public string? Telefon { get; set; }
    public DateTime OlusturmaTarihi { get; set; }
}

/// <summary>
/// Yan ürünler hareketi tek satır (Tarım Kredi Raporu için)
/// </summary>
public class YanUrunHareket
{
    public DateTime Tarih { get; set; }
    public string FisTuru { get; set; } = "";
    public string FisNumarasi { get; set; } = "";
    public string FaturaNo { get; set; } = "";
    public string CariHesapKodu { get; set; } = "";
    public string CariHesapUnvani { get; set; } = "";
    public string MalzemeKodu { get; set; } = "";
    public string MalzemeAciklamasi { get; set; } = "";
    public decimal GirisMiktari { get; set; }
    public decimal GirisFiyati { get; set; }
    public decimal GirisTutari { get; set; }
    public decimal CikisMiktari { get; set; }
    public decimal CikisFiyati { get; set; }
    public decimal CikisTutari { get; set; }

    public bool Iade => FisTuru.Contains("İade", StringComparison.OrdinalIgnoreCase);
    public decimal NetMiktar => Iade ? -GirisMiktari : CikisMiktari;
    public decimal NetTutar  => Iade ? -GirisTutari  : CikisTutari;
    public decimal Fiyat => Iade
        ? (GirisMiktari > 0 ? GirisTutari / GirisMiktari : 0m)
        : (CikisMiktari > 0 ? CikisTutari / CikisMiktari : 0m);

    /// <summary>Pozitif görünüm (ekran/Excel'de "-" gösterilmesin)</summary>
    public decimal MiktarGorunen => Math.Abs(NetMiktar);
    public decimal TutarGorunen => Math.Abs(NetTutar);

    /// <summary>"Toptan Satış İrsaliyesi" → "Toptan Satış Faturası"</summary>
    public string FisTuruGorunen => (FisTuru ?? "").Replace("İrsaliyesi", "Faturası", StringComparison.OrdinalIgnoreCase);

    /// <summary>Ekranda gösterilecek no: FaturaNo varsa o, yoksa fiş (irsaliye) no</summary>
    public string NumaraGorunen => !string.IsNullOrWhiteSpace(FaturaNo) ? FaturaNo : FisNumarasi;
}

/// <summary>
/// Firma bazlı alt toplam
/// </summary>
public class TarimKrediFirmaOzet
{
    public string CariHesapUnvani { get; set; } = "";
    public string CariHesapKodu { get; set; } = "";
    public List<YanUrunHareket> Hareketler { get; set; } = new();
    public decimal ToplamMiktar => Hareketler.Sum(h => h.NetMiktar);
    public decimal ToplamTutar  => Hareketler.Sum(h => h.NetTutar);
}

/// <summary>
/// Bölge bazlı rapor
/// </summary>
public class TarimKrediBolgeRapor
{
    public string Bolge { get; set; } = "";
    public List<TarimKrediFirmaOzet> Firmalar { get; set; } = new();
    public decimal ToplamMiktar => Firmalar.Sum(f => f.ToplamMiktar);
    public decimal ToplamTutar  => Firmalar.Sum(f => f.ToplamTutar);
    public int FirmaSayisi => Firmalar.Count;
    public int HareketSayisi => Firmalar.Sum(f => f.Hareketler.Count);
}
