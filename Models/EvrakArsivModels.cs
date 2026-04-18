namespace RaporlamaPortali.Models;

public enum BelgeKategorisi
{
    Tesis = 1,
    Mustahsil = 2
}

public class Tesis
{
    public int Id { get; set; }
    public string Ad { get; set; } = "";
    public bool Aktif { get; set; } = true;
    public DateTime OlusturmaTarihi { get; set; } = DateTime.Now;
}

public class BelgeTipi
{
    public int Id { get; set; }
    public string Ad { get; set; } = "";
    public BelgeKategorisi Kategori { get; set; }
    public bool Aktif { get; set; } = true;
    public DateTime OlusturmaTarihi { get; set; } = DateTime.Now;
}

public class TesisEvraki
{
    public int Id { get; set; }
    public int DefterYili { get; set; }
    public int DefterSiraNo { get; set; }
    public string DefterNo => $"{DefterYili}/{DefterSiraNo:D4}";

    public int TesisId { get; set; }
    public int BelgeTipiId { get; set; }

    public string DosyaAdi { get; set; } = "";
    public string DosyaYolu { get; set; } = "";
    public string? MimeType { get; set; }
    public long DosyaBoyutu { get; set; }

    public DateTime? EvrakTarihi { get; set; }
    public string? EvrakNo { get; set; }
    public DateTime TebellugTarihi { get; set; }
    public DateTime? GecerlilikBaslangic { get; set; }
    public DateTime? GecerlilikBitis { get; set; }

    public string? Aciklama { get; set; }
    public DateTime YuklemeTarihi { get; set; } = DateTime.Now;

    public string? TesisAdi { get; set; }
    public string? BelgeTipiAdi { get; set; }
}

public class MustahsilEvraki
{
    public int Id { get; set; }
    public int DefterYili { get; set; }
    public int DefterSiraNo { get; set; }
    public string DefterNo => $"{DefterYili}/{DefterSiraNo:D4}";

    public int KampanyaYili { get; set; }
    public string? TcKimlikNo { get; set; }
    public string MustahsilAdSoyadi { get; set; } = "";
    public string? MustahsilNo { get; set; }
    public string? HesapNo { get; set; }

    public int BelgeTipiId { get; set; }

    public string DosyaAdi { get; set; } = "";
    public string DosyaYolu { get; set; } = "";
    public string? MimeType { get; set; }
    public long DosyaBoyutu { get; set; }

    public DateTime? EvrakTarihi { get; set; }
    public string? EvrakNo { get; set; }
    public DateTime TebellugTarihi { get; set; }

    public string? Aciklama { get; set; }
    public DateTime YuklemeTarihi { get; set; } = DateTime.Now;

    public string? BelgeTipiAdi { get; set; }
}

public class MustahsilOzet
{
    public string TcKimlikNo { get; set; } = "";
    public string AdSoyadi { get; set; } = "";
    public string? MustahsilNo { get; set; }
    public string? HesapNo { get; set; }
}
