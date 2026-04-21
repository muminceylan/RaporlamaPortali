namespace RaporlamaPortali.Models;

/// <summary>
/// Kategorinin "türü" — hangi arşiv tablosuna kaydedileceğini belirler.
/// </summary>
public enum BelgeKategorisi
{
    Tesis = 1,
    Mustahsil = 2,
    Genel = 3
}

public class Tesis
{
    public int Id { get; set; }
    public string Ad { get; set; } = "";
    public bool Aktif { get; set; } = true;
    public DateTime OlusturmaTarihi { get; set; } = DateTime.Now;
}

/// <summary>
/// Kullanıcı tarafından yönetilen kategori (ör. Tesis, Müstahsil, Tarım Kredi Faturaları).
/// Tur alanı, kategoriye ait evrakın hangi tabloya/forma bağlanacağını söyler.
/// </summary>
public class BelgeKategori
{
    public int Id { get; set; }
    public string Ad { get; set; } = "";
    public BelgeKategorisi Tur { get; set; } = BelgeKategorisi.Genel;
    public bool Aktif { get; set; } = true;
    public DateTime OlusturmaTarihi { get; set; } = DateTime.Now;
}

public class BelgeTipi
{
    public int Id { get; set; }
    public string Ad { get; set; } = "";
    public int KategoriId { get; set; }
    public bool Aktif { get; set; } = true;
    public DateTime OlusturmaTarihi { get; set; } = DateTime.Now;

    // JOIN ile doldurulur
    public string? KategoriAdi { get; set; }
    public BelgeKategorisi? KategoriTuru { get; set; }
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

/// <summary>
/// Genel türündeki kategorilere ait evrak (ör. Tarım Kredi Faturaları).
/// Firma/Cari + Fatura No + Tarih + Tutar temel alanları ile tutulur.
/// </summary>
public class GenelEvraki
{
    public int Id { get; set; }
    public int DefterYili { get; set; }
    public int DefterSiraNo { get; set; }
    public string DefterNo => $"{DefterYili}/{DefterSiraNo:D4}";

    public int KategoriId { get; set; }
    public int BelgeTipiId { get; set; }

    public string CariUnvani { get; set; } = "";
    public string? FaturaNo { get; set; }
    public DateTime? EvrakTarihi { get; set; }
    public decimal? Tutar { get; set; }

    public string DosyaAdi { get; set; } = "";
    public string DosyaYolu { get; set; } = "";
    public string? MimeType { get; set; }
    public long DosyaBoyutu { get; set; }

    public DateTime TebellugTarihi { get; set; }
    public string? Aciklama { get; set; }
    public DateTime YuklemeTarihi { get; set; } = DateTime.Now;

    public string? KategoriAdi { get; set; }
    public string? BelgeTipiAdi { get; set; }
}

public class MustahsilOzet
{
    public string TcKimlikNo { get; set; } = "";
    public string AdSoyadi { get; set; } = "";
    public string? MustahsilNo { get; set; }
    public string? HesapNo { get; set; }
}
