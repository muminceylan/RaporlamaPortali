namespace RaporlamaPortali.Models;

/// <summary>
/// İşletme Malzemeleri Raporu konfigürasyonu. JSON dosyada
/// (<c>C:\RaporlamaPortaliData\isletme_malzemeleri.json</c>) saklanır.
/// 3 sabit grup (Yakıtlar / Torbalar Filtre Bezleri / Kimyasallar) için
/// Logo malzeme kod önekleri ve isteğe bağlı manuel açıklamalar tutar.
/// </summary>
public class IsletmeMalzemeleriKonfig
{
    public IsletmeMalzemeGrubu Yakitlar   { get; set; } = new() { GrupAdi = "Yakıtlar" };
    public IsletmeMalzemeGrubu Torbalar   { get; set; } = new() { GrupAdi = "Torbalar Filtre Bezleri" };
    public IsletmeMalzemeGrubu Kimyasallar{ get; set; } = new() { GrupAdi = "Kimyasallar" };

    /// <summary>Malzeme bazında manuel açıklamalar (her kod için kısa not).</summary>
    public Dictionary<string, string> MalzemeAciklamalari { get; set; } = new();

    /// <summary>Son kullanılan kampanya tarih aralığı — sayfa açılırken hatırla.</summary>
    public DateTime? KampanyaBaslangic { get; set; }
    public DateTime? KampanyaBitis     { get; set; }
}

public class IsletmeMalzemeGrubu
{
    public string GrupAdi { get; set; } = "";

    /// <summary>
    /// Logo kod önekleri. Excel'deki "208.01" veya tam Logo formatı "S.208.01" kabul edilir.
    /// Sorgu sırasında her öneki için <c>CODE LIKE '%{onek}%'</c> uygulanır, böylece
    /// S./K. öneki ile başlayan tüm varyantlar tek seferde gelir.
    /// </summary>
    public List<string> KodOnekleri { get; set; } = new();
}

/// <summary>
/// İşletme Malzemeleri Raporu — bir malzemenin tek satırı.
/// Excel "MALZEME 2025-2026 YILI.XLS" formatıyla birebir uyumlu kolonlar.
/// </summary>
public class IsletmeMalzemeSatiri
{
    public string MalzemeKodu      { get; set; } = "";
    public string MalzemeAdi       { get; set; } = "";
    public string Birim            { get; set; } = "";

    /// <summary>Devir tarihinden ÖNCE birikmiş net stok (V1 view kümülatif).</summary>
    public decimal DevirStogu      { get; set; }

    /// <summary>Devir tarihi seçilen "günlük gelen" sütunu için (ör. 2.06.2026).</summary>
    public DateTime? GunlukGelenTarihi { get; set; }
    public decimal GunlukGelen     { get; set; }

    /// <summary>Devir tarihinden bugüne kadar tüm girişler (IOCODE=1 satın alma).</summary>
    public decimal GelenToplam     { get; set; }

    /// <summary>Kampanya başlangıç-bitiş aralığında çıkışlar (IOCODE=4).</summary>
    public decimal KampanyaSarfiyati { get; set; }

    /// <summary>Devir sonrası, kampanya aralığı dışındaki çıkışlar.</summary>
    public decimal KampanyaDisiSarf { get; set; }

    /// <summary>Toplam sarfiyat = Kampanya + Kampanya dışı.</summary>
    public decimal ToplamSarfiyat  { get; set; }

    /// <summary>Mevcut stok = GNTOTST anlık (Devir + Gelen − Sarf doğrulaması).</summary>
    public decimal Mevcut          { get; set; }

    /// <summary>Kullanıcının JSON konfigürasyonda tuttuğu manuel açıklama.</summary>
    public string Aciklama         { get; set; } = "";
}
