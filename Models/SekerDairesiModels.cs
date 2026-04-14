namespace RaporlamaPortali.Models;

/// <summary>
/// Bir sonraki ayda bulunan Toptan Satış İade İrsaliyesi'nin mevcut döneme yansıması
/// </summary>
public class SatisIadeDipnot
{
    /// <summary>İadenin LOGO'daki orijinal kategorisi (A_KOTASI vb.)</summary>
    public string KaynakKategori    { get; set; } = "";
    public string KaynakKategoriAdi { get; set; } = "";

    /// <summary>Gerçekte hangi kategoriden düşüldüğü (orijinal kategoride satış yoksa ticari kategoriye yönlendi)</summary>
    public string HedefKategori    { get; set; } = "";
    public string HedefKategoriAdi { get; set; } = "";

    /// <summary>Düşülen miktar (satış yeterliyse = iade miktarı, değilse = mevcut satış kadar)</summary>
    public decimal Miktar  { get; set; }
    public decimal Tutar   { get; set; }

    /// <summary>Satış yetersiz olduğu için karşılanamayan kısım (0 ise tam karşılandı)</summary>
    public decimal KarsılanamayenMiktar { get; set; }
    public decimal KarsılanamayenTutar  { get; set; }

    /// <summary>"Ekim 2025" gibi</summary>
    public string SonrakiAyAdi { get; set; } = "";

    /// <summary>Kaynak ≠ Hedef ise true (farklı kategoriye yönlendirildi)</summary>
    public bool Yonlendirildi { get; set; }
}

/// <summary>
/// Stok Sorgula (VBA Module2 StokSorgula) - Ambar/malzeme bazlı stok satırı
/// </summary>
public class StokDetayRow
{
    public int AmbarNo { get; set; }
    public string Ambar { get; set; } = "";
    public string MalzemeKodu { get; set; } = "";
    public string MalzemeAdi { get; set; } = "";
    public string AnaBirim { get; set; } = "";
    public decimal Stok { get; set; }
    public decimal? Kg { get; set; }
}

/// <summary>
/// Stok Sorgula - Malzeme bazlı özet (OzetSayfaOlustur karşılığı)
/// </summary>
public class StokOzetRow
{
    public string MalzemeKodu { get; set; } = "";
    public string MalzemeAdi { get; set; } = "";
    public decimal ToplamKg { get; set; }
}

/// <summary>
/// Sade Şeker Analizi (VBA Module3 SadeSekerAnaliziYap) - Kategori bazlı analiz
/// </summary>
public class SekerKategoriAnaliz
{
    public string Kategori { get; set; } = "";
    public string KategoriAdi { get; set; } = "";

    // Dönem başı stok (kullanıcı girişi)
    public decimal DonemBasiMiktar { get; set; }
    public decimal DonemBasiTutar { get; set; }

    // Giriş kalemleri
    public decimal UretimMiktar { get; set; }
    public decimal UretimTutar { get; set; }
    public decimal SatinAlmaMiktar { get; set; }
    public decimal SatinAlmaTutar { get; set; }
    public decimal SatisIadeMiktar { get; set; }   // Satıştan iade (giriş)
    public decimal SatisIadeTutar { get; set; }
    public decimal HammaddeGirisMiktar { get; set; }
    public decimal HammaddeGirisTutar { get; set; }
    public decimal ReceteFarkMiktar { get; set; }
    public decimal ReceteFarkTutar { get; set; }
    public decimal SayimFazlasiMiktar { get; set; }
    public decimal SayimFazlasiTutar { get; set; }

    // Çıkış kalemleri
    public decimal SatisMiktar { get; set; }
    public decimal SatisTutar { get; set; }
    public decimal SarfMiktar { get; set; }
    public decimal SarfTutar { get; set; }
    public decimal FireMiktar { get; set; }
    public decimal FireTutar { get; set; }
    public decimal YemekhaneMiktar { get; set; }
    public decimal YemekhaneTutar { get; set; }
    public decimal PromsMiktar { get; set; }       // PROMS ve Teknik Masraf Sarf
    public decimal PromsTutar { get; set; }
    public decimal HammaddeCikisMiktar { get; set; }
    public decimal HammaddeCikisTutar { get; set; }
    public decimal SatinAlmaIadeMiktar { get; set; }
    public decimal SatinAlmaIadeTutar { get; set; }

    // Hesaplanan dönem sonu stok
    public decimal ToplamGirisMiktar =>
        DonemBasiMiktar + UretimMiktar + SatinAlmaMiktar + SatisIadeMiktar +
        HammaddeGirisMiktar + ReceteFarkMiktar + SayimFazlasiMiktar;

    public decimal ToplamGirisTutar =>
        DonemBasiTutar + UretimTutar + SatinAlmaTutar + SatisIadeTutar +
        HammaddeGirisTutar + ReceteFarkTutar + SayimFazlasiTutar;

    public decimal ToplamCikisMiktar =>
        SatisMiktar + SarfMiktar + FireMiktar + YemekhaneMiktar +
        PromsMiktar + SatinAlmaIadeMiktar + HammaddeCikisMiktar;

    public decimal ToplamCikisTutar =>
        SatisTutar + SarfTutar + FireTutar + YemekhaneTutar +
        PromsTutar + SatinAlmaIadeTutar + HammaddeCikisTutar;

    public decimal DonemSonuMiktar => ToplamGirisMiktar - ToplamCikisMiktar;
    public decimal DonemSonuTutar => ToplamGirisTutar - ToplamCikisTutar;
}
