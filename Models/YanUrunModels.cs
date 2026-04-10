namespace RaporlamaPortali.Models;

/// <summary>
/// Yan ürün satış/üretim/iade özet verisi
/// </summary>
public class YanUrunOzet
{
    public string MalzemeKodu { get; set; } = "";
    public string MalzemeAdi { get; set; } = "";
    public string Kategori { get; set; } = ""; // MELAS, YAS_KUSPE, KURU_KUSPE, ALKOL, DIGER
    
    // Devir Stok (01.09.2025 bazlı)
    public decimal DevirStok { get; set; }
    
    // Satın Alma
    public decimal SatinAlmaMiktari { get; set; }
    public decimal SatinAlmaTutari { get; set; }
    
    // Üretim
    public decimal UretimMiktari { get; set; }
    
    // Satış
    public decimal SatisMiktari { get; set; }
    public decimal SatisTutari { get; set; }
    
    // İade
    public decimal IadeMiktari { get; set; }
    public decimal IadeTutari { get; set; }
    
    // Satınalma İade
    public decimal SatinAlmaIadeMiktari { get; set; }
    
    // Tüketim (Melas için alkol üretiminde kullanılan)
    public decimal TuketimMiktari { get; set; }
    
    // Hesaplanan alanlar
    public decimal OrtalamaFiyat => SatisMiktari > 0 ? SatisTutari / SatisMiktari : 0;
    public decimal NetSatis => SatisMiktari - IadeMiktari;
    public decimal NetTutar => SatisTutari - IadeTutari;
    
    /// <summary>
    /// Stok = Devir + Üretim + Satın Alma + İade - SatınalmaİAde - Satış - Tüketim
    /// </summary>
    public decimal Stok => DevirStok + UretimMiktari + SatinAlmaMiktari + IadeMiktari - SatinAlmaIadeMiktari - SatisMiktari - TuketimMiktari;
    
    // TON cinsinden değerler (1000'e bölünmüş)
    public decimal DevirStokTon => Math.Round(DevirStok / 1000, 2);
    public decimal UretimTon => Math.Round(UretimMiktari / 1000, 2);
    public decimal SatinAlmaTon => Math.Round(SatinAlmaMiktari / 1000, 2);
    public decimal SatisTon => Math.Round(SatisMiktari / 1000, 2);
    public decimal IadeTon => Math.Round(IadeMiktari / 1000, 2);
    public decimal SatinAlmaIadeTon => Math.Round(SatinAlmaIadeMiktari / 1000, 2);
    public decimal TuketimTon => Math.Round(TuketimMiktari / 1000, 2);
    public decimal StokTon => Math.Round(Stok / 1000, 2);
}

/// <summary>
/// Alkol ürünleri için özel model
/// </summary>
public class AlkolOzet
{
    public string MalzemeKodu { get; set; } = "";
    public string MalzemeAdi { get; set; } = "";
    public string AlkolTuru { get; set; } = ""; // GIDA, DENATURE, KOLONYA, vb.
    
    // Devir Stok
    public decimal DevirStok { get; set; }
    
    public decimal SatinAlmaMiktari { get; set; }
    public decimal UretimMiktari { get; set; }
    public decimal SatisMiktari { get; set; }
    public decimal SatisTutari { get; set; }
    public decimal IadeMiktari { get; set; }
    public decimal IadeTutari { get; set; }
    
    public decimal OrtalamaFiyat => SatisMiktari > 0 ? SatisTutari / SatisMiktari : 0;
    
    /// <summary>
    /// Stok = Devir + Üretim + Satın Alma + İade - Satış
    /// </summary>
    public decimal Stok => DevirStok + UretimMiktari + SatinAlmaMiktari + IadeMiktari - SatisMiktari;
    
    // TON (Litre) cinsinden değerler
    public decimal DevirStokTon => Math.Round(DevirStok / 1000, 2);
    public decimal UretimTon => Math.Round(UretimMiktari / 1000, 2);
    public decimal SatinAlmaTon => Math.Round(SatinAlmaMiktari / 1000, 2);
    public decimal SatisTon => Math.Round(SatisMiktari / 1000, 2);
    public decimal IadeTon => Math.Round(IadeMiktari / 1000, 2);
    public decimal StokTon => Math.Round(Stok / 1000, 2);
}

/// <summary>
/// Rapor filtre kriterleri
/// </summary>
public class RaporFiltre
{
    public DateTime BaslangicTarihi { get; set; } = new DateTime(DateTime.Now.Year, 1, 1);
    public DateTime BitisTarihi { get; set; } = DateTime.Now;
}

/// <summary>
/// Stok hareket detayı
/// </summary>
public class StokHareket
{
    public DateTime Tarih { get; set; }
    public string FisTuru { get; set; } = "";
    public string FisNo { get; set; } = "";
    public string CariKodu { get; set; } = "";
    public string CariAdi { get; set; } = "";
    public string MalzemeKodu { get; set; } = "";
    public string MalzemeAdi { get; set; } = "";
    public decimal GirisMiktari { get; set; }
    public decimal GirisTutari { get; set; }
    public decimal CikisMiktari { get; set; }
    public decimal CikisTutari { get; set; }
}

/// <summary>
/// Malzeme tanımları - hangi kod hangi ürüne karşılık geliyor
/// VIEW: INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_2025
/// </summary>
public static class MalzemeTanimlari
{
    public static readonly Dictionary<string, (string Ad, string Kategori)> YanUrunler = new()
    {
        // Melas
        { "S.706.04.0001", ("Melas", "MELAS") },
        
        // Yaş Küspe
        { "S.706.04.0002", ("Dökme Yaş Küspe", "YAS_KUSPE") },
        { "S.706.04.0008", ("25 Kg Yaş Küspe", "YAS_KUSPE") },
        { "S.706.04.0009", ("Tonluk Yaş Küspe", "YAS_KUSPE") },
        
        // Kuru Küspe - ÖNEMLİ: Dökme Kuru Küspe özel formül gerektirir
        // Dökme Kuru Küspe Üretim = Toplam Üretim - 50 Kg Çuvallı Üretim
        { "S.706.04.0003", ("50 Kg Çuvallı Kuru Küspe", "KURU_KUSPE") },
        { "S.706.04.0004", ("Dökme Kuru Küspe", "KURU_KUSPE") },
        { "S.706.04.0012", ("Peletlenmemiş Kuru Küspe", "KURU_KUSPE") },
        
        // Diğer
        { "S.706.04.0006", ("Kuyruk", "DIGER") },
        { "S.706.04.0007", ("Toprak", "DIGER") },
        { "Y_100153", ("Iskarta Patates", "DIGER") }
    };
    
    /// <summary>
    /// Alkol ürünleri - VIEW: INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Alkol_2025
    /// ÖNEMLİ: Gıda Alkolü özel formül gerektirir
    /// Gıda Alkolü Üretim = Toplam Üretim - Denatüre Edilmiş Alkol Üretim
    /// </summary>
    public static readonly Dictionary<string, (string Ad, string Tur)> AlkolUrunleri = new()
    {
        { "E.M.01.01.0001", ("Gıda Alkolü (Denatüre Edilmemiş)", "GIDA") },
        { "E.M.01.01.0005", ("Denatüre Edilmiş Alkol", "DENATURE") },
        { "E.M.01.01.0004", ("Teknik Alkol", "TEKNIK") },
        { "E.M.01.01.0002", ("Şilempe", "SILEMPE") },
        { "E.M.01.01.0003", ("Fuzel Yağı", "FUZEL") }
    };
    
    // Özel formül için malzeme kodları
    public const string DOKME_KURU_KUSPE = "S.706.04.0004";
    public const string CUVALLI_KURU_KUSPE = "S.706.04.0003";
    public const string GIDA_ALKOLU = "E.M.01.01.0001";
    public const string DENATURE_ALKOL = "E.M.01.01.0005";
    public const string MELAS = "S.706.04.0001";
    
    /// <summary>
    /// 01.09.2025 00:00 itibariyle devir stokları (KG cinsinden)
    /// </summary>
    public static readonly DateTime DEVIR_STOK_TARIHI = new DateTime(2025, 9, 1);
    
    public static readonly Dictionary<string, decimal> DevirStoklari = new()
    {
        // YAN ÜRÜNLER (KG cinsinden)
        { "S.706.04.0001", 21061182m },      // Melas: 21.061,182 Ton
        { "S.706.04.0002", 0m },              // Dökme Yaş Küspe
        { "S.706.04.0008", 0m },              // 25 Kg Yaş Küspe
        { "S.706.04.0009", 0m },              // Tonluk Yaş Küspe
        { "S.706.04.0003", 15000m },          // 50 Kg Çuvallı Kuru Küspe: 15 Ton
        { "S.706.04.0004", 0m },              // Dökme Kuru Küspe: 0 Ton
        { "S.706.04.0012", 0m },              // Peletlenmemiş Kuru Küspe
        { "S.706.04.0006", 0m },              // Kuyruk
        { "S.706.04.0007", 0m },              // Toprak
        { "Y_100153", 0m },                   // Iskarta Patates
        
        // ETİL ALKOL (KG/Litre cinsinden)
        { "E.M.01.01.0001", 1845469m },       // Gıda Alkolü: 1.845,469 Ton
        { "E.M.01.01.0005", 75965m },         // Denatüre Alkol: 75,965 Ton
        { "E.M.01.01.0004", 24937m },         // Teknik Alkol: 24,937 Ton
        { "E.M.01.01.0002", 977570m },        // Şilempe: 977,570 Ton
        { "E.M.01.01.0003", 1400m }           // Fuzel Yağı: 1,400 Ton
    };
    
    /// <summary>
    /// Malzeme için devir stok miktarını döndürür
    /// </summary>
    public static decimal GetDevirStok(string malzemeKodu)
    {
        return DevirStoklari.TryGetValue(malzemeKodu, out var stok) ? stok : 0m;
    }
}

/// <summary>
/// Şeker satış özet verisi
/// VIEW: INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Seker_2025
/// </summary>
public class SekerSatisOzet
{
    public string Kategori { get; set; } = ""; // A_CUVAL, A_PAKET, B_KOTASI, C_KOTASI, TICARI_CUVAL, TICARI_PAKET
    public string KategoriAdi { get; set; } = "";
    
    // VBA'daki 9 değer (KG cinsinden)
    public decimal DevirStok { get; set; }           // 0: Devir
    public decimal UretimMiktari { get; set; }       // 1: Üretim
    public decimal SatinAlmaMiktari { get; set; }    // 2: Satın Alma
    public decimal IadeMiktari { get; set; }         // 3: Satıştan İade
    public decimal SatinAlmaIadeMiktari { get; set; } // 4: Satınalma İade
    public decimal SatisMiktari { get; set; }        // 5: Satış
    public decimal PromosyonMiktari { get; set; }    // 6: Promosyon
    public decimal SarfMiktari { get; set; }         // 7: Sarf
    
    // Tutar alanları (opsiyonel)
    public decimal UretimTutari { get; set; }
    public decimal SatinAlmaTutari { get; set; }
    public decimal SatisTutari { get; set; }
    public decimal IadeTutari { get; set; }
    
    // Hesaplanan alanlar
    public decimal OrtalamaFiyat => SatisMiktari > 0 ? SatisTutari / SatisMiktari : 0;
    
    /// <summary>
    /// Stok = Devir + Üretim + SatınAlma + SatıştanİAde - SatınalmaİAde - Satış - Promosyon - Sarf
    /// VBA'daki formül
    /// </summary>
    public decimal Stok => DevirStok + UretimMiktari + SatinAlmaMiktari + IadeMiktari 
                          - SatinAlmaIadeMiktari - SatisMiktari - PromosyonMiktari - SarfMiktari;
    
    // TON cinsinden değerler (1000'e bölünmüş)
    public decimal DevirStokTon => Math.Round(DevirStok / 1000, 2);
    public decimal UretimTon => Math.Round(UretimMiktari / 1000, 2);
    public decimal SatinAlmaTon => Math.Round(SatinAlmaMiktari / 1000, 2);
    public decimal IadeTon => Math.Round(IadeMiktari / 1000, 2);
    public decimal SatinAlmaIadeTon => Math.Round(SatinAlmaIadeMiktari / 1000, 2);
    public decimal SatisTon => Math.Round(SatisMiktari / 1000, 2);
    public decimal PromosyonTon => Math.Round(PromosyonMiktari / 1000, 2);
    public decimal SarfTon => Math.Round(SarfMiktari / 1000, 2);
    public decimal StokTon => Math.Round(Stok / 1000, 2);
}

/// <summary>
/// Şeker kategorileri - A Kotası, B Kotası, C Kotası, Paketli Şeker
/// VBA makrosundaki UrunKategorisiBelirle fonksiyonundan alındı
/// </summary>
public static class SekerKategorileri
{
    public const string A_KOTASI = "A_KOTASI";
    public const string B_KOTASI = "B_KOTASI";  
    public const string C_KOTASI = "C_KOTASI";
    public const string PAKETLI_SEKER = "PAKETLI_SEKER";
    public const string KONYA_TICARI = "KONYA_TICARI";
    public const string HARIC = "HARIC"; // Türk Şeker - hesaplamaya dahil edilmez
    
    // Türk Şeker Ticari Mal kodları - HARİÇ tutulacak
    private static readonly string[] HaricKodlar = new[]
    {
        "T.T.0.0.0", "T.S.0.0.0", 
        "T.S.9.1.03.1.1000.20", "T.S.9.1.03.1.3000.06", 
        "T.S.9.1.03.1.5000.04", "T.T.9.1.03.1.5000.04"
    };
    
    // A Kotası Şeker kodları
    private static readonly string[] AKotasiKodlar = new[]
    {
        "S.T.0.0.0", "S.T.0.0.4", "S.705.00.0005"
    };
    
    // B Kotası Şeker kodları
    private static readonly string[] BKotasiKodlar = new[]
    {
        "S.T.0.0.8", "S.705.00.0001"
    };
    
    // C Kotası Şeker kodları
    private static readonly string[] CKotasiKodlar = new[]
    {
        "S.T.0.0.7", "S.705.00.0008"
    };
    
    // Konya Şeker Ticari Mal kodu
    private const string KonyaTicariKod = "S.T.1.0.0";
    
    /// <summary>
    /// Malzeme koduna göre şeker kategorisini belirler
    /// VBA'daki UrunKategorisiBelirle fonksiyonunun C# karşılığı
    /// </summary>
    public static string GetKategori(string malzemeKodu)
    {
        if (string.IsNullOrEmpty(malzemeKodu)) return PAKETLI_SEKER;
        
        var kod = malzemeKodu.Trim();
        
        // Türk Şeker Ticari Mal - HARİÇ tutulacak
        if (HaricKodlar.Contains(kod))
            return HARIC;
        
        // Konya Şeker Ticari Mal
        if (kod == KonyaTicariKod)
            return KONYA_TICARI;
        
        // A Kotası Şeker
        if (AKotasiKodlar.Contains(kod))
            return A_KOTASI;
        
        // B Kotası Şeker
        if (BKotasiKodlar.Contains(kod))
            return B_KOTASI;
        
        // C Kotası Şeker
        if (CKotasiKodlar.Contains(kod))
            return C_KOTASI;
        
        // Diğer tüm kodlar Paketli Şeker
        return PAKETLI_SEKER;
    }
    
    public static string GetKategoriAdi(string kategori)
    {
        return kategori switch
        {
            A_KOTASI => "A Kotası Şeker",
            B_KOTASI => "B Kotası Şeker",
            C_KOTASI => "C Kotası Şeker",
            PAKETLI_SEKER => "Paketli Şeker",
            KONYA_TICARI => "Konya Şeker Ticari Mal",
            HARIC => "Türk Şeker (Hariç)",
            _ => kategori
        };
    }
}
