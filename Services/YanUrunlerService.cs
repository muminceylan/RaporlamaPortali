using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Yan Ürünler (Melas, Küspe, Alkol) rapor servisi
/// Excel'deki hazır VIEW'ları kullanır:
/// - INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_2025 (Yan ürünler)
/// - INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Alkol_2025 (Alkol)
/// 
/// ÖZEL FORMÜLLER:
/// 1. Dökme Kuru Küspe Üretimi = Toplam Üretim - 50 Kg Çuvallı Kuru Küspe Üretimi
/// 2. Gıda Alkolü Üretimi = Toplam Üretim - Denatüre Edilmiş Alkol Üretimi
/// </summary>
public class YanUrunlerService
{
    private readonly DatabaseService _db;

    public YanUrunlerService(DatabaseService db)
    {
        _db = db;
    }

    /// <summary>
    /// Yan ürünlerin özet raporunu getirir
    /// ÖZEL FORMÜL: Dökme Kuru Küspe üretimi düzeltmesi yapılır
    /// STOK: Devir + Üretim + Satın Alma + İade - Satış - Tüketim
    /// </summary>
    public async Task<List<YanUrunOzet>> GetYanUrunlerOzetAsync(DateTime baslangic, DateTime bitis)
    {
        bitis = SistemTarihi.Clamp(bitis);
        var sonuclar = new List<YanUrunOzet>();

        // Tüm tanımlı yan ürünler için veri çek
        foreach (var (malzemeKodu, (malzemeAdi, kategori)) in MalzemeTanimlari.YanUrunler)
        {
            var ozet = await GetMalzemeOzetAsync(malzemeKodu, baslangic, bitis);
            ozet.MalzemeAdi = malzemeAdi;
            ozet.Kategori = kategori;
            
            // Devir stok hesapla
            ozet.DevirStok = await HesaplaDevirStokAsync(malzemeKodu, baslangic);
            
            sonuclar.Add(ozet);
        }

        // Melas için alkol üretiminde tüketilen miktarı hesapla
        var melas = sonuclar.FirstOrDefault(x => x.MalzemeKodu == MalzemeTanimlari.MELAS);
        if (melas != null)
        {
            melas.TuketimMiktari = await GetMelasTuketimiAsync(baslangic, bitis);
        }

        // ÖZEL FORMÜL: Dökme Kuru Küspe Üretimi = Toplam - Çuvallı Üretim
        var dokmeKuru = sonuclar.FirstOrDefault(x => x.MalzemeKodu == MalzemeTanimlari.DOKME_KURU_KUSPE);
        var cuvalliKuru = sonuclar.FirstOrDefault(x => x.MalzemeKodu == MalzemeTanimlari.CUVALLI_KURU_KUSPE);
        
        if (dokmeKuru != null && cuvalliKuru != null)
        {
            // Çuvallı küspe üretimi, dökme küspe üretiminden düşülür
            dokmeKuru.UretimMiktari = Math.Max(0, dokmeKuru.UretimMiktari - cuvalliKuru.UretimMiktari);
        }

        return sonuclar;
    }

    /// <summary>
    /// Alkol ürünlerinin özet raporunu getirir
    /// ÖZEL FORMÜL: Gıda Alkolü üretimi düzeltmesi yapılır
    /// </summary>
    public async Task<List<AlkolOzet>> GetAlkolOzetAsync(DateTime baslangic, DateTime bitis)
    {
        bitis = SistemTarihi.Clamp(bitis);
        var sonuclar = new List<AlkolOzet>();

        foreach (var (malzemeKodu, (malzemeAdi, alkolTuru)) in MalzemeTanimlari.AlkolUrunleri)
        {
            var hareket = await GetAlkolMalzemeOzetAsync(malzemeKodu, baslangic, bitis);
            
            // Devir stok hesapla
            var devirStok = await HesaplaAlkolDevirStokAsync(malzemeKodu, baslangic);
            
            sonuclar.Add(new AlkolOzet
            {
                MalzemeKodu = malzemeKodu,
                MalzemeAdi = malzemeAdi,
                AlkolTuru = alkolTuru,
                DevirStok = devirStok,
                SatinAlmaMiktari = hareket.SatinAlmaMiktari,
                UretimMiktari = hareket.UretimMiktari,
                SatisMiktari = hareket.SatisMiktari,
                SatisTutari = hareket.SatisTutari,
                IadeMiktari = hareket.IadeMiktari,
                IadeTutari = hareket.IadeTutari
            });
        }

        // ÖZEL FORMÜL: Gıda Alkolü Üretimi = Toplam - Denatüre Alkol Üretimi
        var gidaAlkol = sonuclar.FirstOrDefault(x => x.MalzemeKodu == MalzemeTanimlari.GIDA_ALKOLU);
        var denatüreAlkol = sonuclar.FirstOrDefault(x => x.MalzemeKodu == MalzemeTanimlari.DENATURE_ALKOL);
        
        if (gidaAlkol != null && denatüreAlkol != null)
        {
            // Denatüre alkol üretimi, gıda alkolü üretiminden düşülür
            gidaAlkol.UretimMiktari = Math.Max(0, gidaAlkol.UretimMiktari - denatüreAlkol.UretimMiktari);
        }

        return sonuclar;
    }

    /// <summary>
    /// Yan ürün için devir stok hesaplar
    /// Başlangıç tarihi 01.09.2025'ten sonra ise ara dönem hareketlerini hesaplar
    /// </summary>
    private async Task<decimal> HesaplaDevirStokAsync(string malzemeKodu, DateTime baslangicTarihi)
    {
        var bazDevirStok = MalzemeTanimlari.GetDevirStok(malzemeKodu);
        
        // Eğer başlangıç tarihi devir stok tarihinden önce veya eşitse, baz devir stoğu döndür
        if (baslangicTarihi <= MalzemeTanimlari.DEVIR_STOK_TARIHI)
        {
            return bazDevirStok;
        }
        
        // Ara dönem hareketlerini hesapla (devir tarihi ile başlangıç tarihi arası)
        var araDonemBitis = baslangicTarihi.AddDays(-1);
        var araDonemHareket = await GetMalzemeOzetAsync(malzemeKodu, MalzemeTanimlari.DEVIR_STOK_TARIHI, araDonemBitis);
        
        // Melas için tüketimi de hesapla
        decimal tuketim = 0;
        if (malzemeKodu == MalzemeTanimlari.MELAS)
        {
            tuketim = await GetMelasTuketimiAsync(MalzemeTanimlari.DEVIR_STOK_TARIHI, araDonemBitis);
        }
        
        // Dökme Kuru Küspe için özel üretim hesabı
        if (malzemeKodu == MalzemeTanimlari.DOKME_KURU_KUSPE)
        {
            var cuvalliHareket = await GetMalzemeOzetAsync(MalzemeTanimlari.CUVALLI_KURU_KUSPE, MalzemeTanimlari.DEVIR_STOK_TARIHI, araDonemBitis);
            araDonemHareket.UretimMiktari = Math.Max(0, araDonemHareket.UretimMiktari - cuvalliHareket.UretimMiktari);
        }
        
        // Yeni devir = Baz + Üretim + Satın Alma + İade - Satış - Tüketim
        return bazDevirStok 
            + araDonemHareket.UretimMiktari 
            + araDonemHareket.SatinAlmaMiktari 
            + araDonemHareket.IadeMiktari 
            - araDonemHareket.SatisMiktari
            - tuketim;
    }

    /// <summary>
    /// Alkol ürünü için devir stok hesaplar
    /// </summary>
    private async Task<decimal> HesaplaAlkolDevirStokAsync(string malzemeKodu, DateTime baslangicTarihi)
    {
        var bazDevirStok = MalzemeTanimlari.GetDevirStok(malzemeKodu);
        
        if (baslangicTarihi <= MalzemeTanimlari.DEVIR_STOK_TARIHI)
        {
            return bazDevirStok;
        }
        
        var araDonemBitis = baslangicTarihi.AddDays(-1);
        var araDonemHareket = await GetAlkolMalzemeOzetAsync(malzemeKodu, MalzemeTanimlari.DEVIR_STOK_TARIHI, araDonemBitis);
        
        // Gıda Alkolü için özel üretim hesabı
        if (malzemeKodu == MalzemeTanimlari.GIDA_ALKOLU)
        {
            var denatüreHareket = await GetAlkolMalzemeOzetAsync(MalzemeTanimlari.DENATURE_ALKOL, MalzemeTanimlari.DEVIR_STOK_TARIHI, araDonemBitis);
            araDonemHareket.UretimMiktari = Math.Max(0, araDonemHareket.UretimMiktari - denatüreHareket.UretimMiktari);
        }
        
        return bazDevirStok 
            + araDonemHareket.UretimMiktari 
            + araDonemHareket.SatinAlmaMiktari 
            + araDonemHareket.IadeMiktari 
            - araDonemHareket.SatisMiktari;
    }

    /// <summary>
    /// Melas tüketimini hesaplar (Alkol üretimi için)
    /// </summary>
    private async Task<decimal> GetMelasTuketimiAsync(DateTime baslangic, DateTime bitis)
    {
        // Alkol VIEW'ından Melas sarfı - "Proms." fiş türü (alkol üretimi için tüketilen)
        var sql = @"
            SELECT ISNULL(SUM(ABS(CIKIS_MIKTARI)), 0)
            FROM INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Alkol_2025 WITH(NOLOCK)
            WHERE MALZEME_KODU = @MelasKodu
              AND FIS_TURU LIKE '%Proms%'
              AND TARIH >= @Baslangic
              AND TARIH <= @Bitis";

        using var conn = _db.CreateConnection();
        return await conn.ExecuteScalarAsync<decimal>(sql, new 
        { 
            MelasKodu = MalzemeTanimlari.MELAS,
            Baslangic = baslangic, 
            Bitis = bitis 
        });
    }

    /// <summary>
    /// Tek bir malzeme için özet verileri çeker (INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_2025 VIEW'ından)
    /// </summary>
    private async Task<YanUrunOzet> GetMalzemeOzetAsync(string malzemeKodu, DateTime baslangic, DateTime bitis)
    {
        var ozet = new YanUrunOzet { MalzemeKodu = malzemeKodu };

        // VIEW yapısı: TARIH, FIS_TURU, MALZEME_KODU, GIRIS_MIKTARI, GIRIS_TUTARI, CIKIS_MIKTARI, CIKIS_TUTARI
        // FIS_TURU değerleri: "Toptan Satış İrsaliyesi", "Üretimden Giriş Fişi", "Toptan Satış İade İrsaliyesi", "Satınalma İrsaliyesi"
        
        var sql = @"
            SELECT 
                FIS_TURU,
                TOPLAM_GIRIS = ISNULL(SUM(ABS(GIRIS_MIKTARI)), 0),
                TOPLAM_GIRIS_TUTAR = ISNULL(SUM(ABS(GIRIS_TUTARI)), 0),
                TOPLAM_CIKIS = ISNULL(SUM(ABS(CIKIS_MIKTARI)), 0),
                TOPLAM_CIKIS_TUTAR = ISNULL(SUM(ABS(CIKIS_TUTARI)), 0)
            FROM INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_2025 WITH(NOLOCK)
            WHERE MALZEME_KODU = @MalzemeKodu
              AND TARIH >= @Baslangic
              AND TARIH <= @Bitis
            GROUP BY FIS_TURU";

        using var conn = _db.CreateConnection();
        var hareketler = await conn.QueryAsync<dynamic>(sql, new 
        { 
            MalzemeKodu = malzemeKodu, 
            Baslangic = baslangic, 
            Bitis = bitis 
        });

        foreach (var hareket in hareketler)
        {
            string fisTuru = hareket.FIS_TURU?.ToString() ?? "";
            decimal girisMiktar = (decimal)(hareket.TOPLAM_GIRIS ?? 0m);
            decimal girisTutar = (decimal)(hareket.TOPLAM_GIRIS_TUTAR ?? 0m);
            decimal cikisMiktar = (decimal)(hareket.TOPLAM_CIKIS ?? 0m);
            decimal cikisTutar = (decimal)(hareket.TOPLAM_CIKIS_TUTAR ?? 0m);

            if (fisTuru.Contains("Satınalma") || fisTuru.Contains("Satın Alma") || fisTuru.Contains("Alım"))
            {
                ozet.SatinAlmaMiktari += girisMiktar;
                ozet.SatinAlmaTutari += girisTutar;
            }
            else if (fisTuru.Contains("Üretim") || fisTuru.Contains("Üretimden"))
            {
                ozet.UretimMiktari += girisMiktar;
            }
            else if (fisTuru.Contains("Satış İade") || fisTuru.Contains("İade"))
            {
                ozet.IadeMiktari += girisMiktar;
                ozet.IadeTutari += girisTutar;
            }
            else if (fisTuru.Contains("Satış") || fisTuru.Contains("Toptan"))
            {
                ozet.SatisMiktari += cikisMiktar;
                ozet.SatisTutari += cikisTutar;
            }
        }

        return ozet;
    }

    /// <summary>
    /// Alkol malzemesi için özet verileri çeker (INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Alkol_2025 VIEW'ından)
    /// </summary>
    private async Task<YanUrunOzet> GetAlkolMalzemeOzetAsync(string malzemeKodu, DateTime baslangic, DateTime bitis)
    {
        var ozet = new YanUrunOzet { MalzemeKodu = malzemeKodu };

        var sql = @"
            SELECT 
                FIS_TURU,
                TOPLAM_GIRIS = ISNULL(SUM(ABS(GIRIS_MIKTARI)), 0),
                TOPLAM_GIRIS_TUTAR = ISNULL(SUM(ABS(GIRIS_TUTARI)), 0),
                TOPLAM_CIKIS = ISNULL(SUM(ABS(CIKIS_MIKTARI)), 0),
                TOPLAM_CIKIS_TUTAR = ISNULL(SUM(ABS(CIKIS_TUTARI)), 0)
            FROM INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Alkol_2025 WITH(NOLOCK)
            WHERE MALZEME_KODU = @MalzemeKodu
              AND TARIH >= @Baslangic
              AND TARIH <= @Bitis
            GROUP BY FIS_TURU";

        using var conn = _db.CreateConnection();
        var hareketler = await conn.QueryAsync<dynamic>(sql, new 
        { 
            MalzemeKodu = malzemeKodu, 
            Baslangic = baslangic, 
            Bitis = bitis 
        });

        foreach (var hareket in hareketler)
        {
            string fisTuru = hareket.FIS_TURU?.ToString() ?? "";
            decimal girisMiktar = (decimal)(hareket.TOPLAM_GIRIS ?? 0m);
            decimal girisTutar = (decimal)(hareket.TOPLAM_GIRIS_TUTAR ?? 0m);
            decimal cikisMiktar = (decimal)(hareket.TOPLAM_CIKIS ?? 0m);
            decimal cikisTutar = (decimal)(hareket.TOPLAM_CIKIS_TUTAR ?? 0m);

            if (fisTuru.Contains("Satınalma") || fisTuru.Contains("Satın Alma") || fisTuru.Contains("Alım"))
            {
                ozet.SatinAlmaMiktari += girisMiktar;
                ozet.SatinAlmaTutari += girisTutar;
            }
            else if (fisTuru.Contains("Üretim") || fisTuru.Contains("Üretimden"))
            {
                ozet.UretimMiktari += girisMiktar;
            }
            else if (fisTuru.Contains("Satış İade") || fisTuru.Contains("İade"))
            {
                ozet.IadeMiktari += girisMiktar;
                ozet.IadeTutari += girisTutar;
            }
            else if (fisTuru.Contains("Satış") || fisTuru.Contains("Toptan"))
            {
                ozet.SatisMiktari += cikisMiktar;
                ozet.SatisTutari += cikisTutar;
            }
        }

        return ozet;
    }

    /// <summary>
    /// Alkol üretimi için tüketilen melas miktarını hesaplar
    /// FIS_TURU = "Proms. ve Teknik Malzeme" olanların çıkış miktarı
    /// </summary>
    public async Task<decimal> GetAlkolIcinTuketilenMelasAsync(DateTime baslangic, DateTime bitis)
    {
        bitis = SistemTarihi.Clamp(bitis);
        // Melas sarfı - "Proms. ve Teknik Malzeme" fiş türü
        var sql = @"
            SELECT ISNULL(SUM(ABS(CIKIS_MIKTARI)), 0)
            FROM INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Alkol_2025 WITH(NOLOCK)
            WHERE MALZEME_KODU = @MelasKodu
              AND FIS_TURU LIKE '%Proms%'
              AND TARIH >= @Baslangic
              AND TARIH <= @Bitis";

        using var conn = _db.CreateConnection();
        return await conn.ExecuteScalarAsync<decimal>(sql, new 
        { 
            MelasKodu = MalzemeTanimlari.MELAS,
            Baslangic = baslangic, 
            Bitis = bitis 
        });
    }

    /// <summary>
    /// Detaylı stok hareketlerini getirir
    /// </summary>
    public async Task<List<StokHareket>> GetStokHareketleriAsync(
        string malzemeKodu,
        DateTime baslangic,
        DateTime bitis,
        int? fisturuFiltre = null)
    {
        bitis = SistemTarihi.Clamp(bitis);
        var sql = @"
            SELECT 
                Tarih = TARIH,
                FisTuru = FIS_TURU,
                FisNo = FIS_NUMARASI,
                CariKodu = ISNULL(CARI_HESAP_KODU, ''),
                CariAdi = ISNULL(CARI_HESAP_UNVANI, ''),
                MalzemeKodu = MALZEME_KODU,
                MalzemeAdi = MALZEME_ACIKLAMASI,
                GirisMiktari = ISNULL(ABS(GIRIS_MIKTARI), 0),
                GirisTutari = ISNULL(ABS(GIRIS_TUTARI), 0),
                CikisMiktari = ISNULL(ABS(CIKIS_MIKTARI), 0),
                CikisTutari = ISNULL(ABS(CIKIS_TUTARI), 0)
            FROM INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_2025 WITH(NOLOCK)
            WHERE MALZEME_KODU = @MalzemeKodu
              AND TARIH >= @Baslangic
              AND TARIH <= @Bitis
            ORDER BY TARIH DESC";

        using var conn = _db.CreateConnection();
        var sonuc = await conn.QueryAsync<StokHareket>(sql, new 
        { 
            MalzemeKodu = malzemeKodu, 
            Baslangic = baslangic, 
            Bitis = bitis
        });

        return sonuc.ToList();
    }

    /// <summary>
    /// Kategori bazlı toplamları hesaplar (Yaş Küspe, Kuru Küspe toplamları)
    /// </summary>
    public YanUrunOzet HesaplaKategoriToplam(List<YanUrunOzet> ozetler, string kategori)
    {
        var kategoridekiler = ozetler.Where(x => x.Kategori == kategori).ToList();
        
        return new YanUrunOzet
        {
            MalzemeKodu = "",
            MalzemeAdi = kategori switch
            {
                "YAS_KUSPE" => "YAŞ KÜSPE TOPLAM",
                "KURU_KUSPE" => "KURU KÜSPE TOPLAM",
                "MELAS" => "MELAS TOPLAM",
                _ => $"{kategori} TOPLAM"
            },
            Kategori = kategori,
            DevirStok = kategoridekiler.Sum(x => x.DevirStok),
            SatinAlmaMiktari = kategoridekiler.Sum(x => x.SatinAlmaMiktari),
            SatinAlmaTutari = kategoridekiler.Sum(x => x.SatinAlmaTutari),
            UretimMiktari = kategoridekiler.Sum(x => x.UretimMiktari),
            SatisMiktari = kategoridekiler.Sum(x => x.SatisMiktari),
            SatisTutari = kategoridekiler.Sum(x => x.SatisTutari),
            IadeMiktari = kategoridekiler.Sum(x => x.IadeMiktari),
            IadeTutari = kategoridekiler.Sum(x => x.IadeTutari),
            TuketimMiktari = kategoridekiler.Sum(x => x.TuketimMiktari)
        };
    }
}
