using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Şeker Dairesi rapor servisi
/// VBA Module2 (StokSorgula) ve Module3 (SadeSekerAnaliziYap) makrolarının C# karşılığı
/// Veri kaynağı: INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_2025 ve LV_211_01_STINVTOT
/// </summary>
public class SekerDairesiService
{
    private readonly DatabaseService _db;

    // Hariç tutulan ürün kodları (VBA'da sifreliKodlar içinde)
    private static readonly HashSet<string> HaricKodlar = new(StringComparer.OrdinalIgnoreCase)
    {
        "T.T.0.0.0", "T.S.0.0.0",
        "T.S.9.1.03.1.1000.20", "T.S.9.1.03.1.3000.06",
        "T.S.9.1.03.1.5000.04", "T.T.9.1.03.1.5000.04"
    };

    public SekerDairesiService(DatabaseService db)
    {
        _db = db;
    }

    // ─────────────────────────────────────────────────────────────────────
    // STOK SORGULA  (VBA Module2 – CommandButton2)
    // ─────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Verilen tarihe ait ambar/malzeme bazlı şeker stok sorgusunu çalıştırır.
    /// VBA StokSorgula makrosunun birebir SQL portu.
    /// </summary>
    public async Task<(List<StokDetayRow> Detay, List<StokOzetRow> Ozet)> GetStokSorgulaAsync(DateTime tarih)
    {
        var sql = @"
            SELECT
                QRY.AMBAR_NO,
                QRY.AMBAR,
                QRY.MALZEME_KODU,
                QRY.MALZEME_ADI,
                QRY.ANA_BIRIM,
                QRY.STOK,
                KG = QRY.STOK * ITMUKG.CONVFACT1 / ITMUKG.CONVFACT2
            FROM (
                SELECT
                    ITEMREF     = ITM.LOGICALREF,
                    UNITSETREF  = ITM.UNITSETREF,
                    AMBAR_NO    = GNT.INVENNO,
                    AMBAR       = WHO.NAME,
                    MALZEME_KODU = ITM.CODE,
                    MALZEME_ADI  = ITM.NAME,
                    ANA_BIRIM    = UNI.CODE,
                    STOK         = ROUND(SUM(GNT.ONHAND), 2)
                FROM LV_211_01_STINVTOT GNT (NOLOCK)
                LEFT JOIN LG_211_ITEMS   ITM (NOLOCK) ON ITM.LOGICALREF = GNT.STOCKREF
                LEFT JOIN L_CAPIWHOUSE   WHO (NOLOCK) ON WHO.FIRMNR = 211 AND NR = GNT.INVENNO
                LEFT JOIN LG_211_UNITSETL UNI (NOLOCK) ON UNI.UNITSETREF = ITM.UNITSETREF AND UNI.LINENR = 1
                WHERE GNT.INVENNO <> -1
                  AND ITM.STGRPCODE IN (N'ŞEKER', N'SEKERFASON')
                  AND ITM.CODE NOT IN (
                      N'T.T.0.0.0', N'T.S.0.0.0',
                      N'T.S.9.1.03.1.1000.20', N'T.S.9.1.03.1.3000.06',
                      N'T.S.9.1.03.1.5000.04', N'T.T.9.1.03.1.5000.04'
                  )
                  AND GNT.DATE_ <= @Tarih
                GROUP BY ITM.LOGICALREF, ITM.UNITSETREF, GNT.INVENNO,
                         WHO.NAME, ITM.CODE, ITM.NAME, UNI.CODE
            ) QRY
            LEFT JOIN LG_211_ITMUNITA ITMUKG WITH(NOLOCK)
                ON ITMUKG.ITEMREF = QRY.ITEMREF
               AND UNITLINEREF = (
                    SELECT LOGICALREF FROM LG_211_UNITSETL UNI WITH(NOLOCK)
                    WHERE UNI.UNITSETREF = QRY.UNITSETREF AND UNI.CODE = N'KG'
               )
            WHERE STOK <> 0
            ORDER BY QRY.AMBAR_NO, QRY.MALZEME_KODU";

        using var conn = _db.CreateConnection();
        var detayList = (await conn.QueryAsync<StokDetayRow>(sql, new { Tarih = tarih.Date })).ToList();

        // Malzeme bazlı özet (VBA OzetSayfaOlustur karşılığı)
        var ozetList = detayList
            .Where(r => r.Kg.HasValue)
            .GroupBy(r => new { r.MalzemeKodu, r.MalzemeAdi })
            .Select(g => new StokOzetRow
            {
                MalzemeKodu = g.Key.MalzemeKodu,
                MalzemeAdi  = g.Key.MalzemeAdi,
                ToplamKg    = g.Sum(r => r.Kg ?? 0m)
            })
            .OrderBy(r => r.MalzemeKodu)
            .ToList();

        return (detayList, ozetList);
    }

    // ─────────────────────────────────────────────────────────────────────
    // SADE ŞEKER ANALİZİ  (VBA Module3 – CommandButton1)
    // ─────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Tarih aralığındaki şeker hareketlerini A/B/C Kotası, Paketli ve Konya Ticari
    /// kategorilerine göre analiz eder.
    /// Dönem başı stok değerleri kullanıcı tarafından girilir (VBA'daki InputBox'lar).
    /// </summary>
    // ── Kampanya başı devir stokları (01.09.2025 sabahı itibarıyla)
    public static readonly DateTime KampanyaBaslangic = new DateTime(2025, 9, 1);

    private static readonly Dictionary<string, decimal> KampanyaDonemBasi = new()
    {
        ["A_KOTASI"]       = 0m,
        ["B_KOTASI"]       = 6_501_000m,
        ["C_KOTASI"]       = 3_075_811m,
        ["PAKETLI"]        = 0m,
        ["KONYA_TICARI"]   = 0m,
        ["TICARI_KRISTAL"] = 792_780m,
        ["TICARI_PAKET"]   = 50_082m,
    };

    public async Task<(List<SekerKategoriAnaliz> Analiz, List<SatisIadeDipnot> Dipnotlar)> GetSadeSekerAnaliziAsync(
        DateTime baslangic,
        DateTime bitis)
    {
        bitis = SistemTarihi.Clamp(bitis);
        // Dönem başı stoklarını belirle:
        // – Eğer seçilen başlangıç kampanya başıysa → sabit değerler kullan
        // – Değilse → 01.09.2025'ten (baslangic-1) arası çalıştır, dönem sonunu al
        Dictionary<string, decimal> donemBasi;

        if (baslangic.Date <= KampanyaBaslangic.Date)
        {
            donemBasi = new Dictionary<string, decimal>(KampanyaDonemBasi);
        }
        else
        {
            // 01.09.2025 → baslangic-1 dönemini hesapla (LOGO bazlı, üst tablo ile tutarlı)
            var oncekiSonuc = await HesaplaDonemSonuAsync(KampanyaBaslangic, baslangic.AddDays(-1), KampanyaDonemBasi);
            donemBasi = oncekiSonuc.ToDictionary(
                k => k.Kategori,
                k => k.DonemSonuMiktar);
        }
        var sql = @"
            SELECT
                TARIH,
                FIS_TURU,
                MALZEME_KODU,
                GIRIS_MIKTARI     = ISNULL(GIRIS_MIKTAR_KG,  0),
                CIKIS_MIKTARI     = ISNULL(CIKIS_MIKTARI_KG, 0),
                GIRIS_TUTARI      = ISNULL(GIRIS_TUTARI,      0),
                CIKIS_TUTARI      = ISNULL(CIKIS_TUTARI,      0)
            FROM INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Seker_2025 WITH(NOLOCK)
            WHERE TARIH >= @Baslangic
              AND TARIH <= @Bitis
            ORDER BY TARIH ASC";

        using var conn = _db.CreateConnection();
        var hareketler = await conn.QueryAsync<dynamic>(sql, new { Baslangic = baslangic.Date, Bitis = bitis.Date });

        // Kategori sözlüğü – devir stokları dinamik olarak belirlendi
        var sonuc = new Dictionary<string, SekerKategoriAnaliz>
        {
            ["A_KOTASI"]       = new() { Kategori = "A_KOTASI",       KategoriAdi = "A Kotası Şeker",             DonemBasiMiktar = donemBasi.GetValueOrDefault("A_KOTASI") },
            ["B_KOTASI"]       = new() { Kategori = "B_KOTASI",       KategoriAdi = "B Kotası Şeker",             DonemBasiMiktar = donemBasi.GetValueOrDefault("B_KOTASI") },
            ["C_KOTASI"]       = new() { Kategori = "C_KOTASI",       KategoriAdi = "C Kotası Şeker",             DonemBasiMiktar = donemBasi.GetValueOrDefault("C_KOTASI") },
            ["PAKETLI"]        = new() { Kategori = "PAKETLI",        KategoriAdi = "Paketli Şeker",              DonemBasiMiktar = donemBasi.GetValueOrDefault("PAKETLI") },
            ["KONYA_TICARI"]   = new() { Kategori = "KONYA_TICARI",   KategoriAdi = "Konya Şeker Ticari Mal",     DonemBasiMiktar = donemBasi.GetValueOrDefault("KONYA_TICARI") },
            ["TICARI_KRISTAL"] = new() { Kategori = "TICARI_KRISTAL", KategoriAdi = "Ticari Mal Kristal Toz Şeker", DonemBasiMiktar = donemBasi.GetValueOrDefault("TICARI_KRISTAL") },
            ["TICARI_PAKET"]   = new() { Kategori = "TICARI_PAKET",   KategoriAdi = "Ticari Mal Paket Şeker",     DonemBasiMiktar = donemBasi.GetValueOrDefault("TICARI_PAKET") },
        };

        // Ticari stok takibi: satışları/sarfları doğru kategoriye yönlendirmek için
        decimal ticariKristalMevcut = donemBasi.GetValueOrDefault("TICARI_KRISTAL");
        decimal ticariPaketMevcut   = donemBasi.GetValueOrDefault("TICARI_PAKET");

        // A Kotası giriş takibi: Paket Şeker üretimini yönlendirmek için
        // (Dönem Başı + A Kotası üretim + A Kotası satış iade toplamı)
        decimal aKotasiGirisMevcut = donemBasi.GetValueOrDefault("A_KOTASI");

        foreach (var h in hareketler)
        {
            string fisTuru    = h.FIS_TURU?.ToString() ?? "";
            string malzemeKod = h.MALZEME_KODU?.ToString() ?? "";
            decimal girisKg   = Convert.ToDecimal(h.GIRIS_MIKTARI ?? 0m);
            decimal cikisKg   = Convert.ToDecimal(h.CIKIS_MIKTARI ?? 0m);
            decimal girisTl   = Convert.ToDecimal(h.GIRIS_TUTARI  ?? 0m);
            decimal cikisTl   = Convert.ToDecimal(h.CIKIS_TUTARI  ?? 0m);

            // Hariç tutulan kodları atla
            if (HaricKodlar.Contains(malzemeKod)) continue;

            string kategori = MalzemeKategorisi("", malzemeKod);
            if (!sonuc.TryGetValue(kategori, out var kat)) continue;

            string operasyon = FisTuruOperasyon(fisTuru);

            // Satın alma → Ticari Mal kategorilerine yönlendir, stok sayacını güncelle
            if (operasyon == "SATINALMA")
            {
                string hedef = kategori == "A_KOTASI" ? "TICARI_KRISTAL"
                             : kategori == "PAKETLI"  ? "TICARI_PAKET"
                             : kategori;
                if (sonuc.TryGetValue(hedef, out var satinalmaKat))
                {
                    satinalmaKat.SatinAlmaMiktar += girisKg;
                    satinalmaKat.SatinAlmaTutar  += girisTl;
                    if (hedef == "TICARI_KRISTAL") ticariKristalMevcut += girisKg;
                    else if (hedef == "TICARI_PAKET") ticariPaketMevcut += girisKg;
                }
                continue;
            }

            // Satış + Sarf + Yemekhane → Ticari Mal stoğu varsa önce oradan, bittikten sonra ana kategoriden
            if (operasyon == "SATIS" || operasyon == "SARF" || operasyon == "YEMEKHANE")
            {
                if (kategori == "A_KOTASI" && ticariKristalMevcut > 0)
                {
                    decimal ticariPayi  = Math.Min(cikisKg, ticariKristalMevcut);
                    decimal kotaPayi    = cikisKg - ticariPayi;
                    decimal ticariTutar = cikisKg > 0 ? cikisTl * ticariPayi / cikisKg : 0;
                    decimal kotaTutar   = cikisKg > 0 ? cikisTl * kotaPayi   / cikisKg : 0;
                    switch (operasyon)
                    {
                        case "SATIS":
                            sonuc["TICARI_KRISTAL"].SatisMiktar += ticariPayi; sonuc["TICARI_KRISTAL"].SatisTutar += ticariTutar;
                            if (kotaPayi > 0) { sonuc["A_KOTASI"].SatisMiktar += kotaPayi; sonuc["A_KOTASI"].SatisTutar += kotaTutar; }
                            break;
                        case "SARF":
                            sonuc["TICARI_KRISTAL"].SarfMiktar += ticariPayi; sonuc["TICARI_KRISTAL"].SarfTutar += ticariTutar;
                            if (kotaPayi > 0) { sonuc["A_KOTASI"].SarfMiktar += kotaPayi; sonuc["A_KOTASI"].SarfTutar += kotaTutar; }
                            break;
                        case "YEMEKHANE":
                            sonuc["TICARI_KRISTAL"].YemekhaneMiktar += ticariPayi; sonuc["TICARI_KRISTAL"].YemekhaneTutar += ticariTutar;
                            if (kotaPayi > 0) { sonuc["A_KOTASI"].YemekhaneMiktar += kotaPayi; sonuc["A_KOTASI"].YemekhaneTutar += kotaTutar; }
                            break;
                    }
                    ticariKristalMevcut -= ticariPayi;
                    continue;
                }
                if (kategori == "PAKETLI" && ticariPaketMevcut > 0)
                {
                    decimal ticariPayi  = Math.Min(cikisKg, ticariPaketMevcut);
                    decimal kotaPayi    = cikisKg - ticariPayi;
                    decimal ticariTutar = cikisKg > 0 ? cikisTl * ticariPayi / cikisKg : 0;
                    decimal kotaTutar   = cikisKg > 0 ? cikisTl * kotaPayi   / cikisKg : 0;
                    switch (operasyon)
                    {
                        case "SATIS":
                            sonuc["TICARI_PAKET"].SatisMiktar += ticariPayi; sonuc["TICARI_PAKET"].SatisTutar += ticariTutar;
                            if (kotaPayi > 0) { sonuc["PAKETLI"].SatisMiktar += kotaPayi; sonuc["PAKETLI"].SatisTutar += kotaTutar; }
                            break;
                        case "SARF":
                            sonuc["TICARI_PAKET"].SarfMiktar += ticariPayi; sonuc["TICARI_PAKET"].SarfTutar += ticariTutar;
                            if (kotaPayi > 0) { sonuc["PAKETLI"].SarfMiktar += kotaPayi; sonuc["PAKETLI"].SarfTutar += kotaTutar; }
                            break;
                        case "YEMEKHANE":
                            sonuc["TICARI_PAKET"].YemekhaneMiktar += ticariPayi; sonuc["TICARI_PAKET"].YemekhaneTutar += ticariTutar;
                            if (kotaPayi > 0) { sonuc["PAKETLI"].YemekhaneMiktar += kotaPayi; sonuc["PAKETLI"].YemekhaneTutar += kotaTutar; }
                            break;
                    }
                    ticariPaketMevcut -= ticariPayi;
                    continue;
                }
                // Ticari stok tükenmiş veya diğer kategoriler → doğrudan kat'a
                switch (operasyon)
                {
                    case "SATIS":     kat.SatisMiktar     += cikisKg; kat.SatisTutar     += cikisTl; break;
                    case "SARF":      kat.SarfMiktar      += cikisKg; kat.SarfTutar      += cikisTl; break;
                    case "YEMEKHANE": kat.YemekhaneMiktar += cikisKg; kat.YemekhaneTutar += cikisTl; break;
                }
                continue;
            }

            switch (operasyon)
            {
                case "URETIM":
                    if (kategori == "A_KOTASI")
                    {
                        aKotasiGirisMevcut += girisKg;
                        kat.UretimMiktar   += girisKg;
                        kat.UretimTutar    += girisTl;
                    }
                    else if (kategori == "PAKETLI")
                    {
                        decimal kotaPayi   = Math.Min(girisKg, aKotasiGirisMevcut);
                        decimal ticariPayi = girisKg - kotaPayi;
                        if (kotaPayi > 0)
                        {
                            sonuc["PAKETLI"].UretimMiktar     += kotaPayi;
                            sonuc["PAKETLI"].UretimTutar      += girisKg > 0 ? girisTl * kotaPayi / girisKg : 0;
                            aKotasiGirisMevcut -= kotaPayi;
                        }
                        if (ticariPayi > 0)
                        {
                            sonuc["TICARI_PAKET"].UretimMiktar += ticariPayi;
                            sonuc["TICARI_PAKET"].UretimTutar  += girisKg > 0 ? girisTl * ticariPayi / girisKg : 0;
                            ticariPaketMevcut += ticariPayi; // Üretim satış yönlendirme sayacına da eklenir
                        }
                    }
                    else
                    {
                        kat.UretimMiktar += girisKg;
                        kat.UretimTutar  += girisTl;
                    }
                    break;
                case "SATIS_IADE":
                    if (kategori == "A_KOTASI") aKotasiGirisMevcut += girisKg;
                    kat.SatisIadeMiktar += girisKg;
                    kat.SatisIadeTutar  += girisTl;
                    break;
                case "HAMMADDE_GIRIS":
                    // Ayrı sütun gösterilmez ama toplam stoka dahildir
                    kat.HammaddeGirisMiktar += girisKg;
                    kat.HammaddeGirisTutar  += girisTl;
                    break;
                case "RECETE_FARK_GIRIS":
                    kat.ReceteFarkMiktar += girisKg;
                    kat.ReceteFarkTutar  += girisTl;
                    break;
                case "SAYIM_FAZLASI":
                    kat.SayimFazlasiMiktar += girisKg;
                    kat.SayimFazlasiTutar  += girisTl;
                    break;
                case "FIRE":
                    kat.FireMiktar += cikisKg;
                    kat.FireTutar  += cikisTl;
                    break;
                case "PROMS":
                    kat.PromsMiktar += cikisKg;
                    kat.PromsTutar  += cikisTl;
                    break;
                case "HAMMADDE_CIKIS":
                    kat.HammaddeCikisMiktar += cikisKg;
                    kat.HammaddeCikisTutar  += cikisTl;
                    break;
                case "SATINALMA_IADE":
                    kat.SatinAlmaIadeMiktar += cikisKg;
                    kat.SatinAlmaIadeTutar  += cikisTl;
                    break;
                // case "DIGER": yoksay
            }
        }

        // ── Sonraki ay tamamlandıysa Toptan Satış İadelerini mevcut döneme yansıt ──────
        var dipnotlar = new List<SatisIadeDipnot>();
        var nextAyIlk = new DateTime(bitis.AddMonths(1).Year, bitis.AddMonths(1).Month, 1);
        var nextAySon = new DateTime(nextAyIlk.Year, nextAyIlk.Month, DateTime.DaysInMonth(nextAyIlk.Year, nextAyIlk.Month));

        if (DateTime.Today > nextAySon) // Sonraki ay tamamen geçmiş
        {
            var iadeler = await GetNextAySatisIadeleriAsync(nextAyIlk, nextAySon);
            string ayAdi = nextAyIlk.ToString("MMMM yyyy", new System.Globalization.CultureInfo("tr-TR"));

            // Kategori bazında topla
            var gruplar = iadeler
                .GroupBy(x => x.Kategori)
                .Select(g => (Kategori: g.Key, Miktar: g.Sum(x => x.Miktar), Tutar: g.Sum(x => x.Tutar)))
                .ToList();

            foreach (var (kategori, miktar, tutar) in gruplar)
            {
                if (miktar <= 0) continue;

                // Mevcut dönemde bu kategoriden satış var mı?
                bool satisVar = sonuc.TryGetValue(kategori, out var katData) && katData.SatisMiktar > 0;

                string hedef;
                if (satisVar)
                {
                    hedef = kategori;
                }
                else
                {
                    // Satış yoksa → kristal kategoriler → TICARI_KRISTAL, paket → TICARI_PAKET
                    bool isPaket = kategori == "PAKETLI" || kategori == "TICARI_PAKET";
                    hedef = isPaket ? "TICARI_PAKET" : "TICARI_KRISTAL";
                }

                if (sonuc.TryGetValue(hedef, out var hedefKat))
                {
                    // sonuc'u (üst tablo) değiştirme — iade düzeltmesi sadece alt tablo için
                    sonuc.TryGetValue(kategori, out var kSrc);
                    dipnotlar.Add(new SatisIadeDipnot
                    {
                        KaynakKategori       = kategori,
                        KaynakKategoriAdi    = kSrc?.KategoriAdi ?? kategori,
                        HedefKategori        = hedef,
                        HedefKategoriAdi     = hedefKat.KategoriAdi,
                        Miktar               = miktar,
                        Tutar                = tutar,
                        SonrakiAyAdi         = ayAdi,
                        Yonlendirildi        = hedef != kategori
                    });
                }
            }
        }

        return (sonuc.Values.ToList(), dipnotlar);
    }

    /// <summary>
    /// Belirli tarihlerdeki Toptan Satış İade İrsaliyelerini kategori bazında döndürür.
    /// </summary>
    private async Task<List<(string Kategori, decimal Miktar, decimal Tutar)>> GetNextAySatisIadeleriAsync(
        DateTime baslangic, DateTime bitis)
    {
        var sql = @"
            SELECT MALZEME_KODU,
                   GIRIS_MIKTARI = ISNULL(GIRIS_MIKTAR_KG, 0),
                   GIRIS_TUTARI  = ISNULL(GIRIS_TUTARI,    0)
            FROM INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Seker_2025 WITH(NOLOCK)
            WHERE TARIH    >= @Baslangic
              AND TARIH    <= @Bitis
              AND FIS_TURU  = N'Toptan Satış İade İrsaliyesi'";

        using var conn = _db.CreateConnection();
        var rows = await conn.QueryAsync<dynamic>(sql, new { Baslangic = baslangic.Date, Bitis = bitis.Date });

        var result = new List<(string Kategori, decimal Miktar, decimal Tutar)>();
        foreach (var r in rows)
        {
            string malzemeKod = r.MALZEME_KODU?.ToString() ?? "";
            if (HaricKodlar.Contains(malzemeKod)) continue;
            string kategori = MalzemeKategorisi("", malzemeKod);
            decimal miktar  = Convert.ToDecimal(r.GIRIS_MIKTARI ?? 0m);
            decimal tutar   = Convert.ToDecimal(r.GIRIS_TUTARI  ?? 0m);
            result.Add((kategori, miktar, tutar));
        }
        return result;
    }

    /// <summary>
    /// Belirli bir dönem için hesaplama yapar — dönem başı stok aktarımı için kullanılır.
    /// </summary>
    private async Task<List<SekerKategoriAnaliz>> HesaplaDonemSonuAsync(
        DateTime baslangic, DateTime bitis, Dictionary<string, decimal> donemBasi)
    {
        var sql = @"
            SELECT TARIH, FIS_TURU, MALZEME_KODU,
                GIRIS_MIKTARI = ISNULL(GIRIS_MIKTAR_KG,  0),
                CIKIS_MIKTARI = ISNULL(CIKIS_MIKTARI_KG, 0),
                GIRIS_TUTARI  = ISNULL(GIRIS_TUTARI,      0),
                CIKIS_TUTARI  = ISNULL(CIKIS_TUTARI,      0)
            FROM INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Seker_2025 WITH(NOLOCK)
            WHERE TARIH >= @Baslangic AND TARIH <= @Bitis
            ORDER BY TARIH ASC";

        var ara = new Dictionary<string, SekerKategoriAnaliz>
        {
            ["A_KOTASI"]       = new() { Kategori = "A_KOTASI",       KategoriAdi = "A Kotası Şeker",             DonemBasiMiktar = donemBasi.GetValueOrDefault("A_KOTASI") },
            ["B_KOTASI"]       = new() { Kategori = "B_KOTASI",       KategoriAdi = "B Kotası Şeker",             DonemBasiMiktar = donemBasi.GetValueOrDefault("B_KOTASI") },
            ["C_KOTASI"]       = new() { Kategori = "C_KOTASI",       KategoriAdi = "C Kotası Şeker",             DonemBasiMiktar = donemBasi.GetValueOrDefault("C_KOTASI") },
            ["PAKETLI"]        = new() { Kategori = "PAKETLI",        KategoriAdi = "Paketli Şeker",              DonemBasiMiktar = donemBasi.GetValueOrDefault("PAKETLI") },
            ["KONYA_TICARI"]   = new() { Kategori = "KONYA_TICARI",   KategoriAdi = "Konya Şeker Ticari Mal",     DonemBasiMiktar = donemBasi.GetValueOrDefault("KONYA_TICARI") },
            ["TICARI_KRISTAL"] = new() { Kategori = "TICARI_KRISTAL", KategoriAdi = "Ticari Mal Kristal Toz Şeker", DonemBasiMiktar = donemBasi.GetValueOrDefault("TICARI_KRISTAL") },
            ["TICARI_PAKET"]   = new() { Kategori = "TICARI_PAKET",   KategoriAdi = "Ticari Mal Paket Şeker",     DonemBasiMiktar = donemBasi.GetValueOrDefault("TICARI_PAKET") },
        };

        using var conn = _db.CreateConnection();
        var hareketler = await conn.QueryAsync<dynamic>(sql, new { Baslangic = baslangic.Date, Bitis = bitis.Date });

        decimal ticariKristalMevcut = donemBasi.GetValueOrDefault("TICARI_KRISTAL");
        decimal ticariPaketMevcut   = donemBasi.GetValueOrDefault("TICARI_PAKET");
        decimal aKotasiGirisMevcut  = donemBasi.GetValueOrDefault("A_KOTASI");

        foreach (var h in hareketler)
        {
            string fisTuru    = h.FIS_TURU?.ToString() ?? "";
            string malzemeKod = h.MALZEME_KODU?.ToString() ?? "";
            decimal girisKg   = Convert.ToDecimal(h.GIRIS_MIKTARI ?? 0m);
            decimal cikisKg   = Convert.ToDecimal(h.CIKIS_MIKTARI ?? 0m);
            decimal girisTl   = Convert.ToDecimal(h.GIRIS_TUTARI  ?? 0m);
            decimal cikisTl   = Convert.ToDecimal(h.CIKIS_TUTARI  ?? 0m);

            if (HaricKodlar.Contains(malzemeKod)) continue;

            string kategori  = MalzemeKategorisi("", malzemeKod);
            if (!ara.TryGetValue(kategori, out var kat)) continue;

            string operasyon = FisTuruOperasyon(fisTuru);

            // Satın alma → Ticari Mal kategorilerine yönlendir
            if (operasyon == "SATINALMA")
            {
                string hedef = kategori == "A_KOTASI" ? "TICARI_KRISTAL"
                             : kategori == "PAKETLI"  ? "TICARI_PAKET"
                             : kategori;
                if (ara.TryGetValue(hedef, out var satinalmaKat))
                {
                    satinalmaKat.SatinAlmaMiktar += girisKg;
                    satinalmaKat.SatinAlmaTutar  += girisTl;
                    if (hedef == "TICARI_KRISTAL") ticariKristalMevcut += girisKg;
                    else if (hedef == "TICARI_PAKET") ticariPaketMevcut += girisKg;
                }
                continue;
            }

            // Satış + Sarf + Yemekhane → Ticari Mal stoğu varsa önce oradan, bittikten sonra ana kategoriden
            if (operasyon == "SATIS" || operasyon == "SARF" || operasyon == "YEMEKHANE")
            {
                if (kategori == "A_KOTASI" && ticariKristalMevcut > 0)
                {
                    decimal ticariPayi  = Math.Min(cikisKg, ticariKristalMevcut);
                    decimal kotaPayi    = cikisKg - ticariPayi;
                    decimal ticariTutar = cikisKg > 0 ? cikisTl * ticariPayi / cikisKg : 0;
                    decimal kotaTutar   = cikisKg > 0 ? cikisTl * kotaPayi   / cikisKg : 0;
                    switch (operasyon)
                    {
                        case "SATIS":
                            ara["TICARI_KRISTAL"].SatisMiktar += ticariPayi; ara["TICARI_KRISTAL"].SatisTutar += ticariTutar;
                            if (kotaPayi > 0) { ara["A_KOTASI"].SatisMiktar += kotaPayi; ara["A_KOTASI"].SatisTutar += kotaTutar; }
                            break;
                        case "SARF":
                            ara["TICARI_KRISTAL"].SarfMiktar += ticariPayi; ara["TICARI_KRISTAL"].SarfTutar += ticariTutar;
                            if (kotaPayi > 0) { ara["A_KOTASI"].SarfMiktar += kotaPayi; ara["A_KOTASI"].SarfTutar += kotaTutar; }
                            break;
                        case "YEMEKHANE":
                            ara["TICARI_KRISTAL"].YemekhaneMiktar += ticariPayi; ara["TICARI_KRISTAL"].YemekhaneTutar += ticariTutar;
                            if (kotaPayi > 0) { ara["A_KOTASI"].YemekhaneMiktar += kotaPayi; ara["A_KOTASI"].YemekhaneTutar += kotaTutar; }
                            break;
                    }
                    ticariKristalMevcut -= ticariPayi;
                    continue;
                }
                if (kategori == "PAKETLI" && ticariPaketMevcut > 0)
                {
                    decimal ticariPayi  = Math.Min(cikisKg, ticariPaketMevcut);
                    decimal kotaPayi    = cikisKg - ticariPayi;
                    decimal ticariTutar = cikisKg > 0 ? cikisTl * ticariPayi / cikisKg : 0;
                    decimal kotaTutar   = cikisKg > 0 ? cikisTl * kotaPayi   / cikisKg : 0;
                    switch (operasyon)
                    {
                        case "SATIS":
                            ara["TICARI_PAKET"].SatisMiktar += ticariPayi; ara["TICARI_PAKET"].SatisTutar += ticariTutar;
                            if (kotaPayi > 0) { ara["PAKETLI"].SatisMiktar += kotaPayi; ara["PAKETLI"].SatisTutar += kotaTutar; }
                            break;
                        case "SARF":
                            ara["TICARI_PAKET"].SarfMiktar += ticariPayi; ara["TICARI_PAKET"].SarfTutar += ticariTutar;
                            if (kotaPayi > 0) { ara["PAKETLI"].SarfMiktar += kotaPayi; ara["PAKETLI"].SarfTutar += kotaTutar; }
                            break;
                        case "YEMEKHANE":
                            ara["TICARI_PAKET"].YemekhaneMiktar += ticariPayi; ara["TICARI_PAKET"].YemekhaneTutar += ticariTutar;
                            if (kotaPayi > 0) { ara["PAKETLI"].YemekhaneMiktar += kotaPayi; ara["PAKETLI"].YemekhaneTutar += kotaTutar; }
                            break;
                    }
                    ticariPaketMevcut -= ticariPayi;
                    continue;
                }
                switch (operasyon)
                {
                    case "SATIS":     kat.SatisMiktar     += cikisKg; kat.SatisTutar     += cikisTl; break;
                    case "SARF":      kat.SarfMiktar      += cikisKg; kat.SarfTutar      += cikisTl; break;
                    case "YEMEKHANE": kat.YemekhaneMiktar += cikisKg; kat.YemekhaneTutar += cikisTl; break;
                }
                continue;
            }

            switch (operasyon)
            {
                case "URETIM":
                    if (kategori == "A_KOTASI")
                    {
                        aKotasiGirisMevcut += girisKg;
                        kat.UretimMiktar   += girisKg;
                        kat.UretimTutar    += girisTl;
                    }
                    else if (kategori == "PAKETLI")
                    {
                        decimal kotaPayi   = Math.Min(girisKg, aKotasiGirisMevcut);
                        decimal ticariPayi = girisKg - kotaPayi;
                        if (kotaPayi > 0)
                        {
                            ara["PAKETLI"].UretimMiktar     += kotaPayi;
                            ara["PAKETLI"].UretimTutar      += girisKg > 0 ? girisTl * kotaPayi / girisKg : 0;
                            aKotasiGirisMevcut -= kotaPayi;
                        }
                        if (ticariPayi > 0)
                        {
                            ara["TICARI_PAKET"].UretimMiktar += ticariPayi;
                            ara["TICARI_PAKET"].UretimTutar  += girisKg > 0 ? girisTl * ticariPayi / girisKg : 0;
                            ticariPaketMevcut += ticariPayi; // Üretim satış yönlendirme sayacına da eklenir
                        }
                    }
                    else
                    {
                        kat.UretimMiktar += girisKg;
                        kat.UretimTutar  += girisTl;
                    }
                    break;
                case "SATIS_IADE":
                    if (kategori == "A_KOTASI") aKotasiGirisMevcut += girisKg;
                    kat.SatisIadeMiktar += girisKg;
                    kat.SatisIadeTutar  += girisTl;
                    break;
                // HAMMADDE_GIRIS hesaplamaya dahil edilmiyor
                case "RECETE_FARK_GIRIS": kat.ReceteFarkMiktar += girisKg; kat.ReceteFarkTutar   += girisTl; break;
                case "SAYIM_FAZLASI":   kat.SayimFazlasiMiktar += girisKg; kat.SayimFazlasiTutar += girisTl; break;
                case "FIRE":            kat.FireMiktar         += cikisKg; kat.FireTutar         += cikisTl; break;
                case "PROMS":           kat.PromsMiktar        += cikisKg; kat.PromsTutar        += cikisTl; break;
                case "HAMMADDE_CIKIS":  kat.HammaddeCikisMiktar+= cikisKg; kat.HammaddeCikisTutar+= cikisTl; break;
                case "SATINALMA_IADE":  kat.SatinAlmaIadeMiktar+= cikisKg; kat.SatinAlmaIadeTutar+= cikisTl; break;
            }
        }

        return ara.Values.ToList();
    }

    // ─────────────────────────────────────────────────────────────────────
    // BAŞKANLIK AY SONU STOK  (portal ile uyumlu dönem sonu hesabı)
    // ─────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Bir ay için hesaplanan LOGO dönem sonu stoklarını Başkanlık tablo formatına
    /// dönüştürür ve sonraki ay iadelerini (dipnotlar) stoka ekler.
    /// KONYA_TICARI stoğu TICARI_KRISTAL'e dahil edilir (portaldaki gibi).
    /// iadePerHedef: { "A_KOTASI" → 62000, "TICARI_KRISTAL" → 0, ... } gibi
    /// </summary>
    private static Dictionary<string, decimal> HesaplaBaskanlikDonemSonu(
        Dictionary<string, SekerKategoriAnaliz> sonuc,
        Dictionary<string, decimal> iadePerHedef)
    {
        sonuc.TryGetValue("A_KOTASI",       out var a);
        sonuc.TryGetValue("B_KOTASI",       out var b);
        sonuc.TryGetValue("C_KOTASI",       out var c);
        sonuc.TryGetValue("PAKETLI",        out var pak);
        sonuc.TryGetValue("TICARI_KRISTAL", out var tk);
        sonuc.TryGetValue("KONYA_TICARI",   out var knya);
        sonuc.TryGetValue("TICARI_PAKET",   out var tpak);

        return new Dictionary<string, decimal>
        {
            ["A_KOTASI"]       = (a?.DonemSonuMiktar    ?? 0m) + iadePerHedef.GetValueOrDefault("A_KOTASI"),
            ["PAKETLI"]        = (pak?.DonemSonuMiktar   ?? 0m) + iadePerHedef.GetValueOrDefault("PAKETLI"),
            ["B_KOTASI"]       = (b?.DonemSonuMiktar    ?? 0m) + iadePerHedef.GetValueOrDefault("B_KOTASI"),
            ["C_KOTASI"]       = (c?.DonemSonuMiktar    ?? 0m) + iadePerHedef.GetValueOrDefault("C_KOTASI"),
            // KONYA stoğu TICARI_KRISTAL'e dahil edilir (portal görünümüyle uyumlu)
            ["TICARI_KRISTAL"] = (tk?.DonemSonuMiktar   ?? 0m) + (knya?.DonemSonuMiktar ?? 0m)
                                 + iadePerHedef.GetValueOrDefault("TICARI_KRISTAL"),
            ["KONYA_TICARI"]   = 0m,
            ["TICARI_PAKET"]   = (tpak?.DonemSonuMiktar ?? 0m) + iadePerHedef.GetValueOrDefault("TICARI_PAKET"),
        };
    }

    /// <summary>
    /// Kampanya başından verilen bitiş tarihine kadar ay ay işleyerek her ay sonu
    /// Başkanlık dönem sonu hesaplar.
    /// Her ay için sonraki ayın satış iadelerini de stoka ekler (portal mantığıyla uyumlu).
    /// </summary>
    private async Task<Dictionary<string, decimal>> HesaplaDonemSonuBaskanlikZincirliAsync(
        DateTime baslangic, DateTime bitis, Dictionary<string, decimal> baslangicDonemBasi)
    {
        // bkBasi  : Başkanlık-adjusted çalışan başlangıç stokları (döndürülen değer)
        // logoBasi: LOGO-bazlı çalışan başlangıç stokları (delta hesabı için; AY SONU
        //           formülüyle tutarlı kalması amacıyla Başkanlık başıyla değil LOGO
        //           başıyla işlenir — böylece A→TK yönlendirmesi bkBasi'yi tüketmez)
        var bkBasi   = new Dictionary<string, decimal>(baslangicDonemBasi);
        var logoBasi = new Dictionary<string, decimal>(baslangicDonemBasi);

        var ayBaslangic = new DateTime(baslangic.Year, baslangic.Month, 1);
        var sonAy       = new DateTime(bitis.Year,     bitis.Month,     1);

        while (ayBaslangic <= sonAy)
        {
            int ayGun     = DateTime.DaysInMonth(ayBaslangic.Year, ayBaslangic.Month);
            var ayBitis   = new DateTime(ayBaslangic.Year, ayBaslangic.Month, ayGun);
            var gercekBas = ayBaslangic < baslangic ? baslangic : ayBaslangic;
            var gercekBit = ayBitis     > bitis     ? bitis     : ayBitis;

            // Bu ayın hareketlerini LOGO başıyla hesapla → sadece net delta alınacak
            var ayAnaliz = await HesaplaDonemSonuAsync(gercekBas, gercekBit, logoBasi);
            var ayDict   = ayAnaliz.ToDictionary(k => k.Kategori);

            // Sonraki ayın iadelerini al (portal da dönem sonuna bu iadeleri yansıtır)
            var nextAyIlk = new DateTime(ayBitis.AddDays(1).Year, ayBitis.AddDays(1).Month, 1);
            var nextAySon = new DateTime(nextAyIlk.Year, nextAyIlk.Month,
                DateTime.DaysInMonth(nextAyIlk.Year, nextAyIlk.Month));

            var iadePerHedef = new Dictionary<string, decimal>();
            if (DateTime.Today > nextAySon) // Sonraki ay tamamen geçmiş → iadeleri fetch et
            {
                var iadeler = await GetNextAySatisIadeleriAsync(nextAyIlk, nextAySon);
                foreach (var (kategori, miktar, _) in iadeler)
                {
                    if (miktar <= 0) continue;
                    bool satisVar = ayDict.TryGetValue(kategori, out var kd) && kd.SatisMiktar > 0;
                    string hedef = satisVar ? kategori
                                 : (kategori == "PAKETLI" || kategori == "TICARI_PAKET")
                                   ? "TICARI_PAKET" : "TICARI_KRISTAL";
                    iadePerHedef[hedef] = iadePerHedef.GetValueOrDefault(hedef) + miktar;
                }
            }

            // A→TK ve PAKET→TICARI_PAKET yönlendirme (Başkanlık: TM stoğu varsa önce oradan tüketilir)
            // LOGO çalışması logoBasi kadar yönlendirmeyi zaten yaptı.
            // Ekstra = bkBasi ile logoBasi arasındaki fark kadar daha fazla tüketim.
            var aKatC   = ayDict.TryGetValue("A_KOTASI", out var ak)   ? ak  : null;
            var pakKatC = ayDict.TryGetValue("PAKETLI",  out var pk)   ? pk  : null;
            decimal pakUretimC   = pakKatC?.UretimMiktar ?? 0m;
            decimal aSarfExtraC  = Math.Max(0m, (aKatC?.SarfMiktar ?? 0m) - pakUretimC);
            decimal aYemPromsC   = (aKatC?.YemekhaneMiktar ?? 0m) + (aKatC?.PromsMiktar ?? 0m);
            decimal aSatisBrutC  = (aKatC?.SatisMiktar ?? 0m) + aSarfExtraC + aYemPromsC;
            decimal pakYemPromsC = (pakKatC?.YemekhaneMiktar ?? 0m) + (pakKatC?.PromsMiktar ?? 0m);
            decimal pakSatisBrutC = (pakKatC?.SatisMiktar ?? 0m) + (pakKatC?.SarfMiktar ?? 0m) + pakYemPromsC;

            // logoBasi bu ay çalışmaya başlamadan önceki LOGO başlangıç değerleri
            decimal logoTKC = logoBasi.GetValueOrDefault("TICARI_KRISTAL");
            decimal logoTPC = logoBasi.GetValueOrDefault("TICARI_PAKET");
            decimal R_TK = Math.Max(0m, Math.Min(aSatisBrutC,  bkBasi.GetValueOrDefault("TICARI_KRISTAL") - logoTKC));
            decimal R_TP = Math.Max(0m, Math.Min(pakSatisBrutC, bkBasi.GetValueOrDefault("TICARI_PAKET")   - logoTPC));

            // Başkanlık: A/Paket Satış İadesi TM'den yapılan satışın iadesidir;
            // TM kapatıldıktan sonra A'ya değil, TM'e geri döner → A/Paket deltasından çıkart.
            decimal aSatisIadeC  = aKatC?.SatisIadeMiktar  ?? 0m;
            decimal pakSatisIadeC = pakKatC?.SatisIadeMiktar ?? 0m;

            // Başkanlık yeni bası = eskiBk + LOGO delta ± yönlendirme + iadeler
            // A Kotası: R_TK kadar daha az çıktı sayıldığı için delta'ya R_TK eklenir
            // TK      : R_TK kadar fazla çıktı sayıldığı için delta'dan R_TK düşülür
            bkBasi["A_KOTASI"] = bkBasi.GetValueOrDefault("A_KOTASI")
                + (aKatC != null ? aKatC.DonemSonuMiktar - aKatC.DonemBasiMiktar : 0m)
                - aSatisIadeC   // Satış iadesi A'ya eklenmez (TM kaynaklı satışın iadesi)
                + R_TK + iadePerHedef.GetValueOrDefault("A_KOTASI");

            bkBasi["PAKETLI"]  = bkBasi.GetValueOrDefault("PAKETLI")
                + (pakKatC != null ? pakKatC.DonemSonuMiktar - pakKatC.DonemBasiMiktar : 0m)
                - pakSatisIadeC // Paket satış iadesi de Paket'e eklenmez
                + R_TP + iadePerHedef.GetValueOrDefault("PAKETLI");

            foreach (var cat in new[] { "B_KOTASI", "C_KOTASI" })
            {
                decimal delta = ayDict.TryGetValue(cat, out var kd2)
                    ? kd2.DonemSonuMiktar - kd2.DonemBasiMiktar : 0m;
                bkBasi[cat] = bkBasi.GetValueOrDefault(cat) + delta + iadePerHedef.GetValueOrDefault(cat);
            }
            decimal tkDelta   = ayDict.TryGetValue("TICARI_KRISTAL", out var tkKd)
                                ? tkKd.DonemSonuMiktar   - tkKd.DonemBasiMiktar   : 0m;
            decimal knyaDelta = ayDict.TryGetValue("KONYA_TICARI",   out var knyaKd)
                                ? knyaKd.DonemSonuMiktar - knyaKd.DonemBasiMiktar : 0m;
            bkBasi["TICARI_KRISTAL"] = bkBasi.GetValueOrDefault("TICARI_KRISTAL")
                                       + tkDelta + knyaDelta - R_TK
                                       + iadePerHedef.GetValueOrDefault("TICARI_KRISTAL");
            bkBasi["KONYA_TICARI"] = 0m;

            decimal tpakDelta = ayDict.TryGetValue("TICARI_PAKET", out var tpKd)
                                ? tpKd.DonemSonuMiktar - tpKd.DonemBasiMiktar : 0m;
            bkBasi["TICARI_PAKET"] = bkBasi.GetValueOrDefault("TICARI_PAKET")
                                     + tpakDelta - R_TP
                                     + iadePerHedef.GetValueOrDefault("TICARI_PAKET");

            // LOGO zincirini ilerlet (bir sonraki ayın LOGO donemBasi için)
            foreach (var cat in new[] { "A_KOTASI", "B_KOTASI", "C_KOTASI", "PAKETLI",
                                        "KONYA_TICARI", "TICARI_KRISTAL", "TICARI_PAKET" })
                logoBasi[cat] = ayDict.TryGetValue(cat, out var kd3)
                    ? kd3.DonemSonuMiktar : logoBasi.GetValueOrDefault(cat);

            // TM stok negatife düşmesin (KONYA delta aşımı veya yuvarlama hataları)
            if (bkBasi.GetValueOrDefault("TICARI_KRISTAL") < 0) bkBasi["TICARI_KRISTAL"] = 0m;
            if (bkBasi.GetValueOrDefault("TICARI_PAKET")   < 0) bkBasi["TICARI_PAKET"]   = 0m;

            // Başkanlık kuralı: TM stoğu A/Paket stoğunu mahsup eder (netting).
            // TM büyükse A = 0 (TM tüm A'yı karşılar), küçükse A'dan TM kadar düş, TM = 0.
            // Bu kural koşulsuz sıfırlamayı önler — 500 kg TM 17M kg A'yı sıfırlamamalı.
            decimal aV  = bkBasi.GetValueOrDefault("A_KOTASI");
            decimal tkV = bkBasi.GetValueOrDefault("TICARI_KRISTAL");
            if (tkV > 0 && aV > 0)
            {
                decimal net = Math.Min(tkV, aV);
                bkBasi["A_KOTASI"]       = aV  - net;
                bkBasi["TICARI_KRISTAL"] = tkV - net;
            }
            decimal pakV = bkBasi.GetValueOrDefault("PAKETLI");
            decimal tpV  = bkBasi.GetValueOrDefault("TICARI_PAKET");
            if (tpV > 0 && pakV > 0)
            {
                decimal netP = Math.Min(tpV, pakV);
                bkBasi["PAKETLI"]      = pakV - netP;
                bkBasi["TICARI_PAKET"] = tpV  - netP;
            }

            ayBaslangic = ayBaslangic.AddMonths(1);
        }

        return bkBasi;
    }

    /// <summary>
    /// Verilen ay için Başkanlık tablosuna uygun dönem başı stoklarını döndürür.
    /// Önceki aylardaki iade düzeltmeleri zincirli olarak aktarılır.
    /// </summary>
    public async Task<Dictionary<string, decimal>> GetBaskanlikDonemBasiAsync(DateTime baslangic)
    {
        if (baslangic.Date <= KampanyaBaslangic.Date)
            return new Dictionary<string, decimal>(KampanyaDonemBasi);

        return await HesaplaDonemSonuBaskanlikZincirliAsync(
            KampanyaBaslangic, baslangic.AddDays(-1), KampanyaDonemBasi);
    }

    // ─────────────────────────────────────────────────────────────────────
    // Yardımcı metodlar
    // ─────────────────────────────────────────────────────────────────────

    /// <summary>
    /// Malzeme kodundan kategori belirler (SekerSatisService ile aynı mantık)
    /// </summary>
    private static string MalzemeKategorisi(string ad, string kod)
    {
        // Önce malzeme adına göre eşleştir (en güvenilir yöntem)
        var adUpper = (ad ?? "").ToUpperInvariant();
        if (adUpper.Contains("A KOTASI") || adUpper.Contains("A KOTA")) return "A_KOTASI";
        if (adUpper.Contains("B KOTASI") || adUpper.Contains("B KOTA")) return "B_KOTASI";
        if (adUpper.Contains("C KOTASI") || adUpper.Contains("C KOTA")) return "C_KOTASI";
        if (adUpper.Contains("KONYA") && adUpper.Contains("TİCARİ")) return "KONYA_TICARI";
        if (adUpper.Contains("KONYA") && adUpper.Contains("TICARI")) return "KONYA_TICARI";

        // Adda bulunamazsa koda göre dene
        if (string.IsNullOrWhiteSpace(kod)) return "PAKETLI";
        kod = kod.Trim();
        if (kod is "S.T.0.0.0" or "S.T.0.0.4" or "S.705.00.0005") return "A_KOTASI";
        if (kod is "S.T.0.0.8" or "S.705.00.0001")                  return "B_KOTASI";
        if (kod is "S.T.0.0.7" or "S.705.00.0008")                  return "C_KOTASI";
        if (kod == "S.T.1.0.0")                                      return "KONYA_TICARI";

        return "PAKETLI";
    }

    /// <summary>
    /// Fiş türünden operasyon tipi belirler
    /// </summary>
    private static string FisTuruOperasyon(string fisTuru)
    {
        return fisTuru.Trim() switch
        {
            "Üretimden Giriş Fişi"              => "URETIM",
            "HAMMADDE ÇEVRİM GİRİŞİ"            => "HAMMADDE_GIRIS",
            "HAMMADDE ÇEVRİM ÇIKIŞI"             => "HAMMADDE_CIKIS",
            "Satınalma İrsaliyesi"               => "SATINALMA",
            "Satınalma Faturası"                 => "SATINALMA",
            "Toptan Satış İrsaliyesi"            => "SATIS",
            "Toptan Satış Faturası"              => "SATIS",
            "Toptan Satış İade İrsaliyesi"       => "SATIS_IADE",
            "Toptan Satış İade Faturası"         => "SATIS_IADE",
            "Satınalma İade İrsaliyesi"          => "SATINALMA_IADE",
            "Satınalma İade Faturası"            => "SATINALMA_IADE",
            "Sarf Fişi"                          => "SARF",
            "YEMEKHANE KULLANIMI"                => "YEMEKHANE",
            "PROMS ve TEKNİK M ve SARF"          => "PROMS",
            "Fire Fişi"                          => "FIRE",
            "Reçete Fark Fişi"                   => "RECETE_FARK_GIRIS",
            "Sayım Fazlası Fişi"                 => "SAYIM_FAZLASI",
            "Sayım Farkı Fişi"                   => "SAYIM_FAZLASI",
            _                                    => "DIGER"
        };
    }
}
