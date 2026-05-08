using System.Text.Json;
using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Kantar (SabNetKANTAR) ile Logo (INF_UT_Kısıtlı_Malzeme_Raporu) arasında
/// ürün bazında hareket karşılaştırması yapar. Excel makrosunun birebir karşılığı.
/// Eşleştirme listesi C:\RaporlamaPortaliData\kantar_logo_eslesmeleri.json içinde tutulur.
/// </summary>
public class KantarLogoKarsilastirmaService
{
    private readonly DatabaseService _db;

    private static readonly object _lock = new();
    private static List<KantarLogoEslesme>? _cache;

    private static readonly JsonSerializerOptions _jsonOpts = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    private static readonly List<KantarLogoEslesme> _varsayilan = new()
    {
        new() { KantarKodu = "2862", LogoKodu = "S.706.04.0002", MalzemeAdi = "Yaş Küspe Dökme (Bedelli)" },
        new() { KantarKodu = "2926", LogoKodu = "S.706.04.0009", MalzemeAdi = "Yaş Küspe Poşet (1000 Kg)" },
        new() { KantarKodu = "2924", LogoKodu = "S.706.04.0008", MalzemeAdi = "Yaş Küspe Poşet (25 Kg)" },
        new() { KantarKodu = "2908", LogoKodu = "S.706.04.0003", MalzemeAdi = "Çuvallı Kuru Pancar Küspesi" },
        new() { KantarKodu = "2909", LogoKodu = "S.706.04.0004", MalzemeAdi = "Dökme Kuru Pancar Küspesi" },
        new() { KantarKodu = "2866", LogoKodu = "S.706.04.0012", MalzemeAdi = "Peletlenmemiş Kuru Küspe" },
        new() { KantarKodu = "2885", LogoKodu = "S.706.04.0006", MalzemeAdi = "Kuyruk" },
        new() { KantarKodu = "2893", LogoKodu = "S.706.04.0007", MalzemeAdi = "Toprak" },
        // Excel mantığı: Melas ve Iskarta Patates için Kantar birim fiyat / tutardan %1 KDV düş
        new() { KantarKodu = "2927", LogoKodu = "S.706.04.0001", MalzemeAdi = "Melas", KdvDus = true },
        new() { KantarKodu = "2829", LogoKodu = "Y_100153",      MalzemeAdi = "Iskarta Patates", KdvDus = true },
    };

    private static List<KantarLogoEslesme> Eslesmeler
    {
        get
        {
            if (_cache != null) return _cache;
            lock (_lock)
            {
                if (_cache != null) return _cache;
                _cache = Yukle();
                return _cache;
            }
        }
    }

    private static List<KantarLogoEslesme> Yukle()
    {
        try
        {
            var yol = AppDataPaths.KantarLogoEslesmeleriJson;
            if (File.Exists(yol))
            {
                var json = File.ReadAllText(yol);
                var liste = JsonSerializer.Deserialize<List<KantarLogoEslesme>>(json, _jsonOpts);
                if (liste != null) return liste;
            }
        }
        catch { }
        // İlk açılışta veya bozuk dosyada varsayılan listeyi kaydet
        var ilk = _varsayilan.Select(e => new KantarLogoEslesme
        {
            KantarKodu = e.KantarKodu,
            LogoKodu = e.LogoKodu,
            MalzemeAdi = e.MalzemeAdi,
            KdvDus = e.KdvDus
        }).ToList();
        try
        {
            Directory.CreateDirectory(AppDataPaths.DataRoot);
            File.WriteAllText(AppDataPaths.KantarLogoEslesmeleriJson, JsonSerializer.Serialize(ilk, _jsonOpts));
        }
        catch { }
        return ilk;
    }

    private static void Kaydet()
    {
        try
        {
            Directory.CreateDirectory(AppDataPaths.DataRoot);
            File.WriteAllText(AppDataPaths.KantarLogoEslesmeleriJson, JsonSerializer.Serialize(Eslesmeler, _jsonOpts));
        }
        catch { }
    }

    public static IReadOnlyList<KantarLogoEslesme> Eslestirmeler => Eslesmeler;

    /// <summary>Yeni bir eşleştirme ekler. Aynı KantarKodu varsa false döner.</summary>
    public static bool Ekle(KantarLogoEslesme yeni)
    {
        if (yeni == null) return false;
        yeni.KantarKodu = (yeni.KantarKodu ?? "").Trim();
        yeni.LogoKodu = (yeni.LogoKodu ?? "").Trim();
        yeni.MalzemeAdi = (yeni.MalzemeAdi ?? "").Trim();
        if (string.IsNullOrWhiteSpace(yeni.KantarKodu) || string.IsNullOrWhiteSpace(yeni.LogoKodu))
            return false;
        lock (_lock)
        {
            var liste = Eslesmeler;
            if (liste.Any(e => string.Equals(e.KantarKodu, yeni.KantarKodu, StringComparison.OrdinalIgnoreCase)))
                return false;
            liste.Add(yeni);
            Kaydet();
        }
        return true;
    }

    /// <summary>Var olan eşleştirmeyi günceller (KantarKodu eşleşene göre).</summary>
    public static bool Guncelle(string kantarKodu, KantarLogoEslesme yeni)
    {
        if (yeni == null) return false;
        lock (_lock)
        {
            var liste = Eslesmeler;
            var mevcut = liste.FirstOrDefault(e => string.Equals(e.KantarKodu, kantarKodu, StringComparison.OrdinalIgnoreCase));
            if (mevcut == null) return false;
            mevcut.KantarKodu = (yeni.KantarKodu ?? "").Trim();
            mevcut.LogoKodu = (yeni.LogoKodu ?? "").Trim();
            mevcut.MalzemeAdi = (yeni.MalzemeAdi ?? "").Trim();
            mevcut.KdvDus = yeni.KdvDus;
            Kaydet();
        }
        return true;
    }

    public static bool Sil(string kantarKodu)
    {
        lock (_lock)
        {
            var liste = Eslesmeler;
            var n = liste.RemoveAll(e => string.Equals(e.KantarKodu, kantarKodu, StringComparison.OrdinalIgnoreCase));
            if (n == 0) return false;
            Kaydet();
            return true;
        }
    }

    public static void VarsayilanaDondur()
    {
        lock (_lock)
        {
            _cache = _varsayilan.Select(e => new KantarLogoEslesme
            {
                KantarKodu = e.KantarKodu,
                LogoKodu = e.LogoKodu,
                MalzemeAdi = e.MalzemeAdi,
                KdvDus = e.KdvDus
            }).ToList();
            Kaydet();
        }
    }

    public KantarLogoKarsilastirmaService(DatabaseService db)
    {
        _db = db;
    }

    public async Task<List<KantarHamSatir>> KantarHamAsync(DateTime baslangic, DateTime bitis, CancellationToken ct = default)
    {
        bitis = SistemTarihi.Clamp(bitis);
        // Kantar DB'de Tarih int olarak Excel date serial biçiminde saklanır (1900 leap year bug dahil).
        // Doğru dönüşüm: DATEADD(day, Tarih - 2, '1900-01-01'). CAST(Tarih AS datetime) 2 gün kaydırır.
        var sql = $@"
SELECT
    FisNo          = ISNULL(FisNo,''),
    UrunKodu       = ISNULL(UrunKodu,''),
    StokAdi        = ISNULL(StokAdi,''),
    PlakaNo        = ISNULL(PlakaNo,''),
    SoforAdiSoyadi = ISNULL(SoforAdiSoyadi,''),
    Tarih          = DATEADD(day, Tarih - 2, CAST('1900-01-01' AS datetime)),
    Net            = ISNULL(Net, 0),
    BirimFiyat     = ISNULL(BirimFiyat, 0),
    Tutar          = ISNULL(Tutar, 0)
FROM {_db.KantarTablo} WITH(NOLOCK)
WHERE Tarih >= DATEDIFF(day, '1900-01-01', @bas) + 2
  AND Tarih <  DATEDIFF(day, '1900-01-01', @bit) + 3
ORDER BY Tarih, FisNo";
        using var c = _db.CreateKantarConnection();
        var rows = await c.QueryAsync<KantarHamSatir>(new CommandDefinition(sql, new { bas = baslangic.Date, bit = bitis.Date }, cancellationToken: ct));
        return rows.ToList();
    }

    public async Task<List<LogoHamSatir>> LogoHamAsync(DateTime baslangic, DateTime bitis, CancellationToken ct = default)
    {
        bitis = SistemTarihi.Clamp(bitis);
        var sql = $@"
SELECT
    Tarih        = TARIH,
    FisTuru      = ISNULL(FIS_TURU,''),
    FisNumarasi  = ISNULL(FIS_NUMARASI,''),
    CariKodu     = ISNULL(CARI_HESAP_KODU,''),
    CariUnvani   = ISNULL(CARI_HESAP_UNVANI,''),
    MalzemeKodu  = ISNULL(MALZEME_KODU,''),
    MalzemeAdi   = ISNULL(MALZEME_ACIKLAMASI,''),
    CikisMiktari = ISNULL(CIKIS_MIKTARI, 0),
    CikisFiyati  = ISNULL(CIKIS_FIYATI, 0),
    CikisTutari  = ISNULL(CIKIS_TUTARI, 0)
FROM {_db.LogoMalzemeTablo} WITH(NOLOCK)
WHERE TARIH >= @bas
  AND TARIH <  DATEADD(day, 1, @bit)
  AND (MALZEME_KODU LIKE 'S.706%' OR MALZEME_KODU LIKE 'Y\_%' ESCAPE '\')
ORDER BY TARIH, FIS_NUMARASI";
        using var c = _db.CreateConnection();
        var rows = await c.QueryAsync<LogoHamSatir>(new CommandDefinition(sql, new { bas = baslangic.Date, bit = bitis.Date }, cancellationToken: ct));
        return rows.ToList();
    }

    /// <summary>
    /// Belirli bir tarih + logo kodu için fiş bazında Kantar ↔ Logo eşleştirmesi yapar.
    /// Eşleştirme anahtarı: (Net ≈ Miktar) AND (Tutar ≈ Tutar). KDV düşülen ürünler için Kantar tutarı /1.01 uygulanır.
    /// </summary>
    public async Task<List<FisEslestirmesi>> FisDetayAsync(DateTime tarih, string logoKodu, CancellationToken ct = default)
    {
        var esl = Eslesmeler.FirstOrDefault(e => string.Equals(e.LogoKodu, logoKodu, StringComparison.OrdinalIgnoreCase));
        if (esl == null) return new();

        // Kantar fişleri (o tarih + o ürün)
        var kantarSql = $@"
SELECT
    FisNo          = ISNULL(FisNo,''),
    UrunKodu       = ISNULL(UrunKodu,''),
    StokAdi        = ISNULL(StokAdi,''),
    PlakaNo        = ISNULL(PlakaNo,''),
    SoforAdiSoyadi = ISNULL(SoforAdiSoyadi,''),
    Tarih          = DATEADD(day, Tarih - 2, CAST('1900-01-01' AS datetime)),
    Net            = ISNULL(Net, 0),
    BirimFiyat     = ISNULL(BirimFiyat, 0),
    Tutar          = ISNULL(Tutar, 0)
FROM {_db.KantarTablo} WITH(NOLOCK)
WHERE Tarih = DATEDIFF(day, '1900-01-01', @t) + 2
  AND UrunKodu = @u
ORDER BY FisNo";
        using var kc = _db.CreateKantarConnection();
        var kantarList = (await kc.QueryAsync<KantarHamSatir>(new CommandDefinition(kantarSql, new { t = tarih.Date, u = esl.KantarKodu }, cancellationToken: ct))).ToList();

        // Logo fişleri (o tarih + o malzeme kodu)
        var logoSql = $@"
SELECT
    Tarih        = TARIH,
    FisTuru      = ISNULL(FIS_TURU,''),
    FisNumarasi  = ISNULL(FIS_NUMARASI,''),
    CariKodu     = ISNULL(CARI_HESAP_KODU,''),
    CariUnvani   = ISNULL(CARI_HESAP_UNVANI,''),
    MalzemeKodu  = ISNULL(MALZEME_KODU,''),
    MalzemeAdi   = ISNULL(MALZEME_ACIKLAMASI,''),
    CikisMiktari = ISNULL(CIKIS_MIKTARI, 0),
    CikisFiyati  = ISNULL(CIKIS_FIYATI, 0),
    CikisTutari  = ISNULL(CIKIS_TUTARI, 0)
FROM {_db.LogoMalzemeTablo} WITH(NOLOCK)
WHERE TARIH = @t
  AND MALZEME_KODU = @m
ORDER BY FIS_NUMARASI";
        using var lc = _db.CreateConnection();
        var logoList = (await lc.QueryAsync<LogoHamSatir>(new CommandDefinition(logoSql, new { t = tarih.Date, m = esl.LogoKodu }, cancellationToken: ct))).ToList();

        // Logo miktar/tutar mutlak değere çevir
        var logoItems = logoList.Select(l => new
        {
            Row = l,
            Miktar = Math.Abs(l.CikisMiktari),
            BirimFiyat = Math.Abs(l.CikisFiyati),
            Tutar = Math.Abs(l.CikisTutari)
        }).ToList();

        // KDV düşümü için Kantar tutar/birim fiyatını düz
        var kantarItems = kantarList.Select(k =>
        {
            decimal birim = k.BirimFiyat, tutar = k.Tutar;
            if (esl.KdvDus) { birim = birim / 1.01m; tutar = tutar / 1.01m; }
            return new { Row = k, Miktar = k.Net, BirimFiyat = birim, Tutar = tutar };
        }).ToList();

        var sonuc = new List<FisEslestirmesi>();
        var logoKullanildi = new bool[logoItems.Count];
        const decimal tol = 0.5m;

        foreach (var k in kantarItems)
        {
            int eslesen = -1;
            for (int i = 0; i < logoItems.Count; i++)
            {
                if (logoKullanildi[i]) continue;
                var l = logoItems[i];
                if (Math.Abs(k.Miktar - l.Miktar) < tol && Math.Abs(k.Tutar - l.Tutar) < tol)
                {
                    eslesen = i;
                    break;
                }
            }

            if (eslesen >= 0)
            {
                logoKullanildi[eslesen] = true;
                var l = logoItems[eslesen];
                sonuc.Add(new FisEslestirmesi
                {
                    KantarFisNo = k.Row.FisNo,
                    KantarPlaka = k.Row.PlakaNo,
                    KantarSofor = k.Row.SoforAdiSoyadi,
                    KantarNet = k.Row.Net,
                    KantarBirimFiyat = k.Row.BirimFiyat,
                    KantarTutar = k.Row.Tutar,
                    LogoFisNumarasi = l.Row.FisNumarasi,
                    LogoCariKodu = l.Row.CariKodu,
                    LogoCariUnvani = l.Row.CariUnvani,
                    LogoMiktar = l.Miktar,
                    LogoBirimFiyat = l.BirimFiyat,
                    LogoTutar = l.Tutar,
                    Durum = "EŞLEŞTİ"
                });
            }
            else
            {
                sonuc.Add(new FisEslestirmesi
                {
                    KantarFisNo = k.Row.FisNo,
                    KantarPlaka = k.Row.PlakaNo,
                    KantarSofor = k.Row.SoforAdiSoyadi,
                    KantarNet = k.Row.Net,
                    KantarBirimFiyat = k.Row.BirimFiyat,
                    KantarTutar = k.Row.Tutar,
                    Durum = "LOGO'DA YOK"
                });
            }
        }

        for (int i = 0; i < logoItems.Count; i++)
        {
            if (logoKullanildi[i]) continue;
            var l = logoItems[i];
            sonuc.Add(new FisEslestirmesi
            {
                LogoFisNumarasi = l.Row.FisNumarasi,
                LogoCariKodu = l.Row.CariKodu,
                LogoCariUnvani = l.Row.CariUnvani,
                LogoMiktar = l.Miktar,
                LogoBirimFiyat = l.BirimFiyat,
                LogoTutar = l.Tutar,
                Durum = "KANTAR'DA YOK"
            });
        }

        return sonuc
            .OrderBy(s => s.Durum == "EŞLEŞTİ" ? 0 : 1)
            .ThenBy(s => s.KantarFisNo ?? s.LogoFisNumarasi)
            .ToList();
    }

    public async Task<KarsilastirmaSonuc> KarsilastirmaYapAsync(DateTime baslangic, DateTime bitis, CancellationToken ct = default)
    {
        bitis = SistemTarihi.Clamp(bitis);
        var kantar = await KantarHamAsync(baslangic, bitis, ct);
        var logo = await LogoHamAsync(baslangic, bitis, ct);

        var eslesmeDict = Eslesmeler.ToDictionary(e => e.KantarKodu, e => e);
        var logoKoduSet = Eslesmeler.Select(e => e.LogoKodu).ToHashSet(StringComparer.OrdinalIgnoreCase);

        // Anahtar: tarih (gün) + logo kodu
        var gunlukMap = new Dictionary<(DateTime, string), KarsilastirmaSatiri>();

        foreach (var k in kantar)
        {
            if (string.IsNullOrWhiteSpace(k.UrunKodu)) continue;
            if (!eslesmeDict.TryGetValue(k.UrunKodu.Trim(), out var esl)) continue;

            decimal birim = k.BirimFiyat;
            decimal tutar = k.Tutar;
            if (esl.KdvDus)
            {
                birim = birim / 1.01m;
                tutar = tutar / 1.01m;
            }

            var key = (k.Tarih.Date, esl.LogoKodu);
            if (!gunlukMap.TryGetValue(key, out var satir))
            {
                satir = new KarsilastirmaSatiri
                {
                    Tarih = k.Tarih.Date,
                    LogoKodu = esl.LogoKodu,
                    MalzemeAdi = esl.MalzemeAdi
                };
                gunlukMap[key] = satir;
            }
            satir.KantarMiktar += k.Net;
            satir.KantarBirimFiyat += birim;
            satir.KantarTutar += tutar;
        }

        foreach (var l in logo)
        {
            if (string.IsNullOrWhiteSpace(l.MalzemeKodu)) continue;
            var kod = l.MalzemeKodu.Trim();
            // VBA: Left(kod,5)='S.706' OR Left(kod,2)='Y_' — zaten SQL'de filtreli
            // Sadece eşleştirme tablosundaki Logo kodlarını kapsa (VBA tüm S.706'ları eşleşmeye ekliyor ama sadece kantar eşleştirmesi olanlara denk gelebiliyor)
            decimal miktar = Math.Abs(l.CikisMiktari);
            decimal birim = Math.Abs(l.CikisFiyati);
            decimal tutar = Math.Abs(l.CikisTutari);

            var key = (l.Tarih.Date, kod);
            if (!gunlukMap.TryGetValue(key, out var satir))
            {
                satir = new KarsilastirmaSatiri
                {
                    Tarih = l.Tarih.Date,
                    LogoKodu = kod,
                    MalzemeAdi = l.MalzemeAdi
                };
                gunlukMap[key] = satir;
            }
            satir.LogoMiktar += miktar;
            satir.LogoBirimFiyat += birim;
            satir.LogoTutar += tutar;
        }

        var gunluk = gunlukMap.Values
            .OrderBy(s => s.Tarih)
            .ThenBy(s => s.LogoKodu)
            .ToList();

        // Toplu rapor — VBA mantığı: bugünün tarihi hariç
        var bugun = DateTime.Today;
        var topluDict = Eslesmeler.ToDictionary(
            e => e.LogoKodu,
            e => new TopluKarsilastirma { LogoKodu = e.LogoKodu, MalzemeAdi = e.MalzemeAdi });

        foreach (var g in gunluk)
        {
            if (g.Tarih >= bugun) continue;
            if (!topluDict.TryGetValue(g.LogoKodu, out var t)) continue;
            t.KantarMiktar += g.KantarMiktar;
            t.KantarTutar += g.KantarTutar;
            t.LogoMiktar += g.LogoMiktar;
            t.LogoTutar += g.LogoTutar;
        }

        var toplu = topluDict.Values
            .Where(t => t.KantarMiktar != 0 || t.LogoMiktar != 0 || t.KantarTutar != 0 || t.LogoTutar != 0)
            .OrderBy(t => t.LogoKodu)
            .ToList();

        return new KarsilastirmaSonuc
        {
            GunlukRapor = gunluk,
            TopluRapor = toplu,
            KantarSatirSayisi = kantar.Count,
            LogoSatirSayisi = logo.Count,
            BaslangicTarihi = baslangic,
            BitisTarihi = bitis
        };
    }
}
