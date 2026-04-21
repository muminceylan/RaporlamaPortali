using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// SabNetPMHS (Avans / Makbuz / CariHareketler) ile Logo'daki
/// INF_Mustahsil_Hareketleri_211 view'ı arasında TC bazında müstahsil
/// hareketlerini karşılaştırır. Excel dosyasına ihtiyaç yoktur —
/// Logo verisi doğrudan SQL'den çekilir.
/// TC kimlik numarası = SUBSTRING(CARI_HESAP_KODU, 2, 11)
/// </summary>
public class MustahsilKarsilastirmaService
{
    private readonly DatabaseService _db;
    private readonly ILogger<MustahsilKarsilastirmaService> _logger;

    // Logo HAREKET_TURU ↔ SabNet tablosu eşleştirmeleri
    public const string HT_AVANS = "Toptan satış faturası";
    public const string HT_MAKBUZ = "Müstahsil makbuzu";
    public const string HT_VIRMAN = "Virman Fişi";

    public MustahsilKarsilastirmaService(DatabaseService db, ILogger<MustahsilKarsilastirmaService> logger)
    {
        _db = db;
        _logger = logger;
    }

    public async Task<MustahsilKarsilastirmaSonuc> KarsilastirAsync(DateTime baslangic, DateTime bitis, string kampanyaYili, CancellationToken ct = default)
    {
        var sonuc = new MustahsilKarsilastirmaSonuc
        {
            BaslangicTarihi = baslangic,
            BitisTarihi = bitis,
            KampanyaYili = kampanyaYili
        };

        // Logo + SabNet toplamlarını paralel çek
        // + Diğer yıllara ait SabNet satırlarını çek (2025 gibi) — Logo'da eşleşenleri
        //   özet hesabından düşmek için kullanılır (önce cari yıl eşleşmesi denendikten sonra).
        var exclTask = DigerYilExclusionAsync(baslangic, bitis, kampanyaYili, ct);
        var sabAvansTask = SabAvansToplamAsync(baslangic, bitis, kampanyaYili, ct);
        var sabMakbuzTask = SabMakbuzToplamAsync(baslangic, bitis, kampanyaYili, ct);
        var sabCariTask = SabCariToplamAsync(baslangic, bitis, kampanyaYili, ct);
        var curAvansRowsTask = CurrentYearRowsAvansAsync(baslangic, bitis, kampanyaYili, ct);
        var curMakbuzRowsTask = CurrentYearRowsMakbuzAsync(baslangic, bitis, kampanyaYili, ct);
        var curVirmanRowsTask = CurrentYearRowsVirmanAsync(baslangic, bitis, kampanyaYili, ct);
        var ciftciAdTask = CiftciAdMapAsync(ct);
        await Task.WhenAll(exclTask, sabAvansTask, sabMakbuzTask, sabCariTask,
                           curAvansRowsTask, curMakbuzRowsTask, curVirmanRowsTask, ciftciAdTask);

        var excl = exclTask.Result;
        var logo = await LogoToplamAsync(baslangic, bitis, excl,
            curAvansRowsTask.Result, curMakbuzRowsTask.Result, curVirmanRowsTask.Result, ct);
        var sabAvans = sabAvansTask.Result;
        var sabMakbuz = sabMakbuzTask.Result;
        var sabCari = sabCariTask.Result;
        var ciftciAd = ciftciAdTask.Result;

        sonuc.LogoSatirSayisi = logo.Sum(x => x.Value.Adet);
        sonuc.SabAvansSatirSayisi = sabAvans.Sum(x => x.Value.Adet);
        sonuc.SabMakbuzSatirSayisi = sabMakbuz.Sum(x => x.Value.Adet);
        sonuc.SabCariSatirSayisi = sabCari.Sum(x => x.Value.BorcAdet + x.Value.AlacakAdet);

        // TC kümesini birleştir
        var tumTc = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var k in logo.Keys) tumTc.Add(k);
        foreach (var k in sabAvans.Keys) tumTc.Add(k);
        foreach (var k in sabMakbuz.Keys) tumTc.Add(k);
        foreach (var k in sabCari.Keys) tumTc.Add(k);

        var kayitlar = new List<TcOzet>(tumTc.Count);
        foreach (var tc in tumTc)
        {
            var ozet = new TcOzet { TcKimlikNo = tc };

            // Ad-soyad öncelik sırası:
            //   1) PMHS_CiftciKarti (master) — en doğru kaynak
            //   2) PMHS_MustahsilMakbuzu.AdiSoyadi
            //   3) Logo CARI_HESAP_UNVANI
            if (ciftciAd.TryGetValue(tc, out var ad) && !string.IsNullOrWhiteSpace(ad))
                ozet.AdiSoyadi = ad;

            if (logo.TryGetValue(tc, out var lg))
            {
                if (string.IsNullOrWhiteSpace(ozet.AdiSoyadi)) ozet.AdiSoyadi = lg.AdiSoyadi;
                ozet.LogoToptanSatisFaturasi = lg.ToptanSatisFaturasi;
                ozet.LogoMustahsilMakbuzu = lg.MustahsilMakbuzu;
                ozet.LogoVirmanBorc = lg.VirmanBorc;
                ozet.LogoVirmanAlacak = lg.VirmanAlacak;
                ozet.LogoSatirSayisi = lg.Adet;
            }
            if (sabAvans.TryGetValue(tc, out var a))
            {
                ozet.SabAvans = a.Tutar;
            }
            if (sabMakbuz.TryGetValue(tc, out var m))
            {
                ozet.SabMakbuz = m.Tutar;
                if (string.IsNullOrWhiteSpace(ozet.AdiSoyadi)) ozet.AdiSoyadi = m.AdiSoyadi;
            }
            if (sabCari.TryGetValue(tc, out var c))
            {
                ozet.SabCariBorc = c.Borc;
                ozet.SabCariAlacak = c.Alacak;
            }
            kayitlar.Add(ozet);
        }

        sonuc.Kayitlar = kayitlar
            .OrderByDescending(x => !x.Eslesiyor)
            .ThenBy(x => x.AdiSoyadi)
            .ToList();
        return sonuc;
    }

    // ===== Çiftçi master adları (PMHS_CiftciKarti) =====
    // TC → AdiSoyadi lookup; tüm karşılaştırmalarda birincil isim kaynağı
    private async Task<Dictionary<string, string>> CiftciAdMapAsync(CancellationToken ct)
    {
        const string sql = @"
SELECT TcKimlikNo = LTRIM(RTRIM(ISNULL(TcKimlikNo,''))),
       AdiSoyadi = LTRIM(RTRIM(ISNULL(AdiSoyadi,'')))
FROM PMHS_CiftciKarti WITH(NOLOCK)
WHERE ISNULL(TcKimlikNo,'') <> ''";
        using var c = _db.CreatePmhsConnection();
        var rows = await c.QueryAsync<(string TcKimlikNo, string AdiSoyadi)>(
            new CommandDefinition(sql, cancellationToken: ct));
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in rows)
        {
            if (string.IsNullOrWhiteSpace(r.TcKimlikNo)) continue;
            if (!map.ContainsKey(r.TcKimlikNo))
                map[r.TcKimlikNo] = r.AdiSoyadi;
        }
        return map;
    }

    // ===== SabNet sorguları =====
    // Tarih kolonu Excel date serial (int). Dönüşüm: DATEADD(day, Tarih - 2, '1900-01-01')

    private record AvansItem(decimal Tutar, int Adet, string AdiSoyadi);
    private record MakbuzItem(decimal Tutar, int Adet, string AdiSoyadi);
    private record CariItem(decimal Borc, decimal Alacak, int BorcAdet, int AlacakAdet);

    private async Task<Dictionary<string, AvansItem>> SabAvansToplamAsync(DateTime bas, DateTime bit, string kampanyaYili, CancellationToken ct)
    {
        int basSerial = (int)(bas.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        int bitSerial = (int)(bit.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        const string sql = @"
SELECT TcKimlikNo = ISNULL(TcKimlikNo,''),
       Tutar = SUM(ISNULL(Tutar,0)),
       Adet = COUNT(*),
       AdiSoyadi = MAX(ISNULL(Kaydeden,''))
FROM PMHS_AvansFormu WITH(NOLOCK)
WHERE FormTarihi >= @bas AND FormTarihi <= @bit
  AND ISNULL(TcKimlikNo,'') <> ''
  AND ISNULL(SozlesmeYili,'') = @yil
GROUP BY TcKimlikNo";
        using var c = _db.CreatePmhsConnection();
        var rows = await c.QueryAsync<dynamic>(new CommandDefinition(sql, new { bas = basSerial, bit = bitSerial, yil = kampanyaYili }, cancellationToken: ct));
        return rows.ToDictionary(
            r => (string)r.TcKimlikNo,
            r => new AvansItem((decimal)r.Tutar, (int)r.Adet, (string)r.AdiSoyadi),
            StringComparer.OrdinalIgnoreCase);
    }

    private async Task<Dictionary<string, MakbuzItem>> SabMakbuzToplamAsync(DateTime bas, DateTime bit, string kampanyaYili, CancellationToken ct)
    {
        int basSerial = (int)(bas.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        int bitSerial = (int)(bit.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        const string sql = @"
SELECT TcKimlikNo = ISNULL(TcKimlikNo,''),
       Tutar = SUM(ISNULL(NetHakedis,0)),
       Adet = COUNT(*),
       AdiSoyadi = MAX(ISNULL(AdiSoyadi,''))
FROM PMHS_MustahsilMakbuzu WITH(NOLOCK)
WHERE Tarih >= @bas AND Tarih <= @bit
  AND ISNULL(TcKimlikNo,'') <> ''
  AND ISNULL(KampanyaYili,'') = @yil
GROUP BY TcKimlikNo";
        using var c = _db.CreatePmhsConnection();
        var rows = await c.QueryAsync<dynamic>(new CommandDefinition(sql, new { bas = basSerial, bit = bitSerial, yil = kampanyaYili }, cancellationToken: ct));
        return rows.ToDictionary(
            r => (string)r.TcKimlikNo,
            r => new MakbuzItem((decimal)r.Tutar, (int)r.Adet, (string)r.AdiSoyadi),
            StringComparer.OrdinalIgnoreCase);
    }

    private async Task<Dictionary<string, CariItem>> SabCariToplamAsync(DateTime bas, DateTime bit, string kampanyaYili, CancellationToken ct)
    {
        int basSerial = (int)(bas.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        int bitSerial = (int)(bit.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        const string sql = @"
SELECT TcKimlikNo = ISNULL(TcKimlikNo,''),
       Borc = SUM(CASE WHEN BA='BORÇ' THEN ISNULL(Tutar,0) ELSE 0 END),
       Alacak = SUM(CASE WHEN BA='ALACAK' THEN ISNULL(Tutar,0) ELSE 0 END),
       BorcAdet = SUM(CASE WHEN BA='BORÇ' THEN 1 ELSE 0 END),
       AlacakAdet = SUM(CASE WHEN BA='ALACAK' THEN 1 ELSE 0 END)
FROM PMHS_CariHareketler WITH(NOLOCK)
WHERE Tarih >= @bas AND Tarih <= @bit
  AND ISNULL(TcKimlikNo,'') <> ''
  AND ISNULL(SozlesmeYili,'') = @yil
GROUP BY TcKimlikNo";
        using var c = _db.CreatePmhsConnection();
        var rows = await c.QueryAsync<dynamic>(new CommandDefinition(sql, new { bas = basSerial, bit = bitSerial, yil = kampanyaYili }, cancellationToken: ct));
        return rows.ToDictionary(
            r => (string)r.TcKimlikNo,
            r => new CariItem((decimal)r.Borc, (decimal)r.Alacak, (int)r.BorcAdet, (int)r.AlacakAdet),
            StringComparer.OrdinalIgnoreCase);
    }

    // ===== Logo sorgusu =====
    // INF_Mustahsil_Hareketleri_211 view: CARI_HESAP_UNVANI, MODUL, TARIH, HAREKET_TURU,
    // CARI_HESAP_KODU (TC = SUBSTRING(CARI_HESAP_KODU, 2, 11)), TOPLAM
    private record LogoItem(decimal ToptanSatisFaturasi, decimal MustahsilMakbuzu, decimal VirmanBorc, decimal VirmanAlacak, int Adet, string AdiSoyadi);

    // Seçilen yıldan farklı yıllara ait SabNet satırlarını TC + kategori + tutar
    // setleri halinde döner. Logo özetinde aynı TC + kategori + tutara sahip
    // satırlar "diğer yıl kampanyasına ait" kabul edilir ve toplamdan düşülür.
    private record DigerYilExcl(
        Dictionary<string, List<decimal>> Avans,
        Dictionary<string, List<decimal>> Makbuz,
        Dictionary<string, List<decimal>> Virman);

    private async Task<DigerYilExcl> DigerYilExclusionAsync(DateTime bas, DateTime bit, string kampanyaYili, CancellationToken ct)
    {
        // NOT: Tarih filtresi yok. SabNet'te bir önceki/diğer kampanyaya ait
        // kayıtların Logo'ya hangi tarihte işlendiği değişebildiği için
        // (ör. 2025 kampanya virmanı Logo'da 2026 Mart'ta işlenebilir),
        // eleme yalnızca TC + Tutar + SozlesmeYili≠@yil kriterine göre yapılır.
        const string sqlAvans = @"
SELECT TcKimlikNo = ISNULL(TcKimlikNo,''), Tutar = ISNULL(Tutar,0)
FROM PMHS_AvansFormu WITH(NOLOCK)
WHERE ISNULL(TcKimlikNo,'') <> ''
  AND ISNULL(SozlesmeYili,'') <> ''
  AND ISNULL(SozlesmeYili,'') <> @yil";
        const string sqlMakbuz = @"
SELECT TcKimlikNo = ISNULL(TcKimlikNo,''), Tutar = ISNULL(NetHakedis,0)
FROM PMHS_MustahsilMakbuzu WITH(NOLOCK)
WHERE ISNULL(TcKimlikNo,'') <> ''
  AND ISNULL(KampanyaYili,'') <> ''
  AND ISNULL(KampanyaYili,'') <> @yil";
        const string sqlCari = @"
SELECT TcKimlikNo = ISNULL(TcKimlikNo,''),
       BA = ISNULL(BA,''),
       Tutar = ISNULL(Tutar,0)
FROM PMHS_CariHareketler WITH(NOLOCK)
WHERE ISNULL(TcKimlikNo,'') <> ''
  AND ISNULL(SozlesmeYili,'') <> ''
  AND ISNULL(SozlesmeYili,'') <> @yil";

        using var c = _db.CreatePmhsConnection();
        var prm = new { yil = kampanyaYili };
        var aRows = await c.QueryAsync<(string TcKimlikNo, decimal Tutar)>(new CommandDefinition(sqlAvans, prm, cancellationToken: ct));
        var mRows = await c.QueryAsync<(string TcKimlikNo, decimal Tutar)>(new CommandDefinition(sqlMakbuz, prm, cancellationToken: ct));
        var vRows = await c.QueryAsync<(string TcKimlikNo, string BA, decimal Tutar)>(new CommandDefinition(sqlCari, prm, cancellationToken: ct));

        var avans = new Dictionary<string, List<decimal>>(StringComparer.OrdinalIgnoreCase);
        var makbuz = new Dictionary<string, List<decimal>>(StringComparer.OrdinalIgnoreCase);
        var virman = new Dictionary<string, List<decimal>>(StringComparer.OrdinalIgnoreCase);

        foreach (var r in aRows)
            (avans.TryGetValue(r.TcKimlikNo, out var l) ? l : avans[r.TcKimlikNo] = new()).Add(r.Tutar);
        foreach (var r in mRows)
            (makbuz.TryGetValue(r.TcKimlikNo, out var l) ? l : makbuz[r.TcKimlikNo] = new()).Add(r.Tutar);
        foreach (var r in vRows)
        {
            decimal imzali = string.Equals(r.BA, "ALACAK", StringComparison.OrdinalIgnoreCase) ? -r.Tutar : r.Tutar;
            (virman.TryGetValue(r.TcKimlikNo, out var l) ? l : virman[r.TcKimlikNo] = new()).Add(imzali);
        }
        return new DigerYilExcl(avans, makbuz, virman);
    }

    private static bool TryExclude(Dictionary<string, List<decimal>> map, string tc, decimal tutar)
    {
        if (!map.TryGetValue(tc, out var list)) return false;
        for (int i = 0; i < list.Count; i++)
        {
            if (Math.Abs(list[i] - tutar) < EslesmeToleransi)
            {
                list.RemoveAt(i);
                return true;
            }
        }
        return false;
    }

    // Cari kampanyaya ait SabNet satır tutarları (TC bazında liste).
    // Logo satırlarını önce BU listeye göre "tüketip" eşleştirmek için kullanılır.
    private async Task<Dictionary<string, List<decimal>>> CurrentYearRowsAvansAsync(DateTime bas, DateTime bit, string kampanyaYili, CancellationToken ct)
    {
        int basSerial = (int)(bas.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        int bitSerial = (int)(bit.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        const string sql = @"
SELECT TcKimlikNo = ISNULL(TcKimlikNo,''), Tutar = ISNULL(Tutar,0)
FROM PMHS_AvansFormu WITH(NOLOCK)
WHERE FormTarihi >= @bas AND FormTarihi <= @bit
  AND ISNULL(TcKimlikNo,'') <> ''
  AND ISNULL(SozlesmeYili,'') = @yil";
        using var c = _db.CreatePmhsConnection();
        var rows = await c.QueryAsync<(string TcKimlikNo, decimal Tutar)>(
            new CommandDefinition(sql, new { bas = basSerial, bit = bitSerial, yil = kampanyaYili }, cancellationToken: ct));
        var map = new Dictionary<string, List<decimal>>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in rows)
            (map.TryGetValue(r.TcKimlikNo, out var l) ? l : map[r.TcKimlikNo] = new()).Add(r.Tutar);
        return map;
    }

    private async Task<Dictionary<string, List<decimal>>> CurrentYearRowsMakbuzAsync(DateTime bas, DateTime bit, string kampanyaYili, CancellationToken ct)
    {
        int basSerial = (int)(bas.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        int bitSerial = (int)(bit.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        const string sql = @"
SELECT TcKimlikNo = ISNULL(TcKimlikNo,''), Tutar = ISNULL(NetHakedis,0)
FROM PMHS_MustahsilMakbuzu WITH(NOLOCK)
WHERE Tarih >= @bas AND Tarih <= @bit
  AND ISNULL(TcKimlikNo,'') <> ''
  AND ISNULL(KampanyaYili,'') = @yil";
        using var c = _db.CreatePmhsConnection();
        var rows = await c.QueryAsync<(string TcKimlikNo, decimal Tutar)>(
            new CommandDefinition(sql, new { bas = basSerial, bit = bitSerial, yil = kampanyaYili }, cancellationToken: ct));
        var map = new Dictionary<string, List<decimal>>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in rows)
            (map.TryGetValue(r.TcKimlikNo, out var l) ? l : map[r.TcKimlikNo] = new()).Add(r.Tutar);
        return map;
    }

    private async Task<Dictionary<string, List<decimal>>> CurrentYearRowsVirmanAsync(DateTime bas, DateTime bit, string kampanyaYili, CancellationToken ct)
    {
        int basSerial = (int)(bas.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        int bitSerial = (int)(bit.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        const string sql = @"
SELECT TcKimlikNo = ISNULL(TcKimlikNo,''),
       BA = ISNULL(BA,''),
       Tutar = ISNULL(Tutar,0)
FROM PMHS_CariHareketler WITH(NOLOCK)
WHERE Tarih >= @bas AND Tarih <= @bit
  AND ISNULL(TcKimlikNo,'') <> ''
  AND ISNULL(SozlesmeYili,'') = @yil";
        using var c = _db.CreatePmhsConnection();
        var rows = await c.QueryAsync<(string TcKimlikNo, string BA, decimal Tutar)>(
            new CommandDefinition(sql, new { bas = basSerial, bit = bitSerial, yil = kampanyaYili }, cancellationToken: ct));
        var map = new Dictionary<string, List<decimal>>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in rows)
        {
            decimal imzali = string.Equals(r.BA, "ALACAK", StringComparison.OrdinalIgnoreCase) ? -r.Tutar : r.Tutar;
            (map.TryGetValue(r.TcKimlikNo, out var l) ? l : map[r.TcKimlikNo] = new()).Add(imzali);
        }
        return map;
    }

    private async Task<Dictionary<string, LogoItem>> LogoToplamAsync(
        DateTime bas, DateTime bit, DigerYilExcl? excl,
        Dictionary<string, List<decimal>> curAvans,
        Dictionary<string, List<decimal>> curMakbuz,
        Dictionary<string, List<decimal>> curVirman,
        CancellationToken ct)
    {
        const string sql = @"
SELECT TcKimlikNo = LTRIM(RTRIM(SUBSTRING(CARI_HESAP_KODU, 2, 11))),
       HAREKET_TURU,
       Tutar = ISNULL(TOPLAM, 0),
       Unvan = ISNULL(CARI_HESAP_UNVANI, '')
FROM INF_Mustahsil_Hareketleri_211 WITH(NOLOCK)
WHERE TARIH >= @bas AND TARIH < DATEADD(day, 1, @bit)
  AND CARI_HESAP_KODU IS NOT NULL
  AND LEN(CARI_HESAP_KODU) >= 11
  AND HAREKET_TURU IN (@ht1, @ht2, @ht3)";

        using var c = _db.CreateConnection();
        var rows = await c.QueryAsync<(string TcKimlikNo, string HAREKET_TURU, decimal Tutar, string Unvan)>(
            new CommandDefinition(sql,
                new { bas = bas.Date, bit = bit.Date, ht1 = HT_AVANS, ht2 = HT_MAKBUZ, ht3 = HT_VIRMAN },
                cancellationToken: ct));

        // Cari yıl listelerini bu metot içinde tüketeceğiz — orijinali bozmamak için derin kopyala
        static Dictionary<string, List<decimal>> CopyMap(Dictionary<string, List<decimal>> src) =>
            src.ToDictionary(kv => kv.Key, kv => new List<decimal>(kv.Value), StringComparer.OrdinalIgnoreCase);
        var avCur = CopyMap(curAvans);
        var mkCur = CopyMap(curMakbuz);
        var vrCur = CopyMap(curVirman);

        var map = new Dictionary<string, (decimal tsf, decimal mm, decimal vb, decimal va, int adet, string ad)>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in rows)
        {
            if (string.IsNullOrWhiteSpace(r.TcKimlikNo) || r.TcKimlikNo.Length < 10) continue;

            bool isAvans = string.Equals(r.HAREKET_TURU, HT_AVANS, StringComparison.OrdinalIgnoreCase);
            bool isMakbuz = string.Equals(r.HAREKET_TURU, HT_MAKBUZ, StringComparison.OrdinalIgnoreCase);
            bool isVirman = string.Equals(r.HAREKET_TURU, HT_VIRMAN, StringComparison.OrdinalIgnoreCase);

            // 1) Önce cari yıl SabNet satırıyla eşleş (varsa). Eşleştiyse Logo'da kalır.
            bool currentMatched = false;
            if (isAvans) currentMatched = TryExclude(avCur, r.TcKimlikNo, r.Tutar);
            else if (isMakbuz) currentMatched = TryExclude(mkCur, r.TcKimlikNo, r.Tutar);
            else if (isVirman) currentMatched = TryExclude(vrCur, r.TcKimlikNo, r.Tutar);

            // 2) Cari yılla eşleşmediyse, diğer yıl SabNet satırına denk geliyor mu? — varsa Logo'dan düş
            if (!currentMatched && excl != null)
            {
                if (isAvans && TryExclude(excl.Avans, r.TcKimlikNo, r.Tutar)) continue;
                if (isMakbuz && TryExclude(excl.Makbuz, r.TcKimlikNo, r.Tutar)) continue;
                if (isVirman && TryExclude(excl.Virman, r.TcKimlikNo, r.Tutar)) continue;
            }

            if (!map.TryGetValue(r.TcKimlikNo, out var v))
                v = (0, 0, 0, 0, 0, r.Unvan);
            else if (string.IsNullOrWhiteSpace(v.ad))
                v.ad = r.Unvan;

            if (isAvans) v.tsf += r.Tutar;
            else if (isMakbuz) v.mm += r.Tutar;
            else if (isVirman)
            {
                if (r.Tutar >= 0) v.vb += r.Tutar;
                else v.va += Math.Abs(r.Tutar);
            }
            v.adet++;
            map[r.TcKimlikNo] = v;
        }

        return map.ToDictionary(
            kv => kv.Key,
            kv => new LogoItem(kv.Value.tsf, kv.Value.mm, kv.Value.vb, kv.Value.va, kv.Value.adet, kv.Value.ad),
            StringComparer.OrdinalIgnoreCase);
    }

    // ===== Satır-seviyesi detay (tek TC) =====
    // Logo ve SabNet satırlarını Tutar üzerinden greedy eşleştirir.
    // Eşleşmeyen Logo → "SABNET'TE YOK", eşleşmeyen SabNet → "LOGO'DA YOK"

    private const decimal EslesmeToleransi = 0.01m;

    private record LogoSatir(DateTime Tarih, string IslemNo, string BelgeNo, decimal Tutar, string Unvan);
    private record SabSatir(DateTime Tarih, string No, decimal Tutar, string Aciklama);

    public async Task<MustahsilTcDetay> TcDetayAsync(string tc, DateTime bas, DateTime bit, string kampanyaYili, CancellationToken ct = default)
    {
        var detay = new MustahsilTcDetay { TcKimlikNo = tc };
        if (string.IsNullOrWhiteSpace(tc)) return detay;

        // Paralel çek
        var logoTask = LogoSatirlarAsync(tc, bas, bit, ct);
        var avansTask = SabAvansSatirlariAsync(tc, bas, bit, kampanyaYili, ct);
        var makbuzTask = SabMakbuzSatirlariAsync(tc, bas, bit, kampanyaYili, ct);
        var cariTask = SabCariSatirlariAsync(tc, bas, bit, kampanyaYili, ct);
        var exclTask = DigerYilExclusionAsync(bas, bit, kampanyaYili, ct);
        await Task.WhenAll(logoTask, avansTask, makbuzTask, cariTask, exclTask);

        detay.AdiSoyadi = logoTask.Result.FirstOrDefault()?.Unvan ?? "";

        var logoByCat = await LogoSatirlarByCategoryAsync(tc, bas, bit, ct);
        var excl = exclTask.Result;

        // Her kategori için TC'nin diğer yıl exclusion listesini (varsa) kopyala — Eslestir bunu tüketecek
        static List<decimal>? TcList(Dictionary<string, List<decimal>> map, string tc) =>
            map.TryGetValue(tc, out var l) ? new List<decimal>(l) : null;

        detay.Avans  = Eslestir(logoByCat.Avans,  avansTask.Result,  TcList(excl.Avans,  tc));
        detay.Makbuz = Eslestir(logoByCat.Makbuz, makbuzTask.Result, TcList(excl.Makbuz, tc));
        detay.Virman = Eslestir(logoByCat.Virman, cariTask.Result,   TcList(excl.Virman, tc));
        return detay;
    }

    private static bool TryExcludeList(List<decimal>? list, decimal tutar)
    {
        if (list == null) return false;
        for (int i = 0; i < list.Count; i++)
        {
            if (Math.Abs(list[i] - tutar) < EslesmeToleransi)
            {
                list.RemoveAt(i);
                return true;
            }
        }
        return false;
    }

    private List<MustahsilDetayEslesme> Eslestir(List<LogoSatir> logo, List<SabSatir> sab, List<decimal>? excl)
    {
        var sonuc = new List<MustahsilDetayEslesme>();
        var kullanilan = new bool[sab.Count];

        // Logo'yu dolaş — önce cari yıl SabNet satırına eşleşmeye çalış, değilse diğer yıl exclusion'a bak
        foreach (var l in logo)
        {
            int idx = -1;
            for (int i = 0; i < sab.Count; i++)
            {
                if (kullanilan[i]) continue;
                if (Math.Abs(sab[i].Tutar - l.Tutar) < EslesmeToleransi)
                {
                    idx = i;
                    break;
                }
            }
            if (idx >= 0)
            {
                kullanilan[idx] = true;
                sonuc.Add(new MustahsilDetayEslesme
                {
                    LogoTarihi = l.Tarih,
                    LogoIslemNo = l.IslemNo,
                    LogoBelgeNo = l.BelgeNo,
                    LogoTutar = l.Tutar,
                    SabNetTarihi = sab[idx].Tarih,
                    SabNetNo = sab[idx].No,
                    SabNetTutar = sab[idx].Tutar,
                    SabNetAciklama = sab[idx].Aciklama,
                    Durum = "EŞLEŞTİ"
                });
                continue;
            }

            // Cari yıl eşleşmedi — diğer yıl (excl) listesinde var mı? Varsa bu Logo satırı
            // diğer kampanyaya ait kabul edilir ve gösterimden sessizce düşer.
            if (TryExcludeList(excl, l.Tutar))
                continue;

            sonuc.Add(new MustahsilDetayEslesme
            {
                LogoTarihi = l.Tarih,
                LogoIslemNo = l.IslemNo,
                LogoBelgeNo = l.BelgeNo,
                LogoTutar = l.Tutar,
                Durum = "SABNET'TE YOK"
            });
        }

        // Kullanılmamış SabNet satırları
        for (int i = 0; i < sab.Count; i++)
        {
            if (kullanilan[i]) continue;
            sonuc.Add(new MustahsilDetayEslesme
            {
                SabNetTarihi = sab[i].Tarih,
                SabNetNo = sab[i].No,
                SabNetTutar = sab[i].Tutar,
                SabNetAciklama = sab[i].Aciklama,
                Durum = "LOGO'DA YOK"
            });
        }

        return sonuc
            .OrderBy(x => x.Eslesti ? 1 : 0)
            .ThenBy(x => x.LogoTarihi ?? x.SabNetTarihi ?? DateTime.MinValue)
            .ToList();
    }

    private record LogoKategoriler(List<LogoSatir> Avans, List<LogoSatir> Makbuz, List<LogoSatir> Virman);

    private async Task<LogoKategoriler> LogoSatirlarByCategoryAsync(string tc, DateTime bas, DateTime bit, CancellationToken ct)
    {
        const string sql = @"
SELECT HAREKET_TURU,
       TARIH,
       ISLEM_NO    = ISNULL(ISLEM_NO, ''),
       BELGE_NO    = ISNULL(BELGE_NO, ''),
       Tutar       = ISNULL(TOPLAM, 0),
       Unvan       = ISNULL(CARI_HESAP_UNVANI, '')
FROM INF_Mustahsil_Hareketleri_211 WITH(NOLOCK)
WHERE TARIH >= @bas AND TARIH < DATEADD(day, 1, @bit)
  AND CARI_HESAP_KODU IS NOT NULL
  AND LTRIM(RTRIM(SUBSTRING(CARI_HESAP_KODU, 2, 11))) = @tc
  AND HAREKET_TURU IN (@ht1, @ht2, @ht3)
ORDER BY TARIH";
        using var c = _db.CreateConnection();
        var rows = await c.QueryAsync<(string HAREKET_TURU, DateTime TARIH, string ISLEM_NO, string BELGE_NO, decimal Tutar, string Unvan)>(
            new CommandDefinition(sql,
                new { bas = bas.Date, bit = bit.Date, tc, ht1 = HT_AVANS, ht2 = HT_MAKBUZ, ht3 = HT_VIRMAN },
                cancellationToken: ct));

        var avans = new List<LogoSatir>();
        var makbuz = new List<LogoSatir>();
        var virman = new List<LogoSatir>();
        foreach (var r in rows)
        {
            var sat = new LogoSatir(r.TARIH, r.ISLEM_NO, r.BELGE_NO, r.Tutar, r.Unvan);
            if (string.Equals(r.HAREKET_TURU, HT_AVANS, StringComparison.OrdinalIgnoreCase)) avans.Add(sat);
            else if (string.Equals(r.HAREKET_TURU, HT_MAKBUZ, StringComparison.OrdinalIgnoreCase)) makbuz.Add(sat);
            else if (string.Equals(r.HAREKET_TURU, HT_VIRMAN, StringComparison.OrdinalIgnoreCase)) virman.Add(sat);
        }
        return new LogoKategoriler(avans, makbuz, virman);
    }

    // Yalnızca başlık/Unvan yakalamak için kullanılan basit liste
    private async Task<List<LogoSatir>> LogoSatirlarAsync(string tc, DateTime bas, DateTime bit, CancellationToken ct)
    {
        const string sql = @"
SELECT TOP 1
       TARIH,
       ISLEM_NO = ISNULL(ISLEM_NO, ''),
       BELGE_NO = ISNULL(BELGE_NO, ''),
       Tutar    = ISNULL(TOPLAM, 0),
       Unvan    = ISNULL(CARI_HESAP_UNVANI, '')
FROM INF_Mustahsil_Hareketleri_211 WITH(NOLOCK)
WHERE TARIH >= @bas AND TARIH < DATEADD(day, 1, @bit)
  AND CARI_HESAP_KODU IS NOT NULL
  AND LTRIM(RTRIM(SUBSTRING(CARI_HESAP_KODU, 2, 11))) = @tc";
        using var c = _db.CreateConnection();
        var rows = await c.QueryAsync<(DateTime TARIH, string ISLEM_NO, string BELGE_NO, decimal Tutar, string Unvan)>(
            new CommandDefinition(sql, new { bas = bas.Date, bit = bit.Date, tc }, cancellationToken: ct));
        return rows.Select(r => new LogoSatir(r.TARIH, r.ISLEM_NO, r.BELGE_NO, r.Tutar, r.Unvan)).ToList();
    }

    private async Task<List<SabSatir>> SabAvansSatirlariAsync(string tc, DateTime bas, DateTime bit, string kampanyaYili, CancellationToken ct)
    {
        int basSerial = (int)(bas.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        int bitSerial = (int)(bit.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        const string sql = @"
SELECT FormNo = ISNULL(FormNo,''),
       FormTarihi,
       Tutar = ISNULL(Tutar, 0),
       Aciklama = ISNULL(Aciklama, '')
FROM PMHS_AvansFormu WITH(NOLOCK)
WHERE FormTarihi >= @bas AND FormTarihi <= @bit
  AND ISNULL(TcKimlikNo,'') = @tc
  AND ISNULL(SozlesmeYili,'') = @yil
ORDER BY FormTarihi";
        using var c = _db.CreatePmhsConnection();
        var rows = await c.QueryAsync<(string FormNo, int FormTarihi, decimal Tutar, string Aciklama)>(
            new CommandDefinition(sql, new { bas = basSerial, bit = bitSerial, tc, yil = kampanyaYili }, cancellationToken: ct));
        return rows.Select(r => new SabSatir(SerialToDate(r.FormTarihi), r.FormNo, r.Tutar, r.Aciklama)).ToList();
    }

    private async Task<List<SabSatir>> SabMakbuzSatirlariAsync(string tc, DateTime bas, DateTime bit, string kampanyaYili, CancellationToken ct)
    {
        int basSerial = (int)(bas.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        int bitSerial = (int)(bit.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        const string sql = @"
SELECT MakbuzNo = ISNULL(MakbuzNo,''),
       Tarih,
       Tutar = ISNULL(NetHakedis, 0),
       Aciklama = ISNULL(AdiSoyadi, '')
FROM PMHS_MustahsilMakbuzu WITH(NOLOCK)
WHERE Tarih >= @bas AND Tarih <= @bit
  AND ISNULL(TcKimlikNo,'') = @tc
  AND ISNULL(KampanyaYili,'') = @yil
ORDER BY Tarih";
        using var c = _db.CreatePmhsConnection();
        var rows = await c.QueryAsync<(string MakbuzNo, int Tarih, decimal Tutar, string Aciklama)>(
            new CommandDefinition(sql, new { bas = basSerial, bit = bitSerial, tc, yil = kampanyaYili }, cancellationToken: ct));
        return rows.Select(r => new SabSatir(SerialToDate(r.Tarih), r.MakbuzNo, r.Tutar, r.Aciklama)).ToList();
    }

    private async Task<List<SabSatir>> SabCariSatirlariAsync(string tc, DateTime bas, DateTime bit, string kampanyaYili, CancellationToken ct)
    {
        int basSerial = (int)(bas.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        int bitSerial = (int)(bit.Date - new DateTime(1900, 1, 1)).TotalDays + 2;
        const string sql = @"
SELECT FisNo = ISNULL(FisNo,''),
       Tarih,
       BA = ISNULL(BA, ''),
       Tutar = ISNULL(Tutar, 0),
       Aciklama = ISNULL(Aciklama, '')
FROM PMHS_CariHareketler WITH(NOLOCK)
WHERE Tarih >= @bas AND Tarih <= @bit
  AND ISNULL(TcKimlikNo,'') = @tc
  AND ISNULL(SozlesmeYili,'') = @yil
ORDER BY Tarih";
        using var c = _db.CreatePmhsConnection();
        var rows = await c.QueryAsync<(string FisNo, int Tarih, string BA, decimal Tutar, string Aciklama)>(
            new CommandDefinition(sql, new { bas = basSerial, bit = bitSerial, tc, yil = kampanyaYili }, cancellationToken: ct));

        // BA='ALACAK' ise Logo'da negatif tutar olarak görünüyor; işareti buna göre çevir
        return rows.Select(r =>
        {
            decimal imzali = string.Equals(r.BA, "ALACAK", StringComparison.OrdinalIgnoreCase) ? -r.Tutar : r.Tutar;
            var aciklama = $"{r.BA}: {r.Aciklama}".Trim();
            return new SabSatir(SerialToDate(r.Tarih), r.FisNo, imzali, aciklama);
        }).ToList();
    }

    private static DateTime SerialToDate(int serial) =>
        new DateTime(1900, 1, 1).AddDays(serial - 2);
}
