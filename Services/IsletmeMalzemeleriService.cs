using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// İşletme Malzemeleri Raporu için Logo Tiger'dan veri çeker:
///  - Devir Stoğu: <c>LV_211_01_STINVTOT_V1</c> kümülatif (SUM ONHAND DATE_ &lt; devirTarihi)
///  - Mevcut Stok: <c>LV_211_01_GNTOTST</c> anlık
///  - Giren / Sarfiyat: <c>LG_211_01_STLINE</c> IOCODE=1 / IOCODE=4
///
/// INVENNO=-1 satırı global toplamı temsil eder; tüm sorgularda <c>INVENNO &gt;= 0</c>
/// filtresi uygulanır. Bkz. [[feedback_logo_ambar_minus1]].
/// </summary>
public class IsletmeMalzemeleriService
{
    private readonly DatabaseService _db;

    public IsletmeMalzemeleriService(DatabaseService db) => _db = db;

    public async Task<List<IsletmeMalzemeSatiri>> RaporGetirAsync(
        IEnumerable<string> kodOnekleri,
        DateTime devirTarihi,
        DateTime? kampanyaBas,
        DateTime? kampanyaBitis,
        DateTime? gunlukGelenTarihi,
        IEnumerable<int>? ambarFiltre,
        IReadOnlyDictionary<string, string> aciklamalar,
        CancellationToken ct = default)
    {
        var onekListesi = kodOnekleri
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => "%" + x.Trim() + "%")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (onekListesi.Count == 0) return new();

        var devirGunBasi     = devirTarihi.Date;
        var kampBasGunBasi   = kampanyaBas?.Date;
        var kampBitGunSonu   = kampanyaBitis?.Date.AddDays(1).AddSeconds(-1);
        var gunlukBas        = gunlukGelenTarihi?.Date;
        var gunlukSon        = gunlukGelenTarihi?.Date.AddDays(1).AddSeconds(-1);

        var ambarListesi = (ambarFiltre ?? Array.Empty<int>())
            .Where(x => x >= 0)
            .Distinct()
            .ToList();
        bool tumAmbarlar = ambarListesi.Count == 0;

        var itTbl   = _db.GetTableName("ITEMS");        // LG_211_ITEMS
        var stTbl   = _db.GetPeriodTableName("STLINE"); // LG_211_01_STLINE
        var v1View  = _db.GetViewName("STINVTOT_V1");   // LV_211_01_STINVTOT_V1
        var gntView = _db.GetViewName("GNTOTST");       // LV_211_01_GNTOTST
        var unTbl   = _db.GetTableName("UNITSETL");     // LG_211_UNITSETL

        // Ambar filtresi — INVENNO=-1 her durumda hariç.
        string AmbarKosul(string alias) =>
            tumAmbarlar ? $"{alias}.INVENNO >= 0" : $"{alias}.INVENNO IN @ambarlar";

        // Kod öneki LIKE ifadeleri (inline parametre)
        string KodKosul(string alias) =>
            "(" + string.Join(" OR ",
                Enumerable.Range(0, onekListesi.Count).Select(i => $"{alias}.CODE LIKE @onek{i}")) + ")";

        var sql = $@"
WITH Kartlar AS (
    SELECT I.LOGICALREF, I.CODE, I.NAME, I.UNITSETREF
    FROM {itTbl} I WITH(NOLOCK)
    WHERE I.ACTIVE = 0 AND {KodKosul("I")}
),
DevirCte AS (
    SELECT V1.STOCKREF, SUM(V1.ONHAND) AS Miktar
    FROM {v1View} V1 WITH(NOLOCK)
    INNER JOIN Kartlar K ON K.LOGICALREF = V1.STOCKREF
    WHERE V1.DATE_ < @devir AND {AmbarKosul("V1")}
    GROUP BY V1.STOCKREF
),
MevcutCte AS (
    SELECT GT.STOCKREF, SUM(GT.ONHAND) AS Miktar
    FROM {gntView} GT WITH(NOLOCK)
    INNER JOIN Kartlar K ON K.LOGICALREF = GT.STOCKREF
    WHERE {AmbarKosul("GT")}
    GROUP BY GT.STOCKREF
),
HareketCte AS (
    SELECT
        SL.STOCKREF,
        Gelen        = SUM(CASE WHEN SL.IOCODE = 1 AND SL.DATE_ >= @devir THEN SL.AMOUNT ELSE 0 END),
        GunlukGelen  = SUM(CASE WHEN SL.IOCODE = 1
                                 AND @gunBas IS NOT NULL
                                 AND SL.DATE_ BETWEEN @gunBas AND @gunSon
                                THEN SL.AMOUNT ELSE 0 END),
        ToplamSarf   = SUM(CASE WHEN SL.IOCODE = 4 AND SL.DATE_ >= @devir THEN SL.AMOUNT ELSE 0 END),
        KampanyaSarf = SUM(CASE WHEN SL.IOCODE = 4
                                 AND @kampBas IS NOT NULL AND @kampBit IS NOT NULL
                                 AND SL.DATE_ BETWEEN @kampBas AND @kampBit
                                THEN SL.AMOUNT ELSE 0 END)
    FROM {stTbl} SL WITH(NOLOCK)
    INNER JOIN Kartlar K ON K.LOGICALREF = SL.STOCKREF
    WHERE SL.CANCELLED = 0
    GROUP BY SL.STOCKREF
)
SELECT
    Kod          = K.CODE,
    Ad           = K.NAME,
    Birim        = (SELECT TOP 1 U.CODE
                    FROM {unTbl} U WITH(NOLOCK)
                    WHERE U.UNITSETREF = K.UNITSETREF AND U.MAINUNIT = 1),
    Devir        = ISNULL(D.Miktar, 0),
    Mevcut       = ISNULL(M.Miktar, 0),
    Gelen        = ISNULL(H.Gelen, 0),
    GunlukGelen  = ISNULL(H.GunlukGelen, 0),
    ToplamSarf   = ISNULL(H.ToplamSarf, 0),
    KampanyaSarf = ISNULL(H.KampanyaSarf, 0)
FROM Kartlar K
LEFT JOIN DevirCte   D ON D.STOCKREF = K.LOGICALREF
LEFT JOIN MevcutCte  M ON M.STOCKREF = K.LOGICALREF
LEFT JOIN HareketCte H ON H.STOCKREF = K.LOGICALREF
ORDER BY K.CODE";

        var parametre = new DynamicParameters();
        parametre.Add("@devir",   devirGunBasi);
        parametre.Add("@gunBas",  gunlukBas);
        parametre.Add("@gunSon",  gunlukSon);
        parametre.Add("@kampBas", kampBasGunBasi);
        parametre.Add("@kampBit", kampBitGunSonu);
        for (int i = 0; i < onekListesi.Count; i++)
            parametre.Add($"@onek{i}", onekListesi[i]);
        if (!tumAmbarlar) parametre.Add("@ambarlar", ambarListesi);

        using var conn = _db.CreateConnection();
        var rows = (await conn.QueryAsync<dynamic>(new CommandDefinition(
            sql, parametre, commandTimeout: 180, cancellationToken: ct))).ToList();

        var sonuc = new List<IsletmeMalzemeSatiri>(rows.Count);
        foreach (var r in rows)
        {
            string kod      = ((string?)r.Kod ?? "").Trim();
            decimal devir   = Convert.ToDecimal(r.Devir   ?? 0);
            decimal mevcut  = Convert.ToDecimal(r.Mevcut  ?? 0);
            decimal gelen   = Convert.ToDecimal(r.Gelen   ?? 0);
            decimal gunluk  = Convert.ToDecimal(r.GunlukGelen ?? 0);
            decimal toplamS = Convert.ToDecimal(r.ToplamSarf  ?? 0);
            decimal kampS   = Convert.ToDecimal(r.KampanyaSarf ?? 0);

            decimal kampDisi = (kampanyaBas.HasValue && kampanyaBitis.HasValue)
                                ? toplamS - kampS
                                : toplamS;

            aciklamalar.TryGetValue(kod, out var aciklama);

            sonuc.Add(new IsletmeMalzemeSatiri
            {
                MalzemeKodu        = kod,
                MalzemeAdi         = ((string?)r.Ad ?? "").Trim(),
                Birim              = ((string?)r.Birim ?? "").Trim(),
                DevirStogu         = devir,
                GunlukGelenTarihi  = gunlukGelenTarihi,
                GunlukGelen        = gunluk,
                GelenToplam        = gelen,
                KampanyaSarfiyati  = kampS,
                KampanyaDisiSarf   = kampDisi,
                ToplamSarfiyat     = toplamS,
                Mevcut             = mevcut,
                Aciklama           = aciklama ?? ""
            });
        }

        return sonuc;
    }

    public async Task<List<AmbarSecenek>> AmbarSecenekleriAsync(CancellationToken ct = default)
    {
        const string sql = @"
SELECT
    CW.NR              AS AmbarNo,
    ISNULL(CW.NAME,'') AS Ad
FROM L_CAPIWHOUSE CW WITH(NOLOCK)
WHERE CW.FIRMNR = @firma AND CW.NR >= 0
ORDER BY CW.NR";
        using var conn = _db.CreateConnection();
        return (await conn.QueryAsync<AmbarSecenek>(
            new CommandDefinition(sql, new { firma = _db.FirmaNo }, cancellationToken: ct))).ToList();
    }
}
