using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Logo'daki yıllık finans view'lerinden (INF_MD_FINANS_PROJE_RAPORU_211_YYYY)
/// veri çeker. Her yıl için ayrı view var — bu servis tarih aralığındaki yılların
/// view'lerini dinamik olarak UNION ALL ile birleştirip tek tablo gibi sunar.
/// </summary>
public class FinansRaporService
{
    private readonly DatabaseService _db;
    private const string ViewPrefix = "INF_MD_FINANS_PROJE_RAPORU_";

    public FinansRaporService(DatabaseService db) => _db = db;

    /// <summary>
    /// Sunucudaki mevcut yıllık view'leri keşfeder (ör. 2021..2026).
    /// Her çağrıda sys.views'a sorulur — yeni yıl view'i açılınca otomatik görünür.
    /// </summary>
    public async Task<List<int>> MevcutYillarAsync(CancellationToken ct = default)
    {
        var prefix = $"{ViewPrefix}{_db.FirmaNo}_";
        var sql = @"SELECT name FROM sys.views WITH(NOLOCK)
                    WHERE name LIKE @p + '____'
                    ORDER BY name";
        using var conn = _db.CreateConnection();
        var isimler = await conn.QueryAsync<string>(
            new CommandDefinition(sql, new { p = prefix }, cancellationToken: ct));

        var yillar = new List<int>();
        foreach (var ad in isimler)
        {
            var son4 = ad[^4..];
            if (int.TryParse(son4, out var y)) yillar.Add(y);
        }
        return yillar;
    }

    public async Task<List<FinansRaporSatiri>> GetHareketlerAsync(
        DateTime baslangic, DateTime bitis, CancellationToken ct = default)
    {
        // Sessiz global üst sınır
        bitis = SistemTarihi.Clamp(bitis);
        if (bitis < baslangic) return new();

        var mevcutYillar = (await MevcutYillarAsync(ct)).ToHashSet();
        var baslangicYil = baslangic.Year;
        var bitisYil     = bitis.Year;

        // Aralıktaki ve sunucuda mevcut olan yılların view'lerini UNION ALL ile birleştir
        var parcalar = new List<string>();
        for (int y = baslangicYil; y <= bitisYil; y++)
            if (mevcutYillar.Contains(y))
                parcalar.Add($"SELECT * FROM {ViewPrefix}{_db.FirmaNo}_{y} WITH(NOLOCK)");

        if (parcalar.Count == 0) return new();

        var birlesik = string.Join("\nUNION ALL\n", parcalar);
        var sql = $@"
SELECT
    LOGICALREF             = V.LOGICALREF,
    FIS_TURU               = ISNULL(V.FIS_TURU, ''),
    FIRMA                  = ISNULL(V.FIRMA, ''),
    HAREKET_TURU           = ISNULL(V.HAREKET_TURU, ''),
    MODUL                  = ISNULL(V.MODUL, ''),
    CH_KOD                 = ISNULL(V.CH_KOD, ''),
    CH_UNVANI              = ISNULL(V.CH_UNVANI, ''),
    BANKA_HESAP_KODU       = ISNULL(V.BANKA_HESAP_KODU, ''),
    BANKA_HESAP_ACIKLAMASI = ISNULL(V.BANKA_HESAP_ACIKLAMASI, ''),
    PROJE_KODU             = ISNULL(V.PROJE_KODU, ''),
    PROJE_ADI              = ISNULL(V.PROJE_ADI, ''),
    TARIH                  = V.TARIH,
    YIL                    = V.YIL,
    AY                     = V.AY,
    GUN                    = V.GUN,
    HARAKET_OZEL_KODU      = ISNULL(V.HARAKET_OZEL_KODU, ''),
    ISLEM_NO               = ISNULL(CAST(V.ISLEM_NO AS NVARCHAR(50)), ''),
    HAVALE                 = ISNULL(V.HAVALE, 0),
    CEK                    = ISNULL(V.CEK, 0),
    DEVIR                  = ISNULL(V.DEVIR, 0),
    DIGER                  = ISNULL(V.DIGER, 0),
    SIGN                   = ISNULL(V.SIGN, 0),
    SPEC_ODE               = ISNULL(V.SPECODE, '')
FROM ({birlesik}) V
WHERE V.TARIH >= @Baslangic AND V.TARIH <= @Bitis
ORDER BY V.TARIH, V.LOGICALREF";

        using var conn = _db.CreateConnection();
        var cmd = new CommandDefinition(sql,
            new { Baslangic = baslangic.Date, Bitis = bitis.Date },
            commandTimeout: 120, cancellationToken: ct);
        var rows = await conn.QueryAsync<FinansRaporSatiri>(cmd);
        return rows.AsList();
    }
}
