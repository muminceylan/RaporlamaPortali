using Dapper;
using Microsoft.Data.SqlClient;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_2025 view'ünün mantığını kullanıcı tarafından
/// verilen parametrik malzeme kodları için çalıştırır. Aynı TRCODE eşlemesi, aynı kolonlar.
/// </summary>
public class MalzemeHareketService
{
    private readonly DatabaseService _db;

    public MalzemeHareketService(DatabaseService db) => _db = db;

    public async Task<List<MalzemeHareketSatiri>> GetHareketlerAsync(
        IEnumerable<string> malzemeKodlari,
        DateTime baslangic,
        DateTime bitis,
        CancellationToken ct = default)
    {
        var kodlar = malzemeKodlari
            .Where(k => !string.IsNullOrWhiteSpace(k))
            .Select(k => k.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (kodlar.Count == 0) return new();

        // Sessiz global üst sınır — stok sayfası dışında her yerde uygulanıyor.
        bitis = SistemTarihi.Clamp(bitis);

        var stlineTbl  = _db.GetPeriodTableName("STLINE");   // LG_211_01_STLINE
        var stficheTbl = _db.GetPeriodTableName("STFICHE");  // LG_211_01_STFICHE
        var itemsTbl   = _db.GetTableName("ITEMS");          // LG_211_ITEMS
        var clcardTbl  = _db.GetTableName("CLCARD");         // LG_211_CLCARD

        // View ile aynı mantık: TRCODE 1/3/8/13, CANCELLED=0, LPRODSTAT=0,
        // STFICHEREF<>0; UINFO2/UINFO1 ile alt birim → ana birim dönüşümü; VATMATRAH tutar.
        var sql = $@"
SELECT
    YIL               = YEAR(STLINE.DATE_),
    AY                = MONTH(STLINE.DATE_),
    TARIH             = STLINE.DATE_,
    FIS_TURU          = CASE STLINE.TRCODE
                            WHEN 1  THEN N'Satınalma İrsaliyesi'
                            WHEN 3  THEN N'Toptan Satış İade İrsaliyesi'
                            WHEN 8  THEN N'Toptan Satış İrsaliyesi'
                            WHEN 13 THEN N'Üretimden Giriş Fişi'
                        END,
    FIS_NUMARASI      = ISNULL(STFICHE.FICHENO, ''),
    CARI_HESAP_KODU   = ISNULL(CLCARD.CODE, ''),
    CARI_HESAP_UNVANI = ISNULL(CLCARD.DEFINITION_, ''),
    MALZEME_KODU      = ITEMS.CODE,
    MALZEME_ACIKLAMASI= ITEMS.NAME,
    GIRIS_MIKTARI     = ISNULL(CASE WHEN STLINE.TRCODE IN (1,3,13)
                              THEN (STLINE.AMOUNT * STLINE.UINFO2 / STLINE.UINFO1) END, 0),
    GIRIS_FIYATI      = ISNULL(CASE WHEN STLINE.TRCODE IN (1,3,13)
                              THEN STLINE.PRICE END, 0),
    GIRIS_TUTARI      = ISNULL(CASE WHEN STLINE.TRCODE IN (1,3,13)
                              THEN STLINE.VATMATRAH END, 0),
    CIKIS_MIKTARI     = ISNULL(CASE WHEN STLINE.TRCODE IN (8)
                              THEN (STLINE.AMOUNT * STLINE.UINFO2 / STLINE.UINFO1) * -1 END, 0),
    CIKIS_FIYATI      = ISNULL(CASE WHEN STLINE.TRCODE IN (8)
                              THEN STLINE.PRICE * -1 END, 0),
    CIKIS_TUTARI      = ISNULL(CASE WHEN STLINE.TRCODE IN (8)
                              THEN STLINE.VATMATRAH * -1 END, 0)
FROM {stlineTbl} STLINE WITH(NOLOCK)
LEFT JOIN {itemsTbl}   ITEMS   WITH(NOLOCK) ON STLINE.STOCKREF   = ITEMS.LOGICALREF
LEFT JOIN {stficheTbl} STFICHE WITH(NOLOCK) ON STLINE.STFICHEREF = STFICHE.LOGICALREF
LEFT JOIN {clcardTbl}  CLCARD  WITH(NOLOCK) ON STLINE.CLIENTREF  = CLCARD.LOGICALREF
WHERE STLINE.CANCELLED = 0
  AND STLINE.LPRODSTAT = 0
  AND STLINE.STFICHEREF <> 0
  AND STLINE.TRCODE IN (1,3,8,13)
  AND STLINE.DATE_ >= @Baslangic
  AND STLINE.DATE_ <= @Bitis
  AND ITEMS.CODE IN @Kodlar
ORDER BY STLINE.DATE_, STFICHE.FICHENO";

        using var conn = _db.CreateConnection();
        var cmd = new CommandDefinition(sql,
            new { Baslangic = baslangic.Date, Bitis = bitis.Date, Kodlar = kodlar },
            cancellationToken: ct);

        var rows = await conn.QueryAsync<MalzemeHareketSatiri>(cmd);
        return rows.AsList();
    }

    /// <summary>
    /// ITEMS tablosunda kod veya açıklamaya göre hızlı arama — autocomplete için.
    /// </summary>
    public async Task<List<(string Kod, string Ad)>> MalzemeAraAsync(string? query, int limit = 30, CancellationToken ct = default)
    {
        var itemsTbl = _db.GetTableName("ITEMS");
        var q        = (query ?? "").Trim();
        var param    = "%" + q + "%";

        var sql = $@"
SELECT TOP (@Limit) CODE, NAME
FROM {itemsTbl} WITH(NOLOCK)
WHERE ACTIVE = 0
  AND (@Q = '' OR CODE LIKE @Param OR NAME LIKE @Param)
ORDER BY CODE";

        using var conn = _db.CreateConnection();
        var rows = await conn.QueryAsync(
            new CommandDefinition(sql, new { Limit = limit, Q = q, Param = param }, cancellationToken: ct));
        var list = new List<(string Kod, string Ad)>();
        foreach (var r in rows)
            list.Add(((string)(r.CODE ?? ""), (string)(r.NAME ?? "")));
        return list;
    }
}
