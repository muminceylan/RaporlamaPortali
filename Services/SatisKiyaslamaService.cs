using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Tek bir cariye ait iki tarih aralığını yan yana kıyaslayan satış raporu.
/// Kaynak: <c>LG_211_01_INVOICE</c> (TRCODE=8 toptan satış, CANCELLED=0)
/// + <c>LG_211_01_STLINE</c> (IOCODE=4, TRCODE=8 satış çıkış satırları)
/// + <c>LG_211_ITEMS</c> (malzeme kart) + <c>LG_211_CLCARD</c> (cari kart).
/// </summary>
public class SatisKiyaslamaService
{
    private readonly DatabaseService _db;

    public SatisKiyaslamaService(DatabaseService db) => _db = db;

    /// <summary>
    /// Cari kodu/ünvan parçası ile en fazla 30 öneri döner (autocomplete).
    /// </summary>
    public async Task<List<CariBilgi>> CariAraAsync(string arama, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(arama)) return new();

        var clcTbl = _db.GetTableName("CLCARD");
        var sql = $@"
SELECT TOP 30
    LogicalRef = C.LOGICALREF,
    Kod        = C.CODE,
    Unvan      = ISNULL(C.DEFINITION_, ''),
    Vkn        = ISNULL(C.TCKNO, ''),
    Sehir      = ISNULL(C.CITY, '')
FROM {clcTbl} C WITH(NOLOCK)
WHERE C.ACTIVE = 0
  AND (C.CODE LIKE @q OR C.DEFINITION_ LIKE @q)
ORDER BY C.CODE";
        using var conn = _db.CreateConnection();
        return (await conn.QueryAsync<CariBilgi>(
            new CommandDefinition(sql, new { q = "%" + arama.Trim() + "%" },
                                  cancellationToken: ct))).ToList();
    }

    /// <summary>
    /// Cari koduyla tek bir cari kartı bulur (tam eşleşme).
    /// </summary>
    public async Task<CariBilgi?> CariBulAsync(string kod, CancellationToken ct = default)
    {
        var clcTbl = _db.GetTableName("CLCARD");
        var sql = $@"
SELECT TOP 1
    LogicalRef = C.LOGICALREF,
    Kod        = C.CODE,
    Unvan      = ISNULL(C.DEFINITION_, ''),
    Vkn        = ISNULL(C.TCKNO, ''),
    Sehir      = ISNULL(C.CITY, '')
FROM {clcTbl} C WITH(NOLOCK)
WHERE C.CODE = @kod";
        using var conn = _db.CreateConnection();
        return (await conn.QueryAsync<CariBilgi>(
            new CommandDefinition(sql, new { kod }, cancellationToken: ct))).FirstOrDefault();
    }

    /// <summary>
    /// İki tarih aralığını malzeme bazında kıyaslar + dönem özetlerini çıkarır.
    /// </summary>
    public async Task<SatisKiyaslamaSonuc> KiyaslaAsync(
        string cariKod,
        DateTime donemABas, DateTime donemABit,
        DateTime donemBBas, DateTime donemBBit,
        CancellationToken ct = default)
    {
        var sonuc = new SatisKiyaslamaSonuc
        {
            DonemA = new DonemOzeti { Baslangic = donemABas.Date, Bitis = donemABit.Date },
            DonemB = new DonemOzeti { Baslangic = donemBBas.Date, Bitis = donemBBit.Date }
        };

        var cari = await CariBulAsync(cariKod, ct);
        if (cari == null) return sonuc;
        sonuc.Cari = cari;

        var itTbl  = _db.GetTableName("ITEMS");
        var invTbl = _db.GetPeriodTableName("INVOICE");
        var stTbl  = _db.GetPeriodTableName("STLINE");
        var unTbl  = _db.GetTableName("UNITSETL");

        // Tek round-trip — 2 result set: malzeme bazlı kıyas + dönem özetleri
        var sql = $@"
DECLARE @AB datetime = @aBas, @AT datetime = @aBit;
DECLARE @BB datetime = @bBas, @BT datetime = @bBit;
DECLARE @CRef int = @ref;

-- 1) MALZEME BAZINDA KIYASLAMA
WITH SL AS (
    SELECT
        SL.STOCKREF,
        Donem = CASE WHEN INV.DATE_ BETWEEN @AB AND @AT THEN 'A'
                     WHEN INV.DATE_ BETWEEN @BB AND @BT THEN 'B' END,
        SL.AMOUNT,
        SL.LINENET
    FROM {stTbl} SL WITH(NOLOCK)
    INNER JOIN {invTbl} INV WITH(NOLOCK) ON INV.LOGICALREF = SL.INVOICEREF
    WHERE SL.CANCELLED = 0 AND INV.CANCELLED = 0
      AND SL.IOCODE = 4 AND SL.TRCODE = 8
      AND INV.TRCODE = 8 AND INV.CLIENTREF = @CRef
      AND ((INV.DATE_ BETWEEN @AB AND @AT) OR (INV.DATE_ BETWEEN @BB AND @BT))
)
SELECT
    MalzemeKodu = I.CODE,
    MalzemeAdi  = I.NAME,
    Birim       = ISNULL((SELECT TOP 1 U.CODE FROM {unTbl} U WITH(NOLOCK)
                          WHERE U.UNITSETREF = I.UNITSETREF AND U.MAINUNIT = 1), ''),
    A_Miktar   = SUM(CASE WHEN SL.Donem = 'A' THEN SL.AMOUNT  ELSE 0 END),
    A_NetTutar = SUM(CASE WHEN SL.Donem = 'A' THEN SL.LINENET ELSE 0 END),
    B_Miktar   = SUM(CASE WHEN SL.Donem = 'B' THEN SL.AMOUNT  ELSE 0 END),
    B_NetTutar = SUM(CASE WHEN SL.Donem = 'B' THEN SL.LINENET ELSE 0 END)
FROM SL
INNER JOIN {itTbl} I WITH(NOLOCK) ON I.LOGICALREF = SL.STOCKREF
GROUP BY I.CODE, I.NAME, I.UNITSETREF
ORDER BY SUM(SL.LINENET) DESC;

-- 2) DÖNEM ÖZETLERİ (fiş bazında)
SELECT
    Donem    = CASE WHEN INV.DATE_ BETWEEN @AB AND @AT THEN 'A' ELSE 'B' END,
    FisSayisi = COUNT(*),
    NetTutar  = SUM(INV.NETTOTAL),
    KdvTutar  = SUM(INV.TOTALVAT),
    BrutTutar = SUM(INV.GROSSTOTAL)
FROM {invTbl} INV WITH(NOLOCK)
WHERE INV.CANCELLED = 0 AND INV.TRCODE = 8 AND INV.CLIENTREF = @CRef
  AND ((INV.DATE_ BETWEEN @AB AND @AT) OR (INV.DATE_ BETWEEN @BB AND @BT))
GROUP BY CASE WHEN INV.DATE_ BETWEEN @AB AND @AT THEN 'A' ELSE 'B' END;";

        using var conn = _db.CreateConnection();
        using var grid = await conn.QueryMultipleAsync(new CommandDefinition(sql, new
        {
            aBas = donemABas.Date,
            aBit = donemABit.Date.AddDays(1).AddSeconds(-1),
            bBas = donemBBas.Date,
            bBit = donemBBit.Date.AddDays(1).AddSeconds(-1),
            @ref = cari.LogicalRef
        }, commandTimeout: 180, cancellationToken: ct));

        var malzemeler = (await grid.ReadAsync<dynamic>()).ToList();
        var ozetler    = (await grid.ReadAsync<dynamic>()).ToList();

        foreach (var r in malzemeler)
        {
            decimal aM = Convert.ToDecimal(r.A_Miktar   ?? 0);
            decimal aT = Convert.ToDecimal(r.A_NetTutar ?? 0);
            decimal bM = Convert.ToDecimal(r.B_Miktar   ?? 0);
            decimal bT = Convert.ToDecimal(r.B_NetTutar ?? 0);

            // İki dönemde de sıfırsa malzemeyi atla
            if (aM == 0 && aT == 0 && bM == 0 && bT == 0) continue;

            sonuc.Satirlar.Add(new SatisKiyaslamaSatiri
            {
                MalzemeKodu = ((string?)r.MalzemeKodu ?? "").Trim(),
                MalzemeAdi  = ((string?)r.MalzemeAdi  ?? "").Trim(),
                Birim       = ((string?)r.Birim       ?? "").Trim(),
                A_Miktar    = aM,
                A_NetTutar  = aT,
                B_Miktar    = bM,
                B_NetTutar  = bT
            });

            sonuc.DonemA.ToplamMiktar += aM;
            sonuc.DonemB.ToplamMiktar += bM;
        }

        foreach (var o in ozetler)
        {
            string d = (string)o.Donem;
            var hedef = d == "A" ? sonuc.DonemA : sonuc.DonemB;
            hedef.FisSayisi = Convert.ToInt32(o.FisSayisi ?? 0);
            hedef.NetTutar  = Convert.ToDecimal(o.NetTutar  ?? 0);
            hedef.KdvTutar  = Convert.ToDecimal(o.KdvTutar  ?? 0);
            hedef.BrutTutar = Convert.ToDecimal(o.BrutTutar ?? 0);
        }

        return sonuc;
    }
}
