using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Tek bir cariye ait iki dönem ALIM (TRCODE=1 mal alım + TRCODE=14 alınan hizmet) kıyaslaması.
/// İki fatura tipi tek raporda birleştirilir, özet bilgisinde ayrı kalem olarak listelenir.
/// Hizmet faturaları da STLINE'a düştüğü için tek STLINE+INVOICE join yeterli — SRVCARD'a
/// dokunmaya gerek yok (Doğuş Afyon Şeker konfigürasyonu 17 Haz 2026 itibarıyla).
/// </summary>
public class AlisKiyaslamaService
{
    private readonly DatabaseService _db;
    private readonly SatisKiyaslamaService _satisService;

    /// <summary>İncelenen alış TRCODE'ları:
    ///   1  = Mal Alım Faturası
    ///   4  = Sabit Kıymet / Duran Varlık Alım Faturası (Doğuş'ta araç/makine alımları buradan)
    ///   14 = Alınan Hizmet Faturası
    /// </summary>
    private static readonly int[] AlisTrCodes = { 1, 4, 14 };

    public AlisKiyaslamaService(DatabaseService db, SatisKiyaslamaService satisService)
    {
        _db = db;
        _satisService = satisService;
    }

    /// <summary>Cari arama — SatisKiyaslamaService ile aynı backend.</summary>
    public Task<List<CariBilgi>> CariAraAsync(string arama, CancellationToken ct = default)
        => _satisService.CariAraAsync(arama, ct);

    public async Task<AlisKiyaslamaSonuc> KiyaslaAsync(
        string cariKod,
        DateTime donemABas, DateTime donemABit,
        DateTime donemBBas, DateTime donemBBit,
        CancellationToken ct = default)
    {
        var sonuc = new AlisKiyaslamaSonuc
        {
            DonemA = new AlisDonemOzeti { Baslangic = donemABas.Date, Bitis = donemABit.Date },
            DonemB = new AlisDonemOzeti { Baslangic = donemBBas.Date, Bitis = donemBBit.Date }
        };

        var cari = await _satisService.CariBulAsync(cariKod, ct);
        if (cari == null) return sonuc;
        sonuc.Cari = cari;

        var itTbl  = _db.GetTableName("ITEMS");
        var invTbl = _db.GetPeriodTableName("INVOICE");
        var stTbl  = _db.GetPeriodTableName("STLINE");
        var unTbl  = _db.GetTableName("UNITSETL");

        // Tek round-trip — 2 result set: (1) malzeme bazlı kıyas, (2) dönem özetleri TRCODE bazlı
        var sql = $@"
DECLARE @AB datetime = @aBas, @AT datetime = @aBit;
DECLARE @BB datetime = @bBas, @BT datetime = @bBit;
DECLARE @CRef int = @ref;

-- 1) MALZEME BAZINDA KIYASLAMA — STLINE × INVOICE birleşik
-- IOCODE filtresi yok: TRCODE=4 (sabit kıymet) IOCODE=0 ile düşer, TRCODE=1/14 IOCODE=1 ile.
WITH SL AS (
    SELECT
        SL.STOCKREF,
        FaturaTipi = CASE WHEN INV.TRCODE = 1  THEN 'Mal Alım'
                          WHEN INV.TRCODE = 4  THEN 'Sabit Kıymet'
                          WHEN INV.TRCODE = 14 THEN 'Hizmet'
                          ELSE 'Diğer' END,
        Donem = CASE WHEN INV.DATE_ BETWEEN @AB AND @AT THEN 'A'
                     WHEN INV.DATE_ BETWEEN @BB AND @BT THEN 'B' END,
        SL.AMOUNT,
        SL.LINENET
    FROM {stTbl} SL WITH(NOLOCK)
    INNER JOIN {invTbl} INV WITH(NOLOCK) ON INV.LOGICALREF = SL.INVOICEREF
    WHERE SL.CANCELLED = 0 AND INV.CANCELLED = 0
      AND INV.TRCODE IN (1, 4, 14) AND INV.CLIENTREF = @CRef
      AND ((INV.DATE_ BETWEEN @AB AND @AT) OR (INV.DATE_ BETWEEN @BB AND @BT))
)
SELECT
    MalzemeKodu = I.CODE,
    MalzemeAdi  = I.NAME,
    Birim       = ISNULL((SELECT TOP 1 U.CODE FROM {unTbl} U WITH(NOLOCK)
                          WHERE U.UNITSETREF = I.UNITSETREF AND U.MAINUNIT = 1), ''),
    FaturaTipi  = SL.FaturaTipi,
    A_Miktar    = SUM(CASE WHEN SL.Donem = 'A' THEN SL.AMOUNT  ELSE 0 END),
    A_NetTutar  = SUM(CASE WHEN SL.Donem = 'A' THEN SL.LINENET ELSE 0 END),
    B_Miktar    = SUM(CASE WHEN SL.Donem = 'B' THEN SL.AMOUNT  ELSE 0 END),
    B_NetTutar  = SUM(CASE WHEN SL.Donem = 'B' THEN SL.LINENET ELSE 0 END)
FROM SL
INNER JOIN {itTbl} I WITH(NOLOCK) ON I.LOGICALREF = SL.STOCKREF
GROUP BY I.CODE, I.NAME, I.UNITSETREF, SL.FaturaTipi
ORDER BY SUM(SL.LINENET) DESC;

-- 2) DÖNEM ÖZETLERİ — TRCODE bazında ayrı (Mal Alım vs Hizmet ayırımı)
SELECT
    Donem    = CASE WHEN INV.DATE_ BETWEEN @AB AND @AT THEN 'A' ELSE 'B' END,
    TrCode   = INV.TRCODE,
    FisSayisi = COUNT(*),
    NetTutar  = SUM(INV.NETTOTAL),
    KdvTutar  = SUM(INV.TOTALVAT),
    BrutTutar = SUM(INV.GROSSTOTAL)
FROM {invTbl} INV WITH(NOLOCK)
WHERE INV.CANCELLED = 0 AND INV.TRCODE IN (1, 4, 14) AND INV.CLIENTREF = @CRef
  AND ((INV.DATE_ BETWEEN @AB AND @AT) OR (INV.DATE_ BETWEEN @BB AND @BT))
GROUP BY CASE WHEN INV.DATE_ BETWEEN @AB AND @AT THEN 'A' ELSE 'B' END, INV.TRCODE;";

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

            if (aM == 0 && aT == 0 && bM == 0 && bT == 0) continue;

            sonuc.Satirlar.Add(new AlisKiyaslamaSatiri
            {
                MalzemeKodu = ((string?)r.MalzemeKodu ?? "").Trim(),
                MalzemeAdi  = ((string?)r.MalzemeAdi  ?? "").Trim(),
                Birim       = ((string?)r.Birim       ?? "").Trim(),
                FaturaTipi  = ((string?)r.FaturaTipi  ?? "").Trim(),
                A_Miktar    = aM,
                A_NetTutar  = aT,
                B_Miktar    = bM,
                B_NetTutar  = bT
            });
        }

        foreach (var o in ozetler)
        {
            string d = (string)o.Donem;
            int trcode = Convert.ToInt32(o.TrCode);
            var hedef = d == "A" ? sonuc.DonemA : sonuc.DonemB;
            int fis  = Convert.ToInt32(o.FisSayisi ?? 0);
            decimal net = Convert.ToDecimal(o.NetTutar ?? 0);
            decimal kdv = Convert.ToDecimal(o.KdvTutar ?? 0);
            decimal brt = Convert.ToDecimal(o.BrutTutar ?? 0);

            if (trcode == 1)
            {
                hedef.SatinAlmaFis  = fis;
                hedef.SatinAlmaNet  = net;
                hedef.SatinAlmaKdv  = kdv;
                hedef.SatinAlmaBrut = brt;
            }
            else if (trcode == 4)
            {
                hedef.SabitKiymetFis  = fis;
                hedef.SabitKiymetNet  = net;
                hedef.SabitKiymetKdv  = kdv;
                hedef.SabitKiymetBrut = brt;
            }
            else if (trcode == 14)
            {
                hedef.HizmetFis  = fis;
                hedef.HizmetNet  = net;
                hedef.HizmetKdv  = kdv;
                hedef.HizmetBrut = brt;
            }
        }

        return sonuc;
    }
}
