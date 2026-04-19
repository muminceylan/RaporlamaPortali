using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

public class LogoIslemleriService
{
    private readonly DatabaseService _db;

    public LogoIslemleriService(DatabaseService db) => _db = db;

    // Logo Tiger KSLINES.TRCODE → insan okunabilir fiş türü
    private static readonly Dictionary<int, string> FisTuruMap = new()
    {
        { 11, "Nakit Tahsilat" },
        { 12, "Nakit Ödeme" },
        { 13, "Özel İşlem Nakit Tahsilat" },
        { 14, "Özel İşlem Nakit Ödeme" },
        { 21, "CH Tahsilat" },
        { 22, "CH Ödeme" },
        { 23, "Özel İşlem CH Tahsilat" },
        { 24, "Özel İşlem CH Ödeme" },
        { 31, "Muhasebe Fişi" },
        { 32, "Kur Farkı Fişi" },
        { 34, "Alınan Hizmet Faturası" },
        { 38, "Banka Kasa İşlem Fişi" },
        { 41, "Döviz Alış" },
        { 42, "Döviz Satış" },
        { 51, "Verilen Hizmet Faturası" },
        { 61, "Banka Gelen Havale" },
        { 71, "Müşteri Çeki İşlem" },
        { 72, "Müşteri Senedi İşlem" },
        { 73, "Verilen Çek" },
        { 74, "Borç Senedi" },
        { 75, "Çek/Senet İade" },
        { 77, "Kredi Kartı İşlem" },
    };

    public static string FisTuruAdi(int trcode) =>
        FisTuruMap.TryGetValue(trcode, out var v) ? v : $"Fiş Türü {trcode}";

    public async Task<List<KasaHareketi>> KasaHareketleriAsync(
        DateTime? baslangic, DateTime? bitis,
        bool iptalDahilEt = false,
        CancellationToken ct = default)
    {
        var ksTbl = _db.GetPeriodTableName("KSLINES");  // LG_211_01_KSLINES
        var clTbl = _db.GetTableName("CLCARD");         // LG_211_CLCARD
        var kcTbl = _db.GetTableName("KSCARD");         // LG_211_KSCARD

        // Afyon Fabrika kasası (100.01.028) — sadece bu kasaya ait fişler
        const string KasaKodu = "100.01.028";

        var sql = $@"
SELECT
    KS.LOGICALREF                              AS LogicalRef,
    KS.DATE_                                   AS Tarih,
    ISNULL(KS.FICHENO,'')                      AS IslemNo,
    ISNULL(NULLIF(KS.DOCODE,''), ISNULL(KS.TRANNO,'')) AS BelgeNo,
    ISNULL(KS.CUSTTITLE,'')                    AS CariUnvani,
    ISNULL(KS.CUSTTITLE3,'')                   AS CariKodu,
    ISNULL(KS.LINEEXP,'')                      AS SatirAciklamasi,
    ISNULL(KS.SPECODE,'')                      AS OzelKodu,
    ISNULL(CAST(KS.TRANGRPNO AS varchar(10)),'') AS TicariIslemGrubu,
    KS.TRCODE                                  AS TrCode,
    CAST(CASE WHEN KS.CANCELLED=1 THEN 1 ELSE 0 END AS BIT) AS Iptal,
    CAST(CASE WHEN KS.ACCOUNTED=1 THEN 1 ELSE 0 END AS BIT) AS Muhasebelesti,
    KS.AMOUNT                                  AS Tutar,
    KS.BRANCH                                  AS IsYeri,
    KS.DEPARTMENT                              AS Bolum,
    KS.TRCURR                                  AS TrCurr,
    KS.TRRATE                                  AS Kur,
    KS.TRNET                                   AS IslemDoviziTutari,
    KS.REPORTNET                               AS RaporlamaDoviziTutari,
    KS.REPORTRATE                              AS RaporlamaDoviziKuru,
    ISNULL(CL.CURCODE,'')                      AS DovizKodu
FROM {ksTbl} KS WITH(NOLOCK)
INNER JOIN {kcTbl} KC WITH(NOLOCK) ON KC.LOGICALREF = KS.CARDREF AND KC.CODE = @kasaKodu
LEFT JOIN L_CURRENCYLIST CL WITH(NOLOCK) ON CL.FIRMNR = @firma AND CL.CURTYPE = KS.TRCURR
WHERE (@bas IS NULL OR KS.DATE_ >= @bas)
  AND (@bit IS NULL OR KS.DATE_ <  DATEADD(day,1,@bit))
  AND (@iptalDahil = 1 OR KS.CANCELLED = 0)
ORDER BY KS.DATE_ DESC, KS.LOGICALREF DESC";

        using var conn = _db.CreateConnection();
        var rows = (await conn.QueryAsync<dynamic>(sql, new
        {
            firma = _db.FirmaNo,
            bas = baslangic,
            bit = bitis,
            iptalDahil = iptalDahilEt ? 1 : 0,
            kasaKodu = KasaKodu
        })).ToList();

        var list = new List<KasaHareketi>(rows.Count);
        foreach (var r in rows)
        {
            int trcurr = Convert.ToInt32(r.TrCurr ?? 0);
            string dovizKodu = r.DovizKodu ?? "";
            string dovizTuru = trcurr == 0 ? "TL" : (string.IsNullOrEmpty(dovizKodu) ? $"CUR{trcurr}" : dovizKodu);
            int trcode = Convert.ToInt32(r.TrCode);
            list.Add(new KasaHareketi
            {
                LogicalRef = Convert.ToInt32(r.LogicalRef),
                Tarih = (DateTime)r.Tarih,
                IslemNo = (string)r.IslemNo,
                BelgeNo = (string)r.BelgeNo,
                CariUnvani = (string)r.CariUnvani,
                CariKodu = (string)r.CariKodu,
                SatirAciklamasi = (string)r.SatirAciklamasi,
                OzelKodu = (string)r.OzelKodu,
                TicariIslemGrubu = (string)r.TicariIslemGrubu,
                TrCode = trcode,
                FisTuru = FisTuruAdi(trcode),
                Iptal = (bool)r.Iptal,
                Muhasebelesti = (bool)r.Muhasebelesti,
                Tutar = Convert.ToDecimal(r.Tutar ?? 0),
                IsYeri = Convert.ToInt32(r.IsYeri ?? 0),
                Bolum = Convert.ToInt32(r.Bolum ?? 0),
                TrCurr = trcurr,
                DovizTuru = dovizTuru,
                Kur = Convert.ToDecimal(r.Kur ?? 0),
                IslemDoviziTutari = Convert.ToDecimal(r.IslemDoviziTutari ?? 0),
                RaporlamaDoviziTutari = Convert.ToDecimal(r.RaporlamaDoviziTutari ?? 0),
                RaporlamaDoviziKuru = Convert.ToDecimal(r.RaporlamaDoviziKuru ?? 0),
            });
        }
        return list;
    }

    public async Task<List<KurBilgisi>> KurlarAsync(
        DateTime? baslangic, DateTime? bitis,
        int? crType = null,
        CancellationToken ct = default)
    {
        var exTbl = $"LG_EXCHANGE_{_db.FirmaNo}";
        var sql = $@"
SELECT
    DX.EDATE                                       AS Tarih,
    DX.CRTYPE                                      AS CrType,
    CASE WHEN DX.CRTYPE = 1 THEN 'USD'
         WHEN DX.CRTYPE = 20 THEN 'EUR'
         ELSE CAST(DX.CRTYPE AS varchar(10)) END   AS DovizKodu,
    CASE WHEN DX.CRTYPE = 1 THEN 'Amerikan Doları'
         WHEN DX.CRTYPE = 20 THEN 'Euro'
         ELSE '' END                               AS DovizAdi,
    DX.RATES1                                      AS Rate1,
    DX.RATES2                                      AS Rate2,
    DX.RATES3                                      AS Rate3,
    DX.RATES4                                      AS Rate4
FROM {exTbl} DX WITH(NOLOCK)
WHERE DX.CRTYPE IN (1, 20)
  AND (@bas IS NULL OR DX.EDATE >= @bas)
  AND (@bit IS NULL OR DX.EDATE <  DATEADD(day,1,@bit))
  AND (@cr  IS NULL OR DX.CRTYPE = @cr)
ORDER BY DX.EDATE DESC, DX.CRTYPE";
        using var conn = _db.CreateConnection();
        var rows = await conn.QueryAsync<KurBilgisi>(sql, new
        {
            firma = _db.FirmaNo,
            bas = baslangic,
            bit = bitis,
            cr = crType
        });
        return rows.ToList();
    }

    public async Task<List<StokSatiri>> StokDurumuAsync(
        int? ambarNo = null,
        bool sifirlariGizle = true,
        CancellationToken ct = default)
    {
        var gntView = _db.GetViewName("GNTOTST");
        var itTbl = _db.GetTableName("ITEMS");

        var sql = $@"
SELECT
    GNT.INVENNO            AS AmbarNo,
    ISNULL(WHO.NAME,'')    AS Ambar,
    ISNULL(ITM.CODE,'')    AS MalzemeKodu,
    ISNULL(ITM.NAME,'')    AS MalzemeAdi,
    ROUND(GNT.ONHAND,2)    AS Stok
FROM {gntView} GNT WITH(NOLOCK)
LEFT JOIN {itTbl} ITM WITH(NOLOCK) ON ITM.LOGICALREF = GNT.STOCKREF
LEFT JOIN L_CAPIWHOUSE WHO WITH(NOLOCK) ON WHO.FIRMNR = @firma AND WHO.NR = GNT.INVENNO
WHERE (@ambar IS NULL OR GNT.INVENNO = @ambar)
  AND (@sifirGizle = 0 OR ROUND(GNT.ONHAND,2) <> 0)
ORDER BY GNT.INVENNO, ITM.CODE";

        using var conn = _db.CreateConnection();
        var rows = (await conn.QueryAsync<dynamic>(new CommandDefinition(sql, new
        {
            firma = _db.FirmaNo,
            ambar = ambarNo,
            sifirGizle = sifirlariGizle ? 1 : 0
        }, commandTimeout: 120, cancellationToken: ct))).ToList();

        var list = new List<StokSatiri>(rows.Count);
        foreach (var r in rows)
        {
            list.Add(new StokSatiri
            {
                AmbarNo = Convert.ToInt32(r.AmbarNo),
                Ambar = (string)r.Ambar,
                MalzemeKodu = (string)r.MalzemeKodu,
                MalzemeAdi = (string)r.MalzemeAdi,
                Stok = Convert.ToDecimal(r.Stok ?? 0),
            });
        }
        return list;
    }

    public async Task<List<AmbarSecenek>> AmbarSecenekleriAsync()
    {
        // Ambar listesi: firmanın tanımlı tüm ambarları (STLINE'da hareket olsun veya olmasın).
        var sql = @"
SELECT
    CW.NR              AS AmbarNo,
    ISNULL(CW.NAME,'') AS Ad
FROM L_CAPIWHOUSE CW WITH(NOLOCK)
WHERE CW.FIRMNR = @firma AND CW.NR >= 0
ORDER BY CW.NR";
        using var conn = _db.CreateConnection();
        return (await conn.QueryAsync<AmbarSecenek>(sql, new { firma = _db.FirmaNo })).ToList();
    }

    /// <summary>Döviz filtresi: Excel USD (1) ve EUR (20) sabitleri.</summary>
    public async Task<List<DovizSecenek>> DovizSecenekleriAsync()
    {
        var exTbl = $"LG_EXCHANGE_{_db.FirmaNo}";
        var sql = $@"
SELECT DISTINCT
    DX.CRTYPE AS CrType,
    CASE WHEN DX.CRTYPE = 1 THEN 'USD' WHEN DX.CRTYPE = 20 THEN 'EUR'
         ELSE CAST(DX.CRTYPE AS varchar(10)) END AS Kod,
    CASE WHEN DX.CRTYPE = 1 THEN 'Amerikan Doları' WHEN DX.CRTYPE = 20 THEN 'Euro'
         ELSE '' END AS Ad
FROM {exTbl} DX WITH(NOLOCK)
WHERE DX.CRTYPE IN (1, 20)
ORDER BY Kod";
        using var conn = _db.CreateConnection();
        return (await conn.QueryAsync<DovizSecenek>(sql)).ToList();
    }
}
