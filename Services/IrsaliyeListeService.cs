using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

// Satınalma (TRCODE=1) ve Toptan Satış (TRCODE=8) irsaliyelerini listeler.
// Ambar filtresi: STFICHE.SOURCEINDEX (satış için bizim ambar) veya DESTINDEX
// (satınalma için bizim ambar) bizim ambar listesinde olmalı.
public sealed class IrsaliyeListeService
{
    private readonly DatabaseService _db;

    public IrsaliyeListeService(DatabaseService db) => _db = db;

    public async Task<List<IrsaliyeOzet>> ListeleAsync(
        int trcode,
        DateTime baslangic,
        DateTime bitis,
        IReadOnlyCollection<int>? ambarFiltresi,
        bool sadeceFaturalanmamis,
        CancellationToken ct = default)
    {
        bitis = SistemTarihi.Clamp(bitis);

        var stficheTbl = _db.GetPeriodTableName("STFICHE");
        var clcardTbl  = _db.GetTableName("CLCARD");

        var filtreSql = (ambarFiltresi != null && ambarFiltresi.Count > 0)
            ? " AND SF.SOURCEINDEX IN @Ambarlar"
            : "";
        var faturaSql = sadeceFaturalanmamis ? " AND SF.BILLED = 0" : "";

        var sql = $@"
SELECT
    LogicalRef  = SF.LOGICALREF,
    FisNo       = ISNULL(SF.FICHENO, ''),
    Tarih       = SF.DATE_,
    Trcode      = SF.TRCODE,
    Faturalandi = CAST(CASE WHEN SF.BILLED = 1 THEN 1 ELSE 0 END AS bit),
    CariKodu    = ISNULL(CL.CODE, ''),
    CariUnvani  = ISNULL(CL.DEFINITION_, ''),
    -- Logo'da hem satınalma (TRCODE=1) hem satış (TRCODE=8) için bizim ambar = SOURCEINDEX.
    -- DESTINDEX genelde 0; sadece transfer fişlerinde (TRCODE=25) kullanılır.
    AmbarNo     = SF.SOURCEINDEX,
    AmbarAdi    = ISNULL(CW.NAME, ''),
    NetTutar    = ISNULL(SF.NETTOTAL, 0),
    BrutTutar   = ISNULL(SF.GROSSTOTAL, 0),
    Aciklama    = NULLIF(LTRIM(RTRIM(ISNULL(SF.GENEXP1, '') + CASE WHEN SF.GENEXP2 IS NOT NULL AND SF.GENEXP2 <> '' THEN ' ' + SF.GENEXP2 ELSE '' END)), '')
FROM {stficheTbl} SF WITH(NOLOCK)
LEFT JOIN {clcardTbl} CL WITH(NOLOCK) ON SF.CLIENTREF = CL.LOGICALREF
LEFT JOIN L_CAPIWHOUSE CW WITH(NOLOCK)
       ON CW.FIRMNR = @Firma
      AND CW.NR     = SF.SOURCEINDEX
WHERE SF.CANCELLED = 0
  AND SF.TRCODE = @Trcode
  AND SF.DATE_ >= @Bas
  AND SF.DATE_ <= @Bit
{filtreSql}
{faturaSql}
ORDER BY SF.DATE_ DESC, SF.FICHENO DESC";

        using var conn = _db.CreateConnection();
        var cmd = new CommandDefinition(sql, new
        {
            Firma     = _db.FirmaNo,
            Trcode    = trcode,
            Bas       = baslangic.Date,
            Bit       = bitis.Date,
            Ambarlar  = ambarFiltresi,
        }, commandTimeout: 120, cancellationToken: ct);

        return (await conn.QueryAsync<IrsaliyeOzet>(cmd)).AsList();
    }

    public async Task<IrsaliyeDetay?> DetayAsync(int logicalRef, CancellationToken ct = default)
    {
        var stficheTbl = _db.GetPeriodTableName("STFICHE");
        var stlineTbl  = _db.GetPeriodTableName("STLINE");
        var clcardTbl  = _db.GetTableName("CLCARD");
        var itemsTbl   = _db.GetTableName("ITEMS");

        // Kayıt bilgisi: CAPIBLOCK_CREADEDDATE/MODIFIEDDATE alanları Logo'da
        // ZATEN tam datetime (saat dahil) tutuyor. HOUR/MIN/SEC ayrı alanlar aynı
        // bilgiyi yedek tutar; bunları üstüne EKLEMEK çift saat hatası yapar.
        var basSql = $@"
SELECT
    LogicalRef     = SF.LOGICALREF,
    FisNo          = ISNULL(SF.FICHENO, ''),
    Tarih          = SF.DATE_,
    Trcode         = SF.TRCODE,
    FisTuru        = CASE SF.TRCODE
                        WHEN 1 THEN N'Satınalma İrsaliyesi'
                        WHEN 8 THEN N'Toptan Satış İrsaliyesi'
                        WHEN 3 THEN N'Toptan Satış İade İrsaliyesi'
                        WHEN 13 THEN N'Üretimden Giriş Fişi'
                        ELSE CAST(SF.TRCODE AS nvarchar(10))
                     END,
    Faturalandi    = CAST(CASE WHEN SF.BILLED = 1 THEN 1 ELSE 0 END AS bit),
    CariKodu       = ISNULL(CL.CODE, ''),
    CariUnvani     = ISNULL(CL.DEFINITION_, ''),
    AmbarNo        = SF.SOURCEINDEX,
    AmbarAdi       = ISNULL(CW.NAME, ''),
    NetTutar       = ISNULL(SF.NETTOTAL, 0),
    BrutTutar      = ISNULL(SF.GROSSTOTAL, 0),
    ToplamIskonto  = ISNULL(SF.TOTALDISCOUNTS, 0),
    ToplamKdv      = ISNULL(SF.TOTALVAT, 0),
    Aciklama1      = SF.GENEXP1,
    Aciklama2      = SF.GENEXP2,
    Aciklama3      = SF.GENEXP3,
    Aciklama4      = SF.GENEXP4,
    Aciklama5      = SF.GENEXP5,
    Aciklama6      = SF.GENEXP6,
    MuafiyetKodu     = NULLIF(LTRIM(RTRIM(ISNULL(SF.VATEXCEPTCODE, ''))), ''),
    MuafiyetAciklama = NULLIF(LTRIM(RTRIM(ISNULL(SF.VATEXCEPTREASON, ''))), ''),
    KullaniciOlusturan  = ISNULL(CU.NAME, ''),
    KullaniciDegistiren = ISNULL(MU.NAME, ''),
    OlusturmaZamani  = SF.CAPIBLOCK_CREADEDDATE,
    DegistirmeZamani = SF.CAPIBLOCK_MODIFIEDDATE
FROM {stficheTbl} SF WITH(NOLOCK)
LEFT JOIN {clcardTbl}    CL WITH(NOLOCK) ON SF.CLIENTREF = CL.LOGICALREF
LEFT JOIN L_CAPIWHOUSE   CW WITH(NOLOCK)
       ON CW.FIRMNR = @Firma
      AND CW.NR     = SF.SOURCEINDEX
LEFT JOIN L_CAPIUSER     CU WITH(NOLOCK) ON CU.NR = SF.CAPIBLOCK_CREATEDBY
LEFT JOIN L_CAPIUSER     MU WITH(NOLOCK) ON MU.NR = SF.CAPIBLOCK_MODIFIEDBY
WHERE SF.LOGICALREF = @Ref";

        // Satır bazlı ambar: STLINE.SOURCEINDEX bizim ambardır (hem satınalma hem satış için).
        var satirSql = $@"
SELECT
    Sira            = ROW_NUMBER() OVER (ORDER BY ST.LOGICALREF),
    MalzemeKodu     = ISNULL(IT.CODE, ''),
    MalzemeAdi      = ISNULL(IT.NAME, ''),
    Miktar          = ISNULL(ST.AMOUNT, 0),
    BirimAdi        = '',
    Fiyat           = ISNULL(ST.PRICE, 0),
    Tutar           = ISNULL(ST.LINENET, 0),
    KdvMatrah       = ISNULL(ST.VATMATRAH, 0),
    AmbarNo         = ISNULL(ST.SOURCEINDEX, 0),
    MuafiyetKodu     = NULLIF(LTRIM(RTRIM(ISNULL(ST.VATEXCEPTCODE, ''))), ''),
    MuafiyetAciklama = NULLIF(LTRIM(RTRIM(ISNULL(ST.VATEXCEPTREASON, ''))), ''),
    SatirAciklama   = NULLIF(ST.LINEEXP, '')
FROM {stlineTbl} ST WITH(NOLOCK)
LEFT JOIN {itemsTbl} IT WITH(NOLOCK) ON ST.STOCKREF = IT.LOGICALREF
WHERE ST.STFICHEREF = @Ref
  AND ST.CANCELLED = 0
  AND ST.LPRODSTAT = 0
ORDER BY ST.LOGICALREF";

        using var conn = _db.CreateConnection();

        var bas = await conn.QueryFirstOrDefaultAsync<IrsaliyeDetay>(
            new CommandDefinition(basSql, new { Firma = _db.FirmaNo, Ref = logicalRef },
                cancellationToken: ct));
        if (bas == null) return null;

        bas.Satirlar = (await conn.QueryAsync<IrsaliyeDetaySatir>(
            new CommandDefinition(satirSql, new { Ref = logicalRef },
                cancellationToken: ct))).AsList();

        return bas;
    }
}
