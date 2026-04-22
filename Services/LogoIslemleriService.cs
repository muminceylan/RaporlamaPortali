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

    // Logo Tiger CLFLINE.TRCODE → insan okunabilir fiş türü (Cari Hesap Fiş Satırı)
    // Tiger 3 standart eşleme — DOGUSNDB (Firma 211) gerçek kayıtlarıyla doğrulandı.
    private static readonly Dictionary<int, string> ClfFisTuruMap = new()
    {
        { 1,  "Nakit Tahsilat" },
        { 2,  "Nakit Ödeme" },
        { 3,  "Borç Dekontu" },
        { 4,  "Alacak Dekontu" },
        { 5,  "Virman Fişi" },
        { 6,  "Kur Farkı İşlem Fişi" },
        { 12, "Özel İşlem Fişi (Borç)" },
        { 13, "Özel İşlem Fişi (Alacak)" },
        { 14, "Açılış Fişi" },
        { 20, "Kredi Kartı İşlem Fişi" },
        { 21, "Gönderilen Havaleler" },
        { 22, "Gelen Havaleler" },
        { 24, "Krediden Ödeme" },
        { 31, "Satınalma Faturası" },
        { 32, "Perakende Satış İade Faturası" },
        { 33, "Toptan Satış İade Faturası" },
        { 34, "Alınan Hizmet Faturası" },
        { 35, "Alınan Proforma Fatura" },
        { 36, "Mahsup Fişi" },
        { 37, "Perakende Satış Faturası" },
        { 38, "Toptan Satış Faturası" },
        { 39, "Verilen Hizmet Faturası" },
        { 40, "Verilen Proforma Fatura" },
        { 41, "Verilen Vade Farkı Faturası" },
        { 42, "Alınan Vade Farkı Faturası" },
        { 43, "Alınan Fiyat Farkı Faturası" },
        { 44, "Verilen Fiyat Farkı Faturası" },
        { 45, "Müstahsil Makbuzu" },
        { 50, "Müstahsil Makbuzu" },
        { 56, "Kur Farkı Fişi" },
        { 61, "Banka Alınan Fatura" },
        { 62, "Banka Verilen Fatura" },
        { 63, "Çek Çıkış (Cari Hesaba)" },
        { 64, "Senet Çıkış (Cari Hesaba)" },
        { 65, "Çek Giriş (Müşteriden)" },
        { 66, "Senet Giriş (Müşteriden)" },
        { 70, "Kredi Kartı Fişi" },
        { 71, "Kredi Kartı İade Fişi" },
        { 72, "Yansıtma Fişi" },
    };

    public static string ClfFisTuruAdi(int trcode) =>
        ClfFisTuruMap.TryGetValue(trcode, out var v) ? v : $"Fiş Türü {trcode}";

    public async Task<List<KasaHareketi>> KasaHareketleriAsync(
        DateTime? baslangic, DateTime? bitis,
        bool iptalDahilEt = false,
        CancellationToken ct = default)
    {
        bitis = SistemTarihi.Clamp(bitis) ?? SistemTarihi.SonDahilGun;
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
        bitis = SistemTarihi.Clamp(bitis) ?? SistemTarihi.SonDahilGun;
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
        }, commandTimeout: 180, cancellationToken: ct))).ToList();

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

    // Cari Hesap Bakiye Sorgulama — Excel "101 211 CARI HESAP BAKIYE.xlsm"
    // VBA'sındaki sorgunun Logo'dan çekim mantığıyla aynı:
    //   CLFLINE (fiş satırları) × CLCARD (cari kartı) birleşimi
    //   SIGN=0 → BORÇ, SIGN=1 → ALACAK (Logo konvansiyonu)
    //   CANCELLED=0, STATUS=0, ACTIVE=0 (iptal olmayan, kapatılmamış, aktif cari)
    //   Tarih aralığı: CLFLINE.DATE_ BETWEEN @bas AND @bit
    //   BAKIYE = BORÇ - ALACAK (aralıktaki fiş satırları üzerinden)
    public async Task<List<CariBakiye>> CariBakiyeAsync(
        DateTime baslangic, DateTime bitis,
        bool sadeceBakiyeOlanlar = true,
        CancellationToken ct = default)
    {
        bitis = SistemTarihi.Clamp(bitis);
        var clfTbl = _db.GetPeriodTableName("CLFLINE");   // LG_211_01_CLFLINE
        var clcTbl = _db.GetTableName("CLCARD");          // LG_211_CLCARD

        var sql = $@"
SELECT
    CARI_HESAP_KODU,
    CARI_HESAP_UNVANI,
    TCKNO,
    VERGINO,
    BORC   = SUM(BORC),
    ALACAK = SUM(ALACAK),
    BAKIYE = SUM(BAKIYE),
    OZEL_KOD, OZEL_KOD2, OZEL_KOD3, OZEL_KOD4, OZEL_KOD5
FROM (
    SELECT
        CLCARD.TCKNO                   AS TCKNO,
        CLCARD.TAXNR                   AS VERGINO,
        CARI_HESAP_KODU   = CLCARD.CODE,
        CARI_HESAP_UNVANI = CLCARD.DEFINITION_,
        ROUND(SUM(CASE CLFLINE.SIGN WHEN 0 THEN CLFLINE.AMOUNT ELSE 0 END),2) AS BORC,
        ROUND(SUM(CASE CLFLINE.SIGN WHEN 1 THEN CLFLINE.AMOUNT ELSE 0 END),2) AS ALACAK,
        ROUND(SUM(CASE CLFLINE.SIGN WHEN 0 THEN CLFLINE.AMOUNT ELSE 0 END),2) -
        ROUND(SUM(CASE CLFLINE.SIGN WHEN 1 THEN CLFLINE.AMOUNT ELSE 0 END),2) AS BAKIYE,
        OZEL_KOD  = CLCARD.SPECODE,
        OZEL_KOD2 = CLCARD.SPECODE2,
        OZEL_KOD3 = CLCARD.SPECODE3,
        OZEL_KOD4 = CLCARD.SPECODE4,
        OZEL_KOD5 = CLCARD.SPECODE5
    FROM {clfTbl} CLFLINE WITH(NOLOCK)
    LEFT OUTER JOIN {clcTbl} CLCARD WITH(NOLOCK)
        ON CLFLINE.CLIENTREF = CLCARD.LOGICALREF
    WHERE CLFLINE.CLIENTREF <> 0
      AND CLFLINE.CANCELLED = 0
      AND CLFLINE.STATUS    = 0
      AND CLFLINE.DATE_ BETWEEN @bas AND @bit
      AND CLFLINE.CLIENTREF IN (SELECT LOGICALREF FROM {clcTbl})
      AND CLCARD.ACTIVE = 0
    GROUP BY CLCARD.TCKNO, CLCARD.TAXNR,
             CLCARD.CODE, CLCARD.DEFINITION_,
             CLCARD.SPECODE, CLCARD.SPECODE2, CLCARD.SPECODE3, CLCARD.SPECODE4, CLCARD.SPECODE5
) AS QRY
GROUP BY TCKNO, VERGINO,
         CARI_HESAP_KODU, CARI_HESAP_UNVANI,
         OZEL_KOD, OZEL_KOD2, OZEL_KOD3, OZEL_KOD4, OZEL_KOD5
{(sadeceBakiyeOlanlar ? "HAVING ROUND(SUM(BAKIYE),2) <> 0" : "")}
ORDER BY CARI_HESAP_KODU";

        using var conn = _db.CreateConnection();
        var rows = (await conn.QueryAsync<dynamic>(new CommandDefinition(sql, new
        {
            bas = baslangic.Date,
            bit = bitis.Date.AddDays(1).AddSeconds(-1)
        }, commandTimeout: 120, cancellationToken: ct))).ToList();

        var list = new List<CariBakiye>(rows.Count);
        foreach (var r in rows)
        {
            list.Add(new CariBakiye
            {
                CariHesapKodu = (string)(r.CARI_HESAP_KODU ?? ""),
                CariHesapUnvani = (string)(r.CARI_HESAP_UNVANI ?? ""),
                TcKimlikNo = (string)(r.TCKNO ?? ""),
                VergiNo = (string)(r.VERGINO ?? ""),
                Borc = Convert.ToDecimal(r.BORC ?? 0),
                Alacak = Convert.ToDecimal(r.ALACAK ?? 0),
                Bakiye = Convert.ToDecimal(r.BAKIYE ?? 0),
                OzelKod = (string)(r.OZEL_KOD ?? ""),
                OzelKod2 = (string)(r.OZEL_KOD2 ?? ""),
                OzelKod3 = (string)(r.OZEL_KOD3 ?? ""),
                OzelKod4 = (string)(r.OZEL_KOD4 ?? ""),
                OzelKod5 = (string)(r.OZEL_KOD5 ?? ""),
            });
        }
        return list;
    }

    /// <summary>
    /// Cari kart seçenekleri (Autocomplete için).
    /// Kaynak: LG_211_CLCARD (ACTIVE=0 → aktif cari kartlar)
    /// </summary>
    public async Task<List<CariSecenek>> CariSecenekleriAsync(CancellationToken ct = default)
    {
        var clcTbl = _db.GetTableName("CLCARD");
        var sql = $@"
SELECT
    Kod         = CODE,
    Unvan       = DEFINITION_,
    TcKimlikNo  = ISNULL(TCKNO,''),
    VergiNo     = ISNULL(TAXNR,'')
FROM {clcTbl} WITH(NOLOCK)
WHERE ACTIVE = 0
ORDER BY CODE";
        using var conn = _db.CreateConnection();
        return (await conn.QueryAsync<CariSecenek>(new CommandDefinition(sql, commandTimeout: 120, cancellationToken: ct))).ToList();
    }

    /// <summary>
    /// Cari hesap hareket (ekstre) listesi.
    /// Kaynak: LG_211_01_CLFLINE (fiş satırı) + LG_211_CLCARD (cari kart)
    ///   (Not: INF_Cari_Hesap_Ekstresi_211 view'ı bazı carileri — özellikle ihracat/yurtdışı —
    ///    kapsam dışında bırakıyor, bu yüzden base tabloları kullanıyoruz.)
    ///   FISNO için LG_211_01_INVOICE / CLFICHE LEFT JOIN (MODULENR bazlı).
    ///   ODEMEPLANI: LG_211_PAYPLANS
    /// SIGN=0 → BORÇ, SIGN=1 → ALACAK (Logo konvansiyonu). BAKIYE servis tarafında kümülatif.
    /// </summary>
    public async Task<(List<CariHareket> Liste, int HamSayi, int TarihAraligiSayi, string Ornek)>
        CariHareketListesiAsync(
            DateTime baslangic, DateTime bitis,
            string? cariKodu = null,
            string? cariUnvan = null,
            CancellationToken ct = default)
    {
        bitis = SistemTarihi.Clamp(bitis);
        var clfTbl = _db.GetPeriodTableName("CLFLINE");  // LG_211_01_CLFLINE
        var clcTbl = _db.GetTableName("CLCARD");         // LG_211_CLCARD
        var invTbl = _db.GetPeriodTableName("INVOICE");  // LG_211_01_INVOICE
        var cfiTbl = _db.GetPeriodTableName("CLFICHE");  // LG_211_01_CLFICHE
        var ppTbl  = $"LG_{_db.FirmaNo:D3}_PAYPLANS";    // LG_211_PAYPLANS
        var ptTbl  = _db.GetPeriodTableName("PAYTRANS"); // LG_211_01_PAYTRANS

        using var conn = _db.CreateConnection();

        var sql = $@"
SELECT
    TARIH      = F.DATE_,
    VADE       = PT.VADE,
    MODULENR   = F.MODULENR,
    ODEMEPLANI = ISNULL(PP.CODE,''),
    TRCODE     = F.TRCODE,
    FISNO      = COALESCE(
        CASE F.MODULENR WHEN 4 THEN INV.FICHENO END,
        CASE F.MODULENR WHEN 5 THEN CLFI.FICHENO END,
        CAST(F.TRANNO AS varchar(16))),
    ACIKLAMA   = ISNULL(F.LINEEXP,''),
    CARIKODU   = C.CODE,
    CARIUNVAN  = C.DEFINITION_,
    BORC       = CASE F.SIGN WHEN 0 THEN F.AMOUNT ELSE 0 END,
    ALACAK     = CASE F.SIGN WHEN 1 THEN F.AMOUNT ELSE 0 END,
    TRCURR     = ISNULL(F.TRCURR, 0),
    TRRATE     = ISNULL(F.TRRATE, 0),
    TRNET      = ISNULL(F.TRNET, 0),
    DOVIZKODU  = ISNULL(CL.CURCODE,'')
FROM {clfTbl} F WITH(NOLOCK)
LEFT JOIN {clcTbl} C WITH(NOLOCK) ON F.CLIENTREF = C.LOGICALREF
LEFT JOIN {ppTbl}  PP WITH(NOLOCK) ON F.PAYDEFREF = PP.LOGICALREF
LEFT JOIN {invTbl} INV  WITH(NOLOCK) ON F.MODULENR = 4 AND F.SOURCEFREF = INV.LOGICALREF
LEFT JOIN {cfiTbl} CLFI WITH(NOLOCK) ON F.MODULENR = 5 AND F.SOURCEFREF = CLFI.LOGICALREF
LEFT JOIN L_CURRENCYLIST CL WITH(NOLOCK) ON CL.FIRMNR = @firma AND CL.CURTYPE = F.TRCURR
LEFT JOIN (
    SELECT MODULENR, FICHEREF, CARDREF, MIN(DATE_) AS VADE
    FROM {ptTbl} WITH(NOLOCK)
    WHERE ISNULL(CANCELLED,0) = 0
    GROUP BY MODULENR, FICHEREF, CARDREF
) PT ON PT.MODULENR = F.MODULENR AND PT.FICHEREF = F.SOURCEFREF AND PT.CARDREF = F.CLIENTREF
WHERE F.CANCELLED = 0
  AND F.CLIENTREF <> 0
  AND F.DATE_ BETWEEN @bas AND @bit
  {(string.IsNullOrWhiteSpace(cariKodu) ? "" : "AND C.CODE = @cariKodu")}
ORDER BY C.CODE, F.DATE_, F.LOGICALREF";

        var rows = (await conn.QueryAsync<dynamic>(new CommandDefinition(sql, new
        {
            firma = _db.FirmaNo,
            bas = baslangic.Date,
            bit = bitis.Date.AddDays(1).AddSeconds(-1),
            cariKodu = cariKodu ?? ""
        }, commandTimeout: 300, cancellationToken: ct))).ToList();

        int hamSayi = rows.Count;
        int tarihAraligiSayi = rows.Count;
        string ornek = "CLFLINE direkt sorgusu";

        var list = new List<CariHareket>(rows.Count);
        string? onceki = null;
        decimal kumule = 0m;
        foreach (var r in rows)
        {
            var kod = ((string)(r.CARIKODU ?? "")).Trim();
            if (onceki != kod) { kumule = 0m; onceki = kod; }

            var borc = Convert.ToDecimal(r.BORC ?? 0);
            var alacak = Convert.ToDecimal(r.ALACAK ?? 0);
            kumule += borc - alacak;

            int trcode = r.TRCODE is null ? 0 : Convert.ToInt32(r.TRCODE);
            int trcurr = r.TRCURR is null ? 0 : Convert.ToInt32(r.TRCURR);
            decimal trrate = Convert.ToDecimal(r.TRRATE ?? 0);
            decimal trnet = Convert.ToDecimal(r.TRNET ?? 0);
            string dovizKodu = ((string)(r.DOVIZKODU ?? "")).Trim();
            if (trcurr != 0 && string.IsNullOrEmpty(dovizKodu))
                dovizKodu = trcurr switch { 1 => "USD", 20 => "EUR", _ => $"CUR{trcurr}" };

            decimal dBorc = 0m, dAlacak = 0m;
            if (trcurr != 0 && trnet != 0)
            {
                if (borc > 0) dBorc = trnet;
                else if (alacak > 0) dAlacak = trnet;
            }

            DateTime? vade = r.VADE == null ? null : (DateTime?)Convert.ToDateTime(r.VADE);
            int modulenr = r.MODULENR == null ? 0 : Convert.ToInt32(r.MODULENR);

            list.Add(new CariHareket
            {
                Tarih       = Convert.ToDateTime(r.TARIH),
                VadeTarihi  = vade,
                ModuleNr    = modulenr,
                OdemePlani  = (string)(r.ODEMEPLANI ?? ""),
                TrCode      = trcode,
                FisTuru     = ClfFisTuruAdi(trcode),
                FisNo       = ((string)(r.FISNO ?? "")).Trim(),
                Aciklama    = (string)(r.ACIKLAMA ?? ""),
                CariKodu    = kod,
                CariUnvan   = (string)(r.CARIUNVAN ?? ""),
                Borc        = borc,
                Alacak      = alacak,
                Bakiye      = kumule,
                TrCurr      = trcurr,
                DovizKodu   = trcurr == 0 ? "" : dovizKodu,
                DovizKur    = trrate,
                DovizBorc   = dBorc,
                DovizAlacak = dAlacak
            });
        }

        return (list, hamSayi, tarihAraligiSayi, ornek);
    }

    /// <summary>
    /// Kur farkı hesaplaması — seçili carinin dövizli hareketlerini alır, para birimi bazında
    /// net FC bakiyeyi güncel kur ile revalüe eder ve defter TL bakiye ile farkını gösterir.
    /// Pozitif kur farkı → BİZ karşı firmaya fatura keseriz (lehimize).
    /// Negatif kur farkı → KARŞI FİRMA bize fatura keser (aleyhimize).
    /// </summary>
    public async Task<(List<KurFarkiSatiri> Ozet, List<CariHareket> Detay)>
        KurFarkiHesaplaAsync(
            DateTime baslangic, DateTime bitis,
            string? cariKodu, string? cariUnvan,
            DateTime hesaplamaTarihi,
            int rateKolonu = 2,
            bool krediKartiSentetikle = true,
            CancellationToken ct = default)
    {
        bitis = SistemTarihi.Clamp(bitis);
        hesaplamaTarihi = SistemTarihi.Clamp(hesaplamaTarihi);
        var (hareketler, _, _, _) = await CariHareketListesiAsync(
            baslangic, bitis, cariKodu, cariUnvan, ct);

        var dovizliler = hareketler.Where(h => h.TrCurr != 0).ToList();
        if (dovizliler.Count == 0)
            return (new List<KurFarkiSatiri>(), new List<CariHareket>());

        var exTbl = $"LG_EXCHANGE_{_db.FirmaNo}";
        var rateCol = rateKolonu switch
        {
            1 => "RATES1",
            2 => "RATES2",
            3 => "RATES3",
            4 => "RATES4",
            _ => "RATES2"
        };

        using var conn = _db.CreateConnection();

        // Kredi Kartı İşlem Fişi (TRCODE=20) ve TL-only satırları sentetik dövize çevir.
        // Hedef döviz: dövizli hareketlerde baskın olan (en çok işlem sayılı) para birimi.
        var sentetikler = new List<CariHareket>();
        int? hedefCurr = null;
        string hedefDovizKodu = "";
        if (krediKartiSentetikle)
        {
            var currGroups = dovizliler
                .GroupBy(x => x.TrCurr)
                .Select(g => new { Curr = g.Key, Sayi = g.Count(), Kod = g.First().DovizKodu })
                .OrderByDescending(x => x.Sayi)
                .ToList();
            if (currGroups.Count > 0)
            {
                hedefCurr = currGroups[0].Curr;
                hedefDovizKodu = currGroups[0].Kod;
            }

            var krediKartiTl = hareketler
                .Where(h => h.TrCode == 20 && h.TrCurr == 0 && (h.Borc != 0 || h.Alacak != 0))
                .ToList();

            if (hedefCurr.HasValue && krediKartiTl.Count > 0)
            {
                var kurSql = $@"
SELECT TOP 1 {rateCol} AS Kur, EDATE AS Tarih
FROM {exTbl} WITH(NOLOCK)
WHERE CRTYPE = @cr AND EDATE <= @t AND {rateCol} > 0
ORDER BY EDATE DESC";

                foreach (var kk in krediKartiTl)
                {
                    var kurRow = await conn.QueryFirstOrDefaultAsync<dynamic>(
                        new CommandDefinition(kurSql,
                            new { cr = hedefCurr.Value, t = kk.Tarih.Date },
                            cancellationToken: ct));
                    if (kurRow == null) continue;

                    decimal kur = Convert.ToDecimal(kurRow.Kur ?? 0);
                    if (kur <= 0) continue;

                    decimal dBorc   = kk.Borc > 0   ? Math.Round(kk.Borc / kur, 2)   : 0m;
                    decimal dAlacak = kk.Alacak > 0 ? Math.Round(kk.Alacak / kur, 2) : 0m;

                    sentetikler.Add(new CariHareket
                    {
                        Tarih         = kk.Tarih,
                        OdemePlani    = kk.OdemePlani,
                        TrCode        = kk.TrCode,
                        FisTuru       = kk.FisTuru,
                        FisNo         = kk.FisNo,
                        Aciklama      = kk.Aciklama,
                        CariKodu      = kk.CariKodu,
                        CariUnvan     = kk.CariUnvan,
                        Borc          = kk.Borc,
                        Alacak        = kk.Alacak,
                        Bakiye        = kk.Bakiye,
                        TrCurr        = hedefCurr.Value,
                        DovizKodu     = hedefDovizKodu,
                        DovizKur      = kur,
                        DovizBorc     = dBorc,
                        DovizAlacak   = dAlacak,
                        SentetikDoviz = true
                    });
                }
            }
        }

        var tumDovizli = dovizliler.Concat(sentetikler).ToList();
        var sonuc = new List<KurFarkiSatiri>();

        foreach (var grp in tumDovizli.GroupBy(x => x.TrCurr).OrderBy(g => g.Key))
        {
            var fcBorc   = grp.Sum(x => x.DovizBorc);
            var fcAlacak = grp.Sum(x => x.DovizAlacak);
            var fcBak    = fcBorc - fcAlacak;

            var tlBorc   = grp.Sum(x => x.Borc);
            var tlAlacak = grp.Sum(x => x.Alacak);
            var tlBak    = tlBorc - tlAlacak;

            var sentList = grp.Where(x => x.SentetikDoviz).ToList();

            var kurSql2 = $@"
SELECT TOP 1 EDATE AS Tarih, {rateCol} AS Kur
FROM {exTbl} WITH(NOLOCK)
WHERE CRTYPE = @cr AND EDATE <= @t AND {rateCol} > 0
ORDER BY EDATE DESC";
            var kurRow2 = await conn.QueryFirstOrDefaultAsync<dynamic>(
                new CommandDefinition(kurSql2,
                    new { cr = grp.Key, t = hesaplamaTarihi.Date },
                    cancellationToken: ct));

            decimal guncelKur = 0m;
            DateTime? guncelKurTarihi = null;
            if (kurRow2 != null)
            {
                guncelKur = Convert.ToDecimal(kurRow2.Kur ?? 0);
                guncelKurTarihi = Convert.ToDateTime(kurRow2.Tarih);
            }

            decimal ortKur = fcBak == 0 ? 0 : Math.Abs(tlBak / fcBak);
            decimal guncelTl = Math.Round(fcBak * guncelKur, 2);
            decimal kurFarki = Math.Round(guncelTl - tlBak, 2);

            string kesen, yon;
            if (Math.Abs(kurFarki) < 0.01m || fcBak == 0 || guncelKur == 0)
            {
                kesen = "YOK";
                yon = "—";
            }
            else if (kurFarki > 0)
            {
                kesen = "BİZ → Karşı Firma";
                yon = "Lehimize (Gelir)";
            }
            else
            {
                kesen = "Karşı Firma → BİZ";
                yon = "Aleyhimize (Gider)";
            }

            string dKod = grp.First().DovizKodu;
            string aciklama = fcBak switch
            {
                0 => "Döviz bakiyesi sıfır — kur farkı oluşmaz",
                > 0 => $"{Math.Abs(fcBak):N2} {dKod} karşı firmadan alacak",
                < 0 => $"{Math.Abs(fcBak):N2} {dKod} karşı firmaya borç"
            };

            sonuc.Add(new KurFarkiSatiri
            {
                TrCurr           = grp.Key,
                DovizKodu        = dKod,
                IslemSayisi      = grp.Count(),
                FcBorc           = fcBorc,
                FcAlacak         = fcAlacak,
                FcBakiye         = fcBak,
                TlBorc           = tlBorc,
                TlAlacak         = tlAlacak,
                TlBakiye         = tlBak,
                OrtalamaKur      = ortKur,
                GuncelKur        = guncelKur,
                GuncelKurTarihi  = guncelKurTarihi,
                GuncelTlDeger    = guncelTl,
                KurFarki         = kurFarki,
                KesenTaraf       = kesen,
                Yon              = yon,
                Aciklama         = aciklama,
                SentetikSayisi   = sentList.Count,
                SentetikFcBorc   = sentList.Sum(x => x.DovizBorc),
                SentetikFcAlacak = sentList.Sum(x => x.DovizAlacak)
            });
        }

        var detay = tumDovizli
            .OrderBy(x => x.TrCurr)
            .ThenBy(x => x.Tarih)
            .ToList();

        return (sonuc, detay);
    }

    /// <summary>
    /// FIFO mahsuplaştırma ile kur farkı:
    ///   - Tahsilat tipleri: Banka/Havale (MODULENR=7), Kredi Kartı (TRCODE=20), Virman (TRCODE=5)
    ///   - Fatura vadesi PAYTRANS'tan; yoksa fatura tarihi
    ///   - Her tahsilat en eski açık faturaya (vade sınırı olmadan) yazılır
    ///   - TL tahsilat → Döviz fatura: tahsilat günü kuruyla FC'ye çevrilir, gerçekleşmiş kur farkı = FC × (tahsilat kuru − fatura kuru)
    ///   - FC tahsilat → Döviz fatura: gerçekleşmiş kur farkı = FC × (tahsilat kuru − fatura kuru)
    ///   - Açık kalan döviz faturalar için güncel kurla gerçekleşmemiş kur farkı
    /// </summary>
    public async Task<KurFarkiFifoSonuc> KurFarkiFifoHesaplaAsync(
        DateTime baslangic, DateTime bitis,
        string? cariKodu, string? cariUnvan,
        DateTime hesaplamaTarihi,
        int rateKolonu = 2,
        CancellationToken ct = default)
    {
        bitis = SistemTarihi.Clamp(bitis);
        hesaplamaTarihi = SistemTarihi.Clamp(hesaplamaTarihi);
        var (hareketler, _, _, _) = await CariHareketListesiAsync(
            baslangic, bitis, cariKodu, cariUnvan, ct);

        var sonuc = new KurFarkiFifoSonuc();
        if (hareketler.Count == 0) return sonuc;

        var exTbl = $"LG_EXCHANGE_{_db.FirmaNo}";
        var rateCol = rateKolonu switch
        {
            1 => "RATES1", 2 => "RATES2", 3 => "RATES3", 4 => "RATES4",
            _ => "RATES2"
        };
        using var conn = _db.CreateConnection();

        async Task<decimal> KurAl(int trcurr, DateTime tarih)
        {
            if (trcurr == 0) return 1m;
            var sql = $@"
SELECT TOP 1 {rateCol} FROM {exTbl} WITH(NOLOCK)
WHERE CRTYPE = @cr AND EDATE <= @t AND {rateCol} > 0
ORDER BY EDATE DESC";
            var r = await conn.ExecuteScalarAsync<decimal?>(
                new CommandDefinition(sql, new { cr = trcurr, t = tarih.Date }, cancellationToken: ct));
            return r ?? 0m;
        }

        bool EslesirMi(CariHareket h)
        {
            // Tahsilat tarafında: Havale (MODULENR=7) OR Kredi Kartı (TRCODE=20) OR Virman (TRCODE=5)
            if (h.Alacak <= 0 && h.Borc <= 0) return false;
            if (h.TrCode == 36) return false; // Mahsup Fişi hariç
            return h.ModuleNr == 7 || h.TrCode == 20 || h.TrCode == 5 || h.TrCurr != 0;
        }

        // Faturalar = SIGN=0 (Borç). Kur farkı fişi hariç, mahsup fişi hariç.
        var faturalar = hareketler
            .Where(h => h.Borc > 0 && h.TrCode != 6 && h.TrCode != 33 && h.TrCode != 36)
            .Select(h => new FifoAcikSatir
            {
                Tarih       = h.Tarih,
                Vade        = h.VadeTarihi ?? h.Tarih,
                FisTuru     = h.FisTuru,
                FisNo       = h.FisNo,
                TrCurr      = h.TrCurr,
                DovizKodu   = h.DovizKodu,
                KalanTl     = h.Borc,
                KalanFc     = h.DovizBorc,
                FaturaKuru  = h.DovizKur,
                Tip         = h.TrCurr == 0 ? "TL Fatura" : $"{h.DovizKodu} Fatura"
            })
            .OrderBy(f => f.Vade).ThenBy(f => f.Tarih)
            .ToList();

        // Tahsilatlar: Havale/KK/Virman + her türlü dövizli alacak
        var tahsilatlar = hareketler
            .Where(h => h.Alacak > 0 && EslesirMi(h))
            .Select(h => new
            {
                Hareket = h,
                KalanTl = h.Alacak,
                KalanFc = h.DovizAlacak
            })
            .OrderBy(t => t.Hareket.Tarih)
            .Select(t => new TahsilatState
            {
                Hareket = t.Hareket,
                KalanTl = t.KalanTl,
                KalanFc = t.KalanFc
            })
            .ToList();

        // FIFO döngüsü
        foreach (var t in tahsilatlar)
        {
            int guard = 0;
            while (t.KalanTl > 0.005m && guard++ < 1000)
            {
                // Sadece tahsilat ile AYNI döviz cinsindeki faturalar mahsup edilir
                // (Logo'nun davranışı bu şekilde — farklı dövizler çapraz mahsuplaştırılmaz)
                var fat = faturalar.FirstOrDefault(f =>
                    f.KalanTl > 0.005m && f.TrCurr == t.Hareket.TrCurr);
                if (fat == null) break;

                decimal mahsupTl, mahsupFc = 0, kurFarki = 0, kullanilanKur = 0;
                string tip, aciklama;

                if (t.Hareket.TrCurr == 0)
                {
                    // TL → TL
                    mahsupTl = Math.Min(t.KalanTl, fat.KalanTl);
                    t.KalanTl -= mahsupTl;
                    fat.KalanTl -= mahsupTl;
                    tip = "TL → TL";
                    aciklama = "Direkt mahsup (kur farkı yok)";
                }
                else
                {
                    // FC → FC (aynı döviz)
                    mahsupFc = Math.Min(t.KalanFc, fat.KalanFc);
                    if (mahsupFc <= 0.005m) break;
                    var oranlanmiTl = fat.KalanFc == 0 ? 0 : Math.Round(mahsupFc * fat.FaturaKuru, 2);
                    mahsupTl = Math.Round(mahsupFc * t.Hareket.DovizKur, 2);
                    t.KalanFc -= mahsupFc;
                    t.KalanTl -= mahsupTl;
                    fat.KalanFc -= mahsupFc;
                    fat.KalanTl -= oranlanmiTl;
                    kurFarki = Math.Round(mahsupFc * (t.Hareket.DovizKur - fat.FaturaKuru), 2);
                    kullanilanKur = t.Hareket.DovizKur;
                    tip = $"{fat.DovizKodu} → {fat.DovizKodu}";
                    aciklama = $"Fatura kuru {fat.FaturaKuru:N4}, tahsilat kuru {t.Hareket.DovizKur:N4}";
                }

                sonuc.Eslesmeler.Add(new FifoEslesme
                {
                    FaturaTarihi         = fat.Tarih,
                    FaturaVade           = fat.Vade,
                    FaturaFisTuru        = fat.FisTuru,
                    FaturaFisNo          = fat.FisNo,
                    FaturaTrCurr         = fat.TrCurr,
                    FaturaDovizKodu      = fat.DovizKodu,
                    FaturaKuru           = fat.FaturaKuru,
                    TahsilatTarihi       = t.Hareket.Tarih,
                    TahsilatFisTuru      = t.Hareket.FisTuru,
                    TahsilatFisNo        = t.Hareket.FisNo,
                    TahsilatTrCurr       = t.Hareket.TrCurr,
                    MahsupTl             = mahsupTl,
                    MahsupFc             = mahsupFc,
                    KullanilanKur        = kullanilanKur,
                    GerceklesmisKurFarki = kurFarki,
                    Tip                  = tip,
                    Aciklama             = aciklama
                });
            }

            if (t.KalanTl > 0.01m || t.KalanFc > 0.01m)
            {
                sonuc.FazlaTahsilatlar.Add(new FifoAcikSatir
                {
                    Tarih     = t.Hareket.Tarih,
                    Vade      = t.Hareket.Tarih,
                    FisTuru   = t.Hareket.FisTuru,
                    FisNo     = t.Hareket.FisNo,
                    TrCurr    = t.Hareket.TrCurr,
                    DovizKodu = t.Hareket.DovizKodu,
                    KalanTl   = Math.Round(t.KalanTl, 2),
                    KalanFc   = Math.Round(t.KalanFc, 2),
                    Tip       = "Fazla Tahsilat"
                });
            }
        }

        // Açık faturalar için gerçekleşmemiş kur farkı
        foreach (var fat in faturalar.Where(f => f.KalanTl > 0.01m || f.KalanFc > 0.01m))
        {
            if (fat.TrCurr != 0 && fat.KalanFc > 0.005m)
            {
                var guncelKur = await KurAl(fat.TrCurr, hesaplamaTarihi);
                fat.GuncelKur = guncelKur;
                fat.GuncelTlDeger = Math.Round(fat.KalanFc * guncelKur, 2);
                fat.GerceklesmemisKurFarki = Math.Round(fat.GuncelTlDeger - fat.KalanTl, 2);
            }
            sonuc.AcikFaturalar.Add(fat);
        }

        sonuc.ToplamGerceklesmis = sonuc.Eslesmeler.Sum(x => x.GerceklesmisKurFarki);
        sonuc.ToplamGerceklesmemis = sonuc.AcikFaturalar.Sum(x => x.GerceklesmemisKurFarki);
        sonuc.ToplamKurFarki = Math.Round(sonuc.ToplamGerceklesmis + sonuc.ToplamGerceklesmemis, 2);

        if (Math.Abs(sonuc.ToplamKurFarki) < 0.01m)
        {
            sonuc.KesenTaraf = "YOK";
            sonuc.Yon = "—";
        }
        else if (sonuc.ToplamKurFarki > 0)
        {
            sonuc.KesenTaraf = "BİZ → Karşı Firma";
            sonuc.Yon = "Lehimize (Gelir)";
        }
        else
        {
            sonuc.KesenTaraf = "Karşı Firma → BİZ";
            sonuc.Yon = "Aleyhimize (Gider)";
        }

        return sonuc;
    }

    private sealed class TahsilatState
    {
        public CariHareket Hareket { get; set; } = null!;
        public decimal KalanTl { get; set; }
        public decimal KalanFc { get; set; }
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
