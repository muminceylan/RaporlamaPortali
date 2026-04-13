using System.Data;
using Dapper;
using Microsoft.Data.SqlClient;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

public class PancarRaporService
{
    private const string ConnStr =
        "Server=192.168.77.7;Database=SabNetPMHS;User Id=reportuser;" +
        "Password=reportuser;TrustServerCertificate=True;Connect Timeout=30;";

    // Aktif kampanya yılı — yeni kampanya başladığında burası güncellenir
    public static int KampanyaYili() => 2026;

    /// <summary>İCMAL özeti (Çiftçi/Kepçe/Mouse/Müteahhit toplamları)</summary>
    public async Task<List<PancarIcmalKayit>> GetIcmalAsync()
    {
        int yil = KampanyaYili();
        using var conn = new SqlConnection(ConnStr);
        try
        {
            var sql = $"SELECT Tip, Aciklama, Net, Tutar FROM OzetMouseKepceCiftciMuteahhit_{yil} ORDER BY Tip, Aciklama";
            return (await conn.QueryAsync<PancarIcmalKayit>(sql)).ToList();
        }
        catch { return new List<PancarIcmalKayit>(); }
    }

    /// <summary>Çiftçi özet listesi (VW_CiftciPancarOzet)</summary>
    public async Task<List<PancarCiftciDetay>> GetCiftciListesiAsync()
    {
        int yil = KampanyaYili();
        using var conn = new SqlConnection(ConnStr);
        try
        {
            var sql = $@"SELECT Bolge, Koy, HesapKodu, TcKimlikNo, AdiSoyadi,
                                TaahhutTon, NetMiktar, APancari, APancariYuzde,
                                KotaFazlasi, OrtalamaPolar
                         FROM VW_CiftciPancarOzet_{yil}
                         ORDER BY Bolge, Koy, AdiSoyadi";
            return (await conn.QueryAsync<PancarCiftciDetay>(sql)).ToList();
        }
        catch { return new List<PancarCiftciDetay>(); }
    }

    /// <summary>Tam detay (bordro hesaplarıyla birlikte)</summary>
    public async Task<List<PancarDetayTam>> GetDetayTamAsync()
    {
        int yil = KampanyaYili();
        using var conn = new SqlConnection(ConnStr);
        try
        {
            var sql = $@"SELECT HesapKodu, TcKimlikNo, AdiSoyadi, BabaAdi,
                                TaahhutTon, TaahhutTonA, NetMiktar, APancari, APancariBedeli,
                                KotaFazlasi, OrtalamaPolar,
                                NakdiAvans, AvansToplami, Hakedis, NetHakedis
                         FROM KarakanRaporu_Detay_{yil}
                         ORDER BY AdiSoyadi";
            return (await conn.QueryAsync<PancarDetayTam>(sql)).ToList();
        }
        catch { return new List<PancarDetayTam>(); }
    }

    /// <summary>Avans hareketleri — AvansGrubu bazında toplamlar (canlı: veri girdikçe yansır)</summary>
    public async Task<List<PancarAvansKayit>> GetAvansAsync()
    {
        int yil = KampanyaYili();
        return await GetAvansForYilAsync(yil);
    }

    private async Task<List<PancarAvansKayit>> GetAvansForYilAsync(int yil)
    {
        using var conn = new SqlConnection(ConnStr);
        try
        {
            var sql = @"
                SELECT
                    AF.KaynakEvrak,
                    ISNULL(AT.AvansGrubu, AF.AvansNo) AS AvansGrubu,
                    SUM(CASE WHEN AF.KdvDahilHaric = 'DAHİL'
                             THEN AF.Tutar - AF.KdvTutari
                             ELSE AF.Tutar END)  AS TutarToplami,
                    SUM(AF.StopajTutari)          AS StopajToplami,
                    SUM(AF.KdvTutari)             AS KdvToplami
                FROM PMHS_AvansFormu AS AF
                LEFT JOIN PMHS_AvansTanimlari AS AT ON AT.AvansKodu = AF.AvansNo
                WHERE AF.SozlesmeYili = @Yil
                  AND AF.KaynakEvrak IN (N'AYNİ AVANS', N'NAKDİ AVANS')
                  AND AF.KaynakBolge <> ''
                GROUP BY AF.KaynakEvrak, ISNULL(AT.AvansGrubu, AF.AvansNo)
                ORDER BY AF.KaynakEvrak DESC, ISNULL(AT.AvansGrubu, AF.AvansNo)";
            return (await conn.QueryAsync<PancarAvansKayit>(sql, new { Yil = yil })).ToList();
        }
        catch { return new List<PancarAvansKayit>(); }
    }

    /// <summary>Özet istatistikler: çiftçi sayısı, taahhüt, net, fire, polar — sadece kampanya yılı</summary>
    public async Task<PancarOzetIstatistik> GetOzetIstatistikAsync()
    {
        int yil = KampanyaYili();
        using var conn = new SqlConnection(ConnStr);
        try
        {
            // Çiftçi sayısı ve toplam taahhüt
            var sqlC = $@"SELECT ISNULL(COUNT(*),0) AS ToplamCiftci,
                                 ISNULL(SUM(TaahhutTon),0) AS ToplamTaahhut
                          FROM VW_CiftciPancarOzet_{yil}";
            var ciftci = await conn.QueryFirstOrDefaultAsync(sqlC);

            // Gelen net, fire oranı, polar — CiftciKantariHareketleri, sadece o kampanya yılı
            var sqlK = @"SELECT ISNULL(SUM([NET]),0)                           AS ToplamNet,
                                ISNULL(AVG(CAST([FİRE ORANI] AS FLOAT)),0)     AS OrtFireOrani,
                                ISNULL(AVG(CAST([POLAR ORANI] AS FLOAT)),0)    AS OrtPolar
                         FROM CiftciKantariHareketleri
                         WHERE [KAMPANYA YILI] = @Yil";
            var kantar = await conn.QueryFirstOrDefaultAsync(sqlK, new { Yil = yil });

            return new PancarOzetIstatistik
            {
                ToplamCiftci  = (int)    (ciftci?.ToplamCiftci  ?? 0),
                ToplamTaahhut = (decimal)(ciftci?.ToplamTaahhut ?? 0m),
                ToplamNet     = (decimal)(kantar?.ToplamNet     ?? 0m),
                OrtFireOrani  = (double) (kantar?.OrtFireOrani  ?? 0.0),
                OrtPolar      = (double) (kantar?.OrtPolar      ?? 0.0),
            };
        }
        catch { return new PancarOzetIstatistik(); }
    }

    /// <summary>KarakanRaporu_Detay'dan finansal özet (KDV, Stopaj, Nakliye, Kota, Bağkur, Borsa...)</summary>
    public async Task<PancarFinansOzet> GetFinansOzetAsync()
    {
        int yil = KampanyaYili();
        using var conn = new SqlConnection(ConnStr);
        try
        {
            var sql = $@"SELECT
                SUM(AvansKdv)            AS AvansKdv,
                ISNULL((SELECT SUM(StopajTutari) FROM PMHS_AvansFormu
                         WHERE SozlesmeYili = {yil}
                           AND KaynakEvrak IN (N'AYNİ AVANS', N'NAKDİ AVANS')
                           AND KaynakBolge <> ''), 0) AS AlimStopaji,
                SUM(OdenenNakliyePrimi)  AS NakliyePrimi,
                SUM(KotaCezasi)          AS KotaCezasi,
                SUM(BagkurBorcu)         AS BagkurBorcu,
                SUM(EBorsaTescil)        AS BorsaTescil,
                SUM(KotaFazlasiOdemesi)  AS KotaFazlasi,
                SUM(CPancariOdemesi)     AS CPancari
            FROM KarakanRaporu_Detay_{yil}";
            return (await conn.QueryFirstOrDefaultAsync<PancarFinansOzet>(sql)) ?? new PancarFinansOzet();
        }
        catch { return new PancarFinansOzet(); }
    }

    /// <summary>KarakanRaporu_Detay özet: pancar türleri, primler, nakliye</summary>
    public async Task<PancarIcmalDetay> GetIcmalDetayAsync()
    {
        int yil = KampanyaYili();
        using var conn = new SqlConnection(ConnStr);
        try
        {
            var sql = $@"
                SELECT
                    ISNULL(SUM(NetMiktar)          / 1000, 0) AS NetMiktarTon,
                    ISNULL(SUM(APancari)           / 1000, 0) AS APancariTon,
                    ISNULL(SUM(APancariBedeli),         0)    AS APancariBedeli,
                    ISNULL(SUM(CPancari)           / 1000, 0) AS CPancariTon,
                    ISNULL(SUM(CPancariBedeli),         0)    AS CPancariBedeli,
                    ISNULL(SUM(KotaFazlasi)        / 1000, 0) AS KotaFazlasiTon,
                    ISNULL(SUM(KotaFazlasiBedeli),      0)    AS KotaFazlasiBedeli,
                    ISNULL(SUM(KuspePrimiA) + SUM(KuspePrimiC), 0) AS KuspePrimi,
                    ISNULL(SUM(KotaTamamlamaPrimi),     0)    AS KotaTamamlamaPrimi,
                    ISNULL(SUM(OdenenNakliyePrimi),     0)    AS MustahsilNakliye
                FROM KarakanRaporu_Detay_{yil}";
            return (await conn.QueryFirstOrDefaultAsync<PancarIcmalDetay>(sql)) ?? new PancarIcmalDetay();
        }
        catch { return new PancarIcmalDetay(); }
    }

    /// <summary>Bölge bazında özet istatistik</summary>
    public async Task<(int ToplamCiftci, decimal ToplamTaahhut, decimal ToplamNet, decimal OrtPolar)> GetOzet()
    {
        int yil = KampanyaYili();
        using var conn = new SqlConnection(ConnStr);
        try
        {
            var sql = $@"SELECT COUNT(*) as C, SUM(TaahhutTon) as T, SUM(NetMiktar) as N, AVG(OrtalamaPolar) as P
                         FROM VW_CiftciPancarOzet_{yil}";
            var row = await conn.QueryFirstOrDefaultAsync(sql);
            return ((int)(row?.C ?? 0), (decimal)(row?.T ?? 0m),
                    (decimal)(row?.N ?? 0m), (decimal)(row?.P ?? 0m));
        }
        catch { return (0, 0, 0, 0); }
    }
}
