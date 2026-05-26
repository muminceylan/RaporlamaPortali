using Microsoft.Data.SqlClient;
using RaporlamaPortali.Models;
using RaporlamaPortali.Services.Logo;

namespace RaporlamaPortali.Services;

/// <summary>
/// SabNet PMHS'ten Ayni Avans satirlarini cekip Logo'ya Toptan Satis Faturasi olarak posting.
/// Referans: F:\30.12.2025 yedek\hakan\Genel Projeler\PMHS\PMHS\Raporlar\AyniAvansEntegrasyonu.vb
/// </summary>
public class SatisFaturaService
{
    private const string SabNetConnStr =
        "Server=192.168.77.7;Database=SabNetPMHS;User Id=reportuser;" +
        "Password=reportuser;TrustServerCertificate=True;Connect Timeout=30;";

    // Logo Unity bağlantı bilgileri (kullanıcı doğruladı: 13 May 2026)
    public static readonly LogoUnityCredentials LogoCred = new()
    {
        Kullanici = "Mümin CEYLAN",
        Sifre     = "2483480",
        FirmaNo   = 211,
        DonemNo   = 1
    };

    private readonly LogoUnityComService _logo;
    public SatisFaturaService(LogoUnityComService logo) { _logo = logo; }

    public async Task<List<AyniAvansSatiri>> GetAyniAvanslarAsync(
        IEnumerable<int>? avansNoFiltre = null,
        int sozlesmeYili = 2026,
        bool sadeceBasilmamis = true)
    {
        var liste = new List<AyniAvansSatiri>();
        await using var conn = new SqlConnection(SabNetConnStr);
        await conn.OpenAsync();

        var sql = @"
SELECT
  AF.FormNo,
  AF.SiraNo,
  AF.SozlesmeYili,
  AF.HesapNo,
  ISNULL(AF.TcKimlikNo,'') AS TcKimlikNo,
  ISNULL(CK.AdiSoyadi,'')  AS AdSoyad,
  ISNULL(CK.EPostaAdresi,'') AS EPostaAdresi,
  ISNULL(CK.GsmNo,'')        AS GsmNo,
  AF.AvansNo,
  ISNULL(AT.AvansAdi,'')        AS AvansAdi,
  ISNULL(AT.AvansStokKodu,'')   AS AvansStokKodu,
  ISNULL(AT.AvansBirimi,'KG')   AS AvansBirimi,
  ISNULL(AT.KdvDahilHaric,'DAHİL') AS KdvDahilHaric,
  ISNULL(AT.KdvOrani,0)         AS KdvOraniTanim,
  ISNULL(AT.Kod_AmbarKodu,'')   AS Kod_AmbarKodu,
  ISNULL(AT.Kod_FabrikaKodu,'') AS Kod_FabrikaKodu,
  ISNULL(AT.Kod_IsyeriKodu,'')  AS Kod_IsyeriKodu,
  ISNULL(AT.Kod_BolumKodu,'')   AS Kod_BolumKodu,
  ISNULL(AT.Kod_TicaretGrubu,'AFYON') AS Kod_TicaretGrubu,
  AF.Miktar,
  AF.BirimFiyat,
  AF.Tutar,
  ISNULL(AF.KdvOrani,0) AS KdvOrani,
  ISNULL(AF.KdvTutari,0) AS KdvTutari,
  ISNULL(AF.KdvDahilHaric,'DAHİL') AS KdvDahilHaricSatir,
  ISNULL(AF.KaynakBolge,'') AS KaynakBolge,
  ISNULL(AF.ErpEvrakNo,'')    AS ErpEvrakNo,
  AF.ErpEvrakTarihi,
  ISNULL(AF.ErpEvrakTipi,'')  AS ErpEvrakTipi,
  CONVERT(date, DATEADD(DAY, AF.FormTarihi-2, '1900-01-01')) AS FormTarihi
FROM PMHS_AvansFormu AF
LEFT JOIN PMHS_AvansTanimlari AT ON AT.AvansKodu = AF.AvansNo
LEFT JOIN PMHS_CiftciKarti   CK ON CK.TcKimlikNo = AF.TcKimlikNo
WHERE AF.SozlesmeYili = @yil
  AND AF.KaynakBolge <> ''
  AND AT.Cinsi = 'AYNİ'
";
        if (avansNoFiltre != null)
        {
            var list = avansNoFiltre.ToList();
            if (list.Count > 0)
                sql += " AND AF.AvansNo IN (" + string.Join(",", list) + ") ";
        }
        if (sadeceBasilmamis)
            sql += " AND (AF.ErpEvrakNo IS NULL OR LTRIM(RTRIM(AF.ErpEvrakNo)) = '') ";

        sql += " ORDER BY AF.AvansNo, AF.FormNo, AF.SiraNo ";

        await using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("@yil", sozlesmeYili);
        cmd.CommandTimeout = 120;
        await using var rd = await cmd.ExecuteReaderAsync();
        while (await rd.ReadAsync())
        {
            liste.Add(new AyniAvansSatiri
            {
                FormNo         = rd["FormNo"]?.ToString() ?? "",
                SiraNo         = rd["SiraNo"] is int i ? i : 0,
                HesapNo        = rd["HesapNo"]?.ToString() ?? "",
                TcKimlikNo     = rd["TcKimlikNo"]?.ToString() ?? "",
                AdSoyad        = rd["AdSoyad"]?.ToString() ?? "",
                EPostaAdresi   = rd["EPostaAdresi"]?.ToString() ?? "",
                GsmNo          = rd["GsmNo"]?.ToString() ?? "",
                AvansNo        = Convert.ToInt32(rd["AvansNo"]),
                AvansAdi       = rd["AvansAdi"]?.ToString() ?? "",
                AvansStokKodu  = rd["AvansStokKodu"]?.ToString() ?? "",
                AvansBirimi    = rd["AvansBirimi"]?.ToString() ?? "KG",
                KdvDahilHaric  = rd["KdvDahilHaric"]?.ToString() ?? "DAHİL",
                KdvOrani       = Convert.ToDecimal(rd["KdvOrani"]),
                Miktar         = Convert.ToDecimal(rd["Miktar"]),
                BirimFiyat     = Convert.ToDecimal(rd["BirimFiyat"]),
                Tutar          = Convert.ToDecimal(rd["Tutar"]),
                KaynakBolge    = rd["KaynakBolge"]?.ToString() ?? "",
                ErpEvrakNo     = rd["ErpEvrakNo"]?.ToString() ?? "",
                FormTarihi     = rd["FormTarihi"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(rd["FormTarihi"]),
                Kod_AmbarKodu      = rd["Kod_AmbarKodu"]?.ToString() ?? "",
                Kod_FabrikaKodu    = rd["Kod_FabrikaKodu"]?.ToString() ?? "",
                Kod_IsyeriKodu     = rd["Kod_IsyeriKodu"]?.ToString() ?? "",
                Kod_BolumKodu      = rd["Kod_BolumKodu"]?.ToString() ?? "",
                Kod_TicaretGrubu   = rd["Kod_TicaretGrubu"]?.ToString() ?? "AFYON",
            });
        }
        return liste;
    }

    public async Task<List<int>> GetAvailableAvansNolarAsync(int sozlesmeYili = 2026)
    {
        var liste = new List<int>();
        await using var conn = new SqlConnection(SabNetConnStr);
        await conn.OpenAsync();
        var sql = @"
SELECT DISTINCT AF.AvansNo
FROM PMHS_AvansFormu AF
LEFT JOIN PMHS_AvansTanimlari AT ON AT.AvansKodu = AF.AvansNo
WHERE AF.SozlesmeYili = @yil AND AF.KaynakBolge <> '' AND AT.Cinsi = 'AYNİ'
ORDER BY AF.AvansNo";
        await using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("@yil", sozlesmeYili);
        await using var rd = await cmd.ExecuteReaderAsync();
        while (await rd.ReadAsync()) liste.Add(Convert.ToInt32(rd[0]));
        return liste;
    }

    public async Task<Dictionary<int, string>> GetAvansAdlariAsync()
    {
        var d = new Dictionary<int, string>();
        await using var conn = new SqlConnection(SabNetConnStr);
        await conn.OpenAsync();
        await using var cmd = new SqlCommand(
            "SELECT AvansKodu, AvansAdi FROM PMHS_AvansTanimlari WHERE Cinsi='AYNİ' ORDER BY AvansAdi", conn);
        await using var rd = await cmd.ExecuteReaderAsync();
        while (await rd.ReadAsync())
            d[Convert.ToInt32(rd[0])] = rd[1]?.ToString() ?? "";
        return d;
    }

    /// <summary>
    /// Logo'ya post + başarılılarda PMHS_AvansFormu güncelle.
    /// </summary>
    public async Task<LogoSatisAktarimSonuc> LogoyaAktarAsync(
        List<AyniAvansSatiri> satirlar,
        DateTime faturaTarihi,
        long erpEvrakNoBaslangic,
        CancellationToken ct = default)
    {
        // Model listesi
        var batch = new List<LogoSatisFaturaModel>();
        long siraNo = erpEvrakNoBaslangic;
        foreach (var s in satirlar)
        {
            var no = siraNo.ToString();
            var m = new LogoSatisFaturaModel
            {
                Tip          = 8,
                FaturaNo     = "B " + no,
                Tarih        = faturaTarihi,
                CariKod      = "S" + s.TcKimlikNo,
                CariUnvan    = s.AdSoyad,
                CariTcKimlik = s.TcKimlikNo,
                CariEPosta   = s.EPostaAdresi,
                CariTelefon  = s.GsmNo,
                TradingGrp   = string.IsNullOrWhiteSpace(s.Kod_TicaretGrubu) ? "AFYON" : s.Kod_TicaretGrubu,
                Fabrika      = ParseIntOr(s.Kod_FabrikaKodu, 17),
                Ambar        = ParseIntOr(s.Kod_AmbarKodu, 147),
                Isyeri       = s.Kod_IsyeriKodu,
                Bolum        = s.Kod_BolumKodu,
                ProjeKodu    = "",
                MasterCode   = s.AvansStokKodu,
                Birim        = s.AvansBirimi,
                Miktar       = s.Miktar,
                BirimFiyat   = HesaplaBirimFiyat(s),
                KdvOran      = s.KdvOrani,
                KdvDahilHaric= s.KdvDahilHaric,
                FormNo       = s.FormNo,
                HesapNo      = s.HesapNo,
            };
            batch.Add(m);
            siraNo++;
        }

        var sonuc = await _logo.AktarSalesAsync(LogoCred, batch, ct);

        // Başarılı satırları PMHS_AvansFormu'nda güncelle
        if (sonuc.LoginBasarili && sonuc.Faturalar.Any(f => f.Basarili))
        {
            await using var conn = new SqlConnection(SabNetConnStr);
            await conn.OpenAsync();
            int erpTarih = (int)faturaTarihi.ToOADate();
            foreach (var r in sonuc.Faturalar.Where(f => f.Basarili))
            {
                try
                {
                    await using var upd = new SqlCommand(
                        @"UPDATE PMHS_AvansFormu
                              SET ErpEvrakTipi='Fatura', ErpEvrakNo=@no, ErpEvrakTarihi=@tar
                            WHERE FormNo=@fno AND HesapNo=@hno", conn);
                    // SabNet'in pattern'iyle birebir: "B 1305260001" (prefix dahil) yazılır.
                    upd.Parameters.AddWithValue("@no", r.FaturaNo);
                    upd.Parameters.AddWithValue("@tar", erpTarih);
                    upd.Parameters.AddWithValue("@fno", r.FormNo);
                    upd.Parameters.AddWithValue("@hno", r.HesapNo);
                    upd.CommandTimeout = 30;
                    await upd.ExecuteNonQueryAsync();
                }
                catch { /* tek satır update başarısız ise diğerleri etkilenmesin */ }
            }
        }

        return sonuc;
    }

    private static decimal HesaplaBirimFiyat(AyniAvansSatiri s)
    {
        if (s.KdvOrani <= 0) return s.BirimFiyat;
        // KDV Dahil ise birim fiyatı KDV'den arındır; Hariç ise olduğu gibi.
        if (string.Equals(s.KdvDahilHaric, "DAHİL", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(s.KdvDahilHaric, "DAHIL", StringComparison.OrdinalIgnoreCase))
        {
            return Math.Round(s.BirimFiyat / (1m + s.KdvOrani / 100m), 6, MidpointRounding.AwayFromZero);
        }
        return s.BirimFiyat;
    }

    private static int ParseIntOr(string s, int def) => int.TryParse(s, out var v) ? v : def;
}
