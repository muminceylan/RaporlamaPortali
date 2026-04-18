using Dapper;
using Microsoft.Data.SqlClient;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

public class MustahsilLookupService
{
    private const string ConnStr =
        "Server=192.168.77.7;Database=SabNetPMHS;User Id=reportuser;" +
        "Password=reportuser;TrustServerCertificate=True;Connect Timeout=30;";

    public List<int> MevcutKampanyaYillari()
    {
        try
        {
            using var conn = new SqlConnection(ConnStr);
            return conn.Query<string>(
                "SELECT DISTINCT SozlesmeYili FROM dbo.PMHS_SozlesmeBilgileri WHERE ISNUMERIC(SozlesmeYili) = 1")
                .Select(s => int.TryParse(s, out var y) ? y : -1)
                .Where(y => y > 0)
                .OrderByDescending(y => y)
                .ToList();
        }
        catch
        {
            return new List<int>();
        }
    }

    public List<MustahsilOzet> Mustahsiller(int kampanyaYili)
    {
        try
        {
            using var conn = new SqlConnection(ConnStr);
            var sql = @"
SELECT
    s.TcKimlikNo                     AS TcKimlikNo,
    ISNULL(c.AdiSoyadi, '')          AS AdSoyadi,
    c.MustahsilHesapKodu             AS MustahsilNo,
    s.HesapNo                        AS HesapNo
FROM (
    SELECT TcKimlikNo, MAX(HesapNo) AS HesapNo
    FROM dbo.PMHS_SozlesmeBilgileri
    WHERE SozlesmeYili = @yil AND TcKimlikNo IS NOT NULL AND TcKimlikNo <> ''
    GROUP BY TcKimlikNo
) s
LEFT JOIN dbo.PMHS_CiftciKarti c ON c.TcKimlikNo = s.TcKimlikNo
ORDER BY ISNULL(c.AdiSoyadi, '')";
            return conn.Query<MustahsilOzet>(sql, new { yil = kampanyaYili.ToString() }).ToList();
        }
        catch
        {
            return new List<MustahsilOzet>();
        }
    }
}
