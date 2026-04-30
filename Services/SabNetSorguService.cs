using System.Text;
using Dapper;
using Microsoft.Data.Sqlite;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// SabNet.db SQLite üzerinde liste sorguları (filtre + sayfalama).
/// </summary>
public class SabNetSorguService
{
    static SabNetSorguService()
    {
        DefaultTypeMap.MatchNamesWithUnderscores = true;
    }

    private readonly SabNetDbService _db;

    public SabNetSorguService(SabNetDbService db) => _db = db;

    public async Task<List<string>> SozlesmeYillariniGetirAsync()
    {
        await using var conn = new SqliteConnection(_db.ConnectionString);
        await conn.OpenAsync();
        var hareket = await conn.QueryAsync<string>(@"
            SELECT DISTINCT SozlesmeYili FROM SabNetKantarHareketleri
            WHERE SozlesmeYili IS NOT NULL AND SozlesmeYili <> ''");
        var log = await conn.QueryAsync<string>(@"
            SELECT DISTINCT SozlesmeYili FROM SabNetKantarHareketleriLog
            WHERE SozlesmeYili IS NOT NULL AND SozlesmeYili <> ''");
        return hareket.Union(log).OrderByDescending(x => x).ToList();
    }

    public async Task<Dictionary<string, string>> FirmaAdiLookupAsync()
    {
        await using var conn = new SqliteConnection(_db.ConnectionString);
        await conn.OpenAsync();
        var rows = await conn.QueryAsync<(string HesapKodu, string Unvan1)>(@"
            SELECT HesapKodu, Unvan1 FROM SabNetCariHesaplar
            WHERE HesapKodu IS NOT NULL AND Unvan1 IS NOT NULL");
        var map = new Dictionary<string, string>();
        foreach (var r in rows)
            if (!map.ContainsKey(r.HesapKodu)) map[r.HesapKodu] = r.Unvan1;
        return map;
    }

    public async Task<Dictionary<string, string>> UrunAdiLookupAsync()
    {
        await using var conn = new SqliteConnection(_db.ConnectionString);
        await conn.OpenAsync();
        var rows = await conn.QueryAsync<(long? StokKodu, string StokAdi)>(@"
            SELECT StokKodu, StokAdi FROM SabNetStokKartlari
            WHERE StokKodu IS NOT NULL AND StokAdi IS NOT NULL");
        var map = new Dictionary<string, string>();
        foreach (var r in rows)
        {
            var key = r.StokKodu!.Value.ToString();
            if (!map.ContainsKey(key)) map[key] = r.StokAdi;
        }
        return map;
    }

    public async Task<(long hareket, long log)> SayilarAsync(string? sozlesmeYili, string? kantar)
    {
        await using var conn = new SqliteConnection(_db.ConnectionString);
        await conn.OpenAsync();

        var (whereH, prmH) = WhereOlustur(sozlesmeYili, kantar, arama: null, isLog: false);
        var (whereL, prmL) = WhereOlustur(sozlesmeYili, kantar, arama: null, isLog: true);

        var h = await conn.ExecuteScalarAsync<long>($"SELECT COUNT(*) FROM SabNetKantarHareketleri {whereH}", prmH);
        var l = await conn.ExecuteScalarAsync<long>($"SELECT COUNT(*) FROM SabNetKantarHareketleriLog {whereL}", prmL);
        return (h, l);
    }

    public async Task<(List<SabNetKantarHareketi> rows, int toplam)> HareketleriGetirAsync(
        string? sozlesmeYili, string? kantar, string? arama, int sayfa, int sayfaBoyutu)
    {
        await using var conn = new SqliteConnection(_db.ConnectionString);
        await conn.OpenAsync();

        var (where, prm) = WhereOlustur(sozlesmeYili, kantar, arama, isLog: false);

        int toplam = await conn.ExecuteScalarAsync<int>($"SELECT COUNT(*) FROM SabNetKantarHareketleri {where}", prm);
        if (toplam == 0) return (new(), 0);

        var dprm = new DynamicParameters(prm);
        dprm.Add("@_offset", (sayfa - 1) * sayfaBoyutu);
        dprm.Add("@_limit", sayfaBoyutu);

        var rows = await conn.QueryAsync<SabNetKantarHareketi>(
            $"SELECT * FROM SabNetKantarHareketleri {where} ORDER BY Row_ID DESC LIMIT @_limit OFFSET @_offset", dprm);
        return (rows.ToList(), toplam);
    }

    public async Task<(List<SabNetKantarHareketiLog> rows, int toplam)> LoglariGetirAsync(
        string? sozlesmeYili, string? kantar, string? arama, int sayfa, int sayfaBoyutu)
    {
        await using var conn = new SqliteConnection(_db.ConnectionString);
        await conn.OpenAsync();

        var (where, prm) = WhereOlustur(sozlesmeYili, kantar, arama, isLog: true);

        int toplam = await conn.ExecuteScalarAsync<int>($"SELECT COUNT(*) FROM SabNetKantarHareketleriLog {where}", prm);
        if (toplam == 0) return (new(), 0);

        var dprm = new DynamicParameters(prm);
        dprm.Add("@_offset", (sayfa - 1) * sayfaBoyutu);
        dprm.Add("@_limit", sayfaBoyutu);

        var rows = await conn.QueryAsync<SabNetKantarHareketiLog>(
            $"SELECT * FROM SabNetKantarHareketleriLog {where} ORDER BY Id DESC LIMIT @_limit OFFSET @_offset", dprm);
        return (rows.ToList(), toplam);
    }

    private static (string where, DynamicParameters prm) WhereOlustur(string? sozlesmeYili, string? kantar, string? arama, bool isLog)
    {
        var sb = new StringBuilder();
        var prm = new DynamicParameters();
        var koşullar = new List<string>();

        if (!string.IsNullOrWhiteSpace(sozlesmeYili))
        {
            koşullar.Add("SozlesmeYili = @yil");
            prm.Add("@yil", sozlesmeYili);
        }

        if (kantar == "DIGER")
        {
            koşullar.Add("KantarKodu IS NOT NULL AND KantarKodu NOT IN ('K1','M1','M2')");
        }
        else if (!string.IsNullOrWhiteSpace(kantar))
        {
            koşullar.Add("KantarKodu = @kantar");
            prm.Add("@kantar", kantar);
        }

        if (!string.IsNullOrWhiteSpace(arama))
        {
            var s = "%" + arama.Trim() + "%";
            if (isLog)
            {
                koşullar.Add(@"(FisNo LIKE @s OR TcKimlikNo LIKE @s OR PlakaNo LIKE @s
                              OR HesapKodu LIKE @s OR LogKaydeden LIKE @s)");
            }
            else
            {
                koşullar.Add(@"(FisNo LIKE @s OR TcKimlikNo LIKE @s OR PlakaNo LIKE @s
                              OR HesapKodu LIKE @s OR RandevuNo LIKE @s)");
            }
            prm.Add("@s", s);
        }

        if (koşullar.Count > 0)
            sb.Append("WHERE ").Append(string.Join(" AND ", koşullar));

        return (sb.ToString(), prm);
    }
}
