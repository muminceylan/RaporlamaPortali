using System.Globalization;
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

    private static readonly CultureInfo TrCulture = CultureInfo.GetCultureInfo("tr-TR");

    /// <summary>
    /// Türkçe locale'inde lowercase yapan SQLite fonksiyonu. WHERE'de filtreyi
    /// büyük/küçük harf duyarsız hale getirmek için kullanılır.
    /// </summary>
    private async Task<SqliteConnection> AcAsync()
    {
        var conn = new SqliteConnection(_db.ConnectionString);
        await conn.OpenAsync();
        conn.CreateFunction<string?, string?>("trlower",
            s => s == null ? null : s.ToLower(TrCulture));
        return conn;
    }

    public async Task<List<string>> SozlesmeYillariniGetirAsync()
    {
        await using var conn = await AcAsync();
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
        await using var conn = await AcAsync();
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
        await using var conn = await AcAsync();
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

    public async Task<(long hareket, long log)> SayilarAsync(
        string? sozlesmeYili, string? kantar,
        DateTime? baslangic = null, DateTime? bitis = null,
        Dictionary<string, string>? sutunFiltreleri = null)
    {
        await using var conn = await AcAsync();

        var (whereH, prmH) = WhereOlustur(sozlesmeYili, kantar, arama: null, baslangic, bitis, sutunFiltreleri, isLog: false);
        var (whereL, prmL) = WhereOlustur(sozlesmeYili, kantar, arama: null, baslangic, bitis, sutunFiltreleri, isLog: true);

        var h = await conn.ExecuteScalarAsync<long>($"SELECT COUNT(*) FROM SabNetKantarHareketleri {whereH}", prmH);
        var l = await conn.ExecuteScalarAsync<long>($"SELECT COUNT(*) FROM SabNetKantarHareketleriLog {whereL}", prmL);
        return (h, l);
    }

    public async Task<(List<SabNetKantarHareketi> rows, int toplam)> HareketleriGetirAsync(
        string? sozlesmeYili, string? kantar, string? arama, int sayfa, int sayfaBoyutu,
        DateTime? baslangic = null, DateTime? bitis = null,
        Dictionary<string, string>? sutunFiltreleri = null)
    {
        await using var conn = await AcAsync();

        var (where, prm) = WhereOlustur(sozlesmeYili, kantar, arama, baslangic, bitis, sutunFiltreleri, isLog: false);

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
        string? sozlesmeYili, string? kantar, string? arama, int sayfa, int sayfaBoyutu,
        DateTime? baslangic = null, DateTime? bitis = null,
        Dictionary<string, string>? sutunFiltreleri = null)
    {
        await using var conn = await AcAsync();

        var (where, prm) = WhereOlustur(sozlesmeYili, kantar, arama, baslangic, bitis, sutunFiltreleri, isLog: true);

        int toplam = await conn.ExecuteScalarAsync<int>($"SELECT COUNT(*) FROM SabNetKantarHareketleriLog {where}", prm);
        if (toplam == 0) return (new(), 0);

        var dprm = new DynamicParameters(prm);
        dprm.Add("@_offset", (sayfa - 1) * sayfaBoyutu);
        dprm.Add("@_limit", sayfaBoyutu);

        var rows = await conn.QueryAsync<SabNetKantarHareketiLog>(
            $"SELECT * FROM SabNetKantarHareketleriLog {where} ORDER BY Id DESC LIMIT @_limit OFFSET @_offset", dprm);
        return (rows.ToList(), toplam);
    }

    private static readonly DateTime _delphiBase = new(1899, 12, 30);
    private static int DelphiSerial(DateTime d) => (int)(d.Date - _delphiBase).TotalDays;

    /// <summary>
    /// UI sütun anahtarı → SQL kolonu ve tipi (text/numeric/dateSerial/saatSerial).
    /// Tip, hangi SQL ifadesinin uygulanacağını belirler.
    /// </summary>
    private enum SutunTipi { Text, Numeric, DateSerial, SaatSerial }

    private static readonly Dictionary<string, (string sql, SutunTipi tip)> _sutunMap = new()
    {
        // Tarih/saat kolonları (Delphi serial int)
        ["Tarih"]        = ("Tarih",        SutunTipi.DateSerial),
        ["KayitTarihi"]  = ("KayitTarihi",  SutunTipi.DateSerial),
        ["KayitSaati"]   = ("KayitSaati",   SutunTipi.SaatSerial),
        ["CikisTarihi"]  = ("CikisTarihi",  SutunTipi.DateSerial),
        ["CikisSaati"]   = ("CikisSaati",   SutunTipi.SaatSerial),
        ["LogKayitTarihi"] = ("LogKayitTarihi", SutunTipi.DateSerial),
        ["LogKayitSaati"]  = ("LogKayitSaati",  SutunTipi.SaatSerial),
        // Text kolonlar
        ["FisNo"]          = ("FisNo",          SutunTipi.Text),
        ["IslemTipi"]      = ("IslemTipi",      SutunTipi.Text),
        ["SozlesmeYili"]   = ("SozlesmeYili",   SutunTipi.Text),
        ["TcKimlikNo"]     = ("TcKimlikNo",     SutunTipi.Text),
        ["HesapKodu"]      = ("HesapKodu",      SutunTipi.Text),
        ["UrunKodu"]       = ("UrunKodu",       SutunTipi.Text),
        ["PlakaNo"]        = ("PlakaNo",        SutunTipi.Text),
        ["SoforAdiSoyadi"] = ("SoforAdiSoyadi", SutunTipi.Text),
        ["Kod5"]           = ("Kod5",           SutunTipi.Text),
        ["BosaltmaYeri"]   = ("BosaltmaYeri",   SutunTipi.Text),
        ["Aciklama"]       = ("Aciklama",       SutunTipi.Text),
        ["KantarKodu"]     = ("KantarKodu",     SutunTipi.Text),
        ["LogIslemTipi"]   = ("LogIslemTipi",   SutunTipi.Text),
        ["LogKaydeden"]    = ("LogKaydeden",    SutunTipi.Text),
        ["LogAciklama"]    = ("LogAciklama",    SutunTipi.Text),
        // Numerik kolonlar
        ["BirimFiyat"]     = ("BirimFiyat",     SutunTipi.Numeric),
        ["Brut"]           = ("Brut",           SutunTipi.Numeric),
        ["Dara"]           = ("Dara",           SutunTipi.Numeric),
        ["Net"]            = ("Net",            SutunTipi.Numeric),
        ["Sevk"]           = ("Sevk",           SutunTipi.Numeric),
        ["Fark"]           = ("Fark",           SutunTipi.Numeric),
        ["FireOrani"]      = ("FireOrani",      SutunTipi.Numeric),
        ["PolarOrani"]     = ("PolarOrani",     SutunTipi.Numeric),
        ["Nakit"]          = ("Nakit",          SutunTipi.Numeric),
        ["KrediKarti"]     = ("KrediKarti",     SutunTipi.Numeric),
        ["Cari"]           = ("Cari",           SutunTipi.Numeric),
        ["Havale"]         = ("Havale",         SutunTipi.Numeric),
        ["Row_ID"]         = ("Row_ID",         SutunTipi.Numeric),
    };

    private static (string where, DynamicParameters prm) WhereOlustur(
        string? sozlesmeYili, string? kantar, string? arama,
        DateTime? baslangic, DateTime? bitis,
        Dictionary<string, string>? sutunFiltreleri, bool isLog)
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
            koşullar.Add("KantarKodu IS NOT NULL AND KantarKodu COLLATE NOCASE NOT IN ('K1','M1','M2')");
        }
        else if (!string.IsNullOrWhiteSpace(kantar))
        {
            koşullar.Add("KantarKodu COLLATE NOCASE = @kantar");
            prm.Add("@kantar", kantar);
        }

        if (baslangic.HasValue)
        {
            koşullar.Add("Tarih >= @tarihBas");
            prm.Add("@tarihBas", DelphiSerial(baslangic.Value));
        }
        if (bitis.HasValue)
        {
            koşullar.Add("Tarih <= @tarihBit");
            prm.Add("@tarihBit", DelphiSerial(bitis.Value));
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

        // Sütun bazlı filtreler — büyük/küçük harf duyarsız (trlower) + Türkçe locale
        if (sutunFiltreleri != null)
        {
            int idx = 0;
            foreach (var (key, val) in sutunFiltreleri)
            {
                if (string.IsNullOrWhiteSpace(val)) continue;
                var pname = "@flt" + idx++;
                var pattern = "%" + val.Trim() + "%";

                if (key == "FirmaAdi")
                {
                    koşullar.Add(
                        $"HesapKodu IN (SELECT HesapKodu FROM SabNetCariHesaplar WHERE trlower(IFNULL(Unvan1,'')) LIKE trlower({pname}))");
                    prm.Add(pname, pattern);
                }
                else if (key == "UrunAdi")
                {
                    koşullar.Add(
                        $"UrunKodu IN (SELECT CAST(StokKodu AS TEXT) FROM SabNetStokKartlari WHERE trlower(IFNULL(StokAdi,'')) LIKE trlower({pname}))");
                    prm.Add(pname, pattern);
                }
                else if (_sutunMap.TryGetValue(key, out var info))
                {
                    var ifade = info.tip switch
                    {
                        SutunTipi.Text =>
                            $"trlower(IFNULL({info.sql},'')) LIKE trlower({pname})",
                        SutunTipi.Numeric =>
                            $"CAST(IFNULL({info.sql},0) AS TEXT) LIKE {pname}",
                        SutunTipi.DateSerial =>
                            // Delphi serial → 'dd.MM.yyyy' string
                            $"strftime('%d.%m.%Y', date('1899-12-30', '+'||IFNULL({info.sql},0)||' days')) LIKE {pname}",
                        SutunTipi.SaatSerial =>
                            // Saniye → 'HH:mm:ss' string
                            $"printf('%02d:%02d:%02d', IFNULL({info.sql},0)/3600, (IFNULL({info.sql},0)%3600)/60, IFNULL({info.sql},0)%60) LIKE {pname}",
                        _ => null
                    };
                    if (ifade != null)
                    {
                        koşullar.Add(ifade);
                        prm.Add(pname, pattern);
                    }
                }
                // bilinmeyen anahtarlar sessizce yok sayılır (SQL injection güvencesi)
            }
        }

        if (koşullar.Count > 0)
            sb.Append("WHERE ").Append(string.Join(" AND ", koşullar));

        return (sb.ToString(), prm);
    }
}
