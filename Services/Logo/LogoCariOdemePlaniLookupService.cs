using System.Collections.Concurrent;
using Dapper;

namespace RaporlamaPortali.Services.Logo;

/// <summary>
/// Logo cari kartında tanımlı ödeme planı (vade) KODunu çeker.
/// CLCARD.PAYMENTREF > 0 → PAYPLANS.CODE join.
///
/// Satış Siparişi aktarımında: cari kartında plan tanımlı ise PAYMENT_CODE
/// olarak bu kodu kullanırız (Logo Unity COM, ARP_CODE üzerinden cari default'unu
/// kendiliğinden uygulamıyor — elle göndermek zorundayız).
/// Boş string = cari'de plan yok → Excel "Peşin/Vadeli" kuralı devreye girer.
/// </summary>
public class LogoCariOdemePlaniLookupService
{
    private readonly DatabaseService _db;
    private readonly ILogger<LogoCariOdemePlaniLookupService> _log;
    private static readonly ConcurrentDictionary<string, string> _cache = new(StringComparer.OrdinalIgnoreCase);

    public LogoCariOdemePlaniLookupService(DatabaseService db, ILogger<LogoCariOdemePlaniLookupService> log)
    {
        _db = db;
        _log = log;
    }

    /// <summary>Tek cari için ödeme planı kodu ("" → tanımlı değil).</summary>
    public async Task<string> PlanKoduAsync(string? cariKodu, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(cariKodu)) return "";
        var key = $"{_db.FirmaNo}:{cariKodu.Trim()}";
        if (_cache.TryGetValue(key, out var v)) return v;

        try
        {
            var sql = $@"
                SELECT TOP 1 p.CODE
                FROM LG_{_db.FirmaNo}_CLCARD c WITH(NOLOCK)
                INNER JOIN LG_{_db.FirmaNo}_PAYPLANS p WITH(NOLOCK) ON p.LOGICALREF = c.PAYMENTREF
                WHERE c.CODE = @code AND c.PAYMENTREF > 0";
            using var conn = _db.CreateConnection();
            var kod = await conn.QueryFirstOrDefaultAsync<string?>(
                new CommandDefinition(sql, new { code = cariKodu.Trim() }, cancellationToken: ct));
            kod = (kod ?? "").Trim();
            _cache[key] = kod;
            return kod;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Cari ödeme planı lookup başarısız (cariKodu={C})", cariKodu);
            return "";
        }
    }

    /// <summary>
    /// Toplu sorgu — bir aktarımdaki tüm fişler için tek SQL çağrısı.
    /// Sonuç: cariKodu → planKodu (boş veya "47" / "55" / "70" gibi).
    /// </summary>
    public async Task<Dictionary<string, string>> ToplulukPlanKodlariAsync(
        IEnumerable<string> cariKodlari, CancellationToken ct = default)
    {
        var sonuc = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var hedef = cariKodlari.Where(k => !string.IsNullOrWhiteSpace(k))
                                .Select(k => k.Trim())
                                .Distinct(StringComparer.OrdinalIgnoreCase)
                                .ToList();
        if (hedef.Count == 0) return sonuc;

        var sorulacak = new List<string>();
        foreach (var k in hedef)
        {
            var key = $"{_db.FirmaNo}:{k}";
            if (_cache.TryGetValue(key, out var v)) sonuc[k] = v;
            else sorulacak.Add(k);
        }
        if (sorulacak.Count == 0) return sonuc;

        try
        {
            var sql = $@"
                SELECT c.CODE, p.CODE AS PLAN_CODE
                FROM LG_{_db.FirmaNo}_CLCARD c WITH(NOLOCK)
                INNER JOIN LG_{_db.FirmaNo}_PAYPLANS p WITH(NOLOCK) ON p.LOGICALREF = c.PAYMENTREF
                WHERE c.PAYMENTREF > 0 AND c.CODE IN @kodlar";
            using var conn = _db.CreateConnection();
            var rows = await conn.QueryAsync<(string CODE, string PLAN_CODE)>(
                new CommandDefinition(sql, new { kodlar = sorulacak }, cancellationToken: ct));
            foreach (var r in rows)
            {
                var code = (r.CODE ?? "").Trim();
                var plan = (r.PLAN_CODE ?? "").Trim();
                sonuc[code] = plan;
                _cache[$"{_db.FirmaNo}:{code}"] = plan;
            }
            // Sorulanlardan dönmeyenler → "" cache + sonuç
            foreach (var k in sorulacak)
                if (!sonuc.ContainsKey(k)) { sonuc[k] = ""; _cache[$"{_db.FirmaNo}:{k}"] = ""; }
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Toplu cari ödeme planı sorgusu başarısız ({Adet} cari)", sorulacak.Count);
        }

        return sonuc;
    }

    public void CacheTemizle() => _cache.Clear();
}
