using System.Collections.Concurrent;
using System.Text.RegularExpressions;
using Dapper;

namespace RaporlamaPortali.Services.Logo;

// Logo muhasebe hesap planından (LG_{firma}_EMUHACC) KDV oranına göre uygun hesap kodunu
// DİNAMİK olarak çeker. Hesap kodlarında son ek formatı sınıftan sınıfa değişiyor
// (191.01.07 = %20 normal vs 192.02.004 = %20 tevkifatlı) — hardcoded map kırılgan.
// Bu servis, ilk çağrıda Logo hesap planını okuyup DEFINITION_ içindeki "%N" pattern'iyle
// oran ↔ kod eşlemesi kurar; sonraki çağrılar cache'ten döner.
//
// Hesap planı değişirse uygulama restart veya CacheTemizle() çağrısı gerekir.
public class LogoKdvHesapLookupService
{
    private readonly DatabaseService _db;
    private readonly ILogger<LogoKdvHesapLookupService> _log;

    // Cache: (firma, tevkifatli) → (oran → hesap kodu)
    private static readonly ConcurrentDictionary<string, IReadOnlyDictionary<decimal, string>> _cache = new();

    // Sınıf prefix'leri. Tevkifatsız = 191.01.* (İNDİRİLECEK KDV), tevkifatlı = 192.02.*
    // (TEVKİFATLI İNDİRİLECEK KDV). Kullanıcı hesap planını farklı düzenlerse buradan tek
    // yerden değişir.
    private const string PREFIX_NORMAL    = "191.01.";
    private const string PREFIX_TEVKIFATLI = "192.02.";

    // DEFINITION_ içinden "%N" veya "%N,M" oranını çeker. "%7,27" → 7.27, "%20" → 20.
    private static readonly Regex _oranRgx = new(@"%(\d+(?:[.,]\d+)?)", RegexOptions.Compiled);

    public LogoKdvHesapLookupService(DatabaseService db, ILogger<LogoKdvHesapLookupService> log)
    {
        _db  = db;
        _log = log;
    }

    /// <summary>
    /// KDV oranına göre Logo "İndirilecek KDV" hesap kodunu döner.
    /// Tevkifatlı ise 192.02.* sınıfından, değilse 191.01.* sınıfından.
    /// DB'den bulunamazsa LogoHesapKoduMaps fallback'i kullanılır.
    /// </summary>
    public async Task<string> KdvHesabiBulAsync(decimal kdvOran, bool tevkifatli, CancellationToken ct = default)
    {
        if (kdvOran <= 0m) return "";

        var map = await MapGetirAsync(tevkifatli, ct);
        if (map.TryGetValue(kdvOran, out var kod) && !string.IsNullOrEmpty(kod))
            return kod;

        // DB'de yoksa hardcoded fallback
        var fallback = LogoHesapKoduMaps.KdvHesabiBul(kdvOran, tevkifatli);
        if (!string.IsNullOrEmpty(fallback))
            _log.LogDebug("KDV {Oran}% ({Tip}) Logo'da bulunamadı → hardcoded fallback: {Kod}",
                kdvOran, tevkifatli ? "tevkifatlı" : "normal", fallback);
        return fallback;
    }

    /// <summary>Cache'i temizler — hesap planı değişikliğinden sonra çağrılabilir.</summary>
    public void CacheTemizle() => _cache.Clear();

    /// <summary>
    /// Belirli sınıf için (oran → hesap kodu) map'ini döner. İlk çağrıda DB'den çeker, cache'ler.
    /// </summary>
    private async Task<IReadOnlyDictionary<decimal, string>> MapGetirAsync(bool tevkifatli, CancellationToken ct)
    {
        var cacheKey = $"{_db.FirmaNo}:{(tevkifatli ? "T" : "N")}";
        if (_cache.TryGetValue(cacheKey, out var cached)) return cached;

        var prefix = tevkifatli ? PREFIX_TEVKIFATLI : PREFIX_NORMAL;
        var map = new Dictionary<decimal, string>();
        try
        {
            var sql = $@"SELECT CODE, DEFINITION_
                         FROM LG_{_db.FirmaNo}_EMUHACC WITH(NOLOCK)
                         WHERE CODE LIKE @p
                         ORDER BY CODE";
            using var conn = _db.CreateConnection();
            var rows = await conn.QueryAsync<(string Code, string Def)>(
                new CommandDefinition(sql, new { p = prefix + "%" }, cancellationToken: ct));

            foreach (var r in rows)
            {
                var code = (r.Code ?? "").Trim();
                var def  = (r.Def  ?? "").Trim();
                if (string.IsNullOrEmpty(code)) continue;

                // Yalnızca yaprak hesapları (en az 4 segment: 191.01.NN veya 192.02.NNN) topla.
                // Sınıf başlığını (191.01) ve ara seviyeleri (191.02 - SATIŞTAN İADELER) ele.
                var trail = code.Substring(prefix.Length);
                if (string.IsNullOrEmpty(trail) || trail.Contains('.')) continue;

                var m = _oranRgx.Match(def);
                if (!m.Success) continue;

                var ham = m.Groups[1].Value.Replace(',', '.');
                if (!decimal.TryParse(ham, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out var oran)) continue;

                // Aynı oran için birden fazla kod varsa İLKİ kazanır (CODE ASC).
                // (Logo bazen alt sınıflar açıyor: 192.02.005 = "İNDİİRLECEK KDV %1" — bu da %1.
                // ORDER BY CODE ile asıl tanımlanmış olan .001 önce gelir.)
                if (!map.ContainsKey(oran))
                    map[oran] = code;
            }

            _log.LogInformation("Logo KDV hesap planı yüklendi: {Tip} {Sayi} oran ({Kodlar})",
                tevkifatli ? "tevkifatlı" : "normal", map.Count,
                string.Join(", ", map.Select(kv => $"%{kv.Key}→{kv.Value}")));
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Logo KDV hesap lookup başarısız (prefix={P}) — hardcoded fallback kullanılacak", prefix);
        }

        _cache[cacheKey] = map;
        return map;
    }
}
