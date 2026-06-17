using System.Collections.Concurrent;
using Dapper;

namespace RaporlamaPortali.Services.Logo;

// Logo malzeme/hizmet kartında tanımlı "Muhasebe Hesapları" sekmesindeki kayıtlardan
// (LG_{firma}_CRDACREF) doğrudan hesap kodunu çeker.
//
// CRDACREF yapısı (Card-Account Reference):
//   - CARDREF    : malzeme kartının LOGICALREF'i (ITEMS.LOGICALREF veya SRVCARD.LOGICALREF)
//   - TRCODE     : 1 = malzeme/hizmet kartı bağlantıları (bizim ilgilendiğimiz), 2 = sabit kıymet, 5 = cari
//   - TYP        : hesap tipi — 1=Alımlar, 3=Satışlar, 5=Sarflar, 10..15=İadeler, 95=Sayım Fazla, 96=Sayım Noksan, 99/136/162=Özel
//   - ACCOUNTREF : muhasebe hesabının LOGICALREF'i (→ EMUHACC.LOGICALREF)
//   - CENTERREF  : opsiyonel masraf merkezi
//
// Aktarım için kullandığımız:
//   - Satın Alma Faturası satırı GL_CODE1 = kartın TYP=1 (Alımlar Hesabı) kaydı
//   - Hizmet kartlarında da aynı tabloda yapı (CARDREF=SRVCARD.LOGICALREF)
//
// Cache: per (firma, faturaTipi, typ, masterCode). Aynı malzeme/hesap tipi kombinasyonu için
// uygulama yaşam süresi boyunca tek sorgu. Hesap değişirse `CacheTemizle()` çağrılır.
public class LogoAccCodesLookupService
{
    private readonly DatabaseService _db;
    private readonly ILogger<LogoAccCodesLookupService> _log;

    private static readonly ConcurrentDictionary<string, string> _cache = new();

    // CRDACREF.TYP değerleri (Logo standart)
    public const int TYP_ALIM        = 1;
    public const int TYP_SATIS       = 3;
    public const int TYP_SATIN_IADE  = 11;
    public const int TYP_SATIS_IADE  = 12;

    public LogoAccCodesLookupService(DatabaseService db, ILogger<LogoAccCodesLookupService> log)
    {
        _db = db;
        _log = log;
    }

    // Satın Alma faturası için "Alımlar Hesabı" (GL_CODE1). faturaTipi: 1=Mal (ITEMS), 4=Hizmet (SRVCARD).
    public Task<string> AlimHesabiAsync(string? masterCode, int faturaTipi, CancellationToken ct = default)
        => KartHesabiAsync(masterCode, faturaTipi, TYP_ALIM, ct);

    // Genel — kart üzerinde herhangi bir TYP'nin hesap kodu.
    public async Task<string> KartHesabiAsync(string? masterCode, int faturaTipi, int typ, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(masterCode)) return "";
        var cacheKey = $"{_db.FirmaNo}:{faturaTipi}:{typ}:{masterCode}";
        if (_cache.TryGetValue(cacheKey, out var cached)) return cached;

        try
        {
            // Logo CRDACREF.TRCODE → kart tipi:
            //   1 = ITEMS (Malzeme), 2 = Sabit Kıymet, 3 = SRVCARD (Hizmet), 5 = CLCARD (Cari)
            // Hizmet faturası (Type=4) → SRVCARD, TRCODE=3. Mal (Type=1) → ITEMS, TRCODE=1.
            var kartTbl  = faturaTipi == 4 ? "SRVCARD" : "ITEMS";
            var trCode   = faturaTipi == 4 ? 3 : 1;
            var sql = $@"
                SELECT TOP 1 e.CODE
                FROM LG_{_db.FirmaNo}_CRDACREF c WITH(NOLOCK)
                INNER JOIN LG_{_db.FirmaNo}_{kartTbl} k WITH(NOLOCK) ON k.LOGICALREF = c.CARDREF
                INNER JOIN LG_{_db.FirmaNo}_EMUHACC e WITH(NOLOCK) ON e.LOGICALREF = c.ACCOUNTREF
                WHERE k.CODE = @code AND c.TRCODE = @trcode AND c.TYP = @typ
                ORDER BY c.LOGICALREF DESC";
            using var conn = _db.CreateConnection();
            var kod = await conn.QueryFirstOrDefaultAsync<string?>(
                new CommandDefinition(sql, new { code = masterCode, trcode = trCode, typ }, cancellationToken: ct));
            kod = (kod ?? "").Trim();
            _cache[cacheKey] = kod;
            if (!string.IsNullOrEmpty(kod))
                _log.LogDebug("CRDACREF: {Code} (faturaTipi={Ft}, TYP={Typ}) → {Hesap}", masterCode, faturaTipi, typ, kod);
            return kod;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "CRDACREF lookup başarısız (masterCode={M}, faturaTipi={Ft}, typ={T})", masterCode, faturaTipi, typ);
            return "";
        }
    }

    /// <summary>
    /// CARİ kartının muhasebe hesap kodu — CRDACREF.TRCODE=5 (CLCARD bağlantısı).
    /// XML'deki GL_CODE alanı için kullanılır (örn 120.00.02.0001).
    /// Cari kartlarında birden fazla TYP olabilir — ilk bulunanı döndürür.
    /// </summary>
    public async Task<string> CariMuhasebeKoduAsync(string? cariKodu, CancellationToken ct = default)
    {
        if (string.IsNullOrWhiteSpace(cariKodu)) return "";
        var cacheKey = $"{_db.FirmaNo}:CARI:5:{cariKodu}";
        if (_cache.TryGetValue(cacheKey, out var cached)) return cached;

        try
        {
            var sql = $@"
                SELECT TOP 1 e.CODE
                FROM LG_{_db.FirmaNo}_CRDACREF c WITH(NOLOCK)
                INNER JOIN LG_{_db.FirmaNo}_CLCARD k WITH(NOLOCK) ON k.LOGICALREF = c.CARDREF
                INNER JOIN LG_{_db.FirmaNo}_EMUHACC e WITH(NOLOCK) ON e.LOGICALREF = c.ACCOUNTREF
                WHERE k.CODE = @code AND c.TRCODE = 5
                ORDER BY c.LOGICALREF DESC";
            using var conn = _db.CreateConnection();
            var kod = await conn.QueryFirstOrDefaultAsync<string?>(
                new CommandDefinition(sql, new { code = cariKodu }, cancellationToken: ct));
            kod = (kod ?? "").Trim();
            _cache[cacheKey] = kod;
            if (!string.IsNullOrEmpty(kod))
                _log.LogDebug("CRDACREF (cari): {Code} → {Hesap}", cariKodu, kod);
            return kod;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Cari muhasebe kodu lookup başarısız (cariKodu={C})", cariKodu);
            return "";
        }
    }

    public void CacheTemizle() => _cache.Clear();
}
