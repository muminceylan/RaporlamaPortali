using Dapper;

namespace RaporlamaPortali.Services.Logo;

/// <summary>
/// Malzeme kartında tanımlı "Satınalma/Satış Fiyatları" listesinden satış fiyatını çeker.
/// LG_{firma}_PRCLIST tablosu:
///   PTYPE=2 → Satış (Logo Tiger 3 v3.0 standardı — empirically verified)
///   CARDREF → ITEMS.LOGICALREF
///   PRICE   → birim fiyat
///   CLIENTCODE → cariye özel fiyat (boşsa GENEL fiyattır)
///   CLSPECODE  → cari özel kod (cari grubuna özel; boşsa GENEL)
///   CAPIBLOCK_CREADEDDATE (Logo'nun standart typo'lu kolon adı) → kayıt tarihi
///   ACTIVE=0 → aktif kayıt
///
/// SEÇİM KURALI (kullanıcı talebi):
///   - CLIENTCODE BOŞ + CLSPECODE BOŞ (cari tanımı olmayan / "genel" fiyat satırları)
///   - bu satırlar arasında KAYIT TARİHİ EN SON OLAN
///   - cariye/cari grubuna özel fiyat satırı varsa bile kullanılmaz
/// </summary>
public class LogoFiyatListesiLookupService
{
    private readonly DatabaseService _db;
    private readonly ILogger<LogoFiyatListesiLookupService> _log;

    public LogoFiyatListesiLookupService(DatabaseService db, ILogger<LogoFiyatListesiLookupService> log)
    {
        _db = db;
        _log = log;
    }

    /// <summary>
    /// Verilen malzeme kodları için, her birinin en yeni satış fiyat liste satırından
    /// PRICE değerini döndürür. Bulunamayan kodlar dictionary'ye eklenmez (0 değil).
    /// </summary>
    public async Task<Dictionary<string, decimal>> MalzemeSatisFiyatlariAsync(
        IEnumerable<string> malzemeKodlari, CancellationToken ct = default)
    {
        var kodlar = malzemeKodlari.Where(k => !string.IsNullOrWhiteSpace(k))
                                   .Select(k => k.Trim())
                                   .Distinct()
                                   .ToList();
        var sonuc = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        if (kodlar.Count == 0) return sonuc;

        try
        {
            // SADECE genel (cari tanımı olmayan) satırlar — CLIENTCODE ve CLSPECODE boş.
            // Bunlar içinde her malzeme için kayıt tarihi en yeni satırın PRICE'ı seçilir.
            var sql = $@"
                SELECT MALZEME_KODU, PRICE FROM (
                    SELECT i.CODE AS MALZEME_KODU,
                           p.PRICE,
                           ROW_NUMBER() OVER (PARTITION BY i.LOGICALREF
                                              ORDER BY p.CAPIBLOCK_CREADEDDATE DESC, p.LOGICALREF DESC) AS RN
                    FROM LG_{_db.FirmaNo}_PRCLIST p WITH(NOLOCK)
                    INNER JOIN LG_{_db.FirmaNo}_ITEMS i WITH(NOLOCK) ON i.LOGICALREF = p.CARDREF
                    WHERE p.PTYPE  = 2     -- Satış fiyatı
                      AND p.ACTIVE = 0     -- 0=aktif, 1=pasif
                      AND (p.CLIENTCODE IS NULL OR LTRIM(RTRIM(p.CLIENTCODE)) = '')
                      AND (p.CLSPECODE  IS NULL OR LTRIM(RTRIM(p.CLSPECODE))  = '')
                      AND i.CODE IN @kodlar
                ) t WHERE RN = 1";

            using var conn = _db.CreateConnection();
            var satirlar = await conn.QueryAsync<(string MalzemeKodu, decimal Price)>(
                new CommandDefinition(sql, new { kodlar }, cancellationToken: ct));

            foreach (var s in satirlar)
                if (!string.IsNullOrWhiteSpace(s.MalzemeKodu))
                    sonuc[s.MalzemeKodu.Trim()] = s.Price;

            _log.LogInformation("Fiyat lookup: {Soruldu} kod soruldu, {Bulundu} bulundu",
                kodlar.Count, sonuc.Count);
            return sonuc;
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Malzeme satış fiyatları sorgusu başarısız");
            return sonuc;
        }
    }
}
