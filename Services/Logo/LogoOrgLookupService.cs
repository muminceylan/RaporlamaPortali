using Dapper;

namespace RaporlamaPortali.Services.Logo;

// Fabrika (L_CAPIFIRM / L_CAPIDIV), Ambar (L_CAPIWHOUSE), Masraf Merkezi (LG_211_EMCENTER)
// dropdown'ları için lookup. Sonuçlar singleton cache'lenir, çünkü bu listeler nadiren değişir.
public class LogoOrgLookupService
{
    private readonly DatabaseService _db;
    private readonly ILogger<LogoOrgLookupService> _log;

    private List<KodAd>? _fabrikalar;
    private List<KodAd>? _ambarlar;
    private List<KodAd>? _masrafMerkezleri;
    private readonly SemaphoreSlim _gate = new(1, 1);

    public LogoOrgLookupService(DatabaseService db, ILogger<LogoOrgLookupService> log)
    {
        _db  = db;
        _log = log;
    }

    public record KodAd(string Kod, string Ad)
    {
        public string Etiket => string.IsNullOrWhiteSpace(Ad) ? Kod : $"{Kod} — {Ad}";
    }

    // Logo'da gerçek "Fabrika" listesi L_CAPIFACTORY tablosundadır (firm bağımsız, FIRMNR ile filtrelenir).
    // Kolonlar: NR (smallint), NAME (nvarchar). Doğuş Çay (FIRMNR=211) için NR=17 AFYON ŞEKER ÜRETİM.
    // L_CAPIDIV "işyeri/şube" listesidir, fabrika değil.
    public async Task<List<KodAd>> FabrikalarAsync()
    {
        if (_fabrikalar != null) return _fabrikalar;
        await _gate.WaitAsync();
        try
        {
            if (_fabrikalar != null) return _fabrikalar;

            var liste = new List<KodAd>();
            try
            {
                await using var con = _db.CreateConnection();
                await con.OpenAsync();

                var rows = await con.QueryAsync<(short Nr, string? Name)>(@"
                    SELECT NR, NAME FROM L_CAPIFACTORY WITH (NOLOCK)
                    WHERE FIRMNR = @firm
                    ORDER BY NR", new { firm = _db.FirmaNo });

                liste = rows.Select(r => new KodAd(r.Nr.ToString(), (r.Name ?? "").Trim())).ToList();
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Fabrika lookup başarısız, varsayılan 17 kullanılacak");
            }

            if (liste.Count == 0)
                liste.Add(new KodAd("17", "Afyon Şeker Fabrikası (default)"));

            _fabrikalar = liste;
            return _fabrikalar;
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task<List<KodAd>> AmbarlarAsync()
    {
        if (_ambarlar != null) return _ambarlar;
        await _gate.WaitAsync();
        try
        {
            if (_ambarlar != null) return _ambarlar;

            var liste = new List<KodAd>();
            try
            {
                await using var con = _db.CreateConnection();
                await con.OpenAsync();

                var rows = await con.QueryAsync<(short Nr, string? Name)>(@"
                    SELECT NR, NAME FROM L_CAPIWHOUSE WITH (NOLOCK)
                    WHERE FIRMNR = @firm AND NR >= 0
                    ORDER BY NR", new { firm = _db.FirmaNo });

                liste = rows.Select(r => new KodAd(r.Nr.ToString(), (r.Name ?? "").Trim())).ToList();
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Ambar lookup başarısız, varsayılan 147 kullanılacak");
            }

            if (liste.Count == 0)
                liste.Add(new KodAd("147", "Hammadde Ambarı (default)"));

            _ambarlar = liste;
            return _ambarlar;
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task<List<KodAd>> MasrafMerkezleriAsync()
    {
        if (_masrafMerkezleri != null) return _masrafMerkezleri;
        await _gate.WaitAsync();
        try
        {
            if (_masrafMerkezleri != null) return _masrafMerkezleri;

            var liste = new List<KodAd>();
            try
            {
                var tbl = _db.GetTableName("EMCENTER");
                await using var con = _db.CreateConnection();
                await con.OpenAsync();

                var rows = await con.QueryAsync<(string? Code, string? Defn)>($@"
                    SELECT CODE, DEFINITION_ FROM {tbl} WITH (NOLOCK)
                    ORDER BY CODE");

                liste = rows.Select(r => new KodAd((r.Code ?? "").Trim(), (r.Defn ?? "").Trim()))
                            .Where(x => !string.IsNullOrEmpty(x.Kod))
                            .ToList();
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Masraf merkezi lookup başarısız, varsayılan 7.04 kullanılacak");
            }

            if (liste.Count == 0)
                liste.Add(new KodAd("7.04", "Genel (default)"));

            _masrafMerkezleri = liste;
            return _masrafMerkezleri;
        }
        finally
        {
            _gate.Release();
        }
    }

    public void CacheTemizle()
    {
        _fabrikalar = null;
        _ambarlar = null;
        _masrafMerkezleri = null;
    }
}
