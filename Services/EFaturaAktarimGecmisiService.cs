using Dapper;
using Microsoft.Data.Sqlite;

namespace RaporlamaPortali.Services;

// Logo'ya başarıyla aktarılan e-Faturaların yerel kaydı.
// SQLite: AppDataPaths.EFaturaTotalsCacheDb içinde aktarim_gecmisi tablosu.
//
// Bu kayıtlar "E-Fatura Aktar" menüsünde aynı faturanın tekrar aktarımını engellemek için
// kullanılır. Kullanıcı "İşaret Kaldır" diyerek kaydı silebilir (faturayı Logo'dan çıkarıp
// tekrar aktarmak isterse).
public class EFaturaAktarimGecmisiService
{
    private readonly ILogger<EFaturaAktarimGecmisiService> _log;
    private static readonly SemaphoreSlim _initLock = new(1, 1);
    private static bool _ready;

    public EFaturaAktarimGecmisiService(ILogger<EFaturaAktarimGecmisiService> log) => _log = log;

    public class Kayit
    {
        public long      EFaturaId      { get; set; }   // POSTBOX.ID
        public string    FaturaNo       { get; set; } = "";
        public string    LogoNumber     { get; set; } = "";  // Logo'da Fiş No (genelde = e-Fatura no)
        public long      LogoRef        { get; set; }        // LG_211_01_INVOICE.LOGICALREF
        public int       FaturaTipi     { get; set; }        // 1=Satın Alma, 4=Hizmet
        public DateTime  AktarimTarihi  { get; set; }
    }

    private async Task EnsureAsync()
    {
        if (_ready) return;
        await _initLock.WaitAsync();
        try
        {
            if (_ready) return;
            Directory.CreateDirectory(AppDataPaths.DataRoot);
            await using var con = new SqliteConnection($"Data Source={AppDataPaths.EFaturaTotalsCacheDb}");
            await con.OpenAsync();
            await con.ExecuteAsync(@"
                CREATE TABLE IF NOT EXISTS aktarim_gecmisi (
                    efatura_id      INTEGER NOT NULL PRIMARY KEY,
                    fatura_no       TEXT    NOT NULL DEFAULT '',
                    logo_number     TEXT    NOT NULL DEFAULT '',
                    logo_ref        INTEGER NOT NULL DEFAULT 0,
                    fatura_tipi     INTEGER NOT NULL DEFAULT 0,
                    aktarim_tarihi  TEXT    NOT NULL
                );
                CREATE INDEX IF NOT EXISTS IX_aktarim_gecmisi_fatura_no ON aktarim_gecmisi(fatura_no);");
            _ready = true;
        }
        finally { _initLock.Release(); }
    }

    // Başarılı aktarım sonrası kayıt. Var olan EFaturaId üzerine yazılır (REPLACE).
    public async Task KaydetAsync(Kayit k, CancellationToken ct = default)
    {
        await EnsureAsync();
        await using var con = new SqliteConnection($"Data Source={AppDataPaths.EFaturaTotalsCacheDb}");
        await con.OpenAsync(ct);
        await con.ExecuteAsync(new CommandDefinition(@"
            INSERT INTO aktarim_gecmisi (efatura_id, fatura_no, logo_number, logo_ref, fatura_tipi, aktarim_tarihi)
            VALUES (@EFaturaId, @FaturaNo, @LogoNumber, @LogoRef, @FaturaTipi, @AktarimTarihi)
            ON CONFLICT(efatura_id) DO UPDATE SET
                fatura_no = excluded.fatura_no,
                logo_number = excluded.logo_number,
                logo_ref = excluded.logo_ref,
                fatura_tipi = excluded.fatura_tipi,
                aktarim_tarihi = excluded.aktarim_tarihi",
            new
            {
                k.EFaturaId, k.FaturaNo, k.LogoNumber, k.LogoRef, k.FaturaTipi,
                AktarimTarihi = k.AktarimTarihi.ToString("yyyy-MM-dd HH:mm:ss"),
            },
            cancellationToken: ct));
    }

    // Belirli EFaturaId'ler için aktarım kayıtları (toplu).
    public async Task<Dictionary<long, Kayit>> ListeAlAsync(IEnumerable<long> efaturaIds, CancellationToken ct = default)
    {
        await EnsureAsync();
        var ids = efaturaIds.Distinct().ToArray();
        var sonuc = new Dictionary<long, Kayit>();
        if (ids.Length == 0) return sonuc;

        await using var con = new SqliteConnection($"Data Source={AppDataPaths.EFaturaTotalsCacheDb}");
        await con.OpenAsync(ct);

        const int chunk = 500;
        for (int o = 0; o < ids.Length; o += chunk)
        {
            var slice = ids.Skip(o).Take(chunk).ToArray();
            var rows = await con.QueryAsync(new CommandDefinition(@"
                SELECT efatura_id AS EFaturaId, fatura_no AS FaturaNo, logo_number AS LogoNumber,
                       logo_ref AS LogoRef, fatura_tipi AS FaturaTipi, aktarim_tarihi AS AktarimTarihiStr
                FROM aktarim_gecmisi WHERE efatura_id IN @ids",
                new { ids = slice }, cancellationToken: ct));
            foreach (var r in rows)
            {
                var kayit = new Kayit
                {
                    EFaturaId  = Convert.ToInt64(r.EFaturaId),
                    FaturaNo   = r.FaturaNo ?? "",
                    LogoNumber = r.LogoNumber ?? "",
                    LogoRef    = Convert.ToInt64(r.LogoRef ?? 0L),
                    FaturaTipi = Convert.ToInt32(r.FaturaTipi ?? 0),
                };
                if (DateTime.TryParse((string?)r.AktarimTarihiStr, out var dt))
                    kayit.AktarimTarihi = dt;
                sonuc[kayit.EFaturaId] = kayit;
            }
        }
        return sonuc;
    }

    // Tek satır lookup — varsa Kayit, yoksa null.
    public async Task<Kayit?> BulAsync(long efaturaId, CancellationToken ct = default)
    {
        var map = await ListeAlAsync(new[] { efaturaId }, ct);
        return map.TryGetValue(efaturaId, out var k) ? k : null;
    }

    // Kullanıcı "tekrar aktarmak istiyorum" deyince — kaydı sil.
    public async Task IsaretiKaldirAsync(long efaturaId, CancellationToken ct = default)
    {
        await EnsureAsync();
        await using var con = new SqliteConnection($"Data Source={AppDataPaths.EFaturaTotalsCacheDb}");
        await con.OpenAsync(ct);
        await con.ExecuteAsync(new CommandDefinition(
            "DELETE FROM aktarim_gecmisi WHERE efatura_id = @id",
            new { id = efaturaId }, cancellationToken: ct));
        _log.LogInformation("Aktarım işareti kaldırıldı: EFaturaId={Id}", efaturaId);
    }
}
