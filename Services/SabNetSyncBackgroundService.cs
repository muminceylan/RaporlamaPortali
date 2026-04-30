using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace RaporlamaPortali.Services;

/// <summary>
/// SabNetKANTAR'dan SabNet.db'ye 60 saniyede bir inkremental senkronizasyon yapar.
/// Sadece bugün/açık trip'lerin kayıtlarını ve yeni log satırlarını çeker.
/// </summary>
public class SabNetSyncBackgroundService : BackgroundService
{
    private readonly IServiceProvider _sp;
    private readonly SabNetSyncDurum _durum;
    private readonly ILogger<SabNetSyncBackgroundService> _log;

    private static readonly TimeSpan _interval = TimeSpan.FromSeconds(60);

    public SabNetSyncBackgroundService(IServiceProvider sp, SabNetSyncDurum durum, ILogger<SabNetSyncBackgroundService> log)
    {
        _sp = sp;
        _durum = durum;
        _log = log;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        // İlk çalıştırmadan önce kısa bekleme (uygulama açılışında DB hazır olsun)
        try { await Task.Delay(TimeSpan.FromSeconds(15), stoppingToken); }
        catch (OperationCanceledException) { return; }

        while (!stoppingToken.IsCancellationRequested)
        {
            if (_durum.OtomatikSyncAcik)
            {
                await SyncCalistirAsync(stoppingToken);
            }

            try { await Task.Delay(_interval, stoppingToken); }
            catch (OperationCanceledException) { break; }
        }
    }

    private async Task SyncCalistirAsync(CancellationToken ct)
    {
        if (!await _durum.Kilit.WaitAsync(0, ct)) return;
        try
        {
            using var scope = _sp.CreateScope();
            var import = scope.ServiceProvider.GetRequiredService<SabNetImportService>();

            var (silinen, eklenenH, sn1) = await import.HareketleriGuncelleAsync(gunGeriye: 1, ct);
            var (eklenenL, sn2) = await import.LoglariGuncelleAsync(ct);

            _durum.SonSyncZamani = DateTime.Now;
            _durum.SonSyncEklenenHareket = eklenenH;
            _durum.SonSyncEklenenLog = eklenenL;
            _durum.SonSyncSureSn = sn1 + sn2;
            _durum.SonHata = null;
            _durum.SonHataZamani = null;
            _durum.Bildir();

            _log.LogInformation("SabNet sync OK: hareket sil={Sil} ekle={EklH} log ekle={EklL} sure={Sn:F1}sn",
                silinen, eklenenH, eklenenL, sn1 + sn2);
        }
        catch (OperationCanceledException) { }
        catch (Exception ex)
        {
            _durum.SonHata = ex.Message;
            _durum.SonHataZamani = DateTime.Now;
            _durum.Bildir();
            _log.LogError(ex, "SabNet sync hata");
        }
        finally
        {
            _durum.Kilit.Release();
        }
    }
}
