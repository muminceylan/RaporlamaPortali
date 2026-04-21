using System.Diagnostics;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// WhatsApp Node.js botunu arka planda yöneten servis.
/// Uygulama başladığında botu otomatik başlatır, kapanınca durdurur.
/// </summary>
public class WhatsAppProcessService : BackgroundService
{
    private readonly ILogger<WhatsAppProcessService> _logger;
    private readonly WhatsAppAyarlariService         _ayarlarService;
    private Process? _nodeProcess;
    private bool     _kullaniciDurdurdu = false;

    public WhatsAppProcessService(
        ILogger<WhatsAppProcessService> logger,
        WhatsAppAyarlariService ayarlarService)
    {
        _logger         = logger;
        _ayarlarService = ayarlarService;
    }

    // Dışarıdan okunabilir durum
    public bool  CalisiyorMu   => _nodeProcess is { HasExited: false };
    public string? SonHata     { get; private set; }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        var ayarlar = _ayarlarService.GetAyarlar();
        if (!ayarlar.OtomatikBaslat)
        {
            _logger.LogInformation("WhatsApp otomatik baslatma kapali.");
            return;
        }

        await Task.Delay(3000, stoppingToken); // Uygulamanın tam açılmasını bekle
        await BotBaslatAsync();

        // Bot durduğunda yeniden başlat (kullanıcı durdurmadıysa)
        while (!stoppingToken.IsCancellationRequested)
        {
            await Task.Delay(10000, stoppingToken);

            if (!_kullaniciDurdurdu && !CalisiyorMu && !stoppingToken.IsCancellationRequested)
            {
                _logger.LogWarning("WhatsApp botu durmus, yeniden baslatiliyor...");
                await BotBaslatAsync();
            }
        }
    }

    public async Task BotBaslatAsync()
    {
        if (CalisiyorMu) return;

        SonHata         = null;
        _kullaniciDurdurdu = false;

        var kodKlasor  = WhatsAppAyarlariService.WhatsAppKodKlasor();
        var veriKlasor = WhatsAppAyarlariService.WhatsAppVeriKlasor();
        Directory.CreateDirectory(veriKlasor);

        if (!Directory.Exists(kodKlasor))
        {
            SonHata = $"WhatsApp kod klasörü bulunamadı: {kodKlasor}";
            _logger.LogError(SonHata);
            return;
        }

        var indexJs = Path.Combine(kodKlasor, "index.js");
        if (!File.Exists(indexJs))
        {
            SonHata = "index.js bulunamadı.";
            _logger.LogError(SonHata);
            return;
        }

        // node_modules yoksa npm install çalıştır
        var nodeModules = Path.Combine(kodKlasor, "node_modules");
        if (!Directory.Exists(nodeModules))
        {
            _logger.LogInformation("node_modules bulunamadi, npm install calistiriliyor...");
            var npmBasarili = await NpmInstallAsync(kodKlasor);
            if (!npmBasarili)
            {
                SonHata = "npm install başarısız. Node.js kurulu mu?";
                _logger.LogError(SonHata);
                return;
            }
        }

        // node.exe bul
        var nodeExe = BulNodeExe();
        if (nodeExe == null)
        {
            SonHata = "node.exe bulunamadı. Node.js kurulu olduğundan emin olun.";
            _logger.LogError(SonHata);
            return;
        }

        _logger.LogInformation("WhatsApp botu baslatiliyor: {Node} {Script}", nodeExe, indexJs);

        var psi = new ProcessStartInfo
        {
            FileName               = nodeExe,
            Arguments              = $"\"{indexJs}\"",
            WorkingDirectory       = kodKlasor,
            UseShellExecute        = false,
            CreateNoWindow         = true,
            RedirectStandardOutput = true,
            RedirectStandardError  = true
        };
        // Node tarafı veri (config, session, log, screenshots) için bu klasörü kullanacak
        psi.Environment["WHATSAPP_DATA_DIR"] = veriKlasor;

        _nodeProcess = new Process
        {
            StartInfo = psi,
            EnableRaisingEvents = true
        };

        _nodeProcess.OutputDataReceived += (_, e) =>
        {
            if (!string.IsNullOrWhiteSpace(e.Data))
                _logger.LogInformation("[WhatsApp] {Mesaj}", e.Data);
        };
        _nodeProcess.ErrorDataReceived += (_, e) =>
        {
            if (!string.IsNullOrWhiteSpace(e.Data))
                _logger.LogWarning("[WhatsApp ERR] {Mesaj}", e.Data);
        };

        try
        {
            _nodeProcess.Start();
            _nodeProcess.BeginOutputReadLine();
            _nodeProcess.BeginErrorReadLine();
            _logger.LogInformation("WhatsApp botu baslatildi. PID: {Pid}", _nodeProcess.Id);
        }
        catch (Exception ex)
        {
            SonHata = "Node başlatma hatası: " + ex.Message;
            _logger.LogError(ex, SonHata);
        }
    }

    public void BotDurdur()
    {
        _kullaniciDurdurdu = true;
        if (_nodeProcess is { HasExited: false })
        {
            try
            {
                _nodeProcess.Kill(entireProcessTree: true);
                _logger.LogInformation("WhatsApp botu durduruldu.");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Bot durdurma hatasi");
            }
        }
        _nodeProcess = null;
    }

    public async Task BotYenidenBaslatAsync()
    {
        BotDurdur();
        await Task.Delay(2000);
        await BotBaslatAsync();
    }

    private async Task<bool> NpmInstallAsync(string klasor)
    {
        try
        {
            var tcs = new TaskCompletionSource<bool>();
            // node_modules'u kaynak WhatsApp klasöründen kopyala (cmd.exe devre dışıysa npm install çalışmaz)
            var kaynakNodeModules = Path.Combine(
                Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) ?? "",
                "WhatsApp", "node_modules");
            if (!Directory.Exists(kaynakNodeModules))
                kaynakNodeModules = Path.Combine(klasor, "..", "node_modules_backup");

            var npm = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName               = "powershell.exe",
                    Arguments              = $"-NoProfile -Command \"& 'C:\\Program Files\\nodejs\\npm.cmd' install\"",
                    WorkingDirectory       = klasor,
                    UseShellExecute        = false,
                    CreateNoWindow         = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError  = true
                },
                EnableRaisingEvents = true
            };
            npm.Exited += (_, _) => tcs.TrySetResult(npm.ExitCode == 0);
            npm.Start();
            npm.BeginOutputReadLine();
            npm.BeginErrorReadLine();

            // 3 dakika zaman aşımı
            using var cts = new CancellationTokenSource(TimeSpan.FromMinutes(3));
            cts.Token.Register(() => tcs.TrySetResult(false));

            return await tcs.Task;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "npm install hatasi");
            return false;
        }
    }

    private static string? BulNodeExe()
    {
        // PATH'ten bul
        foreach (var dir in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';'))
        {
            var full = Path.Combine(dir.Trim(), "node.exe");
            if (File.Exists(full)) return full;
        }
        // Yaygın kurulum yerleri
        string[] olasilar =
        [
            @"C:\Program Files\nodejs\node.exe",
            @"C:\Program Files (x86)\nodejs\node.exe",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                         @"nvm\current\node.exe")
        ];
        return olasilar.FirstOrDefault(File.Exists);
    }

    public override void Dispose()
    {
        BotDurdur();
        base.Dispose();
    }
}
