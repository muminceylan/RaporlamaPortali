using System.Text.Json;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Mail ayarlarını dosyadan okuyup yazan servis
/// Ayarlar mail_ayarlari.json dosyasında saklanır
/// </summary>
public class MailAyarlariService
{
    private readonly string _dosyaYolu;
    private readonly ILogger<MailAyarlariService> _logger;
    private MailAyarlariModel _ayarlar;

    public MailAyarlariService(ILogger<MailAyarlariService> logger)
    {
        _logger = logger;
        _dosyaYolu = Path.Combine(AppContext.BaseDirectory, "mail_ayarlari.json");
        _ayarlar = YukleVeyaOlustur();
    }

    /// <summary>
    /// Mevcut ayarları döndürür
    /// </summary>
    public MailAyarlariModel GetAyarlar()
    {
        return _ayarlar;
    }

    /// <summary>
    /// Ayarları kaydeder
    /// </summary>
    public async Task<bool> KaydetAsync(MailAyarlariModel ayarlar)
    {
        try
        {
            _ayarlar = ayarlar;
            
            var json = JsonSerializer.Serialize(ayarlar, new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });

            await File.WriteAllTextAsync(_dosyaYolu, json);
            _logger.LogInformation("Mail ayarları kaydedildi: {DosyaYolu}", _dosyaYolu);
            return true;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Mail ayarları kaydedilemedi");
            return false;
        }
    }

    /// <summary>
    /// Alıcı ekler
    /// </summary>
    public async Task<bool> AliciEkleAsync(string email)
    {
        if (string.IsNullOrWhiteSpace(email) || !email.Contains('@'))
            return false;

        email = email.Trim().ToLowerInvariant();
        
        if (!_ayarlar.Alicilar.Contains(email))
        {
            _ayarlar.Alicilar.Add(email);
            return await KaydetAsync(_ayarlar);
        }
        return true;
    }

    /// <summary>
    /// Alıcı siler
    /// </summary>
    public async Task<bool> AliciSilAsync(string email)
    {
        if (_ayarlar.Alicilar.Remove(email.Trim().ToLowerInvariant()))
        {
            return await KaydetAsync(_ayarlar);
        }
        return true;
    }

    /// <summary>
    /// Otomatik gönderimi açar/kapatır
    /// </summary>
    public async Task<bool> OtomatikGonderimAyarlaAsync(bool aktif)
    {
        _ayarlar.OtomatikGonderimAktif = aktif;
        return await KaydetAsync(_ayarlar);
    }

    /// <summary>
    /// Gönderim saatini ayarlar
    /// </summary>
    public async Task<bool> GonderimSaatiAyarlaAsync(TimeSpan saat)
    {
        _ayarlar.GonderimSaati = saat;
        return await KaydetAsync(_ayarlar);
    }

    /// <summary>
    /// Dosyadan ayarları yükler veya varsayılan oluşturur
    /// </summary>
    private MailAyarlariModel YukleVeyaOlustur()
    {
        try
        {
            if (File.Exists(_dosyaYolu))
            {
                var json = File.ReadAllText(_dosyaYolu);
                var ayarlar = JsonSerializer.Deserialize<MailAyarlariModel>(json);
                if (ayarlar != null)
                {
                    _logger.LogInformation("Mail ayarları yüklendi: {DosyaYolu}", _dosyaYolu);
                    return ayarlar;
                }
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Mail ayarları yüklenemedi, varsayılan oluşturuluyor");
        }

        // Varsayılan ayarlar
        var varsayilan = new MailAyarlariModel
        {
            SmtpServer = "smtp.office365.com",
            SmtpPort = 587,
            UseSsl = true,
            GondericiMail = "muminceylan@doguscay.com.tr",
            GondericiAdi = "Muhasebe Raporları",
            Sifre = "",
            Alicilar = new List<string>(),
            GonderimSaati = new TimeSpan(12, 0, 0),
            OtomatikGonderimAktif = false
        };

        // Dosyaya kaydet
        try
        {
            var json = JsonSerializer.Serialize(varsayilan, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_dosyaYolu, json);
        }
        catch { }

        return varsayilan;
    }
}

/// <summary>
/// Mail ayarları modeli
/// </summary>
public class MailAyarlariModel
{
    public string SmtpServer { get; set; } = "smtp.office365.com";
    public int SmtpPort { get; set; } = 587;
    public bool UseSsl { get; set; } = true;
    public string GondericiMail { get; set; } = "";
    public string GondericiAdi { get; set; } = "Muhasebe Raporları";
    public string Sifre { get; set; } = "";
    public List<string> Alicilar { get; set; } = new();
    public TimeSpan GonderimSaati { get; set; } = new TimeSpan(12, 0, 0);
    public bool OtomatikGonderimAktif { get; set; } = false;
    public DateTime? SonGonderimZamani { get; set; }
}
