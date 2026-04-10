using System.Net;
using System.Net.Mail;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Outlook/Office 365 üzerinden mail gönderim servisi
/// Tüm alıcılar BCC (gizli) olarak eklenir
/// </summary>
public class MailGonderimService
{
    private readonly MailAyarlari _ayarlar;

    public MailGonderimService(MailAyarlari ayarlar)
    {
        _ayarlar = ayarlar;
    }

    /// <summary>
    /// HTML içerikli mail gönderir
    /// Tüm alıcılar BCC olarak eklenir (gizli)
    /// </summary>
    public async Task<MailSonuc> GonderAsync(string konu, string htmlIcerik, List<string>? ekAlicilar = null)
    {
        var sonuc = new MailSonuc();

        try
        {
            using var smtp = new SmtpClient(_ayarlar.SmtpServer, _ayarlar.SmtpPort)
            {
                EnableSsl = _ayarlar.UseSsl,
                Credentials = new NetworkCredential(_ayarlar.GondericiMail, _ayarlar.Sifre),
                DeliveryMethod = SmtpDeliveryMethod.Network,
                Timeout = 30000 // 30 saniye
            };

            using var mail = new MailMessage
            {
                From = new MailAddress(_ayarlar.GondericiMail, _ayarlar.GondericiAdi),
                Subject = konu,
                Body = htmlIcerik,
                IsBodyHtml = true,
                Priority = MailPriority.Normal
            };

            // Kendine To olarak ekle (BCC'nin çalışması için en az 1 To gerekli)
            mail.To.Add(_ayarlar.GondericiMail);

            // Tüm alıcıları BCC (gizli) olarak ekle
            var tumAlicilar = _ayarlar.Alicilar.ToList();
            if (ekAlicilar != null)
                tumAlicilar.AddRange(ekAlicilar);

            foreach (var alici in tumAlicilar.Distinct())
            {
                if (!string.IsNullOrWhiteSpace(alici) && alici.Contains('@'))
                {
                    mail.Bcc.Add(alici.Trim());
                }
            }

            await smtp.SendMailAsync(mail);

            sonuc.Basarili = true;
            sonuc.Mesaj = "Mail başarıyla gönderildi.";
            sonuc.AliciSayisi = mail.Bcc.Count;
            sonuc.GonderimZamani = DateTime.Now;

            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ✅ Mail gönderildi - {sonuc.AliciSayisi} alıcı (BCC)");
        }
        catch (SmtpException ex)
        {
            sonuc.Basarili = false;
            sonuc.Mesaj = $"SMTP Hatası: {ex.Message}";
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ❌ SMTP Hatası: {ex.Message}");
        }
        catch (Exception ex)
        {
            sonuc.Basarili = false;
            sonuc.Mesaj = $"Hata: {ex.Message}";
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] ❌ Hata: {ex.Message}");
        }

        return sonuc;
    }

    /// <summary>
    /// Mail ayarlarını test eder
    /// </summary>
    public async Task<MailSonuc> TestGonderAsync()
    {
        var testHtml = $@"
<html>
<body style='font-family: Segoe UI, sans-serif;'>
    <h2 style='color: #059669;'>🧪 Mail Test</h2>
    <p>Bu bir test mailidir.</p>
    <p><strong>Gönderim Zamanı:</strong> {DateTime.Now:dd.MM.yyyy HH:mm:ss}</p>
    <p><strong>Sunucu:</strong> {_ayarlar.SmtpServer}:{_ayarlar.SmtpPort}</p>
    <hr>
    <p style='color: #666; font-size: 12px;'>RaporlamaPortali - Otomatik Mail Sistemi</p>
</body>
</html>";

        return await GonderAsync("🧪 Mail Test - RaporlamaPortali", testHtml);
    }
}
