using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace RaporlamaPortali.Services;

/// <summary>
/// Arka planda çalışan zamanlanmış mail gönderim servisi
/// Outlook uygulaması üzerinden mail gönderir (şifre gerektirmez)
/// </summary>
public class ZamanliMailService : BackgroundService
{
    private readonly ILogger<ZamanliMailService> _logger;
    private readonly IServiceProvider _serviceProvider;

    // Pancar maili bugün gönderildi mi?
    private DateTime? _pancarSonGonderim = null;

    public ZamanliMailService(
        ILogger<ZamanliMailService> logger,
        IServiceProvider serviceProvider)
    {
        _logger = logger;
        _serviceProvider = serviceProvider;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _logger.LogInformation("Zamanli Mail Servisi baslatildi.");

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                using var scope = _serviceProvider.CreateScope();
                var ayarlarService = scope.ServiceProvider.GetRequiredService<MailAyarlariService>();
                var ayarlar = ayarlarService.GetAyarlar();

                var simdi = DateTime.Now;

                if (ayarlar.OtomatikGonderimAktif)
                {
                    var bugunGonderimZamani = simdi.Date.Add(ayarlar.GonderimSaati);

                    bool bugunGonderilmedi = ayarlar.SonGonderimZamani == null || 
                                              ayarlar.SonGonderimZamani.Value.Date < simdi.Date;

                    if (simdi >= bugunGonderimZamani && bugunGonderilmedi)
                    {
                        _logger.LogInformation("Gonderim zamani geldi! Mail hazirlaniyor...");
                        
                        var sonuc = await GunlukRaporGonderAsync(scope.ServiceProvider);
                        
                        if (sonuc)
                        {
                            ayarlar.SonGonderimZamani = simdi;
                            await ayarlarService.KaydetAsync(ayarlar);
                        }
                    }
                }

                // Pancar raporu otomatik gönderimi şu an devre dışı — yalnızca manuel gönderilir

                await Task.Delay(TimeSpan.FromMinutes(1), stoppingToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Zamanli mail servisinde hata");
                await Task.Delay(TimeSpan.FromMinutes(5), stoppingToken);
            }
        }

        _logger.LogInformation("Zamanli Mail Servisi durduruldu.");
    }

    private async Task<bool> GunlukRaporGonderAsync(IServiceProvider services)
    {
        try
        {
            var sekerService = services.GetRequiredService<SekerSatisService>();
            var yanUrunlerService = services.GetRequiredService<YanUrunlerService>();
            var htmlService = services.GetRequiredService<HtmlRaporService>();
            var ayarlarService = services.GetRequiredService<MailAyarlariService>();
            var ayarlar = ayarlarService.GetAyarlar();

            var baslangic = new DateTime(2025, 9, 1);
            var bitis = DateTime.Today;

            _logger.LogInformation("Seker verileri cekiliyor...");
            var sekerVerileri = await sekerService.GetSekerSatisOzetAsync(baslangic, bitis);

            _logger.LogInformation("Yan urun verileri cekiliyor...");
            var yanUrunVerileri = await yanUrunlerService.GetYanUrunlerOzetAsync(baslangic, bitis);

            _logger.LogInformation("Alkol verileri cekiliyor...");
            var alkolVerileri = await yanUrunlerService.GetAlkolOzetAsync(baslangic, bitis);
            // AlkolOzet -> YanUrunOzet dönüşümü yaparak listeye ekle
            foreach (var a in alkolVerileri)
            {
                yanUrunVerileri.Add(new RaporlamaPortali.Models.YanUrunOzet
                {
                    MalzemeKodu      = a.MalzemeKodu,
                    MalzemeAdi       = a.MalzemeAdi,
                    Kategori         = "ALKOL",
                    DevirStok        = a.DevirStok,
                    SatinAlmaMiktari = a.SatinAlmaMiktari,
                    UretimMiktari    = a.UretimMiktari,
                    SatisMiktari     = a.SatisMiktari,
                    SatisTutari      = a.SatisTutari,
                    IadeMiktari      = a.IadeMiktari,
                    IadeTutari       = a.IadeTutari
                });
            }

            _logger.LogInformation("HTML rapor olusturuluyor...");
            var htmlIcerik = htmlService.BirlesikRaporHtmlOlustur(sekerVerileri, yanUrunVerileri, baslangic, bitis);

            _logger.LogInformation("Mail gonderiliyor (Outlook)...");
            var konu = "Afyon Seker Fabrikasi Yan Urunler ve Seker Uretim-Satis-Stok Tablosu (" + bitis.ToString("dd.MM.yyyy") + ")";
            var sonuc = await OutlookMailGonderAsync(ayarlar, konu, htmlIcerik);

            if (sonuc.Basarili)
            {
                _logger.LogInformation("Mail basariyla gonderildi! {AliciSayisi} alici (BCC)", sonuc.AliciSayisi);
                return true;
            }
            else
            {
                _logger.LogError("Mail gonderilemedi: {Hata}", sonuc.Mesaj);
                return false;
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Gunluk rapor gonderiminde hata");
            return false;
        }
    }

    private async Task<bool> PancarRaporGonderAsync(IServiceProvider services)
    {
        try
        {
            var pancarService  = services.GetRequiredService<PancarRaporService>();
            var htmlService    = services.GetRequiredService<HtmlRaporService>();
            var ayarlarService = services.GetRequiredService<MailAyarlariService>();
            var ayarlar        = ayarlarService.GetAyarlar();

            if (ayarlar.Alicilar.Count == 0) return false;

            _logger.LogInformation("Pancar icmal verisi cekiliyor...");
            var icmal     = await pancarService.GetIcmalAsync();
            var ciftciler = await pancarService.GetCiftciListesiAsync();
            var avans     = await pancarService.GetAvansAsync();
            var finans    = await pancarService.GetFinansOzetAsync();

            var html  = htmlService.PancarRaporHtmlOlustur(icmal, ciftciler, DateTime.Today, avans, finans);
            var konu  = $"Afyon Şeker Fabrikası Pancar Raporu ({DateTime.Today:dd.MM.yyyy})";
            var sonuc = await OutlookMailGonderAsync(ayarlar, konu, html);

            if (sonuc.Basarili) { _logger.LogInformation("Pancar maili gonderildi."); return true; }
            else { _logger.LogError("Pancar mail hatasi: {H}", sonuc.Mesaj); return false; }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Pancar rapor gonderiminde hata");
            return false;
        }
    }

    // Manuel tetikleme için
    public async Task<bool> PancarRaporManuelGonder(IServiceProvider services)
        => await PancarRaporGonderAsync(services);

    /// <summary>
    /// Outlook uygulaması üzerinden mail gönderir (VBA'daki gibi)
    /// Şifre gerektirmez, bilgisayardaki Outlook hesabını kullanır
    /// </summary>
    public static async Task<MailSonuc> OutlookMailGonderAsync(MailAyarlariModel ayarlar, string konu, string htmlIcerik)
    {
        var sonuc = new MailSonuc();

        try
        {
            if (ayarlar.Alicilar.Count == 0)
            {
                sonuc.Basarili = false;
                sonuc.Mesaj = "Hic alici eklenmemis!";
                return sonuc;
            }

            // Outlook COM nesnesi oluştur
            Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType == null)
            {
                sonuc.Basarili = false;
                sonuc.Mesaj = "Outlook yuklu degil veya bulunamadi!";
                return sonuc;
            }

            // Outlook uygulamasını oluştur veya mevcut olanı kullan
            dynamic? outlookApp = Activator.CreateInstance(outlookType);

            if (outlookApp == null)
            {
                sonuc.Basarili = false;
                sonuc.Mesaj = "Outlook baslatilamadi!";
                return sonuc;
            }

            // Mail oluştur (olMailItem = 0)
            dynamic mailItem = outlookApp.CreateItem(0);

            // Alıcıları BCC olarak ekle (VBA'daki gibi)
            var bccList = string.Join("; ", ayarlar.Alicilar.Where(a => !string.IsNullOrWhiteSpace(a)));
            
            mailItem.To = ""; // To boş bırak
            mailItem.BCC = bccList;
            mailItem.Subject = konu;
            mailItem.HTMLBody = htmlIcerik;

            // Gönder
            await Task.Run(() => mailItem.Send());

            // COM nesnelerini temizle
            if (mailItem != null) Marshal.ReleaseComObject(mailItem);
            if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);

            sonuc.Basarili = true;
            sonuc.AliciSayisi = ayarlar.Alicilar.Count;
            sonuc.GonderimZamani = DateTime.Now;
        }
        catch (COMException ex)
        {
            sonuc.Basarili = false;
            sonuc.Mesaj = "Outlook hatasi: " + ex.Message;
        }
        catch (Exception ex)
        {
            sonuc.Basarili = false;
            sonuc.Mesaj = ex.Message;
        }

        return sonuc;
    }
}

public class MailSonuc
{
    public bool Basarili { get; set; }
    public string Mesaj { get; set; } = "";
    public DateTime GonderimZamani { get; set; }
    public int AliciSayisi { get; set; }
}
