using MudBlazor.Services;
using RaporlamaPortali.Services;
using System.Diagnostics;

var exeDir = Path.GetDirectoryName(Environment.ProcessPath) ?? AppContext.BaseDirectory;
var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args            = args,
    ContentRootPath = exeDir,
    WebRootPath     = Path.Combine(exeDir, "wwwroot")
});

// Kestrel'i sadece localhost'ta çalıştır (güvenlik için)
builder.WebHost.ConfigureKestrel(options =>
{
    options.ListenLocalhost(5050); // http://localhost:5050
});

// Add services to the container.
builder.Services.AddRazorPages();
builder.Services.AddServerSideBlazor();
builder.Services.AddMudServices();

// Database connection
builder.Services.AddSingleton<DatabaseService>();

// Report services
builder.Services.AddScoped<YanUrunlerService>();
builder.Services.AddScoped<SekerSatisService>();
builder.Services.AddScoped<PancarOdemeService>();
builder.Services.AddScoped<ExcelExportService>();
builder.Services.AddScoped<HtmlRaporService>();
builder.Services.AddScoped<PancarRaporService>();

// Mail servisleri
builder.Services.AddSingleton<MailAyarlariService>();
builder.Services.AddSingleton<ZamanliMailService>();
builder.Services.AddHostedService(sp => sp.GetRequiredService<ZamanliMailService>());

// WhatsApp servisleri
builder.Services.AddSingleton<WhatsAppAyarlariService>();
builder.Services.AddSingleton<WhatsAppProcessService>();
builder.Services.AddHostedService(sp => sp.GetRequiredService<WhatsAppProcessService>());

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
}

app.UseStaticFiles();
app.UseRouting();

app.MapBlazorHub();
app.MapFallbackToPage("/_Host");

// WhatsApp entegrasyonu için rapor API endpoint'i
// GET /api/rapor → Yan Ürünler + Şeker HTML raporunu döndürür
app.MapGet("/api/rapor", async (HttpContext context) =>
{
    try
    {
        using var scope = context.RequestServices.CreateScope();
        var yanUrunlerService = scope.ServiceProvider.GetRequiredService<YanUrunlerService>();
        var sekerService      = scope.ServiceProvider.GetRequiredService<SekerSatisService>();
        var htmlService       = scope.ServiceProvider.GetRequiredService<HtmlRaporService>();

        var baslangic = new DateTime(2025, 9, 1);
        var bitis     = DateTime.Today;

        var sekerVerileri    = await sekerService.GetSekerSatisOzetAsync(baslangic, bitis);
        var yanUrunVerileri  = await yanUrunlerService.GetYanUrunlerOzetAsync(baslangic, bitis);
        var alkolVerileri    = await yanUrunlerService.GetAlkolOzetAsync(baslangic, bitis);

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

        var html = htmlService.BirlesikRaporHtmlOlustur(sekerVerileri, yanUrunVerileri, baslangic, bitis);
        context.Response.ContentType = "text/html; charset=utf-8";
        await context.Response.WriteAsync(html);
    }
    catch (Exception ex)
    {
        context.Response.StatusCode = 500;
        await context.Response.WriteAsync("Hata: " + ex.Message);
    }
});

// Pancar raporu API endpoint'i
// GET /api/pancar-raporu → Pancar İCMAL + Çiftçi listesi HTML raporunu döndürür
app.MapGet("/api/pancar-raporu", async (HttpContext context) =>
{
    try
    {
        using var scope = context.RequestServices.CreateScope();
        var pancarService = scope.ServiceProvider.GetRequiredService<PancarRaporService>();
        var htmlService   = scope.ServiceProvider.GetRequiredService<HtmlRaporService>();

        var icmal     = await pancarService.GetIcmalAsync();
        var ciftciler = await pancarService.GetCiftciListesiAsync();
        var avans     = await pancarService.GetAvansAsync();
        var finans    = await pancarService.GetFinansOzetAsync();

        var html = htmlService.PancarRaporHtmlOlustur(icmal, ciftciler, DateTime.Today, avans, finans);
        context.Response.ContentType = "text/html; charset=utf-8";
        await context.Response.WriteAsync(html);
    }
    catch (Exception ex)
    {
        context.Response.StatusCode = 500;
        await context.Response.WriteAsync("Hata: " + ex.Message);
    }
});

// Uygulama başladığında tarayıcıyı otomatik aç
var url = "http://localhost:5050";
Console.WriteLine($"");
Console.WriteLine($"╔══════════════════════════════════════════════════════════╗");
Console.WriteLine($"║                   RAPORLAMA PORTALİ                       ║");
Console.WriteLine($"║            Doğuş Çay - Afyon Şeker Fabrikası              ║");
Console.WriteLine($"╠══════════════════════════════════════════════════════════╣");
Console.WriteLine($"║  Uygulama başlatıldı!                                     ║");
Console.WriteLine($"║  Tarayıcıda açılıyor: {url,-30}    ║");
Console.WriteLine($"╠══════════════════════════════════════════════════════════╣");
Console.WriteLine($"║  📧 Mail ayarları: Sol menü > Mail Ayarları               ║");
Console.WriteLine($"║                                                           ║");
Console.WriteLine($"║  Kapatmak için bu pencereyi kapatın veya Ctrl+C basın.   ║");
Console.WriteLine($"╚══════════════════════════════════════════════════════════╝");
Console.WriteLine($"");

// Tarayıcıyı aç
try
{
    Process.Start(new ProcessStartInfo
    {
        FileName = url,
        UseShellExecute = true
    });
}
catch
{
    Console.WriteLine($"Tarayıcı otomatik açılamadı. Lütfen manuel olarak açın: {url}");
}

app.Run();
