using MudBlazor.Services;
using RaporlamaPortali.Services;
using System.Diagnostics;
using System.Security.Claims;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;

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
builder.Services.AddRazorPages(options =>
{
    // Tüm Razor Pages giriş gerektirsin, sadece Giris sayfası anonim
    options.Conventions.AuthorizeFolder("/");
    options.Conventions.AllowAnonymousToPage("/Giris");
});
builder.Services.AddServerSideBlazor();
builder.Services.AddMudServices();

// Cookie Authentication
builder.Services.AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
    .AddCookie(options =>
    {
        options.LoginPath = "/giris";
        options.Cookie.Name = "RaporPortalAuth";
        options.ExpireTimeSpan = TimeSpan.FromDays(7);
        options.SlidingExpiration = true;
    });
builder.Services.AddAuthorization();

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

// Giriş servisi
builder.Services.AddSingleton<GirisAyarlariService>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
}

app.UseStaticFiles();
app.UseRouting();
app.UseAuthentication();
app.UseAuthorization();

app.MapBlazorHub();
app.MapRazorPages();

// Giriş endpoint'i (anonim)
app.MapPost("/giris-yap", async (HttpContext ctx, GirisAyarlariService girisService) =>
{
    var form = await ctx.Request.ReadFormAsync();
    var kullanici = form["kullanici"].FirstOrDefault() ?? "";
    var sifre     = form["sifre"].FirstOrDefault() ?? "";
    var returnUrl = form["returnUrl"].FirstOrDefault() ?? "/";
    if (!returnUrl.StartsWith("/")) returnUrl = "/";

    if (girisService.GirisKontrol(kullanici, sifre))
    {
        var claims   = new List<Claim> { new Claim(ClaimTypes.Name, kullanici) };
        var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
        await ctx.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, new ClaimsPrincipal(identity));
        return Results.Redirect(returnUrl);
    }
    return Results.Redirect("/giris?hata=1");
}).AllowAnonymous();

// Çıkış endpoint'i
app.MapGet("/cikis", async (HttpContext ctx) =>
{
    await ctx.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);
    return Results.Redirect("/giris");
}).AllowAnonymous();

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
}).AllowAnonymous();

// Pancar raporu API endpoint'i
// GET /api/pancar-raporu → Pancar İCMAL + Çiftçi listesi HTML raporunu döndürür
app.MapGet("/api/pancar-raporu", async (HttpContext context) =>
{
    try
    {
        using var scope = context.RequestServices.CreateScope();
        var pancarService = scope.ServiceProvider.GetRequiredService<PancarRaporService>();
        var htmlService   = scope.ServiceProvider.GetRequiredService<HtmlRaporService>();

        var t1 = pancarService.GetIcmalAsync();
        var t2 = pancarService.GetCiftciListesiAsync();
        var t3 = pancarService.GetAvansAsync();
        var t4 = pancarService.GetFinansOzetAsync();
        var t5 = pancarService.GetIcmalDetayAsync();
        var t6 = pancarService.GetOzetIstatistikAsync();
        await Task.WhenAll(t1, t2, t3, t4, t5, t6);

        var html = htmlService.PancarRaporHtmlOlustur(
            t1.Result, t2.Result, DateTime.Today, t3.Result, t4.Result, t5.Result, t6.Result);
        context.Response.ContentType = "text/html; charset=utf-8";
        await context.Response.WriteAsync(html);
    }
    catch (Exception ex)
    {
        context.Response.StatusCode = 500;
        await context.Response.WriteAsync("Hata: " + ex.Message);
    }
}).AllowAnonymous();

// Blazor fallback — giriş zorunlu
app.MapFallbackToPage("/_Host").RequireAuthorization();

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
