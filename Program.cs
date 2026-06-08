using Dapper;
using MudBlazor.Services;
using RaporlamaPortali.Services;
using System.Diagnostics;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Security.Claims;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;

// Başlatma şifresi sor + şifrelenmiş kasayı (vault) aç. Yanlış/iptal → process kapanır.
// Require() bittiğinde SecretsService de yüklü; SQL şifreleri / mail şifresi / Claude API key
// hep buradan okunacak.
RaporlamaPortali.Services.LaunchAuthService.Require();

// Anthropic API key — ClaudeService env var'dan okuyor, kasadan gelen değeri Process scope'a yansıt
if (!string.IsNullOrEmpty(RaporlamaPortali.Services.SecretsService.AnthropicApiKey))
{
    Environment.SetEnvironmentVariable(
        "ANTHROPIC_API_KEY",
        RaporlamaPortali.Services.SecretsService.AnthropicApiKey,
        EnvironmentVariableTarget.Process);
}

// Alt çizgili SQL kolon adlarını PascalCase property'lere otomatik eşleştir
// Örn: MALZEME_KODU → MalzemeKodu, AMBAR_NO → AmbarNo
DefaultTypeMap.MatchNamesWithUnderscores = true;

// .NET Core/8: cp1254 vb. eski code page'ler default değil — ExcelDataReader BIFF5 (Logo NetRapor)
// dosyalarında Türkçe karakter okumak için bu provider'ı kayıt etmek zorunlu.
System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

var exeDir = Path.GetDirectoryName(Environment.ProcessPath) ?? AppContext.BaseDirectory;

// Veritabanı, evrak arşivi, WhatsApp oturumu vb. kalıcı verileri publish klasörü
// dışına taşı (ilk çalıştırmada otomatik migration)
RaporlamaPortali.Services.AppDataPaths.EnsureAndMigrate();

var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args            = args,
    ContentRootPath = exeDir,
    WebRootPath     = Path.Combine(exeDir, "wwwroot")
});

// Şifrelenmiş kasadan gelen değerleri IConfiguration'a bindir.
// Bu kayıt en sona eklenir → appsettings.json'daki (boş/placeholder) değerleri ezer.
// DatabaseService, MailGonderimService vb. constructor'ları bu sayede şifrelenmiş değeri görür.
{
    var s = RaporlamaPortali.Services.SecretsService.Snapshot();
    var ovr = new Dictionary<string, string?>();
    if (!string.IsNullOrWhiteSpace(s.LogoConnectionString))
        ovr["ConnectionStrings:LogoDB"]   = s.LogoConnectionString;
    if (!string.IsNullOrWhiteSpace(s.KantarConnectionString))
        ovr["ConnectionStrings:KantarDB"] = s.KantarConnectionString;
    if (!string.IsNullOrWhiteSpace(s.PmhsConnectionString))
        ovr["ConnectionStrings:PMHSDB"]   = s.PmhsConnectionString;
    if (!string.IsNullOrWhiteSpace(s.MailPassword))
        ovr["MailAyarlari:Sifre"]         = s.MailPassword;
    if (ovr.Count > 0) builder.Configuration.AddInMemoryCollection(ovr);
}

// IIS out-of-process altında çalışmıyorsa ve ASPNETCORE_URLS override edilmemişse
// 5050'yi hem localhost'ta hem de tüm IPv4 ağ arayüzlerinde dinle.
// VPN üzerinden telefon/diğer cihazlardan erişim mümkün olur.
// Not: ListenAnyIP (wildcard bind) non-interactive modda Windows kısıtlaması
// nedeniyle sessizce başarısız olabildiğinden, her IPv4 arayüzüne tek tek bind ediyoruz.
if (Environment.GetEnvironmentVariable("ASPNETCORE_PORT") == null &&
    Environment.GetEnvironmentVariable("ASPNETCORE_IIS_HTTPPORT") == null &&
    string.IsNullOrEmpty(Environment.GetEnvironmentVariable("ASPNETCORE_URLS")))
{
    builder.WebHost.ConfigureKestrel(options =>
    {
        options.ListenLocalhost(5050); // http://localhost:5050

        try
        {
            var ipv4Addresses = NetworkInterface.GetAllNetworkInterfaces()
                .Where(n => n.OperationalStatus == OperationalStatus.Up)
                .SelectMany(n => n.GetIPProperties().UnicastAddresses)
                .Where(a => a.Address.AddressFamily == AddressFamily.InterNetwork
                            && !IPAddress.IsLoopback(a.Address))
                .Select(a => a.Address)
                .Distinct()
                .ToList();

            foreach (var ip in ipv4Addresses)
            {
                try { options.Listen(ip, 5050); }
                catch { /* port çakışması veya bind izni yok — sessizce geç */ }
            }
        }
        catch { /* interface enumerasyonu başarısız — sadece localhost yetsin */ }
    });
}

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
builder.Services.AddScoped<SekerDairesiService>();
builder.Services.AddScoped<PancarOdemeService>();
builder.Services.AddScoped<ExcelExportService>();
builder.Services.AddScoped<OdemeService>();
builder.Services.AddScoped<OdemeKontrolService>();
builder.Services.AddScoped<OdemeKarsilastirService>();
builder.Services.AddScoped<SatisFaturaService>();
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

// Evrak Arşivi servisleri
builder.Services.AddSingleton<EvrakArsivService>();
builder.Services.AddSingleton<MustahsilLookupService>();

// Logo İşlemleri
builder.Services.AddScoped<LogoIslemleriService>();

// Kontrol: Kantar - Logo karşılaştırması
builder.Services.AddScoped<KantarLogoKarsilastirmaService>();

// Kontrol: SabNet - Logo (Müstahsil) karşılaştırması
builder.Services.AddScoped<MustahsilKarsilastirmaService>();

// Tarım Kredi Raporu (bölge eşleşmesi + yan ürün hareket)
builder.Services.AddSingleton<TarimKrediService>();

// Malzeme Hareket Listesi — kullanıcı tarafından tanımlanan kodlar için STLINE hareket raporu
builder.Services.AddScoped<MalzemeHareketService>();
builder.Services.AddSingleton<MalzemeListeService>();
builder.Services.AddSingleton<AfyonAmbarService>();
builder.Services.AddScoped<IrsaliyeListeService>();

// Finans Raporu — yıllık INF_MD_FINANS_PROJE_RAPORU_211_YYYY view'lerini birleştirir
builder.Services.AddScoped<FinansRaporService>();

// İşletme Malzemeleri Raporu — Yakıtlar / Torbalar / Kimyasallar (V1 + STLINE)
builder.Services.AddSingleton<IsletmeMalzemeleriKonfigService>();
builder.Services.AddScoped<IsletmeMalzemeleriService>();

// Cari Mutabakatı — AI destekli (Anthropic Claude API ile PDF/Excel parse)
builder.Services.AddHttpClient();
builder.Services.AddScoped<ClaudeService>();
builder.Services.AddScoped<MutabakatService>();
builder.Services.AddScoped<KarsiFirmaKurFarkiService>();

// e-Fatura görüntüleme — eFaturaDogusCayDB POSTBOX/ELEMENTS + UBL XML parse
builder.Services.AddScoped<EFaturaService>();

// e-Fatura → Logo Tiger Unity COM aktarımı
builder.Services.AddSingleton<RaporlamaPortali.Services.Logo.LogoUnityComService>();
builder.Services.AddSingleton<RaporlamaPortali.Services.Logo.LogoOrgLookupService>();
builder.Services.AddScoped<RaporlamaPortali.Services.Logo.LogoCariLookupService>();
builder.Services.AddScoped<RaporlamaPortali.Services.Logo.LogoMasterKartLookupService>();
builder.Services.AddScoped<RaporlamaPortali.Services.Logo.LogoBirimSetiLookupService>();
builder.Services.AddScoped<RaporlamaPortali.Services.Logo.LogoAccCodesLookupService>();
builder.Services.AddSingleton<RaporlamaPortali.Services.Logo.LogoKdvHesapLookupService>();
builder.Services.AddSingleton<EFaturaAktarimGecmisiService>();
builder.Services.AddScoped<RaporlamaPortali.Services.Logo.LogoAktarimService>();

// SabNet Kantar — SabNetKANTAR SQL Server'dan SabNet.db SQLite'a aktarım + listeleme
builder.Services.AddSingleton<SabNetDbService>();
builder.Services.AddSingleton<SabNetBaglantiService>();
builder.Services.AddScoped<SabNetImportService>();
builder.Services.AddScoped<SabNetSorguService>();
builder.Services.AddSingleton<SabNetSyncDurum>();
builder.Services.AddHostedService<SabNetSyncBackgroundService>();

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

// Çıkış endpoint'i — yedek alabilmek için WhatsApp/Chromium/Node ve programın
// kendisini tamamen kapatır. Yeniden başlatmaz; kullanıcı .exe'yi elle çalıştırır.
app.MapGet("/cikis", async (HttpContext ctx,
    WhatsAppProcessService wa,
    IHostApplicationLifetime lifetime) =>
{
    await ctx.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);

    // Response gönderildikten sonra arka planda her şeyi kapat
    _ = Task.Run(async () =>
    {
        try { await Task.Delay(1000); } catch { }

        // 1) WhatsApp node + puppeteer chromium alt ağacını kapat
        try { wa.BotDurdur(); } catch { }

        // 2) Defansif: arta kalan node + puppeteer chromium varsa öldür
        try
        {
            foreach (var p in Process.GetProcessesByName("node"))
            {
                try { p.Kill(entireProcessTree: true); } catch { }
            }
            foreach (var p in Process.GetProcessesByName("chrome"))
            {
                try
                {
                    var path = p.MainModule?.FileName ?? "";
                    if (path.IndexOf(@"\.cache\puppeteer\", StringComparison.OrdinalIgnoreCase) >= 0
                        || path.IndexOf(@"\puppeteer\", StringComparison.OrdinalIgnoreCase) >= 0)
                        p.Kill(entireProcessTree: true);
                }
                catch { }
            }
        }
        catch { }

        // 3) RaporlamaPortali'nin kendisini kapat
        try { await Task.Delay(500); } catch { }
        try { lifetime.StopApplication(); } catch { }
        try { await Task.Delay(2000); } catch { }
        Environment.Exit(0);
    });

    const string html = @"<!DOCTYPE html>
<html lang='tr'>
<head>
  <meta charset='utf-8'>
  <title>Kapatılıyor</title>
  <style>
    body{font-family:Segoe UI,Arial,sans-serif;background:#fafafa;color:#222;
         display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0}
    .card{background:#fff;padding:32px 40px;border-radius:8px;
          box-shadow:0 2px 8px rgba(0,0,0,.08);text-align:center;max-width:560px}
    h1{color:#c62828;font-size:1.35rem;margin:0 0 12px}
    p{margin:8px 0;color:#555;line-height:1.55}
    ul{text-align:left;margin:14px 0;color:#444}
    .muted{color:#999;font-size:.85rem;margin-top:18px}
    .ok{color:#2e7d32;font-weight:600}
  </style>
</head>
<body>
  <div class='card'>
    <h1>Yedek için sistem kapatılıyor</h1>
    <p>Yedek almanızı engelleyecek tüm süreçler kapatılıyor:</p>
    <ul>
      <li>WhatsApp botu (node.exe)</li>
      <li>Puppeteer Chromium pencereleri</li>
      <li>RaporlamaPortali.exe (uygulama)</li>
    </ul>
    <p class='ok'>Birkaç saniye içinde tamamlanır — sonra yedeği alabilirsiniz.</p>
    <p class='muted'>Programı tekrar başlatmak için masaüstündeki <b>RaporlamaPortali</b> kısayolunu çalıştırın.</p>
  </div>
</body>
</html>";
    return Results.Content(html, "text/html; charset=utf-8");
}).AllowAnonymous();

// DEBUG: Şeker view'undaki gerçek MALZEME_KODU ve MALZEME_ADI değerlerini göster
app.MapGet("/api/debug-seker", async (HttpContext context) =>
{
    using var scope = context.RequestServices.CreateScope();
    var db = scope.ServiceProvider.GetRequiredService<DatabaseService>();
    using var conn = db.CreateConnection();
    var rows = await conn.QueryAsync(@"
        SELECT DISTINCT TOP 50
            v.MALZEME_KODU,
            MALZEME_ADI = ISNULL(itm.NAME, '-- JOIN YOK --')
        FROM INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Seker_2025 v WITH(NOLOCK)
        LEFT JOIN LG_211_ITEMS itm WITH(NOLOCK) ON itm.CODE = v.MALZEME_KODU
        ORDER BY v.MALZEME_KODU");
    var sb = new System.Text.StringBuilder("<pre>");
    sb.AppendLine("MALZEME_KODU\t\t\tMALZEME_ADI");
    sb.AppendLine(new string('-', 80));
    foreach (var r in rows)
        sb.AppendLine($"{r.MALZEME_KODU,-35}\t{r.MALZEME_ADI}");
    sb.Append("</pre>");
    context.Response.ContentType = "text/html; charset=utf-8";
    await context.Response.WriteAsync(sb.ToString());
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
        var bitis     = RaporlamaPortali.Services.SistemTarihi.Bugun();

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

        bool bulanik = context.Request.Query["bulanik"] == "true";
        var html = htmlService.BirlesikRaporHtmlOlustur(sekerVerileri, yanUrunVerileri, baslangic, bitis, bulanik, kompakt: true);
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

        bool bulanik = context.Request.Query["bulanik"] == "true";
        var html = htmlService.PancarRaporHtmlOlustur(
            t1.Result, t2.Result, DateTime.Today, t3.Result, t4.Result, t5.Result, t6.Result, bulanik, kompakt: true);
        context.Response.ContentType = "text/html; charset=utf-8";
        await context.Response.WriteAsync(html);
    }
    catch (Exception ex)
    {
        context.Response.StatusCode = 500;
        await context.Response.WriteAsync("Hata: " + ex.Message);
    }
}).AllowAnonymous();

// GET /api/cay-durum-raporu → Çay Durum HTML (WhatsApp PNG için)
app.MapGet("/api/cay-durum-raporu", async (HttpContext context) =>
{
    try
    {
        using var scope = context.RequestServices.CreateScope();
        var logo = scope.ServiceProvider.GetRequiredService<LogoIslemleriService>();
        var html = scope.ServiceProvider.GetRequiredService<HtmlRaporService>();

        var veriler = await logo.CayDurumuAsync();
        bool bulanik = context.Request.Query["bulanik"] == "true";
        var output = html.CayDurumHtmlOlustur(veriler, bulanik);
        context.Response.ContentType = "text/html; charset=utf-8";
        await context.Response.WriteAsync(output);
    }
    catch (Exception ex)
    {
        context.Response.StatusCode = 500;
        await context.Response.WriteAsync("Hata: " + ex.Message);
    }
}).AllowAnonymous();

// GET /api/gubre-stok-raporu → Gübre Stok HTML (WhatsApp PNG için)
app.MapGet("/api/gubre-stok-raporu", async (HttpContext context) =>
{
    try
    {
        using var scope = context.RequestServices.CreateScope();
        var logo = scope.ServiceProvider.GetRequiredService<LogoIslemleriService>();
        var html = scope.ServiceProvider.GetRequiredService<HtmlRaporService>();

        var hepsi = await logo.StokDurumuAsync(null, sifirlariGizle: true);
        var onEkler = new[] { "A.G.", "S.707.03" };
        var filtreli = hepsi.Where(s => s.AmbarNo != -1
            && onEkler.Any(ek => (s.MalzemeKodu ?? "").StartsWith(ek, StringComparison.OrdinalIgnoreCase))).ToList();

        bool bulanik = context.Request.Query["bulanik"] == "true";
        var output = html.GubreStokHtmlOlustur(filtreli, bulanik);
        context.Response.ContentType = "text/html; charset=utf-8";
        await context.Response.WriteAsync(output);
    }
    catch (Exception ex)
    {
        context.Response.StatusCode = 500;
        await context.Response.WriteAsync("Hata: " + ex.Message);
    }
}).AllowAnonymous();

// GET /api/gubre-ciro-raporu → Gübre Ciro HTML (WhatsApp PNG için)
app.MapGet("/api/gubre-ciro-raporu", async (HttpContext context) =>
{
    try
    {
        using var scope = context.RequestServices.CreateScope();
        var logo = scope.ServiceProvider.GetRequiredService<LogoIslemleriService>();
        var html = scope.ServiceProvider.GetRequiredService<HtmlRaporService>();

        var bas = new DateTime(2026, 1, 1);
        var malzeme = await logo.GubreCiroMalzemeAsync(bas);

        bool bulanik = context.Request.Query["bulanik"] == "true";
        // WhatsApp sadece Malzeme Ciro sayfasini istiyor — pivot sorgusunu yapmaya gerek yok
        var output = html.GubreCiroHtmlOlustur(malzeme, new(), bas, bulanik, sadeceMalzeme: true);
        context.Response.ContentType = "text/html; charset=utf-8";
        await context.Response.WriteAsync(output);
    }
    catch (Exception ex)
    {
        context.Response.StatusCode = 500;
        await context.Response.WriteAsync("Hata: " + ex.Message);
    }
}).AllowAnonymous();

// DEBUG: Ekim 2025 A_KOTASI ham hareketleri (FIS_TURU bazında)
// GET /api/debug-ekim-akotasi
app.MapGet("/api/debug-ekim-akotasi", async (HttpContext context) =>
{
    using var scope = context.RequestServices.CreateScope();
    var db = scope.ServiceProvider.GetRequiredService<DatabaseService>();
    using var conn = db.CreateConnection();

    // A_KOTASI malzeme kodları
    var aKotasiKodlar = new[] { "S.T.0.0.0", "S.T.0.0.4", "S.705.00.0005" };

    var rows = await conn.QueryAsync<dynamic>(@"
        SELECT
            FIS_TURU,
            MALZEME_KODU,
            ToplamGiris  = SUM(ISNULL(GIRIS_MIKTAR_KG,  0)),
            ToplamCikis  = SUM(ISNULL(CIKIS_MIKTARI_KG, 0)),
            SatirSayisi  = COUNT(*)
        FROM INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_Seker_2025 WITH(NOLOCK)
        WHERE TARIH >= '2025-10-01' AND TARIH <= '2025-10-31'
          AND MALZEME_KODU IN ('S.T.0.0.0','S.T.0.0.4','S.705.00.0005')
        GROUP BY FIS_TURU, MALZEME_KODU
        ORDER BY FIS_TURU, MALZEME_KODU");

    var sb = new System.Text.StringBuilder("<pre style='font-family:monospace;font-size:13px'>");
    sb.AppendLine("=== EKİM 2025 A_KOTASI HAREKETLERİ ===");
    sb.AppendLine($"{"FIS_TURU",-45} {"MALZEME_KODU",-20} {"GIRIS_KG",15:N2} {"CIKIS_KG",15:N2} {"SAYI",6}");
    sb.AppendLine(new string('-', 110));
    decimal topGiris = 0, topCikis = 0;
    foreach (var r in rows)
    {
        decimal g = (decimal)(r.ToplamGiris ?? 0m);
        decimal c = (decimal)(r.ToplamCikis ?? 0m);
        topGiris += g; topCikis += c;
        sb.AppendLine($"{r.FIS_TURU,-45} {r.MALZEME_KODU,-20} {g,15:N2} {c,15:N2} {r.SatirSayisi,6}");
    }
    sb.AppendLine(new string('-', 110));
    sb.AppendLine($"{"TOPLAM",-67} {topGiris,15:N2} {topCikis,15:N2}");
    sb.AppendLine();
    sb.AppendLine($"NET (Giris - Cikis) = {topGiris - topCikis:N2}");
    sb.Append("</pre>");

    context.Response.ContentType = "text/html; charset=utf-8";
    await context.Response.WriteAsync(sb.ToString());
}).AllowAnonymous();

// Şeker Kategorisi Bazlı Analiz (üst tablo – ham LOGO) API endpoint'i
// GET /api/seker-analiz?baslangic=2025-09-01&bitis=2025-09-30
app.MapGet("/api/seker-analiz", async (HttpContext context) =>
{
    try
    {
        using var scope = context.RequestServices.CreateScope();
        var sekerDairesiService = scope.ServiceProvider.GetRequiredService<SekerDairesiService>();
        var htmlService         = scope.ServiceProvider.GetRequiredService<HtmlRaporService>();

        if (!DateTime.TryParse(context.Request.Query["baslangic"], out var baslangic))
            baslangic = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        if (!DateTime.TryParse(context.Request.Query["bitis"], out var bitis))
            bitis = new DateTime(DateTime.Today.Year, DateTime.Today.Month,
                DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month));

        bool bulanik = context.Request.Query["bulanik"] == "true";
        var (analiz, _) = await sekerDairesiService.GetSadeSekerAnaliziAsync(baslangic, bitis);
        var html = htmlService.SekerAnalizHtmlOlustur(analiz, baslangic, bitis, bulanik, kompakt: true);
        context.Response.ContentType = "text/html; charset=utf-8";
        await context.Response.WriteAsync(html);
    }
    catch (Exception ex)
    {
        context.Response.StatusCode = 500;
        await context.Response.WriteAsync("Hata: " + ex.Message);
    }
}).AllowAnonymous();

// Şeker Dairesi Başkanlık raporu API endpoint'i
// GET /api/seker-raporu?baslangic=2025-09-01&bitis=2025-09-30
app.MapGet("/api/seker-raporu", async (HttpContext context) =>
{
    try
    {
        using var scope = context.RequestServices.CreateScope();
        var sekerDairesiService = scope.ServiceProvider.GetRequiredService<SekerDairesiService>();
        var htmlService         = scope.ServiceProvider.GetRequiredService<HtmlRaporService>();

        if (!DateTime.TryParse(context.Request.Query["baslangic"], out var baslangic))
            baslangic = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        if (!DateTime.TryParse(context.Request.Query["bitis"], out var bitis))
            bitis = new DateTime(DateTime.Today.Year, DateTime.Today.Month,
                DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month));

        bool bulanik = context.Request.Query["bulanik"] == "true";
        var (analiz, dipnotlar) = await sekerDairesiService.GetSadeSekerAnaliziAsync(baslangic, bitis);
        var tBas = sekerDairesiService.GetBaskanlikDonemBasiAsync(baslangic);
        var tSon = sekerDairesiService.GetBaskanlikDonemBasiAsync(bitis.AddDays(1));
        await Task.WhenAll(tBas, tSon);
        var html = htmlService.SekerRaporHtmlOlustur(analiz, dipnotlar, baslangic, bitis, tBas.Result, tSon.Result, bulanik, kompakt: true);
        context.Response.ContentType = "text/html; charset=utf-8";
        await context.Response.WriteAsync(html);
    }
    catch (Exception ex)
    {
        context.Response.StatusCode = 500;
        await context.Response.WriteAsync("Hata: " + ex.Message);
    }
}).AllowAnonymous();

// Malzeme Hareket Listesi — Excel'den Web Query / Power Query ile yenilenebilsin diye CSV döndürür.
// GET /api/malzeme-hareket?liste=AfyonYanUrun&baslangic=2024-09-30&bitis=2026-08-31
//   veya
// GET /api/malzeme-hareket?kodlar=S.706.04.0001,S.706.04.0002&baslangic=2024-09-30
// Parametre verilmezse tüm kayıtlı listeleri birleştirir, tarih aralığı 2023-09-18'den bugüne.
app.MapGet("/api/malzeme-hareket", async (HttpContext ctx,
    MalzemeHareketService hareket, MalzemeListeService liste) =>
{
    try
    {
        var kodSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var listeAdi = ctx.Request.Query["liste"].ToString();
        if (!string.IsNullOrWhiteSpace(listeAdi))
        {
            var l = liste.Getir(listeAdi);
            if (l != null) foreach (var k in l.MalzemeKodlari) kodSet.Add(k);
        }

        var kodlarQs = ctx.Request.Query["kodlar"].ToString();
        if (!string.IsNullOrWhiteSpace(kodlarQs))
            foreach (var k in kodlarQs.Split(new[] { ',', ';', '|' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                kodSet.Add(k);

        // Parametre yoksa tüm kayıtlı listeleri birleştir
        if (kodSet.Count == 0)
            foreach (var l in liste.Listele())
                foreach (var k in l.MalzemeKodlari) kodSet.Add(k);

        if (!DateTime.TryParse(ctx.Request.Query["baslangic"], out var bas))
            bas = new DateTime(2023, 9, 18);
        if (!DateTime.TryParse(ctx.Request.Query["bitis"], out var bit))
            bit = RaporlamaPortali.Services.SistemTarihi.Bugun();

        var satirlar = await hareket.GetHareketlerAsync(kodSet, bas, bit);

        // CSV (UTF-8 BOM + ; ayırıcı — Excel Türkçe yerelinde direkt açılır)
        var sb = new System.Text.StringBuilder();
        sb.Append('﻿'); // BOM
        sb.AppendLine("YIL;AY;TARIH;FIS_TURU;FIS_NUMARASI;CARI_HESAP_KODU;CARI_HESAP_UNVANI;MALZEME_KODU;MALZEME_ACIKLAMASI;GIRIS_MIKTARI;GIRIS_FIYATI;GIRIS_TUTARI;CIKIS_MIKTARI;CIKIS_FIYATI;CIKIS_TUTARI");
        string E(string? s) => (s ?? "").Replace(";", ",").Replace("\r", " ").Replace("\n", " ");
        string N(decimal d) => d.ToString(System.Globalization.CultureInfo.InvariantCulture);
        foreach (var s in satirlar)
            sb.AppendLine(string.Join(';',
                s.Yil, s.Ay, s.Tarih.ToString("yyyy-MM-dd"),
                E(s.FisTuru), E(s.FisNumarasi),
                E(s.CariHesapKodu), E(s.CariHesapUnvani),
                E(s.MalzemeKodu), E(s.MalzemeAciklamasi),
                N(s.GirisMiktari), N(s.GirisFiyati), N(s.GirisTutari),
                N(s.CikisMiktari), N(s.CikisFiyati), N(s.CikisTutari)));

        ctx.Response.ContentType = "text/csv; charset=utf-8";
        ctx.Response.Headers["Content-Disposition"] = "inline; filename=\"malzeme-hareket.csv\"";
        await ctx.Response.WriteAsync(sb.ToString());
    }
    catch (Exception ex)
    {
        ctx.Response.StatusCode = 500;
        await ctx.Response.WriteAsync("Hata: " + ex.Message);
    }
}).AllowAnonymous();

// e-Fatura XML indir (auth zorunlu)
app.MapGet("/efatura-xml", async (HttpContext ctx, EFaturaService svc, long id) =>
{
    var pair = await svc.XmlGetirAsync(id);
    if (pair == null) return Results.NotFound();
    var bytes = System.Text.Encoding.UTF8.GetBytes(pair.Value.Xml);
    return Results.File(bytes, "application/xml", $"eFatura_{id}_{pair.Value.DosyaAdi}");
}).RequireAuthorization();

// e-Fatura orijinal ZIP indir
app.MapGet("/efatura-zip", async (HttpContext ctx, EFaturaService svc, long id) =>
{
    var data = await svc.ZipGetirAsync(id);
    if (data == null) return Results.NotFound();
    return Results.File(data, "application/zip", $"eFatura_{id}.zip");
}).RequireAuthorization();

// e-Fatura yazdırılabilir HTML önizleme (Ctrl+P → PDF olarak kaydet)
app.MapGet("/efatura-pdf", async (HttpContext ctx, EFaturaService svc, long id, string? vkn) =>
{
    var detay = await svc.DetayGetirAsync(id, vkn);
    if (detay == null) return Results.NotFound();
    var html = EFaturaHtmlBuilder.Build(detay);
    ctx.Response.ContentType = "text/html; charset=utf-8";
    await ctx.Response.WriteAsync(html);
    return Results.Empty;
}).RequireAuthorization();

// Evrak dosyası indirme / görüntüleme (auth zorunlu)
// GET /evrak-dosya?kategori=tesis|mustahsil|genel&id=123&inline=true
app.MapGet("/evrak-dosya", (HttpContext ctx, EvrakArsivService arsiv,
    string kategori, int id, bool inline) =>
{
    string? fullPath = null;
    string  dosyaAdi = "dosya";
    string? mime     = "application/octet-stream";

    if (kategori.Equals("tesis", StringComparison.OrdinalIgnoreCase))
    {
        var e = arsiv.TesisEvrakGetir(id);
        if (e == null) return Results.NotFound();
        fullPath = arsiv.TamYolaCevir(e.DosyaYolu);
        dosyaAdi = e.DosyaAdi;
        mime     = e.MimeType ?? mime;
    }
    else if (kategori.Equals("mustahsil", StringComparison.OrdinalIgnoreCase))
    {
        var e = arsiv.MustahsilEvrakGetir(id);
        if (e == null) return Results.NotFound();
        fullPath = arsiv.TamYolaCevir(e.DosyaYolu);
        dosyaAdi = e.DosyaAdi;
        mime     = e.MimeType ?? mime;
    }
    else if (kategori.Equals("genel", StringComparison.OrdinalIgnoreCase))
    {
        var e = arsiv.GenelEvrakGetir(id);
        if (e == null) return Results.NotFound();
        fullPath = arsiv.TamYolaCevir(e.DosyaYolu);
        dosyaAdi = e.DosyaAdi;
        mime     = e.MimeType ?? mime;
    }

    if (fullPath == null || !File.Exists(fullPath)) return Results.NotFound();

    var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.Read);
    return inline
        ? Results.File(stream, mime, enableRangeProcessing: true)
        : Results.File(stream, mime, dosyaAdi);
}).RequireAuthorization();

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
