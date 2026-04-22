using Dapper;
using MudBlazor.Services;
using RaporlamaPortali.Services;
using System.Diagnostics;
using System.Security.Claims;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;

// Başlatma şifresi sor. Yanlış/iptal → process kapanır.
RaporlamaPortali.Services.LaunchAuthService.Require();

// Alt çizgili SQL kolon adlarını PascalCase property'lere otomatik eşleştir
// Örn: MALZEME_KODU → MalzemeKodu, AMBAR_NO → AmbarNo
DefaultTypeMap.MatchNamesWithUnderscores = true;

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

// IIS out-of-process altında çalışmıyorsa Kestrel portunu sabitle
if (Environment.GetEnvironmentVariable("ASPNETCORE_PORT") == null &&
    Environment.GetEnvironmentVariable("ASPNETCORE_IIS_HTTPPORT") == null)
{
    builder.WebHost.ConfigureKestrel(options =>
    {
        options.ListenLocalhost(5050); // http://localhost:5050
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

// Finans Raporu — yıllık INF_MD_FINANS_PROJE_RAPORU_211_YYYY view'lerini birleştirir
builder.Services.AddScoped<FinansRaporService>();

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
