using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using ExcelDataReader;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Anthropic Claude API üzerinden cari ekstresi (PDF / Excel) parse'ı.
/// API key ortam değişkeninden okunur (ANTHROPIC_API_KEY); commit edilmez.
/// </summary>
public class ClaudeService
{
    private readonly HttpClient _http;
    private readonly string? _apiKey;
    private const string ApiUrl = "https://api.anthropic.com/v1/messages";
    private const string Model = "claude-sonnet-4-5";

    public ClaudeService(IHttpClientFactory httpFactory)
    {
        _http = httpFactory.CreateClient();
        // Büyük ekstreler (1+ yıl, yüzlerce satır) max_tokens=32000 ile 3-5 dk sürebiliyor.
        // Streaming kullanmadığımız için tek bir HTTP isteğinde yanıtın tamamı beklenir.
        _http.Timeout = TimeSpan.FromMinutes(10);
        _apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY", EnvironmentVariableTarget.User)
                  ?? Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY", EnvironmentVariableTarget.Process)
                  ?? Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY", EnvironmentVariableTarget.Machine);
    }

    public bool ApiKeyVar => !string.IsNullOrWhiteSpace(_apiKey);

    /// <summary>
    /// PDF veya Excel byte'larını parse edip ekstre satırlarını döner.
    /// Excel ise yerel olarak metne çevrilir (token tasarrufu).
    /// PDF ise doğrudan Claude'a document olarak gönderilir.
    /// </summary>
    public async Task<List<MutabakatKayit>> EkstreyiParseEtAsync(
        byte[] dosyaBytes, string dosyaAdi, CancellationToken ct = default)
    {
        if (!ApiKeyVar)
            throw new InvalidOperationException(
                "ANTHROPIC_API_KEY ortam değişkeni tanımlı değil. Mutabakat sayfası için " +
                "console.anthropic.com'dan API key alıp PowerShell ile " +
                "[Environment]::SetEnvironmentVariable(\"ANTHROPIC_API_KEY\", \"sk-...\", \"User\") " +
                "komutuyla kaydedin.");

        var uzanti = (Path.GetExtension(dosyaAdi) ?? "").ToLowerInvariant();
        bool pdf = uzanti == ".pdf";
        bool excel = uzanti is ".xlsx" or ".xls" or ".xlsm";

        if (!pdf && !excel)
            throw new InvalidOperationException("Sadece PDF veya Excel (xlsx/xls/xlsm) yükleyebilirsiniz.");

        // İstek body'sini hazırla
        object userContent;
        if (pdf)
        {
            var b64 = Convert.ToBase64String(dosyaBytes);
            userContent = new object[]
            {
                new {
                    type = "document",
                    source = new { type = "base64", media_type = "application/pdf", data = b64 }
                },
                new { type = "text", text = PromptMetni() }
            };
        }
        else
        {
            // Excel'i metne çevir — Claude'a token olarak çok daha ucuz.
            // .xls dosyaları çoğu zaman aslında HTML raporu olarak gelir (Logo/Netsis/NetRapor),
            // veya gerçek BIFF binary olabilir. ClosedXML sadece .xlsx anlar.
            // Strateji: önce ClosedXML dene, başarısız olursa HTML/düz metin olarak oku.
            string metin;
            try
            {
                metin = ExcelMetneCevir(dosyaBytes);
            }
            catch
            {
                metin = HtmlVeyaMetinOlarakOku(dosyaBytes);
            }

            userContent = new object[]
            {
                new { type = "text", text = PromptMetni() + "\n\n--- EKSTRE İÇERİĞİ (TABLO) ---\n" + metin }
            };
        }

        var body = new
        {
            model = Model,
            max_tokens = 32000,
            messages = new[]
            {
                new { role = "user", content = userContent }
            }
        };

        var json = JsonSerializer.Serialize(body);
        using var req = new HttpRequestMessage(HttpMethod.Post, ApiUrl);
        req.Headers.Add("x-api-key", _apiKey);
        req.Headers.Add("anthropic-version", "2023-06-01");
        req.Content = new StringContent(json, Encoding.UTF8, "application/json");

        using var resp = await _http.SendAsync(req, ct);
        var respText = await resp.Content.ReadAsStringAsync(ct);
        if (!resp.IsSuccessStatusCode)
            throw new HttpRequestException($"Claude API hata ({(int)resp.StatusCode}): {respText}");

        // Yanıttan asistan metnini çıkar
        using var doc = JsonDocument.Parse(respText);
        var contentArr = doc.RootElement.GetProperty("content");
        var sb = new StringBuilder();
        foreach (var blok in contentArr.EnumerateArray())
        {
            if (blok.TryGetProperty("type", out var t) && t.GetString() == "text"
                && blok.TryGetProperty("text", out var txt))
            {
                sb.Append(txt.GetString());
            }
        }
        var asistanMetin = sb.ToString();

        // Asistanın çıktısı içinden JSON bloğu yakala
        var jsonGovde = JsonBlokunuCikar(asistanMetin)
            ?? throw new InvalidOperationException("AI yanıtında JSON bulunamadı:\n" + asistanMetin);

        return JsonuKayitlaraCevir(jsonGovde);
    }

    /// <summary>
    /// Döviz + TL kolonlarını AYRI AYRI çıkaran parse — kur farkı hesaplaması için.
    /// Her satır için hem TL borc/alacak, hem döviz borc/alacak ve döviz cinsi okunur.
    /// </summary>
    public async Task<List<MutabakatKayit>> DovizliEkstreyiParseEtAsync(
        byte[] dosyaBytes, string dosyaAdi, CancellationToken ct = default)
    {
        if (!ApiKeyVar)
            throw new InvalidOperationException(
                "ANTHROPIC_API_KEY ortam değişkeni tanımlı değil.");

        var uzanti = (Path.GetExtension(dosyaAdi) ?? "").ToLowerInvariant();
        bool pdf = uzanti == ".pdf";
        bool excel = uzanti is ".xlsx" or ".xls" or ".xlsm";

        if (!pdf && !excel)
            throw new InvalidOperationException("Sadece PDF veya Excel (xlsx/xls/xlsm) yükleyebilirsiniz.");

        object userContent;
        if (pdf)
        {
            var b64 = Convert.ToBase64String(dosyaBytes);
            userContent = new object[]
            {
                new { type = "document",
                      source = new { type = "base64", media_type = "application/pdf", data = b64 } },
                new { type = "text", text = DovizliPromptMetni() }
            };
        }
        else
        {
            string metin;
            try { metin = ExcelMetneCevir(dosyaBytes); }
            catch { metin = HtmlVeyaMetinOlarakOku(dosyaBytes); }
            userContent = new object[]
            {
                new { type = "text", text = DovizliPromptMetni() + "\n\n--- EKSTRE İÇERİĞİ ---\n" + metin }
            };
        }

        var body = new
        {
            model = Model,
            max_tokens = 32000,
            messages = new[] { new { role = "user", content = userContent } }
        };

        var json = JsonSerializer.Serialize(body);
        using var req = new HttpRequestMessage(HttpMethod.Post, ApiUrl);
        req.Headers.Add("x-api-key", _apiKey);
        req.Headers.Add("anthropic-version", "2023-06-01");
        req.Content = new StringContent(json, Encoding.UTF8, "application/json");

        using var resp = await _http.SendAsync(req, ct);
        var respText = await resp.Content.ReadAsStringAsync(ct);
        if (!resp.IsSuccessStatusCode)
            throw new HttpRequestException($"Claude API hata ({(int)resp.StatusCode}): {respText}");

        using var doc = JsonDocument.Parse(respText);
        var contentArr = doc.RootElement.GetProperty("content");
        var sb = new StringBuilder();
        foreach (var blok in contentArr.EnumerateArray())
        {
            if (blok.TryGetProperty("type", out var t) && t.GetString() == "text"
                && blok.TryGetProperty("text", out var txt))
                sb.Append(txt.GetString());
        }
        var asistanMetin = sb.ToString();

        var jsonGovde = JsonBlokunuCikar(asistanMetin)
            ?? throw new InvalidOperationException("AI yanıtında JSON bulunamadı:\n" + asistanMetin);

        return DovizliJsonuKayitlaraCevir(jsonGovde);
    }

    private static string DovizliPromptMetni() =>
        @"Aşağıda DÖVİZLİ bir CARİ EKSTRESİ var (Türkçe muhasebe — Logo, Netsis, Mikro vb.).
Bu ekstrede HEM TL HEM DÖVİZ (USD/EUR) kolonları aynı satırda yer alıyor. Görevin: HER hareket
satırını JSON dizisi olarak döndürmek; bir satırda HEM TL HEM DÖVİZ alanlarını AYRI AYRI doldurmak.
SADECE bir JSON dizisi yaz; başka açıklama / kod fence yazma.

Çıktı şeması — her eleman:
{
  ""tarih"": ""YYYY-MM-DD"",
  ""borc"": 0.00,            // TL borç (sayı, ondalık nokta). Yoksa 0.
  ""alacak"": 0.00,          // TL alacak. Yoksa 0.
  ""dovizCinsi"": ""USD"",   // ""USD"" / ""EUR"" / ""GBP"" — döviz yoksa boş """"
  ""dovizBorc"": 0.00,       // Döviz borç. Yoksa 0.
  ""dovizAlacak"": 0.00,     // Döviz alacak. Yoksa 0.
  ""aciklama"": ""..."",
  ""fisNo"": ""...""
}

═══════════════════════════════════════════════════════════════════════════════
İŞ AKIŞIN:
  ① ÜST BAŞLIK satırını oku: ""USD"", ""DOLAR"", ""EUR"", ""EURO"", ""DÖVİZ"" yazan blok = döviz kolonları.
  ② SUB-BAŞLIK satırını oku: ""BORÇ"", ""ALACAK"", ""BAKİYE"" etiketleri burada.
  ③ Bir satırda iki ""BORÇ/ALACAK"" çifti varsa: BÜYÜK rakamlı çift = TL, KÜÇÜK rakamlı çift = DÖVİZ.
     (Türkçe ekstrede TL tutarı dövizin 30-50 KATIDIR — 1 USD ≈ 32 TL, 1 EUR ≈ 35 TL).
  ④ Etiketsiz kolon = TL kabul edilir.
  ⑤ ""BAKİYE"" / ""KÜMÜLATİF"" / ""DEVİR"" kolonlarını YOK SAY.
═══════════════════════════════════════════════════════════════════════════════

KRİTİK KURALLAR:

1. **TARİH = İLK TARİH**. Vade tarihi DEĞİL, işlem tarihi.
2. **DÖVİZ CİNSİ tespit et**: Üst başlıkta ""USD"" yazıyorsa dovizCinsi=""USD"", ""EUR"" → ""EUR"".
   Belirleyemiyorsan veya sadece TL varsa boş """" bırak.
3. **HEM TL HEM DÖVİZ**: Bir fatura satırında borc=1457825.98 ve dovizBorc=33418.44 birlikte olmalı
   (kur ≈ 43.625 buradan örtülü olarak çıkar).
4. **HEM Borç hem Alacak aynı yönde olur**: TL borç ile döviz borç aynı satırda; TL alacak ile döviz
   alacak aynı satırda. Ters yön YOK.
5. **HEPSİNİ ÇIKAR** (devreden / açılış / dönem sonu / toplam satırlarını ATLA).
6. **SAYI FORMATI**: ""1.234.567,89"" → 1234567.89.

ÖRNEK (DENKIM TARTI tarzı USD ekstre):
  Veri:    11.02.2026 | FATURA NO 1 | 33.418,44 USD | 1.457.825,98 TL alacak
  Çıktı:   {""tarih"":""2026-02-11"",""borc"":0,""alacak"":1457825.98,""dovizCinsi"":""USD"",
            ""dovizBorc"":0,""dovizAlacak"":33418.44,""aciklama"":""Fatura"",""fisNo"":""1""}

  Veri:    10.04.2026 | BANKA TAHSİLAT | 33.418,44 USD borç | 1.490.141,61 TL borç
  Çıktı:   {""tarih"":""2026-04-10"",""borc"":1490141.61,""alacak"":0,""dovizCinsi"":""USD"",
            ""dovizBorc"":33418.44,""dovizAlacak"":0,""aciklama"":""Banka Tahsilat"",""fisNo"":""""}";

    private static List<MutabakatKayit> DovizliJsonuKayitlaraCevir(string jsonDizi)
    {
        using var doc = JsonDocument.Parse(jsonDizi);
        if (doc.RootElement.ValueKind != JsonValueKind.Array)
            throw new InvalidOperationException("AI yanıtı JSON dizisi değil.");

        var list = new List<MutabakatKayit>();
        foreach (var e in doc.RootElement.EnumerateArray())
        {
            if (e.ValueKind != JsonValueKind.Object) continue;
            var k = new MutabakatKayit
            {
                Tarih       = AlTarih(e, "tarih"),
                Borc        = AlSayi(e, "borc"),
                Alacak      = AlSayi(e, "alacak"),
                Doviz       = AlMetin(e, "dovizCinsi").ToUpperInvariant().Trim(),
                DovizBorc   = AlSayi(e, "dovizBorc"),
                DovizAlacak = AlSayi(e, "dovizAlacak"),
                Aciklama    = AlMetin(e, "aciklama"),
                FisNo       = AlMetin(e, "fisNo"),
            };
            if (k.Tarih == default && k.Borc == 0 && k.Alacak == 0
                && k.DovizBorc == 0 && k.DovizAlacak == 0) continue;
            list.Add(k);
        }
        return list;
    }

    private static string PromptMetni() =>
        @"Aşağıda bir CARİ EKSTRESİ var (Türkçe muhasebe ekstresi — Logo, Netsis, Mikro, NetRapor vb.).
Görevin: TÜM hareket satırlarını JSON dizisi olarak çıkarmak. SADECE tek bir JSON dizisi yaz; başka
açıklama / metin / kod bloğu işaretleyici yazma. Hiçbir satırı atlama (bakiye/devreden hariç).

Çıktı şeması — her eleman:
{
  ""tarih"": ""YYYY-MM-DD"",        // İŞLEM tarihi (ilk tarih kolonu). Vade tarihi DEĞİL!
  ""borc"": 0.00,                    // TL borç. SAYI; ondalık nokta. Yoksa 0.
  ""alacak"": 0.00,                  // TL alacak. SAYI; yoksa 0.
  ""aciklama"": ""..."",            // İşlem açıklaması (FATURANIZ, GELEN HAVALE, vb.)
  ""fisNo"": ""...""                 // fiş / fatura / dekont numarası; yoksa boş string
}

═══════════════════════════════════════════════════════════════════════════════
ÖNCE: KOLON YAPISINI ANLA (PARSE'TAN ÖNCE TABLO BAŞLIĞINI OKU)

Türkçe muhasebe ekstreleri ÇOĞU ZAMAN ÇOK PARA BİRİMLİ tablo halinde gelir:
örn. EURO BORÇ | EURO ALACAK | (boş) | TL BORÇ | TL ALACAK | TL BAKİYE
veya  TL BORÇ | TL ALACAK | USD BORÇ | USD ALACAK | KUR | BAKİYE

İŞ AKIŞIN:
  ① ÜST BAŞLIK SATIRINI (Row 1) bul: ""EURO"", ""USD"", ""DOLAR"", ""EUR"", ""DÖVİZ"", ""KÜMÜLATIF"" yazıyorsa
     ALTINDAKİ kolonlar DÖVİZ/BAKİYE — bunları YOK SAY.
  ② SUB-BAŞLIK SATIRINI (Row 2) bul: ""BORÇ"", ""ALACAK"", ""BAKİYE"", ""TARİH"", ""FİŞ NO"", ""AÇIKLAMA""
     etiketleri buradadır.
  ③ Eğer iki ""BORÇ/ALACAK"" çifti varsa: üstünde EURO/USD yazan ÇİFTİ ATLA, ETİKETSİZ veya TL/TRY
     yazan ÇİFTİ AL.
  ④ ETİKETSİZ KOLON = TL'dir (Türkçe muhasebe geleneği — ana para birimi etiketlenmez, döviz etiketlenir).
═══════════════════════════════════════════════════════════════════════════════

KRİTİK KURALLAR:

1. **TARİH = İLK TARİH KOLONU**. Bir satırda iki tarih varsa (örn. 20.05.2025 ve 20.05.2025 / vade
   29.08.2025), İLKİNİ AL. Vade tarihi işlem tarihi DEĞİLDİR.

2. **TL KOLONLARINI AL — USD/EUR/DÖVİZ DEĞİL**. ÇOK ÖNEMLİ: Bir tabloda iki Borç/Alacak çifti
   varsa, üstünde ""EURO"" / ""USD"" / ""DOLAR"" / ""EUR"" / ""DÖVİZ"" yazan ÇİFTİ ATLA. TL kolonu
   genelde ETİKETSİZ olur — başlığında para birimi yazmaz çünkü ana para birimidir.

   ÖRNEK YAPI:
     Üst başlık:  | | | | | EURO | EURO | KÜMÜLATIF EURO BAKİYE | | | (boş) | (boş) | (boş) |
     Sub başlık:  CH UNVAN|TARİH|VADE|FİŞ NO|FİŞ TÜRÜ|BORÇ|ALACAK|BAKİYE|BORÇ|ALACAK|BAKİYE|AÇIKLAMA
     Veri satırı: TOYOTA|20.05.2025|20.05.2025|000571658|Hav|190725|0|-190725|8245213.40|0|8245213.40|...

     → DOĞRU çıktı: borc=8245213.40, alacak=0  (TL kolonu = 9. kolon, ETİKETSİZ)
     → YANLIŞ olur: borc=190725  (bu EURO Borç, atla)
     → YANLIŞ olur: borc=-190725 (bu EURO Bakiye)

3. **BÜYÜK RAKAM = TL — BU EN GÜVENİLİR KURAL**. İki Borç çifti veya iki Alacak çifti varsa,
   BÜYÜK OLAN HER ZAMAN TL'dir, küçük olan dövizdir. Türkçe muhasebe ekstresinde TL tutarları
   döviz tutarından 30-50 KAT büyüktür (1 EUR ≈ 35 TL, 1 USD ≈ 32 TL).
   - 190.725 vs 8.245.213,40 görünce → 8.245.213,40 TL'dir, AL.
   - 53.940 vs 2.524.413,58 görünce → 2.524.413,58 TL'dir, AL.
   - 99.150 vs 4.696.457,88 görünce → 4.696.457,88 TL'dir, AL.
   Etiket olsa da olmasa da BÜYÜK rakam TL kabul edilir. Tereddüt etme.

4. **BAKİYE / KÜMÜLATİF KOLONLARI YOK SAY**. ""Bakiye"" / ""Kümülatif"" / ""Devir"" / ""Borç Bak.""
   / ""Alacak Bak."" hareket DEĞİL — kümülatif tutardır. Borc/Alacak alanlarına yazma.

5. **HİÇBİR HAREKET SATIRINI ATLAMA**. Tarih içeren her satır bir harekettir; açıklaması ne olursa
   olsun listeye ekle. Sadece şunları atla:
   - DEVREDEN / AÇILIŞ / DÖNEM BAŞI / DÖNEM SONU / TOPLAM / BAKİYE / ARA TOPLAM satırları
   - Sayfa başlık/footer, imza, ""Sayfa 1/X"" satırları
   - Tarih kolonu boş satırlar

6. **SAYI FORMATI**. Türkçe ""1.234.567,89"" → 1234567.89. ""38.700.636,10 USD"" yazısındaki USD
   etiketini yok say. ""8245213.4"" zaten İngilizce ondalık.

7. **BORÇ + ALACAK** aynı anda dolu olamaz; biri 0'dır.

8. **HEPSİNİ ÇIKAR**. Ekstrede 30 hareket varsa 30 satır döndür — eksik bırakma.

ÖRNEK ÇIKTI (yukarıdaki TOYOTA ekstresi için):
[
  {""tarih"":""2025-05-20"",""borc"":8245213.40,""alacak"":0,""aciklama"":""Gönderilen Havale 190.725-euro TOYOTA"",""fisNo"":""000571658""},
  {""tarih"":""2025-06-30"",""borc"":633531.23,""alacak"":0,""aciklama"":""Y.içi Kur Değerleme"",""fisNo"":""00100570""},
  {""tarih"":""2025-07-04"",""borc"":0,""alacak"":2524413.58,""aciklama"":""Satınalma Faturası"",""fisNo"":""TIM2025000000285""}
]";

    /// <summary>
    /// .xls çoğu Türkçe muhasebe raporunda aslında HTML <table> dosyasıdır (Logo NetRapor,
    /// Netsis, Mikro vb. eski rapor çıktıları). Tag'leri ayıklayıp düz tablo metni üretir.
    /// Eğer içerik gerçek binary ise yine de okunabilir bir özet döner.
    /// </summary>
    private static string HtmlVeyaMetinOlarakOku(byte[] data)
    {
        // Encoding tespiti: önce UTF-8, replacement char yoksa kabul. Yoksa Default (Windows-1254 TR).
        string ham = System.Text.Encoding.UTF8.GetString(data);
        if (ham.Contains('�'))
        {
            try { ham = System.Text.Encoding.Default.GetString(data); }
            catch { /* UTF-8 zaten dolu */ }
        }

        // <br>, </tr> → satır sonuna dönüşsün
        ham = System.Text.RegularExpressions.Regex.Replace(ham, @"<br\s*/?>", "\n",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        ham = System.Text.RegularExpressions.Regex.Replace(ham, @"</tr\s*>", "\n",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        ham = System.Text.RegularExpressions.Regex.Replace(ham, @"</td\s*>", " | ",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        ham = System.Text.RegularExpressions.Regex.Replace(ham, @"</th\s*>", " | ",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        // Tüm HTML tag'lerini sil
        ham = System.Text.RegularExpressions.Regex.Replace(ham, @"<[^>]+>", " ");

        // HTML entity'leri çöz
        ham = System.Net.WebUtility.HtmlDecode(ham);

        // Çoklu boşlukları sadeleştir, ardışık boş satırları tekleştir
        ham = System.Text.RegularExpressions.Regex.Replace(ham, @"[ \t]+", " ");
        var satirlar = ham.Split('\n')
            .Select(s => s.Trim().Trim('|').Trim())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .ToList();

        // 100K karakteri geçmesin (token koruması)
        var sb = new StringBuilder();
        foreach (var s in satirlar)
        {
            if (sb.Length + s.Length > 100_000) { sb.AppendLine("... (kesildi)"); break; }
            sb.AppendLine(s);
        }
        return sb.ToString();
    }

    /// <summary>
    /// ExcelDataReader ile BIFF2..12 dahil tüm Excel formatlarını okur (Logo NetRapor BIFF5 dahil).
    /// CodePagesEncodingProvider Program.cs'te kayıtlı olmalı (Türkçe encoding için).
    /// </summary>
    private static string ExcelMetneCevir(byte[] data)
    {
        using var ms = new MemoryStream(data);
        using var reader = ExcelReaderFactory.CreateReader(ms);
        var sb = new StringBuilder();
        do
        {
            sb.AppendLine($"### Sayfa: {reader.Name}");
            while (reader.Read())
            {
                var hucreler = new List<string>(reader.FieldCount);
                bool doluVar = false;
                for (int c = 0; c < reader.FieldCount; c++)
                {
                    var v = reader.GetValue(c);
                    string s;
                    if (v is null) s = "";
                    else if (v is DateTime dt) s = dt.ToString("dd.MM.yyyy");
                    else if (v is double d)
                        s = d == Math.Truncate(d)
                            ? ((long)d).ToString()
                            : d.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    else s = v.ToString() ?? "";
                    if (!string.IsNullOrWhiteSpace(s)) doluVar = true;
                    hucreler.Add(s);
                }
                if (doluVar) sb.AppendLine(string.Join(" | ", hucreler));
            }
            sb.AppendLine();
            if (sb.Length > 100_000) { sb.AppendLine("... (kesildi)"); break; }
        } while (reader.NextResult());
        return sb.ToString();
    }

    private static string? JsonBlokunuCikar(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        // ```json ... ``` bloğunu temizle (alt sondaki ``` eksik olsa da çalışsın)
        var t = s.Trim();
        if (t.StartsWith("```"))
        {
            var ilkSatirSonu = t.IndexOf('\n');
            if (ilkSatirSonu > 0) t = t[(ilkSatirSonu + 1)..];
            var sonFence = t.LastIndexOf("```", StringComparison.Ordinal);
            if (sonFence >= 0) t = t[..sonFence];
            t = t.Trim();
        }

        int bas = t.IndexOf('[');
        if (bas < 0) return null;

        // 1) Normal yol: ilk '[' ve son ']'
        int son = t.LastIndexOf(']');
        if (son > bas)
        {
            var aday = t.Substring(bas, son - bas + 1);
            if (GecerliJsonMu(aday)) return aday;
        }

        // 2) max_tokens'a takılıp yanıt yarıda kesildiyse — son geçerli `}`'i bul,
        //    diziyi orada kapatıp `]` ekle. String içindeki '{' veya '}' karakterleri
        //    derinliği bozmasın diye string-aware tarama yap.
        return KesikJsonuKurtar(t, bas);
    }

    private static bool GecerliJsonMu(string s)
    {
        try { using var _ = JsonDocument.Parse(s); return true; }
        catch { return false; }
    }

    /// <summary>
    /// '[' ile başlayan bir JSON dizisi yarıda kesildiyse (max_tokens),
    /// son tamamlanmış obje sonuna kadar olan kısmı alıp `]` ile kapatır.
    /// String literal içindeki { } karakterleri sayılmaz (kaçışlı çift tırnak farkındalığı).
    /// </summary>
    private static string? KesikJsonuKurtar(string t, int bas)
    {
        int derinlik = 0;
        bool stringIcinde = false;
        bool kacis = false;
        int sonTamObjeSonu = -1;
        bool diziAcildi = false;

        for (int i = bas; i < t.Length; i++)
        {
            char c = t[i];
            if (kacis) { kacis = false; continue; }
            if (stringIcinde)
            {
                if (c == '\\') { kacis = true; continue; }
                if (c == '"') { stringIcinde = false; }
                continue;
            }
            switch (c)
            {
                case '"': stringIcinde = true; break;
                case '[': if (!diziAcildi) diziAcildi = true; break;
                case '{': derinlik++; break;
                case '}':
                    derinlik--;
                    if (derinlik == 0 && diziAcildi) sonTamObjeSonu = i;
                    break;
            }
        }

        if (sonTamObjeSonu < 0) return null;
        return t.Substring(bas, sonTamObjeSonu - bas + 1) + "]";
    }

    private static List<MutabakatKayit> JsonuKayitlaraCevir(string jsonDizi)
    {
        using var doc = JsonDocument.Parse(jsonDizi);
        if (doc.RootElement.ValueKind != JsonValueKind.Array)
            throw new InvalidOperationException("AI yanıtı JSON dizisi değil.");

        var list = new List<MutabakatKayit>();
        foreach (var e in doc.RootElement.EnumerateArray())
        {
            if (e.ValueKind != JsonValueKind.Object) continue;
            var k = new MutabakatKayit
            {
                Tarih    = AlTarih(e, "tarih"),
                Borc     = AlSayi(e, "borc"),
                Alacak   = AlSayi(e, "alacak"),
                Aciklama = AlMetin(e, "aciklama"),
                FisNo    = AlMetin(e, "fisNo"),
            };
            if (k.Tarih == default && k.Borc == 0 && k.Alacak == 0) continue;
            list.Add(k);
        }
        return list;
    }

    private static DateTime AlTarih(JsonElement e, string adi)
    {
        if (!e.TryGetProperty(adi, out var v)) return default;
        var s = v.GetString();
        if (string.IsNullOrWhiteSpace(s)) return default;
        if (DateTime.TryParse(s, System.Globalization.CultureInfo.InvariantCulture,
                              System.Globalization.DateTimeStyles.None, out var dt))
            return dt.Date;
        return default;
    }

    private static decimal AlSayi(JsonElement e, string adi)
    {
        if (!e.TryGetProperty(adi, out var v)) return 0;
        if (v.ValueKind == JsonValueKind.Number && v.TryGetDecimal(out var d)) return d;
        if (v.ValueKind == JsonValueKind.String)
        {
            var s = (v.GetString() ?? "").Replace(".", "").Replace(",", ".");
            if (decimal.TryParse(s, System.Globalization.NumberStyles.Any,
                                 System.Globalization.CultureInfo.InvariantCulture, out var x))
                return x;
        }
        return 0;
    }

    private static string AlMetin(JsonElement e, string adi)
    {
        if (!e.TryGetProperty(adi, out var v)) return "";
        return v.ValueKind == JsonValueKind.String ? (v.GetString() ?? "") : v.ToString();
    }
}
