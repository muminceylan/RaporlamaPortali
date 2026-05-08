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
        _http.Timeout = TimeSpan.FromMinutes(3);
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
            max_tokens = 8192,
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
TÜRK MUHASEBE EKSTRELERİNDE TİPİK KOLON SIRASI (sol→sağ):
  [Tarih] [Fiş No] [Açıklama] [Vade Tarihi] [TL Borç] [TL Alacak] [USD Borç] [USD Alacak] [Kur] [Bakiye]
═══════════════════════════════════════════════════════════════════════════════

KRİTİK KURALLAR:

1. **TARİH = İLK TARİH KOLONU**. Bir satırda iki tarih varsa (örn. 12.01.2026 ve 12.05.2026),
   İLKİNİ AL. Vade tarihi (genelde 4. kolon) işlem tarihi DEĞİLDİR.

2. **TL KOLONLARINI AL — USD/EUR DEĞİL**. Bir satırda hem TL hem USD/EUR varsa, sadece TL'yi al.
   Tipik durum: TL Borç=0, TL Alacak=570.162,85 ; USD Borç=0, USD Alacak=38.700.636,10
   → DOĞRU çıktı: borc=0, alacak=570162.85 (USD asla)

3. **BAKİYE KOLONLARINI YOK SAY**. ""Borç Bak."" / ""Alacak Bak."" / ""Bakiye"" / ""Kümülatif"" /
   ""Devir"" gibi kolonlar HAREKET DEĞİLDİR. Bunları borc/alacak olarak verme.

4. **HİÇBİR HAREKET SATIRINI ATLAMA**. Tarih içeren her satır bir harekettir; satırın açıklaması
   FATURANIZ, GELEN HAVALE, TARAFINIZDAN, ÇEK, vb. ne olursa olsun listeye ekle. Sadece şunları atla:
   - DEVREDEN / AÇILIŞ / DÖNEM BAŞI / DÖNEM SONU / TOPLAM / BAKİYE / ARA TOPLAM satırları
   - Sayfa başlık/footer, imza, ""Sayfa 1/X"" satırları
   - Tarih kolonu boş olan satırlar

5. **SAYI FORMATI**. Türkçe ""1.234.567,89"" → 1234567.89. ""38.700.636,10 USD"" yazısındaki USD
   etiketini yok say (zaten USD kolonu, bu satırda almıyorsun). ""1495492.07"" zaten İngilizce.

6. **BORÇ + ALACAK** aynı anda dolu olamaz; biri 0'dır.

7. **EKSTRE BAŞINDAKİ SATIR SAYISI**. Ekstrede 30 satır varsa 30 hareket çıktın olmalı (devreden hariç).
   Eksik bırakma — ""tipik bir kaç tane göstereyim"" yapma.

ÖRNEK ÇIKTI:
[
  {""tarih"":""2026-01-08"",""borc"":1495492.07,""alacak"":0,""aciklama"":""HV/8943/QNB DOĞUŞ ÇAY 34.737,12 USD KARŞ.ÖDEME"",""fisNo"":""HV8943""},
  {""tarih"":""2026-01-12"",""borc"":0,""alacak"":570162.85,""aciklama"":""FATURANIZ (1)"",""fisNo"":""AYK20260000000051""}
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
        // ```json ... ``` bloğunu temizle
        var t = s.Trim();
        if (t.StartsWith("```"))
        {
            var ilkSatirSonu = t.IndexOf('\n');
            if (ilkSatirSonu > 0) t = t[(ilkSatirSonu + 1)..];
            var sonFence = t.LastIndexOf("```", StringComparison.Ordinal);
            if (sonFence >= 0) t = t[..sonFence];
            t = t.Trim();
        }
        // İlk '[' ve son ']' arasını al
        int bas = t.IndexOf('[');
        int son = t.LastIndexOf(']');
        if (bas < 0 || son <= bas) return null;
        return t.Substring(bas, son - bas + 1);
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
