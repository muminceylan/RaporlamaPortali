using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Anthropic Claude API ile sesli komut → yapısal intent dönüşümü.
/// Tool use mekanizmasıyla Claude'dan JSON yerine doğrudan tool çağrısı isteriz —
/// model komutu hangi sayfanın aksiyonuna eşleştirdiğini tool seçimiyle bildirir.
///
/// API key mevcut ClaudeService ile aynı env var üzerinden okunur (ANTHROPIC_API_KEY).
/// Model: claude-opus-4-7 (Anthropic CLI skill default, kompleks komutlar için).
/// Adaptive thinking + effort=high ile daha doğru anlama.
/// </summary>
public class VoiceCommandService
{
    private readonly HttpClient _http;
    private readonly ILogger<VoiceCommandService> _log;
    private readonly string? _apiKey;
    private const string ApiUrl = "https://api.anthropic.com/v1/messages";
    private const string Model  = "claude-opus-4-7";

    public VoiceCommandService(IHttpClientFactory httpFactory, ILogger<VoiceCommandService> log)
    {
        _http = httpFactory.CreateClient();
        _http.Timeout = TimeSpan.FromSeconds(45);
        _log = log;
        _apiKey = Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY", EnvironmentVariableTarget.User)
                  ?? Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY", EnvironmentVariableTarget.Process)
                  ?? Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY", EnvironmentVariableTarget.Machine);
    }

    public bool ApiKeyVar => !string.IsNullOrWhiteSpace(_apiKey);

    /// <summary>
    /// Kullanıcının söylediği metni (Türkçe) + sayfanın desteklediği aksiyon listesini alır,
    /// Claude'a "hangi aksiyon, hangi parametreler" sorar, intent olarak döner.
    /// </summary>
    /// <param name="komutMetni">STT'den gelen ham Türkçe metin.</param>
    /// <param name="sayfaAdi">Aksiyonların ait olduğu sayfa (Claude'a sunulan context).</param>
    /// <param name="toolListesi">Sayfanın desteklediği aksiyonlar.</param>
    /// <param name="ekContext">Opsiyonel — Claude'a verilen ek bilgi (ör. son fişler listesi).</param>
    public async Task<VoiceIntent> KomutuYorumlaAsync(
        string komutMetni,
        string sayfaAdi,
        IReadOnlyList<SesliKomutTool> toolListesi,
        string? ekContext = null,
        CancellationToken ct = default)
    {
        if (!ApiKeyVar)
            return new VoiceIntent
            {
                Anlasildi = false,
                Hata      = "ANTHROPIC_API_KEY env değişkeni yok. Sırlar sayfasından gir.",
            };

        if (string.IsNullOrWhiteSpace(komutMetni))
            return new VoiceIntent { Anlasildi = false, Hata = "Boş komut." };

        var tools = ToolListesiniClaudeFormatinaCevir(toolListesi);

        var systemPrompt = SystemPromptOlustur(sayfaAdi, ekContext);

        // Anthropic API body
        // Önemli: tool_choice forcing (any / belirli tool) ile thinking aynı anda kullanılamıyor —
        // Anthropic kısıtı. Intent çıkarımı basit sınıflandırma olduğundan thinking'e gerek yok.
        var body = new
        {
            model      = Model,
            max_tokens = 1500,
            system = new object[]
            {
                new {
                    type = "text",
                    text = systemPrompt,
                    cache_control = new { type = "ephemeral" }
                }
            },
            output_config = new { effort = "medium" },
            tools = tools,
            tool_choice = new { type = "any" },
            messages = new object[]
            {
                new { role = "user", content = komutMetni }
            }
        };

        var json = JsonSerializer.Serialize(body, new JsonSerializerOptions
        {
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull,
        });

        using var req = new HttpRequestMessage(HttpMethod.Post, ApiUrl);
        req.Headers.Add("x-api-key", _apiKey);
        req.Headers.Add("anthropic-version", "2023-06-01");
        req.Content = new StringContent(json, Encoding.UTF8, "application/json");

        try
        {
            using var resp = await _http.SendAsync(req, ct);
            var respText = await resp.Content.ReadAsStringAsync(ct);
            if (!resp.IsSuccessStatusCode)
            {
                _log.LogWarning("Claude API hata ({Status}): {Body}", (int)resp.StatusCode, respText);
                return new VoiceIntent
                {
                    Anlasildi = false,
                    Hata      = $"API hata ({(int)resp.StatusCode}): {KisaltHata(respText)}",
                };
            }

            return ResponseuIntenteCevir(respText);
        }
        catch (TaskCanceledException)
        {
            return new VoiceIntent { Anlasildi = false, Hata = "API zaman aşımı." };
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "VoiceCommand API çağrısı hata");
            return new VoiceIntent { Anlasildi = false, Hata = ex.Message };
        }
    }

    private static string SystemPromptOlustur(string sayfaAdi, string? ekContext)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Sen RaporlamaPortali Blazor web uygulamasının sesli asistanısın.");
        sb.AppendLine($"Şu an kullanıcı '{sayfaAdi}' sayfasındadır. Kullanıcı Türkçe konuşuyor.");
        sb.AppendLine();
        sb.AppendLine("GÖREVİN: Kullanıcının söylediği komutu anla ve sana sunulan tool'lardan UYGUN OLANI çağır.");
        sb.AppendLine("Tool parametrelerini komuttan çıkar. Sayısal değerleri Türkçe'den ondalık biçime çevir");
        sb.AppendLine("(\"yirmi beş kilo\" → 25, \"iki yüz elli\" → 250, \"bin\" → 1000).");
        sb.AppendLine();
        sb.AppendLine("Her tool çağrısının `_onay_metni` parametresinde, kullanıcıya gösterilecek 1 cümlelik onay");
        sb.AppendLine("Türkçe açıklaması yaz — örnek: \"Şu işlemi yapacağım: 152 ambardan dünkü üretim fişini kopyala\".");
        sb.AppendLine("Kullanıcı bu cümleyi okuyup onaylayacak.");
        sb.AppendLine();
        sb.AppendLine("KURALLAR:");
        sb.AppendLine("- Komut belirsizse, en yakın tool'u seç ve `_onay_metni`'nde varsayımını belirt.");
        sb.AppendLine("- Komut hiçbir tool'la eşleşmiyorsa, `bilinmeyen_komut` tool'unu çağır ve sebebini yaz.");
        sb.AppendLine("- Sadece tool çağır, normal metin yanıt verme.");
        sb.AppendLine();
        sb.AppendLine("ÇOKLU KOMUT (önemli!):");
        sb.AppendLine("- Kullanıcı tek nefeste birden fazla işlem söylerse (örn '1. satırı 10000 yap birim KG yap'),");
        sb.AppendLine("  `komut_zinciri` tool'unu çağır. İçinde sırayla yapılacak aksiyonları liste olarak ver.");
        sb.AppendLine("- Tek bir işlem varsa direkt o aksiyonun tool'unu çağır, `komut_zinciri`'ne sokma.");

        if (!string.IsNullOrWhiteSpace(ekContext))
        {
            sb.AppendLine();
            sb.AppendLine("--- SAYFA BAĞLAMI ---");
            sb.AppendLine(ekContext);
        }

        return sb.ToString();
    }

    private static List<object> ToolListesiniClaudeFormatinaCevir(IReadOnlyList<SesliKomutTool> tools)
    {
        var liste = new List<object>();

        foreach (var t in tools)
        {
            var props = new Dictionary<string, object>();
            var required = new List<string>();

            foreach (var (ad, p) in t.Parametreler)
            {
                var schema = new Dictionary<string, object>
                {
                    ["type"]        = p.Tip,
                    ["description"] = p.Aciklama,
                };
                if (p.Secenekler != null && p.Secenekler.Count > 0)
                    schema["enum"] = p.Secenekler;
                props[ad] = schema;
                if (p.Zorunlu) required.Add(ad);
            }

            // Her tool'a _onay_metni parametresi otomatik eklenir
            props["_onay_metni"] = new Dictionary<string, object>
            {
                ["type"]        = "string",
                ["description"] = "Kullanıcıya gösterilecek onay cümlesi (Türkçe, 1 cümle).",
            };
            required.Add("_onay_metni");

            liste.Add(new
            {
                name         = t.Ad,
                description  = t.Aciklama,
                input_schema = new
                {
                    type       = "object",
                    properties = props,
                    required   = required,
                }
            });
        }

        // Komut zinciri — kullanıcı tek nefeste birden fazla işlem söylerse
        liste.Add(new
        {
            name        = "komut_zinciri",
            description = "Tek konuşmada birden fazla işlem yapılacağında. İçindeki aksiyonlar sırayla çalıştırılır. "
                        + "Örnek: '1. satırı 10000 yap birim KG yap' → komut_zinciri([miktar_degistir, birim_degistir]).",
            input_schema = new
            {
                type       = "object",
                properties = new Dictionary<string, object>
                {
                    ["adimlar"] = new Dictionary<string, object>
                    {
                        ["type"]        = "array",
                        ["description"] = "Sırayla çalıştırılacak komutlar. Her eleman {action: tool_adi, params: {...}}.",
                        ["items"] = new Dictionary<string, object>
                        {
                            ["type"] = "object",
                            ["properties"] = new Dictionary<string, object>
                            {
                                ["action"] = new Dictionary<string, object>
                                {
                                    ["type"]        = "string",
                                    ["description"] = "Yukarıdaki tool'lardan birinin adı (örn 'miktar_degistir'). 'sayfa_git' ve 'komut_zinciri' KULLANMA."
                                },
                                ["params"] = new Dictionary<string, object>
                                {
                                    ["type"]        = "object",
                                    ["description"] = "O tool'un parametreleri."
                                }
                            },
                            ["required"] = new[] { "action" }
                        }
                    },
                    ["_onay_metni"] = new Dictionary<string, object>
                    {
                        ["type"]        = "string",
                        ["description"] = "Kullanıcıya gösterilecek tüm zincirin özet açıklaması."
                    }
                },
                required = new[] { "adimlar", "_onay_metni" }
            }
        });

        // Genel "anlayamadım" tool'u — Claude bunu çağırırsa UI uyarı gösterir
        liste.Add(new
        {
            name        = "bilinmeyen_komut",
            description = "Kullanıcının söylediği komut bu sayfada desteklenen aksiyonlarla eşleşmiyorsa bunu kullan.",
            input_schema = new
            {
                type       = "object",
                properties = new Dictionary<string, object>
                {
                    ["sebep"] = new Dictionary<string, object>
                    {
                        ["type"]        = "string",
                        ["description"] = "Komutun neden eşleşmediği — örn 'bu sayfada bu işlem yok' veya 'komut belirsiz'.",
                    },
                    ["_onay_metni"] = new Dictionary<string, object>
                    {
                        ["type"]        = "string",
                        ["description"] = "Kullanıcıya gösterilecek 1 cümlelik açıklama.",
                    }
                },
                required = new[] { "sebep", "_onay_metni" }
            }
        });

        return liste;
    }

    private static VoiceIntent ResponseuIntenteCevir(string responseJson)
    {
        try
        {
            using var doc = JsonDocument.Parse(responseJson);
            var content = doc.RootElement.GetProperty("content");
            foreach (var blok in content.EnumerateArray())
            {
                if (!blok.TryGetProperty("type", out var tipEl)) continue;
                if (tipEl.GetString() != "tool_use") continue;

                var toolAdi  = blok.GetProperty("name").GetString() ?? "";
                var input    = blok.GetProperty("input");
                var prms     = new Dictionary<string, object?>();
                string onay  = "";

                foreach (var p in input.EnumerateObject())
                {
                    if (p.Name == "_onay_metni")
                    {
                        onay = p.Value.GetString() ?? "";
                        continue;
                    }
                    prms[p.Name] = JsonElementToObject(p.Value);
                }

                if (toolAdi == "bilinmeyen_komut")
                {
                    return new VoiceIntent
                    {
                        Action     = "bilinmeyen_komut",
                        Anlasildi  = false,
                        Hata       = prms.TryGetValue("sebep", out var s) ? s?.ToString() : "Komut anlaşılamadı.",
                        OnayMetni  = onay,
                    };
                }

                if (toolAdi == "komut_zinciri")
                {
                    var adimlar = new List<VoiceIntent>();
                    if (input.TryGetProperty("adimlar", out var adimlarEl) && adimlarEl.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var a in adimlarEl.EnumerateArray())
                        {
                            if (a.ValueKind != JsonValueKind.Object) continue;
                            if (!a.TryGetProperty("action", out var actEl)) continue;
                            var altAct = actEl.GetString() ?? "";
                            if (string.IsNullOrWhiteSpace(altAct)) continue;
                            var altParams = new Dictionary<string, object?>();
                            if (a.TryGetProperty("params", out var pEl) && pEl.ValueKind == JsonValueKind.Object)
                            {
                                foreach (var pp in pEl.EnumerateObject())
                                    altParams[pp.Name] = JsonElementToObject(pp.Value);
                            }
                            adimlar.Add(new VoiceIntent
                            {
                                Action    = altAct,
                                Params    = altParams,
                                OnayMetni = altAct,
                                Anlasildi = true,
                            });
                        }
                    }
                    return new VoiceIntent
                    {
                        Action    = "komut_zinciri",
                        OnayMetni = string.IsNullOrWhiteSpace(onay) ? $"{adimlar.Count} işlem zinciri" : onay,
                        Adimlar   = adimlar,
                        Anlasildi = true,
                    };
                }

                return new VoiceIntent
                {
                    Action    = toolAdi,
                    Params    = prms,
                    OnayMetni = string.IsNullOrWhiteSpace(onay)
                        ? $"Yapılacak işlem: {toolAdi}"
                        : onay,
                    Anlasildi = true,
                };
            }

            return new VoiceIntent
            {
                Anlasildi = false,
                Hata      = "Claude tool çağrısı yapmadı — yanıtı yorumlanamadı.",
            };
        }
        catch (Exception ex)
        {
            return new VoiceIntent
            {
                Anlasildi = false,
                Hata      = "Yanıt parse hatası: " + ex.Message,
            };
        }
    }

    private static object? JsonElementToObject(JsonElement el) => el.ValueKind switch
    {
        JsonValueKind.String => el.GetString(),
        JsonValueKind.Number => el.TryGetInt64(out var l) ? l : (object)el.GetDouble(),
        JsonValueKind.True   => true,
        JsonValueKind.False  => false,
        JsonValueKind.Null   => null,
        JsonValueKind.Array  => el.EnumerateArray().Select(JsonElementToObject).ToList(),
        _ => el.GetRawText(),
    };

    private static string KisaltHata(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        return s.Length > 300 ? s[..300] + "..." : s;
    }
}
