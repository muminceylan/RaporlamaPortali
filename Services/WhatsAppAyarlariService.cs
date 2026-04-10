using System.Text.Json;
using QRCoder;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

public class WhatsAppAyarlariService
{
    private readonly string _configDosya;
    private readonly string _durumDosya;
    private readonly JsonSerializerOptions _json = new()
    {
        WriteIndented        = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase  // Node.js uyumlu: yetkiliNumaralar
    };
    private readonly JsonSerializerOptions _jsonRead = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public WhatsAppAyarlariService()
    {
        var klasor   = WhatsAppKlasor();
        _configDosya = Path.Combine(klasor, "whatsapp-config.json");
        _durumDosya  = Path.Combine(klasor, "whatsapp-status.json");
    }

    // Exe'nin yanındaki WhatsApp klasörü
    public static string WhatsAppKlasor()
    {
        var exeDir = Path.GetDirectoryName(Environment.ProcessPath)
                     ?? AppContext.BaseDirectory;
        return Path.Combine(exeDir, "WhatsApp");
    }

    public WhatsAppAyarlariModel GetAyarlar()
    {
        try
        {
            if (File.Exists(_configDosya))
            {
                var json = File.ReadAllText(_configDosya);
                return JsonSerializer.Deserialize<WhatsAppAyarlariModel>(json, _jsonRead) ?? new();
            }
        }
        catch { }
        return new WhatsAppAyarlariModel();
    }

    public void Kaydet(WhatsAppAyarlariModel model)
    {
        var klasor = Path.GetDirectoryName(_configDosya)!;
        Directory.CreateDirectory(klasor);
        File.WriteAllText(_configDosya, JsonSerializer.Serialize(model, _json));
    }

    /// <summary>Ham QR string'ini PNG base64 data URL'ye çevirir</summary>
    public static string QrStringToDataUrl(string qrString)
    {
        if (string.IsNullOrWhiteSpace(qrString)) return "";
        try
        {
            using var qrGenerator = new QRCodeGenerator();
            var data    = qrGenerator.CreateQrCode(qrString, QRCodeGenerator.ECCLevel.M);
            var qrCode  = new PngByteQRCode(data);
            var bytes   = qrCode.GetGraphic(6); // 6px per module
            return "data:image/png;base64," + Convert.ToBase64String(bytes);
        }
        catch { return ""; }
    }

    public List<WhatsAppLogKayit> GetLoglar()
    {
        try
        {
            var logDosya = Path.Combine(WhatsAppKlasor(), "whatsapp-log.json");
            if (File.Exists(logDosya))
                return JsonSerializer.Deserialize<List<WhatsAppLogKayit>>(File.ReadAllText(logDosya), _jsonRead) ?? new();
        }
        catch { }
        return new();
    }

    public WhatsAppDurumModel GetDurum()
    {
        try
        {
            if (File.Exists(_durumDosya))
            {
                var json = File.ReadAllText(_durumDosya);
                var dict = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json);
                if (dict != null)
                {
                    return new WhatsAppDurumModel
                    {
                        Durum    = dict.TryGetValue("durum",    out var d) ? d.GetString() ?? "BAGLI_DEGIL" : "BAGLI_DEGIL",
                        QrString = dict.TryGetValue("qrString", out var q) ? q.GetString() ?? "" : "",
                        Guncelleme = dict.TryGetValue("guncelleme", out var g)
                            ? (DateTime.TryParse(g.GetString(), out var dt) ? dt : DateTime.MinValue)
                            : DateTime.MinValue
                    };
                }
            }
        }
        catch { }
        return new WhatsAppDurumModel();
    }
}
