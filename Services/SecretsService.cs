using System.Text.Json;

namespace RaporlamaPortali.Services;

// Vault MDK ile şifrelenmiş secrets.enc dosyasını yönetir.
// İlk çalıştırmada (secrets.enc yoksa) appsettings.json'daki düz metin sırlardan
// migration yapar, sonra dosyayı şifreleyip yazar. Sonraki açılışlarda direkt okur.
// VaultService.Acik == true olmadan çağrılamaz.
public static class SecretsService
{
    private static SecretsModel? _cached;

    public static bool Yuklu => _cached != null;

    public static string LogoConnectionString  => _cached?.LogoConnectionString  ?? "";
    public static string KantarConnectionString => _cached?.KantarConnectionString ?? "";
    public static string PmhsConnectionString   => _cached?.PmhsConnectionString   ?? "";
    public static string AnthropicApiKey        => _cached?.AnthropicApiKey        ?? "";
    public static string MailPassword           => _cached?.MailPassword           ?? "";

    // Vault unlocked ise çağrılır. secrets.enc varsa decrypt eder; yoksa
    // appsettings.json'dan ilk değerleri toplayıp dosyayı oluşturur.
    public static void LoadOrCreate()
    {
        if (File.Exists(AppDataPaths.SecretsEnc))
        {
            _cached = Decrypt();
        }
        else
        {
            _cached = MigrateFromAppSettings();
            SaveEncrypted(_cached);
        }
    }

    public static void Save(SecretsModel m)
    {
        _cached = m;
        SaveEncrypted(m);
    }

    public static SecretsModel Snapshot() => _cached ?? new SecretsModel();

    // ---------- migration ----------

    private static SecretsModel MigrateFromAppSettings()
    {
        var s = new SecretsModel();
        try
        {
            var appsettingsPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");
            if (File.Exists(appsettingsPath))
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(appsettingsPath));
                if (doc.RootElement.TryGetProperty("ConnectionStrings", out var cs))
                {
                    if (cs.TryGetProperty("LogoDB", out var v))   s.LogoConnectionString   = v.GetString() ?? "";
                    if (cs.TryGetProperty("KantarDB", out var v2)) s.KantarConnectionString = v2.GetString() ?? "";
                    if (cs.TryGetProperty("PMHSDB", out var v3))   s.PmhsConnectionString   = v3.GetString() ?? "";
                }
                if (doc.RootElement.TryGetProperty("MailAyarlari", out var m)
                    && m.TryGetProperty("Sifre", out var sf))
                {
                    var pwd = sf.GetString() ?? "";
                    if (pwd != "OUTLOOK_SIFRENIZI_BURAYA_YAZIN") s.MailPassword = pwd;
                }
            }
        }
        catch { /* ilk kurulumda alanlar boş kalsın, kullanıcı sonra UI'dan girer */ }

        // Anthropic API key en son User env var'dan alınmıştı — ilk kurulumda buradan toplayalım.
        s.AnthropicApiKey =
              Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY", EnvironmentVariableTarget.User)
           ?? Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY", EnvironmentVariableTarget.Process)
           ?? Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY", EnvironmentVariableTarget.Machine)
           ?? "";

        return s;
    }

    // ---------- crypto I/O ----------

    private static SecretsModel Decrypt()
    {
        using var doc = JsonDocument.Parse(File.ReadAllText(AppDataPaths.SecretsEnc));
        var root  = doc.RootElement;
        var nonce  = Convert.FromBase64String(root.GetProperty("Nonce").GetString()!);
        var cipher = Convert.FromBase64String(root.GetProperty("Cipher").GetString()!);
        var tag    = Convert.FromBase64String(root.GetProperty("Tag").GetString()!);

        var plain = VaultService.DecryptWithMdk(nonce, cipher, tag);
        return JsonSerializer.Deserialize<SecretsModel>(plain) ?? new SecretsModel();
    }

    private static void SaveEncrypted(SecretsModel m)
    {
        var plain     = JsonSerializer.SerializeToUtf8Bytes(m);
        var (n, c, t) = VaultService.EncryptWithMdk(plain);
        var doc = new
        {
            Version    = 1,
            UpdatedUtc = DateTime.UtcNow.ToString("o"),
            Nonce      = Convert.ToBase64String(n),
            Cipher     = Convert.ToBase64String(c),
            Tag        = Convert.ToBase64String(t),
        };
        File.WriteAllText(AppDataPaths.SecretsEnc,
            JsonSerializer.Serialize(doc, new JsonSerializerOptions { WriteIndented = true }));
    }

    // ---------- model ----------

    public class SecretsModel
    {
        public string LogoConnectionString   { get; set; } = "";
        public string KantarConnectionString { get; set; } = "";
        public string PmhsConnectionString   { get; set; } = "";
        public string AnthropicApiKey        { get; set; } = "";
        public string MailPassword           { get; set; } = "";
    }
}
