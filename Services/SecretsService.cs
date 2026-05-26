using System.Text.Json;
using Microsoft.Data.SqlClient;

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
    public static string LogoUnityKullanici     => _cached?.LogoUnityKullanici     ?? "";
    public static string LogoUnitySifre         => _cached?.LogoUnitySifre         ?? "";

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
        _cached = Clone(m);
        SaveEncrypted(_cached);
    }

    // Sayfa düzenlemesi cache'i kirletmesin diye derin kopya döner.
    public static SecretsModel Snapshot() => Clone(_cached ?? new SecretsModel());

    private static SecretsModel Clone(SecretsModel m) => new()
    {
        LogoConnectionString   = m.LogoConnectionString,
        KantarConnectionString = m.KantarConnectionString,
        PmhsConnectionString   = m.PmhsConnectionString,
        AnthropicApiKey        = m.AnthropicApiKey,
        MailPassword           = m.MailPassword,
        LogoUnityKullanici     = m.LogoUnityKullanici,
        LogoUnitySifre         = m.LogoUnitySifre,
    };

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

                if (doc.RootElement.TryGetProperty("LogoUnity", out var lu))
                {
                    if (lu.TryGetProperty("Kullanici", out var u)) s.LogoUnityKullanici = u.GetString() ?? "";
                    if (lu.TryGetProperty("Sifre",     out var p)) s.LogoUnitySifre     = p.GetString() ?? "";
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
        // Logo Tiger Unity COM API login bilgileri — e-Fatura'yı Logo'ya aktarırken kullanılır.
        public string LogoUnityKullanici     { get; set; } = "";
        public string LogoUnitySifre         { get; set; } = "";
    }

    // Connection string'i parçalara ayırıp UI'da Server/Database/User/Password
    // alanları olarak düzenlemeyi mümkün kılar. SQL server taşındığında
    // (192.168.0.50\DOGUSNDB → 192.168.0.51\DOGUSNDB2) kullanıcı yine kendi başına düzeltebilsin.
    public class ConnectionParts
    {
        public string Server   { get; set; } = "";
        public string Database { get; set; } = "";
        public string UserId   { get; set; } = "";
        public string Password { get; set; } = "";

        public static ConnectionParts Parse(string cs)
        {
            var p = new ConnectionParts();
            if (string.IsNullOrWhiteSpace(cs)) return p;
            try
            {
                var b = new SqlConnectionStringBuilder(cs);
                p.Server   = b.DataSource    ?? "";
                p.Database = b.InitialCatalog ?? "";
                p.UserId   = b.UserID         ?? "";
                p.Password = b.Password       ?? "";
            }
            catch { /* bozuk cs — boş döner, kullanıcı UI'dan girer */ }
            return p;
        }

        public string Build() => new SqlConnectionStringBuilder
        {
            DataSource             = Server,
            InitialCatalog         = Database,
            UserID                 = UserId,
            Password               = Password,
            TrustServerCertificate = true,
            Encrypt                = false,
            ConnectTimeout         = 30,
        }.ConnectionString;

        public async Task<(bool ok, string message)> TestAsync()
        {
            try
            {
                using var con = new SqlConnection(Build());
                await con.OpenAsync();
                return (true, $"Bağlantı başarılı ({con.ServerVersion}).");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }
    }
}
