using System.Security.Cryptography;
using System.Text.Json;

namespace RaporlamaPortali.Services;

public class GirisAyarlariModel
{
    public string KullaniciAdi { get; set; } = "mümin";
    /// <summary>
    /// "base64salt:base64hash" formatında PBKDF2 hash veya (geçiş için) düz metin.
    /// </summary>
    public string Sifre { get; set; } = "";
}

public class GirisAyarlariService
{
    private readonly string _dosyaYolu;
    private GirisAyarlariModel _ayarlar;
    private readonly object _kilit = new();

    public GirisAyarlariService()
    {
        _dosyaYolu = AppDataPaths.GirisAyarlariJson;
        _ayarlar = Yukle();
    }

    public GirisAyarlariModel GetAyarlar() => _ayarlar;

    public bool GirisKontrol(string kullanici, string sifre)
    {
        if (!string.Equals(_ayarlar.KullaniciAdi.Trim(), kullanici.Trim(), StringComparison.OrdinalIgnoreCase))
            return false;

        // Düz metin varsa → doğrula, hash'le ve kaydet (geçiş)
        if (!HashFormatMi(_ayarlar.Sifre))
        {
            if (_ayarlar.Sifre != sifre) return false;
            // Başarılı giriş — şifreyi hash'le
            lock (_kilit) { _ayarlar.Sifre = SifreHashle(sifre); }
            _ = KaydetAsync();
            return true;
        }

        return SifreDogrula(sifre, _ayarlar.Sifre);
    }

    public async Task<bool> SifreDegistirAsync(string yeniKullanici, string yeniSifre)
    {
        lock (_kilit)
        {
            _ayarlar.KullaniciAdi = yeniKullanici.Trim();
            _ayarlar.Sifre = SifreHashle(yeniSifre);
        }
        return await KaydetAsync();
    }

    // ── Hash yardımcıları ────────────────────────────────────────────

    private static bool HashFormatMi(string deger) => deger.Contains(':') && deger.Length > 40;

    private static string SifreHashle(string sifre)
    {
        var salt = RandomNumberGenerator.GetBytes(16);
        var hash = Rfc2898DeriveBytes.Pbkdf2(sifre, salt, 200_000, HashAlgorithmName.SHA256, 32);
        return Convert.ToBase64String(salt) + ":" + Convert.ToBase64String(hash);
    }

    private static bool SifreDogrula(string sifre, string kayitliHash)
    {
        try
        {
            var p = kayitliHash.Split(':');
            if (p.Length != 2) return false;
            var salt = Convert.FromBase64String(p[0]);
            var hash = Convert.FromBase64String(p[1]);
            var test = Rfc2898DeriveBytes.Pbkdf2(sifre, salt, 200_000, HashAlgorithmName.SHA256, 32);
            return CryptographicOperations.FixedTimeEquals(hash, test);
        }
        catch { return false; }
    }

    // ── Dosya işlemleri ──────────────────────────────────────────────

    private async Task<bool> KaydetAsync()
    {
        try
        {
            var json = JsonSerializer.Serialize(_ayarlar, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(_dosyaYolu, json);
            return true;
        }
        catch { return false; }
    }

    private GirisAyarlariModel Yukle()
    {
        try
        {
            if (File.Exists(_dosyaYolu))
            {
                var json = File.ReadAllText(_dosyaYolu);
                return JsonSerializer.Deserialize<GirisAyarlariModel>(json) ?? Varsayilan();
            }
        }
        catch { }
        var v = Varsayilan();
        File.WriteAllText(_dosyaYolu, JsonSerializer.Serialize(v, new JsonSerializerOptions { WriteIndented = true }));
        return v;
    }

    // Varsayılan: şifre düz metin — ilk girişte otomatik hash'lenir
    private static GirisAyarlariModel Varsayilan() => new()
    {
        KullaniciAdi = "mümin",
        Sifre = "Mc162909119324"
    };
}
