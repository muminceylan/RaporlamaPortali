using System.Text.Json;

namespace RaporlamaPortali.Services;

public class GirisAyarlariModel
{
    public string KullaniciAdi { get; set; } = "mümin";
    public string Sifre { get; set; } = "Mc162909119324";
}

public class GirisAyarlariService
{
    private readonly string _dosyaYolu;
    private GirisAyarlariModel _ayarlar;
    private readonly object _kilit = new();

    public GirisAyarlariService()
    {
        _dosyaYolu = Path.Combine(AppContext.BaseDirectory, "giris_ayarlari.json");
        _ayarlar = Yukle();
    }

    public GirisAyarlariModel GetAyarlar() => _ayarlar;

    public bool GirisKontrol(string kullanici, string sifre)
        => string.Equals(_ayarlar.KullaniciAdi.Trim(), kullanici.Trim(), StringComparison.OrdinalIgnoreCase)
           && _ayarlar.Sifre == sifre;

    public async Task<bool> SifreDegistirAsync(string yeniKullanici, string yeniSifre)
    {
        lock (_kilit)
        {
            _ayarlar.KullaniciAdi = yeniKullanici;
            _ayarlar.Sifre = yeniSifre;
        }
        return await KaydetAsync();
    }

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

    private static GirisAyarlariModel Varsayilan() => new()
    {
        KullaniciAdi = "mümin",
        Sifre = "Mc162909119324"
    };
}
