using System.Text.Json;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// SabNetKANTAR SQL Server bağlantı bilgilerini JSON dosyasında saklar.
/// </summary>
public class SabNetBaglantiService
{
    private readonly string _dosyaYolu;
    private SabNetBaglantiAyari _ayar;
    private readonly object _kilit = new();

    public SabNetBaglantiService()
    {
        _dosyaYolu = AppDataPaths.SabNetBaglantiJson;
        _ayar = Yukle();
    }

    public SabNetBaglantiAyari GetAyar()
    {
        lock (_kilit) { return Klonla(_ayar); }
    }

    public string GetConnectionString()
    {
        var a = GetAyar();
        return $"Server={a.Server};Database={a.Database};User Id={a.Username};Password={a.Password};TrustServerCertificate=True;Connection Timeout=30;";
    }

    public async Task<bool> KaydetAsync(SabNetBaglantiAyari yeni)
    {
        lock (_kilit) { _ayar = Klonla(yeni); }
        try
        {
            var json = JsonSerializer.Serialize(_ayar, new JsonSerializerOptions { WriteIndented = true });
            await File.WriteAllTextAsync(_dosyaYolu, json);
            return true;
        }
        catch { return false; }
    }

    private SabNetBaglantiAyari Yukle()
    {
        try
        {
            if (File.Exists(_dosyaYolu))
            {
                var json = File.ReadAllText(_dosyaYolu);
                var a = JsonSerializer.Deserialize<SabNetBaglantiAyari>(json);
                if (a != null) return a;
            }
        }
        catch { }
        var v = new SabNetBaglantiAyari();
        try
        {
            File.WriteAllText(_dosyaYolu, JsonSerializer.Serialize(v, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { }
        return v;
    }

    private static SabNetBaglantiAyari Klonla(SabNetBaglantiAyari a) => new()
    {
        Server   = a.Server,
        Database = a.Database,
        Username = a.Username,
        Password = a.Password
    };
}
