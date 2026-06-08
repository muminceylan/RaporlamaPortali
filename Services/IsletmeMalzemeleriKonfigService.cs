using System.Text.Json;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// İşletme Malzemeleri Raporu konfigürasyonunu JSON dosyada tutar.
/// Singleton — uygulama ömrü boyunca tek instance, thread-safe IO.
/// </summary>
public class IsletmeMalzemeleriKonfigService
{
    private readonly string _path = AppDataPaths.IsletmeMalzemeleriJson;
    private readonly object _kilit = new();
    private IsletmeMalzemeleriKonfig _konfig = new();

    public IsletmeMalzemeleriKonfigService()
    {
        Yukle();
    }

    public IsletmeMalzemeleriKonfig Mevcut
    {
        get { lock (_kilit) return _konfig; }
    }

    private void Yukle()
    {
        lock (_kilit)
        {
            if (!File.Exists(_path))
            {
                _konfig = VarsayilanKonfig();
                Kaydet();
                return;
            }

            try
            {
                var json = File.ReadAllText(_path);
                var k = JsonSerializer.Deserialize<IsletmeMalzemeleriKonfig>(json);
                _konfig = k ?? VarsayilanKonfig();
                // Grup adlarını sabit tut (eski JSON'da boş kalmış olabilir)
                if (string.IsNullOrWhiteSpace(_konfig.Yakitlar.GrupAdi))    _konfig.Yakitlar.GrupAdi    = "Yakıtlar";
                if (string.IsNullOrWhiteSpace(_konfig.Torbalar.GrupAdi))    _konfig.Torbalar.GrupAdi    = "Torbalar Filtre Bezleri";
                if (string.IsNullOrWhiteSpace(_konfig.Kimyasallar.GrupAdi)) _konfig.Kimyasallar.GrupAdi = "Kimyasallar";
            }
            catch
            {
                _konfig = VarsayilanKonfig();
            }
        }
    }

    public void Kaydet()
    {
        lock (_kilit)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_path)!);
            var json = JsonSerializer.Serialize(_konfig, new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });
            File.WriteAllText(_path, json);
        }
    }

    public bool KodOnekiEkle(string grupAnahtari, string onek)
    {
        if (string.IsNullOrWhiteSpace(onek)) return false;
        var grup = GrupSec(grupAnahtari);
        if (grup == null) return false;
        var temiz = onek.Trim();
        if (grup.KodOnekleri.Any(x => string.Equals(x, temiz, StringComparison.OrdinalIgnoreCase)))
            return false;
        grup.KodOnekleri.Add(temiz);
        grup.KodOnekleri.Sort(StringComparer.OrdinalIgnoreCase);
        Kaydet();
        return true;
    }

    public bool KodOnekiCikar(string grupAnahtari, string onek)
    {
        var grup = GrupSec(grupAnahtari);
        if (grup == null) return false;
        var silinen = grup.KodOnekleri.RemoveAll(x => string.Equals(x, onek, StringComparison.OrdinalIgnoreCase));
        if (silinen == 0) return false;
        Kaydet();
        return true;
    }

    public void AciklamaYaz(string malzemeKodu, string aciklama)
    {
        if (string.IsNullOrWhiteSpace(malzemeKodu)) return;
        lock (_kilit)
        {
            if (string.IsNullOrWhiteSpace(aciklama))
                _konfig.MalzemeAciklamalari.Remove(malzemeKodu);
            else
                _konfig.MalzemeAciklamalari[malzemeKodu] = aciklama.Trim();
        }
        Kaydet();
    }

    public void KampanyaTarihleriKaydet(DateTime? bas, DateTime? bit)
    {
        lock (_kilit)
        {
            _konfig.KampanyaBaslangic = bas;
            _konfig.KampanyaBitis     = bit;
        }
        Kaydet();
    }

    private IsletmeMalzemeGrubu? GrupSec(string anahtar) => anahtar?.ToLowerInvariant() switch
    {
        "yakitlar"    => _konfig.Yakitlar,
        "torbalar"    => _konfig.Torbalar,
        "kimyasallar" => _konfig.Kimyasallar,
        _             => null
    };

    /// <summary>
    /// Excel "Malzeme 2025-2026 YILI.XLS" dosyasında geçen kod öneklerini varsayılan olarak
    /// kullanır. Kullanıcı sonradan UI'dan düzenleyebilir.
    /// </summary>
    private static IsletmeMalzemeleriKonfig VarsayilanKonfig() => new()
    {
        Yakitlar = new IsletmeMalzemeGrubu
        {
            GrupAdi     = "Yakıtlar",
            KodOnekleri = new() { "208.01", "204.01" }
        },
        Torbalar = new IsletmeMalzemeGrubu
        {
            GrupAdi     = "Torbalar Filtre Bezleri",
            KodOnekleri = new() { "301.01", "302.01", "399.01", "303.01", "202.01" }
        },
        Kimyasallar = new IsletmeMalzemeGrubu
        {
            GrupAdi     = "Kimyasallar",
            KodOnekleri = new() { "203.01", "299.01", "510.01" }
        }
    };
}
