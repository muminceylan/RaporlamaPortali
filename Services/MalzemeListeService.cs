using System.Text.Json;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Kullanıcının "Malzeme Hareket Listesi" sayfasında kaydettiği malzeme kodu listelerini
/// JSON dosyasında tutar. Kullanıcı yeni liste ekleyebilir, mevcut listeyi açıp çalıştırabilir,
/// silebilir.
/// </summary>
public class MalzemeListeService
{
    private readonly string _dosyaYolu;
    private readonly ILogger<MalzemeListeService> _logger;
    private readonly SemaphoreSlim _kilit = new(1, 1);
    private List<KayitliMalzemeListesi> _listeler;

    public MalzemeListeService(ILogger<MalzemeListeService> logger)
    {
        _logger = logger;
        _dosyaYolu = AppDataPaths.MalzemeListeleriJson;
        _listeler  = YukleYadaOlustur();
    }

    public IReadOnlyList<KayitliMalzemeListesi> Listele() => _listeler;

    public KayitliMalzemeListesi? Getir(string ad)
        => _listeler.FirstOrDefault(l => string.Equals(l.Ad, ad, StringComparison.OrdinalIgnoreCase));

    public async Task<bool> KaydetAsync(string ad, List<string> kodlar)
    {
        if (string.IsNullOrWhiteSpace(ad)) return false;
        var temizKodlar = kodlar
            .Where(k => !string.IsNullOrWhiteSpace(k))
            .Select(k => k.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        await _kilit.WaitAsync();
        try
        {
            var mevcut = _listeler.FirstOrDefault(l => string.Equals(l.Ad, ad, StringComparison.OrdinalIgnoreCase));
            if (mevcut != null)
            {
                mevcut.MalzemeKodlari  = temizKodlar;
                mevcut.GuncellemeTarihi = DateTime.Now;
            }
            else
            {
                _listeler.Add(new KayitliMalzemeListesi
                {
                    Ad              = ad.Trim(),
                    MalzemeKodlari  = temizKodlar,
                    OlusturmaTarihi = DateTime.Now,
                    GuncellemeTarihi= DateTime.Now
                });
            }
            await YazAsync();
            return true;
        }
        finally { _kilit.Release(); }
    }

    public async Task<bool> SilAsync(string ad)
    {
        await _kilit.WaitAsync();
        try
        {
            var silinen = _listeler.RemoveAll(l => string.Equals(l.Ad, ad, StringComparison.OrdinalIgnoreCase));
            if (silinen > 0) await YazAsync();
            return silinen > 0;
        }
        finally { _kilit.Release(); }
    }

    private List<KayitliMalzemeListesi> YukleYadaOlustur()
    {
        try
        {
            if (!File.Exists(_dosyaYolu)) return new();
            var json = File.ReadAllText(_dosyaYolu);
            return JsonSerializer.Deserialize<List<KayitliMalzemeListesi>>(json) ?? new();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Malzeme listeleri yüklenemedi — boş listeyle başlatılıyor");
            return new();
        }
    }

    private async Task YazAsync()
    {
        var json = JsonSerializer.Serialize(_listeler, new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        });
        await File.WriteAllTextAsync(_dosyaYolu, json);
    }
}
