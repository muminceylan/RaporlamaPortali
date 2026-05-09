using System.Text.Json;

namespace RaporlamaPortali.Services;

// Stok Durumu sayfasındaki "Afyon Ambarları" filtresinin altında yatan ambar numaraları
// listesi. Eskiden kaynak kodda sabitti — yeni ambarlar açıldıkça user tek başına ekleyip
// çıkarabilsin diye burada yönetilir. Liste C:\RaporlamaPortaliData\afyon_ambarlari.json
// dosyasında tutulur; ilk açılışta sabit kodlu varsayılan setle seed edilir.
public sealed class AfyonAmbarService
{
    private static readonly int[] Varsayilan =
    {
        142, 143, 144, 145, 146, 147, 148, 149, 150, 151, 152, 157, 159,
        712, 713, 714, 715, 716, 717, 718,
        750, 752, 753, 754, 755, 756,
        800, 801, 802, 803, 804, 805, 806, 807, 808, 810,
        830, 831, 832, 833, 834, 835, 836, 837, 838,
    };

    private readonly object _lock = new();
    private List<int> _liste = new();
    private bool _yuklendi = false;

    public IReadOnlyList<int> Liste
    {
        get
        {
            EnsureLoaded();
            lock (_lock) return _liste.OrderBy(x => x).ToList();
        }
    }

    public bool Ekle(int ambarNo)
    {
        EnsureLoaded();
        lock (_lock)
        {
            if (_liste.Contains(ambarNo)) return false;
            _liste.Add(ambarNo);
            Kaydet();
            return true;
        }
    }

    public bool Cikar(int ambarNo)
    {
        EnsureLoaded();
        lock (_lock)
        {
            if (!_liste.Remove(ambarNo)) return false;
            Kaydet();
            return true;
        }
    }

    public void Sifirla()
    {
        lock (_lock)
        {
            _liste = Varsayilan.ToList();
            Kaydet();
        }
    }

    private void EnsureLoaded()
    {
        if (_yuklendi) return;
        lock (_lock)
        {
            if (_yuklendi) return;
            try
            {
                var path = AppDataPaths.AfyonAmbarlariJson;
                if (File.Exists(path))
                {
                    var json = File.ReadAllText(path);
                    var arr  = JsonSerializer.Deserialize<int[]>(json);
                    _liste   = (arr ?? Varsayilan).Distinct().ToList();
                }
                else
                {
                    _liste = Varsayilan.ToList();
                    Kaydet();
                }
            }
            catch
            {
                _liste = Varsayilan.ToList();
            }
            _yuklendi = true;
        }
    }

    private void Kaydet()
    {
        try
        {
            Directory.CreateDirectory(AppDataPaths.DataRoot);
            var json = JsonSerializer.Serialize(_liste.OrderBy(x => x).ToArray(),
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(AppDataPaths.AfyonAmbarlariJson, json);
        }
        catch { /* yazma hatası uygulamayı düşürmesin */ }
    }
}
