using System.Text.Json;

namespace RaporlamaPortali.Services;

/// <summary>
/// Ödeme dosyalarının şablonlarına yazılan banka hesap bilgileri.
/// Garanti şube değişikliği gibi durumlarda kullanıcı UI'dan değiştirsin diye
/// kalıcı JSON dosyasında tutulur: %DataRoot%/banka_hesap_ayarlari.json
/// </summary>
public class BankaHesapAyarlari
{
    // --- Garanti BBVA ---
    public string GarantiBorcluSube  { get; set; } = "1685";          // A sütunu (örn 1608, 1685…)
    public string GarantiBorcluHesap { get; set; } = "6297391";       // B sütunu

    // --- İş Bankası ---
    public string IsGondernSube  { get; set; } = "1256";              // C sütunu (Gönderen Şube)
    public string IsGondernHesap { get; set; } = "228393";            // D sütunu (Gönderen Hesap)
    public string IsGondernIban  { get; set; } = "TR590006400000112560228393"; // E sütunu

    // --- Ziraat Bankası ---
    public string ZiraatKurumKodu     { get; set; } = "12025";        // D5 hücresi
    public string ZiraatKurumHesapNo  { get; set; } = "1795-43297028-5077"; // D11 hücresi

    // ----- Persistence -----
    private static string FilePath =>
        Path.Combine(AppDataPaths.DataRoot, "banka_hesap_ayarlari.json");

    public static BankaHesapAyarlari Yukle()
    {
        try
        {
            if (!File.Exists(FilePath)) return new BankaHesapAyarlari();
            var json = File.ReadAllText(FilePath);
            var obj  = JsonSerializer.Deserialize<BankaHesapAyarlari>(json);
            return obj ?? new BankaHesapAyarlari();
        }
        catch { return new BankaHesapAyarlari(); }
    }

    public void Kaydet()
    {
        var dir = Path.GetDirectoryName(FilePath);
        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) Directory.CreateDirectory(dir);
        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(FilePath, json);
    }
}
