using System.Globalization;

namespace RaporlamaPortali.Services;

/// <summary>
/// Sesli komut akışında tarih parametrelerini parse eden ortak yardımcı.
/// Claude'a hem ISO date hem de Türkçe kısayol / serbest format kabul ettirebiliyoruz —
/// bu sınıf tüm formatları tek noktadan DateTime'a çevirir.
///
/// Desteklenen girdiler:
///   - ISO:           "2025-01-01", "2025-1-1"
///   - dd.MM.yyyy:    "01.01.2025", "1.1.2025"
///   - dd/MM/yyyy:    "01/01/2025"
///   - dd-MM-yyyy:    "01-01-2025"
///   - Türkçe ay:     "1 ocak 2025", "15 mart", "12 mayıs 2024"
///   - Kısayol token: "bugun"/"bugün", "dun"/"dün", "yarin"/"yarın",
///                    "bu_ay_basi", "yil_basi"/"yıl başı", "ay_basi"/"ay başı"
///   - Aralık enum:   "bu_ay", "gecen_ay", "bu_yil", "bu_hafta",
///                    "son_7_gun", "son_30_gun", "son_90_gun"
/// </summary>
public static class VoiceDateHelper
{
    private static readonly Dictionary<string, int> _aylar = new(StringComparer.OrdinalIgnoreCase)
    {
        ["ocak"] = 1, ["şubat"] = 2, ["subat"] = 2, ["mart"] = 3, ["nisan"] = 4,
        ["mayıs"] = 5, ["mayis"] = 5, ["haziran"] = 6, ["temmuz"] = 7, ["ağustos"] = 8,
        ["agustos"] = 8, ["eylül"] = 9, ["eylul"] = 9, ["ekim"] = 10, ["kasım"] = 11,
        ["kasim"] = 11, ["aralık"] = 12, ["aralik"] = 12,
    };

    /// <summary>
    /// Tek bir tarihi parse eder. Boş/null/anlaşılmayan girdi → null.
    /// </summary>
    public static DateTime? Parse(object? deger)
    {
        if (deger is null) return null;
        var s = deger.ToString();
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = s.Trim().ToLowerInvariant();

        var bugun = DateTime.Today;

        // Kısayol token'ları
        switch (s.Replace(" ", "_"))
        {
            case "bugun":
            case "bugün":             return bugun;
            case "dun":
            case "dün":               return bugun.AddDays(-1);
            case "yarin":
            case "yarın":             return bugun.AddDays(1);
            case "ay_basi":
            case "ay_başı":
            case "bu_ay_basi":
            case "bu_ay_başı":        return new DateTime(bugun.Year, bugun.Month, 1);
            case "yil_basi":
            case "yıl_başı":
            case "bu_yil_basi":
            case "bu_yıl_başı":       return new DateTime(bugun.Year, 1, 1);
            case "gecen_ay_basi":
            case "geçen_ay_başı":     return new DateTime(bugun.Year, bugun.Month, 1).AddMonths(-1);
            case "gecen_yil_basi":
            case "geçen_yıl_başı":    return new DateTime(bugun.Year - 1, 1, 1);
        }

        // ISO ve nokta/slash/tire ayraçları
        var fmtler = new[]
        {
            "yyyy-MM-dd", "yyyy-M-d",
            "dd.MM.yyyy", "d.M.yyyy",
            "dd/MM/yyyy", "d/M/yyyy",
            "dd-MM-yyyy", "d-M-yyyy",
        };
        if (DateTime.TryParseExact(s, fmtler, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out var dt))
            return dt;

        // Genel TryParse (TR culture — Claude bazen "1 Ocak 2025" gibi yazabilir)
        if (DateTime.TryParse(s, new CultureInfo("tr-TR"), DateTimeStyles.None, out dt))
            return dt;

        // "1 ocak 2025" / "15 mart" gibi serbest TR formatı
        var parcalar = s.Split(new[] { ' ', '.', '/', '-' },
            StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parcalar.Length >= 2)
        {
            int? gun = null, ay = null, yil = null;
            foreach (var p in parcalar)
            {
                if (_aylar.TryGetValue(p, out var ayNo)) { ay = ayNo; continue; }
                if (int.TryParse(p, out var n))
                {
                    if (n >= 1900) yil = n;
                    else if (gun == null) gun = n;
                    else if (ay == null && n is >= 1 and <= 12) ay = n;
                }
            }
            if (gun.HasValue && ay.HasValue)
            {
                var y = yil ?? bugun.Year;
                if (gun.Value is >= 1 and <= 31 && ay.Value is >= 1 and <= 12)
                {
                    try { return new DateTime(y, ay.Value, gun.Value); }
                    catch { return null; }
                }
            }
        }

        return null;
    }

    /// <summary>
    /// Aralık enum'unu (bu_ay, gecen_ay, bu_yil, son_7_gun vb.)
    /// (baslangic, bitis) tuple'ına çevirir. Anlaşılmayan girdi → (null, null).
    /// </summary>
    public static (DateTime? Bas, DateTime? Bit) AraligiCevir(string? kod)
    {
        if (string.IsNullOrWhiteSpace(kod)) return (null, null);
        var bugun = DateTime.Today;
        return kod.Trim().ToLowerInvariant() switch
        {
            "bugun" or "bugün"                  => ((DateTime?)bugun, (DateTime?)bugun),
            "dun" or "dün"                      => (bugun.AddDays(-1), bugun.AddDays(-1)),
            "bu_hafta" or "bu hafta"            => (bugun.AddDays(-(((int)bugun.DayOfWeek + 6) % 7)), bugun),
            "bu_ay" or "bu ay"                  => (new DateTime(bugun.Year, bugun.Month, 1), bugun),
            "gecen_ay" or "geçen_ay" or "geçen ay"
                => (new DateTime(bugun.Year, bugun.Month, 1).AddMonths(-1),
                    new DateTime(bugun.Year, bugun.Month, 1).AddDays(-1)),
            "bu_yil" or "bu_yıl" or "bu yıl"    => (new DateTime(bugun.Year, 1, 1), bugun),
            "gecen_yil" or "geçen_yıl" or "geçen yıl"
                => (new DateTime(bugun.Year - 1, 1, 1), new DateTime(bugun.Year - 1, 12, 31)),
            "son_7_gun" or "son 7 gün" or "son 7 gun"
                => (bugun.AddDays(-7), bugun),
            "son_30_gun" or "son 30 gün" or "son 30 gun"
                => (bugun.AddDays(-30), bugun),
            "son_90_gun" or "son 90 gün" or "son 90 gun"
                => (bugun.AddDays(-90), bugun),
            _ => (null, null),
        };
    }

    /// <summary>
    /// VoiceIntent.Params içinden baslangic/bitis okur ve mevcut alanları günceller.
    /// Eksik parametre eski değerini korur (hedef: kullanıcı sadece bitiş söylese de çalışsın).
    /// </summary>
    public static (DateTime? Bas, DateTime? Bit) ParseBaslangicBitis(
        IDictionary<string, object?> prms,
        DateTime? mevcutBas,
        DateTime? mevcutBit)
    {
        prms.TryGetValue("baslangic", out var basVal);
        prms.TryGetValue("bitis",     out var bitVal);
        var b = Parse(basVal);
        var t = Parse(bitVal);
        return (b ?? mevcutBas, t ?? mevcutBit);
    }
}
