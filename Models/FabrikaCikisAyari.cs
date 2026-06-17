using System.Globalization;

namespace RaporlamaPortali.Models;

/// <summary>
/// Excel R1'deki "Fabrika Çıkışı" değerine göre Logo satış siparişi header'ında
/// kullanılacak İŞ YERİ (BRANCH), BÖLÜM (DEPARTMENT), FABRİKA (FACTORY) ve
/// AMBAR (SOURCE_WH) değerleri.
///
/// Kullanıcı tablosu (12.06.2026):
///   PENDİK    → 001 / 000 / 000 / 534
///   KARACABEY → 000 / 000 / 026 / 784
///   AKSARAY   → 000 / 000 / 013 / 113
///   ŞANLIURFA → 001 / 000 / 000 / 570
///   ÖDEMİŞ    → 000 / 000 / 016 / 079
/// </summary>
public class FabrikaCikisAyari
{
    public string Ad         { get; private set; } = "";
    public int    Branch     { get; private set; }   // İŞ YERİ
    public int    Department { get; private set; }   // BÖLÜM
    public int    FactoryNr  { get; private set; }   // FABRİKA
    public int    SourceWh   { get; private set; }   // AMBAR

    private FabrikaCikisAyari(string ad, int branch, int dept, int factory, int wh)
    { Ad = ad; Branch = branch; Department = dept; FactoryNr = factory; SourceWh = wh; }

    // Sabit tablo — anahtar normalize edilmiş (Türkçe karakter + büyük harf).
    private static readonly Dictionary<string, FabrikaCikisAyari> _tablo = new()
    {
        ["PENDIK"]    = new FabrikaCikisAyari("PENDİK",    1, 0,  0, 534),
        ["KARACABEY"] = new FabrikaCikisAyari("KARACABEY", 0, 0, 26, 784),
        ["AKSARAY"]   = new FabrikaCikisAyari("AKSARAY",   0, 0, 13, 113),
        ["SANLIURFA"] = new FabrikaCikisAyari("ŞANLIURFA", 1, 0,  0, 570),
        ["ODEMIS"]    = new FabrikaCikisAyari("ÖDEMİŞ",    0, 0, 16,  79),
    };

    /// <summary>Excel'den okunan ham metni normalize ederek (Türkçe karakter + büyük harf) çözer.</summary>
    public static FabrikaCikisAyari? Coz(string? ham)
    {
        if (string.IsNullOrWhiteSpace(ham)) return null;
        var key = NormalizeAd(ham);
        _tablo.TryGetValue(key, out var a);
        return a;
    }

    public static IEnumerable<string> TanimliAdlar()
    {
        foreach (var v in _tablo.Values) yield return v.Ad;
    }

    private static string NormalizeAd(string s)
    {
        var u = s.Trim().ToUpper(new CultureInfo("tr-TR"));
        return u
            .Replace("İ", "I")
            .Replace("Ş", "S")
            .Replace("Ç", "C")
            .Replace("Ö", "O")
            .Replace("Ü", "U")
            .Replace("Ğ", "G");
    }
}
