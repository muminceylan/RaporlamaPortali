using System.Text.RegularExpressions;

namespace RaporlamaPortali.Services.Logo;

// Fatura no üzerinde K-prefix manipülasyonu.
//
// Senaryo: aynı fatura no Logo'da başka VKN için zaten kayıtlı → duplicate key hatası.
// Çözüm: yıldan (20XX) sonraki İLK karakteri "K" ile değiştir.
//   AAA2026000000011 → AAA2026K00000011  (ilk "0" → "K")
//
// Bu fatura no'nun toplam uzunluğunu KORUR — Logo FICHENO kolon limiti sorun olmaz.
public static class FaturaNoYardimcisi
{
    private static readonly Regex YilSonrasiKarakter = new(@"(20\d{2})(.)", RegexOptions.Compiled);

    // K-prefix var mı?
    public static bool KPrefixliMi(string? faturaNo)
    {
        if (string.IsNullOrWhiteSpace(faturaNo)) return false;
        var m = YilSonrasiKarakter.Match(faturaNo);
        return m.Success && string.Equals(m.Groups[2].Value, "K", StringComparison.OrdinalIgnoreCase);
    }

    // K-prefix ekle (yoksa). Pattern: 20XX + ilk karakter → 20XX + "K"
    public static string KEkle(string? faturaNo)
    {
        if (string.IsNullOrWhiteSpace(faturaNo)) return faturaNo ?? "";
        var m = YilSonrasiKarakter.Match(faturaNo);
        if (!m.Success) return faturaNo;
        var ilkKarIdx = m.Groups[2].Index;
        // Zaten K ise aynısını döner
        if (faturaNo[ilkKarIdx] is 'K' or 'k') return faturaNo;
        return faturaNo[..ilkKarIdx] + "K" + faturaNo[(ilkKarIdx + 1)..];
    }

    // K-prefix çıkar (varsa). Eskiden kayıp olan rakamı geri getiremiyoruz — K yerine "0" koyar.
    // (Lookup amaçlı kullanılır; gerçek kaydı bulmak için.)
    public static string KKaldirSifirla(string? faturaNo)
    {
        if (string.IsNullOrWhiteSpace(faturaNo)) return faturaNo ?? "";
        var m = YilSonrasiKarakter.Match(faturaNo);
        if (!m.Success) return faturaNo;
        var ilkKarIdx = m.Groups[2].Index;
        if (faturaNo[ilkKarIdx] is not ('K' or 'k')) return faturaNo;
        return faturaNo[..ilkKarIdx] + "0" + faturaNo[(ilkKarIdx + 1)..];
    }

    // Lookup için arama varyantları — orijinal + olası K varyantları.
    public static IEnumerable<string> Varyantlar(string? faturaNo)
    {
        if (string.IsNullOrWhiteSpace(faturaNo)) yield break;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { faturaNo };
        yield return faturaNo;
        var kli = KEkle(faturaNo);
        if (seen.Add(kli)) yield return kli;
        var ksiz = KKaldirSifirla(faturaNo);
        if (seen.Add(ksiz)) yield return ksiz;
    }
}
