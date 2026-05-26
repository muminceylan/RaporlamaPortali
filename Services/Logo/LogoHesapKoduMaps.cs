namespace RaporlamaPortali.Services.Logo;

// KDV oranına / tevkifat oranına göre kullanılacak Logo muhasebe hesap kodları.
// Kullanıcı (Mümin CEYLAN) tarafından 06.03.2026 tarihinde teyit edildi.
// Yeni bir oran eklenirse buraya ekle; başka yerde sabit kod tutma.
public static class LogoHesapKoduMaps
{
    // Normal KDV → İndirilecek KDV (sınıf 191.01.xx)
    public static readonly IReadOnlyDictionary<decimal, string> KdvIndirilecek = new Dictionary<decimal, string>
    {
        [1m]  = "191.01.01",
        [10m] = "191.01.06",
        [20m] = "191.01.07",
    };

    // Tevkifatlı KDV → Tevkifatlı İndirilecek KDV (sınıf 192.02.xx)
    // Logo, satırda GL_CODE3 olarak işliyor.
    public static readonly IReadOnlyDictionary<decimal, string> KdvTevkifatliIndirilecek = new Dictionary<decimal, string>
    {
        [1m]  = "192.02.01",
        [10m] = "192.02.03",
        [20m] = "192.02.04",
    };

    // Tevkifat oranı (Pay/Payda) → Sorumlu Sıfatı ile Ödenecek KDV (sınıf 360.10.xx)
    // GL_CODE4 olarak işlenir.
    // Anahtar: "Pay/Payda" string formatında — XML'deki cbc:CalculationRate ile gelen değerden hesaplanır.
    public static readonly IReadOnlyDictionary<string, string> TevkifatSorumlulukKdv = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["1/2"]   = "360.10.02",   // %50
        ["9/10"]  = "360.10.03",   // %90
        ["4/5"]   = "360.10.05",   // %80
        ["8/10"]  = "360.10.05",   // %80
        ["10/10"] = "360.10.04",   // %100
        ["7/10"]  = "360.10.07",   // %70
        ["2/10"]  = "360.10.10",   // %20
        ["3/10"]  = "360.10.13",   // %30
        ["4/10"]  = "360.10.15",   // %40
    };

    // KDV oranına göre indirilecek KDV hesabını döner. Bulunamazsa boş string döner —
    // çağıran taraf bunu hata sayar (kullanıcı görür, taban kodu üretmiyoruz).
    public static string KdvHesabiBul(decimal kdvOran, bool tevkifatli)
    {
        var map = tevkifatli ? KdvTevkifatliIndirilecek : KdvIndirilecek;
        return map.TryGetValue(kdvOran, out var k) ? k : "";
    }

    // "Pay/Payda" string'inden sorumluluk hesap kodunu döner. Pay/payda 0 ise boş.
    public static string SorumlulukHesabiBul(int pay, int payda)
    {
        if (pay <= 0 || payda <= 0) return "";
        var key = $"{pay}/{payda}";
        return TevkifatSorumlulukKdv.TryGetValue(key, out var k) ? k : "";
    }

    // ─────────────────────────────────────────────────────────────────────
    // Malzeme MASTER_CODE prefix → Muhasebe hesabı (GL_CODE1)
    // Logo malzeme kartında muhasebe hesabı tanımlı değilse, fatura satırında
    // GL_CODE1 set edilmeli — yoksa Logo "muhasebe kodu boş" bırakıyor ve
    // fişin muhasebeleşmesi eksik kalıyor.
    // Eşleşme: en uzun prefix kazanır (daha spesifik ilk).
    // ─────────────────────────────────────────────────────────────────────

    // Satın Alma Faturası (TYPE=1) için — Logo'da "Mal Alımları" hesapları (150/153/157.xx)
    public static readonly IReadOnlyList<(string Prefix, string Hesap)> MalzemeMuhasebeSatinAlma = new List<(string, string)>
    {
        ("A.G.01.01.", "150.18.001"),
        ("A.G.02.",    "150.18.002"),
        // Yeni kurallar buraya eklenir. Daha spesifik (uzun) prefix ÜSTTE olmalı —
        // örn. "A.G.01.01.02." kuralı "A.G.01.01." kuralından önce gelmeli.
    };

    // Alınan Hizmet Faturası (TYPE=4) için — Logo'da "Hizmet/Gider" hesapları (740/760.xx)
    // Şu an boş, gerektiğinde doldurulacak.
    public static readonly IReadOnlyList<(string Prefix, string Hesap)> MalzemeMuhasebeHizmet = new List<(string, string)>
    {
    };

    // MASTER_CODE'a ve fatura tipine göre muhasebe hesabını döner. Eşleşme yoksa boş.
    public static string MalzemeMuhasebeHesabiBul(string? masterCode, int faturaTipi)
    {
        if (string.IsNullOrWhiteSpace(masterCode)) return "";
        var liste = faturaTipi == 4 ? MalzemeMuhasebeHizmet : MalzemeMuhasebeSatinAlma;
        return liste
            .Where(p => masterCode!.StartsWith(p.Prefix, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(p => p.Prefix.Length)
            .Select(p => p.Hesap)
            .FirstOrDefault() ?? "";
    }
}

// UBL XML'inden gelen tevkifat oranı (örnek: 0.5, 50.0, 0.9, 90.0) → (pay, payda).
// Logo DEDUCTION_PART1 (pay) ve DEDUCTION_PART2 (payda) tam sayı bekliyor.
public static class TevkifatPayPaydaCevirici
{
    // Standart Türk e-fatura tevkifat oranları için (pay, payda).
    // Bulunamazsa (0, 0) — çağıran taraf bunu hata sayar.
    public static (int pay, int payda) Bol(decimal oran)
    {
        // 0..1 arası geldiyse %'ye çevir (50.0 değil 0.5 olarak gelmiş olabilir)
        var pct = oran <= 1m ? oran * 100m : oran;
        pct = Math.Round(pct, 0);
        return (int)pct switch
        {
            10  => (1, 10),
            20  => (2, 10),
            30  => (3, 10),
            40  => (4, 10),
            50  => (1, 2),
            60  => (6, 10),
            70  => (7, 10),
            80  => (4, 5),
            90  => (9, 10),
            100 => (10, 10),
            _   => (0, 0),
        };
    }
}

// UBL XML'inde gelen tevkifat kodu (satıcı tarafı — 6xx) → alıcı tarafı (3xx) dönüşüm.
//   601 → 301, 602 → 302, 627 → 327, ...
// Mümin CEYLAN'ın açıklaması: "fatura kesen taraf isem 6 ile başlıyor, biz alıcı tarafız o yüzden 3 ile başlamalı".
public static class TevkifatKoduCevirici
{
    // Yalnızca ilk karakter 6 olan, 3 haneli sayısal kodları çevirir. Beklenmeyen format
    // gelirse aynısını döner (Logo da geri bildiriverir o zaman).
    public static string Cevir(string? satiKodu)
    {
        var s = (satiKodu ?? "").Trim();
        if (s.Length == 0) return "";
        if (s.Length >= 1 && s[0] == '6')
            return "3" + s.Substring(1);
        return s;
    }
}
