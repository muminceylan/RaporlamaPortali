using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Cari mutabakat servisi: VKN ile bizdeki tüm carilerin hareketlerini toplar,
/// karşı firmadan gelen ekstre satırlarıyla eşleştirir.
///
/// Mutabakat mantığı (TERS YÖN):
///   Bizim BORÇ → karşı firmanın gözünde ALACAK olur (örn. biz fatura kestik).
///   Bizim ALACAK → karşı firmanın gözünde BORÇ olur (örn. biz tahsilat aldık).
/// Eşleştirme bu mantıkla yapılır: |bizim.Borç - onlar.Alacak| ve
/// |bizim.Alacak - onlar.Borç| 0 olmalı.
/// </summary>
public class MutabakatService
{
    private readonly LogoIslemleriService _logo;

    public MutabakatService(LogoIslemleriService logo) => _logo = logo;

    public async Task<MutabakatSonuc> MutabakatYapAsync(
        string vergiNo,
        DateTime baslangic, DateTime bitis,
        List<MutabakatKayit> karsiTarafEkstresi,
        MutabakatCariTuru cariTuru = MutabakatCariTuru.Hepsi,
        CancellationToken ct = default)
    {
        var sonuc = new MutabakatSonuc
        {
            AranılanVergiNo = vergiNo,
            Baslangic = baslangic.Date,
            Bitis = bitis.Date
        };

        // 1) VKN ile bizdeki cariler — istenirse 12x (alıcı) veya 32x (satıcı) ile sınırla
        var tumCariler = await _logo.VergiNoyaGoreCarilerAsync(vergiNo, ct);
        var cariler = FiltreleCariler(tumCariler, cariTuru);
        if (cariler.Count == 0)
        {
            // Boş sonuç: karşı tarafın hareketleri tamamen "sadece onlarda" olur
            sonuc.SadeceOnlarda = karsiTarafEkstresi.OrderBy(x => x.Tarih).ToList();
            HesaplaToplamlar(sonuc);
            return sonuc;
        }

        // 2) Bizdeki birleşik hareketler
        var bizimHam = await _logo.CarilerinHareketleriAsync(
            cariler.Select(c => c.Kod), baslangic, bitis, ct);

        var bizim = bizimHam.Select(h => new MutabakatKayit
        {
            Tarih     = h.Tarih,
            Borc      = h.Borc,
            Alacak    = h.Alacak,
            Aciklama  = h.Aciklama,
            FisNo     = h.FisNo,
            CariKodu  = h.CariKodu,
            CariUnvan = h.CariUnvan,
            TrCode    = h.TrCode,
            FisTuru   = h.FisTuru
        }).ToList();

        // 3) Cari özeti
        sonuc.EslesenCariler = cariler.Select(c => new EslesenCari
        {
            Kod = c.Kod,
            Unvan = c.Unvan,
            VergiNo = c.VergiNo,
            HareketSayisi = bizim.Count(b => b.CariKodu == c.Kod)
        }).OrderBy(c => c.Kod).ToList();

        // 4) Eşleştirme (bizim ↔ onlar)
        Eslestir(bizim, karsiTarafEkstresi, sonuc);
        HesaplaToplamlar(sonuc);
        return sonuc;
    }

    /// <summary>
    /// 2 aşamalı eşleştirme (sıralama önemli — belge no en kesin bilgidir, önce o):
    ///   1) Belge no eşleşmesi (boş değilse) + ters yön + tutar tam
    ///   2) Aynı tarih + ters yön + tutar tam (belge no eşleşmeyenler için fallback)
    /// Bir kez eşleşen kayıt tekrar kullanılmaz.
    /// </summary>
    private static void Eslestir(
        List<MutabakatKayit> bizim,
        List<MutabakatKayit> onlar,
        MutabakatSonuc sonuc)
    {
        var bizimKalan = bizim.OrderBy(x => x.Tarih).ToList();
        var onlarKalan = onlar.OrderBy(x => x.Tarih).ToList();
        var bizimKullanildi = new HashSet<int>();
        var onlarKullanildi = new HashSet<int>();

        // 1) Belge no (boş olmayan) — en kesin bilgi, önce bunu dene.
        //    Aynı tarih + tutar başka bir satırla yanlış eşleşmesin diye sıralama önemli.
        for (int i = 0; i < bizimKalan.Count; i++)
        {
            if (bizimKullanildi.Contains(i)) continue;
            var b = bizimKalan[i];
            if (string.IsNullOrWhiteSpace(b.FisNo)) continue;
            for (int j = 0; j < onlarKalan.Count; j++)
            {
                if (onlarKullanildi.Contains(j)) continue;
                var o = onlarKalan[j];
                if (string.IsNullOrWhiteSpace(o.FisNo)) continue;
                if (!FisNoEslesir(b.FisNo, o.FisNo)) continue;
                if (!TersYonEslesir(b, o)) continue;

                bizimKullanildi.Add(i); onlarKullanildi.Add(j);
                sonuc.Eslesenler.Add(new MutabakatEslesme
                {
                    Bizim = b, Onlar = o,
                    GunFarki = Math.Abs((b.Tarih.Date - o.Tarih.Date).Days),
                    TutarFarki = (b.Borc - b.Alacak) - (o.Alacak - o.Borc),
                    EslesmeYontemi = "Belge no"
                });
                break;
            }
        }

        // 2) Aynı tarih + tutar — fallback (belge no boş veya farklı olan kayıtlar için)
        for (int i = 0; i < bizimKalan.Count; i++)
        {
            if (bizimKullanildi.Contains(i)) continue;
            var b = bizimKalan[i];
            for (int j = 0; j < onlarKalan.Count; j++)
            {
                if (onlarKullanildi.Contains(j)) continue;
                var o = onlarKalan[j];
                if (b.Tarih.Date != o.Tarih.Date) continue;
                if (!TersYonEslesir(b, o)) continue;

                bizimKullanildi.Add(i); onlarKullanildi.Add(j);
                sonuc.Eslesenler.Add(new MutabakatEslesme
                {
                    Bizim = b, Onlar = o,
                    GunFarki = 0,
                    TutarFarki = (b.Borc - b.Alacak) - (o.Alacak - o.Borc),
                    EslesmeYontemi = "Aynı tarih + tutar"
                });
                break;
            }
        }

        for (int i = 0; i < bizimKalan.Count; i++)
            if (!bizimKullanildi.Contains(i)) sonuc.SadeceBizde.Add(bizimKalan[i]);
        for (int j = 0; j < onlarKalan.Count; j++)
            if (!onlarKullanildi.Contains(j)) sonuc.SadeceOnlarda.Add(onlarKalan[j]);

        sonuc.SadeceBizde = sonuc.SadeceBizde.OrderBy(x => x.Tarih).ToList();
        sonuc.SadeceOnlarda = sonuc.SadeceOnlarda.OrderBy(x => x.Tarih).ToList();
        sonuc.Eslesenler = sonuc.Eslesenler.OrderBy(x => x.Bizim.Tarih).ToList();
    }

    /// <summary>
    /// Bizim BORÇ ↔ onların ALACAK ya da bizim ALACAK ↔ onların BORÇ?
    /// Tutarlar 0.10 TL toleransıyla eşit mi? — KDV yuvarlama farkı, kuruş kayması vb.
    /// </summary>
    private static bool TersYonEslesir(MutabakatKayit b, MutabakatKayit o)
    {
        const decimal eps = 0.10m;
        if (b.Borc > 0 && o.Alacak > 0)
            return Math.Abs(b.Borc - o.Alacak) <= eps;
        if (b.Alacak > 0 && o.Borc > 0)
            return Math.Abs(b.Alacak - o.Borc) <= eps;
        return false;
    }

    /// <summary>
    /// Türk e-Fatura / e-Arşiv ETTN formatı: 3 büyük harf + 4 hane yıl + 9 hane sıra = 16 karakter.
    /// (örn GDL2024000005354, GIB2024000000287, 1OF2023000001936 — son örnek 1 rakam + 2 harf
    /// öneki olduğu için klasik patern dışında, onu da yakalamak için her iki varyantı arıyoruz.)
    /// Verilen metnin içinden bu formatı (varsa) çıkarır.
    /// </summary>
    private static string EFaturaNoCikar(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        var temiz = new string(s.Where(c => char.IsLetterOrDigit(c)).ToArray()).ToUpperInvariant();
        // En yaygın: 3 harf + 13 rakam (GDL2024..., GIB2024...)
        var m = System.Text.RegularExpressions.Regex.Match(temiz, @"[A-Z]{3}\d{13}");
        if (m.Success) return m.Value;
        // Alternatif: 1 rakam + 2 harf + 13 rakam (1OF2023..., 1OFE...)
        m = System.Text.RegularExpressions.Regex.Match(temiz, @"\d[A-Z]{2,3}\d{13}");
        if (m.Success) return m.Value;
        return "";
    }

    /// <summary>
    /// Belge no eşleştirme:
    ///   1) Önce e-Fatura/e-Arşiv ETTN deseni (16 karakter standart) çıkar — varsa tam karşılaştır.
    ///      Bu en güvenli yöntem: PDF'te açıklamayla yapışık kalan fiş no'ları da yakalar
    ///      (örn "GDL2024000005354Toptan Satış" → "GDL2024000005354").
    ///   2) ETTN yoksa normalize edip karşılaştır (boşluk/tire/baştaki sıfır temizle).
    ///   3) Norm eşit değilse, biri diğerinin başlangıcı mı bak (min 8 karakter ortak).
    /// </summary>
    private static bool FisNoEslesir(string a, string b)
    {
        // 1) E-Fatura ETTN — en güçlü kural
        var ea = EFaturaNoCikar(a);
        var eb = EFaturaNoCikar(b);
        if (!string.IsNullOrEmpty(ea) && !string.IsNullOrEmpty(eb))
            return ea == eb;

        string Norm(string s) => new string((s ?? "")
            .Where(c => char.IsLetterOrDigit(c)).ToArray()).TrimStart('0').ToUpperInvariant();
        var na = Norm(a); var nb = Norm(b);
        if (string.IsNullOrEmpty(na) || string.IsNullOrEmpty(nb)) return false;
        if (na == nb) return true;

        // 2) Prefix eşleşme (fiş no + bitişik metin durumu — ETTN dışı formatlar için)
        const int minOrtak = 8;
        if (na.Length >= minOrtak && nb.StartsWith(na)) return true;
        if (nb.Length >= minOrtak && na.StartsWith(nb)) return true;
        return false;
    }

    /// <summary>
    /// VKN ile bulunan tüm carileri seçilen tipe göre filtreler.
    /// 12x = Alıcılar (satış yaptıklarımız), 32x = Satıcılar (alış yaptıklarımız).
    /// </summary>
    private static List<CariSecenek> FiltreleCariler(
        List<CariSecenek> cariler, MutabakatCariTuru tur)
    {
        if (tur == MutabakatCariTuru.Hepsi) return cariler;
        string prefix = tur == MutabakatCariTuru.Alici120 ? "12" : "32";
        return cariler.Where(c => (c.Kod ?? "").TrimStart().StartsWith(prefix)).ToList();
    }

    private static void HesaplaToplamlar(MutabakatSonuc s)
    {
        var bizim = s.Eslesenler.Select(e => e.Bizim).Concat(s.SadeceBizde);
        var onlar = s.Eslesenler.Select(e => e.Onlar).Concat(s.SadeceOnlarda);
        s.BizimToplamBorc = bizim.Sum(x => x.Borc);
        s.BizimToplamAlacak = bizim.Sum(x => x.Alacak);
        s.OnlarToplamBorc = onlar.Sum(x => x.Borc);
        s.OnlarToplamAlacak = onlar.Sum(x => x.Alacak);
    }
}
