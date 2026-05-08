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
        CancellationToken ct = default)
    {
        var sonuc = new MutabakatSonuc
        {
            AranılanVergiNo = vergiNo,
            Baslangic = baslangic.Date,
            Bitis = bitis.Date
        };

        // 1) VKN ile bizdeki cariler
        var cariler = await _logo.VergiNoyaGoreCarilerAsync(vergiNo, ct);
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
    /// 2 aşamalı eşleştirme:
    ///   1) Aynı tarih + ters yön + tutar 0 farkla eşit
    ///   2) Belge no eşleşmesi (boş değilse) + ters yön
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

        // 1) Aynı tarih, ters yön
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

        // 2) Belge no (boş olmayan)
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
    /// Tutarlar 0.01 toleransıyla eşit mi?
    /// </summary>
    private static bool TersYonEslesir(MutabakatKayit b, MutabakatKayit o)
    {
        const decimal eps = 0.01m;
        if (b.Borc > 0 && o.Alacak > 0)
            return Math.Abs(b.Borc - o.Alacak) < eps;
        if (b.Alacak > 0 && o.Borc > 0)
            return Math.Abs(b.Alacak - o.Borc) < eps;
        return false;
    }

    /// <summary>
    /// Belge no'ları normalize ederek karşılaştır (boşluk, tire, başındaki sıfırları temizle).
    /// "F00125" ve "F-125" eşleşir.
    /// </summary>
    private static bool FisNoEslesir(string a, string b)
    {
        string Norm(string s) => new string((s ?? "")
            .Where(c => char.IsLetterOrDigit(c)).ToArray()).TrimStart('0').ToUpperInvariant();
        var na = Norm(a); var nb = Norm(b);
        if (string.IsNullOrEmpty(na) || string.IsNullOrEmpty(nb)) return false;
        return na == nb;
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
