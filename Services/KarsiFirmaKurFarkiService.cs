using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Karşı firmanın dövizli ekstresinden TIPLI kur farkı hesaplaması:
///
/// 1) KARŞI FİRMA İÇİ (FIFO): onların kendi faturaları ile kendi ödemeleri arasındaki
///    kur değişimi. Örtük kur = TL/Döviz; FIFO ile en eski açık fatura kapatılır;
///    kur farkı = (ödeme_kuru − fatura_kuru) × eşleşen_döviz.
///
/// 2) BELGE BAZLI (Logo ↔ Karşı firma): aynı belgenin (fatura/ödeme) bizdeki Logo TL'si ile
///    karşı firmanın aynı belgesindeki TL'si arasındaki fark. Belge no eşleşirse direkt;
///    yoksa aynı tarih+döviz tutarı ile fallback. Bu, "biz farklı kur kullanmışız" hatasını
///    yakalamak içindir.
/// </summary>
public class KarsiFirmaKurFarkiService
{
    private readonly LogoIslemleriService _logo;

    public KarsiFirmaKurFarkiService(LogoIslemleriService logo) => _logo = logo;

    public async Task<KurFarkiSonuc> HesaplaAsync(
        string vergiNo,
        DateTime baslangic, DateTime bitis,
        List<MutabakatKayit> ekstreSatirlari,
        string? secilenDoviz = null,
        CancellationToken ct = default)
    {
        var sonuc = new KurFarkiSonuc
        {
            AranılanVergiNo = vergiNo,
            Baslangic = baslangic.Date,
            Bitis = bitis.Date,
            AIHamKayitlar = ekstreSatirlari.ToList()
        };

        // 1) Karşı firma satırlarından dövizli olanlar
        var dovizliler = ekstreSatirlari
            .Where(k => (k.DovizBorc != 0 || k.DovizAlacak != 0) && !string.IsNullOrWhiteSpace(k.Doviz))
            .ToList();

        if (dovizliler.Count == 0)
        {
            sonuc.Doviz = "";
            await EkleBizdekiCariler(sonuc, vergiNo, ct);
            return sonuc;
        }

        var doviz = !string.IsNullOrWhiteSpace(secilenDoviz)
            ? secilenDoviz!.ToUpperInvariant()
            : dovizliler.GroupBy(k => k.Doviz)
                        .OrderByDescending(g => g.Count())
                        .First().Key;
        sonuc.Doviz = doviz;

        var hedef = dovizliler.Where(k => k.Doviz == doviz).ToList();

        // 2) FIFO hesaplaması (karşı firma içi)
        FifoHesapla(hedef, doviz, sonuc);

        // 3) Logo'dan bizdeki dövizli hareketleri çek + belge bazlı eşleştirme
        await BelgeBazliHesapla(vergiNo, baslangic, bitis, hedef, doviz, sonuc, ct);

        return sonuc;
    }

    private static void FifoHesapla(List<MutabakatKayit> hedef, string doviz, KurFarkiSonuc sonuc)
    {
        var faturalar = hedef
            .Where(k => k.DovizAlacak > 0)
            .OrderBy(k => k.Tarih)
            .Select(k => new FifoSatir(k, k.DovizAlacak, k.Alacak))
            .ToList();

        var odemeler = hedef
            .Where(k => k.DovizBorc > 0)
            .OrderBy(k => k.Tarih)
            .Select(k => new FifoSatir(k, k.DovizBorc, k.Borc))
            .ToList();

        sonuc.ToplamFaturaDoviz = faturalar.Sum(f => f.OrjDoviz);
        sonuc.ToplamOdemeDoviz  = odemeler.Sum(o => o.OrjDoviz);

        foreach (var odeme in odemeler)
        {
            var odemeKalan = odeme.KalanDoviz;
            if (odemeKalan == 0) continue;
            foreach (var fatura in faturalar)
            {
                if (odemeKalan == 0) break;
                if (fatura.KalanDoviz == 0) continue;

                var eslesen = Math.Min(odemeKalan, fatura.KalanDoviz);
                if (fatura.Kur == 0 || odeme.Kur == 0) continue;

                var kurFarki = (odeme.Kur - fatura.Kur) * eslesen;

                sonuc.Eslesmeler.Add(new KurFarkiEslesme
                {
                    Doviz = doviz,
                    FaturaTarihi   = fatura.Kayit.Tarih,
                    FaturaAciklama = fatura.Kayit.Aciklama,
                    FaturaFisNo    = fatura.Kayit.FisNo,
                    FaturaDoviz    = fatura.OrjDoviz,
                    FaturaTL       = fatura.OrjTL,
                    FaturaKur      = Math.Round(fatura.Kur, 4),
                    OdemeTarihi    = odeme.Kayit.Tarih,
                    OdemeAciklama  = odeme.Kayit.Aciklama,
                    OdemeFisNo     = odeme.Kayit.FisNo,
                    OdemeDoviz     = odeme.OrjDoviz,
                    OdemeTL        = odeme.OrjTL,
                    OdemeKur       = Math.Round(odeme.Kur, 4),
                    EslesenDoviz   = eslesen,
                    KurFarki       = Math.Round(kurFarki, 2)
                });

                fatura.KalanDoviz -= eslesen;
                odemeKalan        -= eslesen;
            }
            odeme.KalanDoviz = odemeKalan;
        }

        sonuc.ToplamKurFarki = sonuc.Eslesmeler.Sum(e => e.KurFarki);

        sonuc.AcikFaturalar = faturalar.Where(f => f.KalanDoviz > 0).Select(f => new MutabakatKayit
        {
            Tarih = f.Kayit.Tarih, Aciklama = f.Kayit.Aciklama, FisNo = f.Kayit.FisNo,
            Doviz = doviz, DovizAlacak = f.KalanDoviz,
            Alacak = f.Kur != 0 ? Math.Round(f.KalanDoviz * f.Kur, 2) : 0
        }).ToList();

        sonuc.AcikOdemeler = odemeler.Where(o => o.KalanDoviz > 0).Select(o => new MutabakatKayit
        {
            Tarih = o.Kayit.Tarih, Aciklama = o.Kayit.Aciklama, FisNo = o.Kayit.FisNo,
            Doviz = doviz, DovizBorc = o.KalanDoviz,
            Borc = o.Kur != 0 ? Math.Round(o.KalanDoviz * o.Kur, 2) : 0
        }).ToList();
    }

    private async Task EkleBizdekiCariler(KurFarkiSonuc sonuc, string vergiNo, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(vergiNo)) return;
        try
        {
            var cariler = await _logo.VergiNoyaGoreCarilerAsync(vergiNo, ct);
            sonuc.BizdekiCariler = cariler.Select(c => new EslesenCari
            {
                Kod = c.Kod, Unvan = c.Unvan, VergiNo = c.VergiNo
            }).ToList();
        }
        catch { /* Logo erişilemezse sessiz geç */ }
    }

    private async Task BelgeBazliHesapla(
        string vergiNo, DateTime baslangic, DateTime bitis,
        List<MutabakatKayit> hedefOnlar, string doviz,
        KurFarkiSonuc sonuc, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(vergiNo)) return;

        var cariler = await _logo.VergiNoyaGoreCarilerAsync(vergiNo, ct);
        if (cariler.Count == 0) return;

        sonuc.BizdekiCariler = cariler.Select(c => new EslesenCari
        {
            Kod = c.Kod, Unvan = c.Unvan, VergiNo = c.VergiNo
        }).ToList();

        var hareketler = await _logo.CarilerinHareketleriAsync(
            cariler.Select(c => c.Kod), baslangic, bitis, ct);

        // Bizdeki TÜM hareketleri al (TL havale/kredi kartı ödemeleri dahil — onların ekstresinde
        // bunlar dövize çevrilmiş olabilir, bu durumda TL tutar üzerinden eşleştireceğiz).
        var bizimHepsi = hareketler
            .Where(h => h.Borc != 0 || h.Alacak != 0)
            .Select(h =>
            {
                bool ayniDoviz = h.DovizKodu.Equals(doviz, StringComparison.OrdinalIgnoreCase);
                return new MutabakatKayit
                {
                    Tarih     = h.Tarih,
                    Borc      = h.Borc,
                    Alacak    = h.Alacak,
                    Aciklama  = h.Aciklama,
                    FisNo     = h.FisNo,
                    CariKodu  = h.CariKodu,
                    CariUnvan = h.CariUnvan,
                    Doviz     = ayniDoviz ? h.DovizKodu : "",
                    DovizBorc   = ayniDoviz ? h.DovizBorc   : 0,
                    DovizAlacak = ayniDoviz ? h.DovizAlacak : 0
                };
            })
            .ToList();
        sonuc.BizimDovizliKayitlar = bizimHepsi;

        if (bizimHepsi.Count == 0) return;

        var kullanilanOnlar = new HashSet<int>();

        foreach (var biz in bizimHepsi)
        {
            // Yön TL üzerinden belirlenir (bizde döviz olmayabilir):
            //   biz BORÇ (faturayı biz aldık)  → onların ALACAK
            //   biz ALACAK (biz tahsilat aldık) → onların BORÇ
            bool bizBorcMu = biz.Borc > 0 || biz.DovizBorc > 0;
            decimal bizDoviz = biz.DovizTutar;
            decimal bizTL = biz.Tutar;
            bool bizimDovizYok = bizDoviz == 0;

            int? eslesenIndex = null;
            string? yontem = null;

            // ① Fiş No eşleşmesi
            if (!string.IsNullOrWhiteSpace(biz.FisNo))
            {
                for (int i = 0; i < hedefOnlar.Count; i++)
                {
                    if (kullanilanOnlar.Contains(i)) continue;
                    var o = hedefOnlar[i];
                    if (string.IsNullOrWhiteSpace(o.FisNo)) continue;

                    bool yonUyumlu = bizBorcMu ? o.DovizAlacak > 0 : o.DovizBorc > 0;
                    if (!yonUyumlu) continue;

                    if (FisNoEslesir(biz.FisNo, o.FisNo))
                    {
                        eslesenIndex = i; yontem = "Fiş No"; break;
                    }
                }
            }

            // ② Aynı tarih + DÖVİZ tutarı (bizde döviz varsa)
            if (eslesenIndex == null && !bizimDovizYok)
            {
                for (int i = 0; i < hedefOnlar.Count; i++)
                {
                    if (kullanilanOnlar.Contains(i)) continue;
                    var o = hedefOnlar[i];
                    bool yonUyumlu = bizBorcMu ? o.DovizAlacak > 0 : o.DovizBorc > 0;
                    if (!yonUyumlu) continue;

                    decimal oDoviz = bizBorcMu ? o.DovizAlacak : o.DovizBorc;
                    if (Math.Abs(oDoviz - bizDoviz) > 0.01m) continue;
                    if (o.Tarih.Date != biz.Tarih.Date) continue;

                    eslesenIndex = i; yontem = "Tarih + Döviz"; break;
                }
            }

            // ③ Aynı tarih + TL tutarı (≤ 20.000 TL fark) — TL havale/kredi kartı senaryosu
            if (eslesenIndex == null)
            {
                const decimal TL_TOLERANS = 20000m;
                for (int i = 0; i < hedefOnlar.Count; i++)
                {
                    if (kullanilanOnlar.Contains(i)) continue;
                    var o = hedefOnlar[i];
                    bool yonUyumlu = bizBorcMu ? o.Alacak > 0 : o.Borc > 0;
                    if (!yonUyumlu) continue;

                    decimal oTL = bizBorcMu ? o.Alacak : o.Borc;
                    if (oTL == 0 || bizTL == 0) continue;

                    if (Math.Abs(oTL - bizTL) > TL_TOLERANS) continue;
                    if (o.Tarih.Date != biz.Tarih.Date) continue;

                    eslesenIndex = i; yontem = "Tarih + TL (±20.000)"; break;
                }
            }

            // ④ ±3 gün + DÖVİZ tutarı
            if (eslesenIndex == null && !bizimDovizYok)
            {
                int enYakinGun = int.MaxValue; int? enYakin = null;
                for (int i = 0; i < hedefOnlar.Count; i++)
                {
                    if (kullanilanOnlar.Contains(i)) continue;
                    var o = hedefOnlar[i];
                    bool yonUyumlu = bizBorcMu ? o.DovizAlacak > 0 : o.DovizBorc > 0;
                    if (!yonUyumlu) continue;

                    decimal oDoviz = bizBorcMu ? o.DovizAlacak : o.DovizBorc;
                    if (Math.Abs(oDoviz - bizDoviz) > 0.01m) continue;

                    int gunFark = Math.Abs((o.Tarih.Date - biz.Tarih.Date).Days);
                    if (gunFark <= 3 && gunFark < enYakinGun)
                    {
                        enYakinGun = gunFark; enYakin = i;
                    }
                }
                if (enYakin != null) { eslesenIndex = enYakin; yontem = $"±{enYakinGun} gün + Döviz"; }
            }

            // ⑤ ±3 gün + TL tutarı (≤ 20.000 TL fark)
            if (eslesenIndex == null)
            {
                const decimal TL_TOLERANS = 20000m;
                int enYakinGun = int.MaxValue; int? enYakin = null;
                for (int i = 0; i < hedefOnlar.Count; i++)
                {
                    if (kullanilanOnlar.Contains(i)) continue;
                    var o = hedefOnlar[i];
                    bool yonUyumlu = bizBorcMu ? o.Alacak > 0 : o.Borc > 0;
                    if (!yonUyumlu) continue;

                    decimal oTL = bizBorcMu ? o.Alacak : o.Borc;
                    if (oTL == 0 || bizTL == 0) continue;
                    if (Math.Abs(oTL - bizTL) > TL_TOLERANS) continue;

                    int gunFark = Math.Abs((o.Tarih.Date - biz.Tarih.Date).Days);
                    if (gunFark <= 3 && gunFark < enYakinGun)
                    {
                        enYakinGun = gunFark; enYakin = i;
                    }
                }
                if (enYakin != null) { eslesenIndex = enYakin; yontem = $"±{enYakinGun} gün + TL (±20.000)"; }
            }

            if (eslesenIndex == null) continue;

            kullanilanOnlar.Add(eslesenIndex.Value);
            var onlar = hedefOnlar[eslesenIndex.Value];
            decimal onlarTL    = bizBorcMu ? onlar.Alacak    : onlar.Borc;
            decimal onlarDoviz = bizBorcMu ? onlar.DovizAlacak : onlar.DovizBorc;

            // Bizde döviz yoksa, eşleşen onların satırının döviz bilgisini bizim kayda yaz
            // (kullanıcı isteği: "bizim hareketleri de karşı firma ekstresine göre o satırı dövizli say")
            if (bizimDovizYok && onlarDoviz > 0)
            {
                biz.Doviz = doviz;
                if (bizBorcMu) biz.DovizBorc = onlarDoviz;
                else           biz.DovizAlacak = onlarDoviz;
            }

            // Bizde döviz yoksa, referans olarak onların döviz tutarını kullan
            decimal kurHesabiDoviz = bizimDovizYok ? onlarDoviz : bizDoviz;
            decimal bizimKur = kurHesabiDoviz != 0 ? Math.Round(bizTL / kurHesabiDoviz, 4) : 0;
            decimal onlarKur = kurHesabiDoviz != 0 ? Math.Round(onlarTL / kurHesabiDoviz, 4) : 0;

            string yonEtiketi;
            if (bizimDovizYok)
                yonEtiketi = bizBorcMu ? "TL HAREKET → Dövizleştirildi (biz borçlu)" : "TL TAHSİLAT → Dövizleştirildi (biz aldık)";
            else
                yonEtiketi = bizBorcMu ? "FATURA (biz aldık)" : "ÖDEME (biz tahsil)";

            sonuc.BelgeBazliFarklar.Add(new BelgeBazliKurFarki
            {
                Tarih          = biz.Tarih,
                FisNo          = biz.FisNo,
                Aciklama       = biz.Aciklama,
                Yon            = yonEtiketi,
                Doviz          = doviz,
                DovizTutar     = kurHesabiDoviz,
                BizimTL        = bizTL,
                BizimKur       = bizimKur,
                BizimCariKodu  = biz.CariKodu,
                OnlarTL        = onlarTL,
                OnlarKur       = onlarKur,
                TLFarki        = Math.Round(bizTL - onlarTL, 2),
                KurFarki       = Math.Round(bizimKur - onlarKur, 4),
                EslesmeYontemi = yontem ?? ""
            });
        }

        sonuc.BelgeBazliToplamTLFarki = sonuc.BelgeBazliFarklar.Sum(b => b.TLFarki);
    }

    private static bool FisNoEslesir(string a, string b)
    {
        if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b)) return false;
        var na = a.Trim().TrimStart('0');
        var nb = b.Trim().TrimStart('0');
        if (na.Equals(nb, StringComparison.OrdinalIgnoreCase)) return true;
        // Bir taraf "TIM2025000000285" gibi prefix'liyse, diğeri sadece numerik olabilir
        return na.EndsWith(nb, StringComparison.OrdinalIgnoreCase)
            || nb.EndsWith(na, StringComparison.OrdinalIgnoreCase);
    }

    private sealed class FifoSatir
    {
        public MutabakatKayit Kayit { get; }
        public decimal OrjDoviz { get; }
        public decimal OrjTL { get; }
        public decimal Kur { get; }
        public decimal KalanDoviz { get; set; }

        public FifoSatir(MutabakatKayit k, decimal doviz, decimal tl)
        {
            Kayit      = k;
            OrjDoviz   = doviz;
            OrjTL      = tl;
            Kur        = doviz != 0 ? Math.Round(tl / doviz, 4) : 0;
            KalanDoviz = doviz;
        }
    }
}
