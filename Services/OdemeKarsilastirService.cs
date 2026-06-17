using ClosedXML.Excel;
using RaporlamaPortali.Models;
using System.Globalization;
using System.Text;

namespace RaporlamaPortali.Services;

/// <summary>
/// Bankaya gönderilen ödeme dosyaları ile Logo Finans Raporu'ndaki gerçek ödemeleri karşılaştırır.
/// - Eşleşen: dosyada var + Logo'da aynı tutarda aynı kişiye ödeme yapılmış
/// - Ödenmedi: dosyada var ama Logo'da yok (banka IBAN/ad hatasıyla geri çevirmiş olabilir)
/// - Mükerrer: Logo'da aynı kişiye birden fazla ödeme yapılmış
/// </summary>
public class OdemeKarsilastirService
{
    private readonly FinansRaporService _finans;
    private readonly OdemeService       _odeme;
    public OdemeKarsilastirService(FinansRaporService finans, OdemeService odeme)
    { _finans = finans; _odeme = odeme; }

    public async Task<OdemeKarsilastirSonuc> KarsilastirAsync(
        byte[]? ziraatBytes, byte[]? garantiBytes, byte[]? isBytes,
        DateTime logoBaslangic, DateTime logoBitis,
        decimal tutarToleransi = 0.01m,
        CancellationToken ct = default)
    {
        var sonuc = new OdemeKarsilastirSonuc
        {
            LogoBaslangic = logoBaslangic,
            LogoBitis     = logoBitis
        };

        // 1) Banka dosyalarını oku
        var dosyaSatirlari = new List<DosyaKayit>();
        if (ziraatBytes != null)
            dosyaSatirlari.AddRange(ZiraatOku(ziraatBytes, sonuc.Hatalar));
        if (garantiBytes != null)
            dosyaSatirlari.AddRange(GarantiOku(garantiBytes, sonuc.Hatalar));
        if (isBytes != null)
            dosyaSatirlari.AddRange(IsOku(isBytes, sonuc.Hatalar));

        sonuc.DosyaSatirSayisi = dosyaSatirlari.Count;
        sonuc.DosyaToplam      = dosyaSatirlari.Sum(x => x.Tutar);

        if (dosyaSatirlari.Count == 0)
        {
            sonuc.Hatalar.Add("Hiçbir banka dosyası yüklenmedi ya da içlerinde geçerli satır bulunamadı.");
            return sonuc;
        }

        // 2) Logo'dan tarih aralığındaki ödemeleri çek (sadece BİZİM gönderdiğimiz çıkışlar)
        var logoHareketler = await _finans.GetHareketlerAsync(logoBaslangic, logoBitis, ct);
        // NOT: Bu view'da SIGN tek başına yön ayrımı yapmıyor — hem "Gelen Havale" hem
        // "Gönderilen Havale" SIGN=-1 olarak geliyor. Ayrım FIS_TURU metninde:
        //   "Gönderilen Havale" → bizden çıkan (ödeme)
        //   "Gelen Havale"      → bize gelen (iade / tahsilat). DAHİL DEĞİL.
        var odemeler = logoHareketler
            .Where(h => IsCikisOdeme(h.FisTuru))
            .Select(h => new LogoOdeme
            {
                CariKod  = h.ChKod ?? "",
                Unvan    = h.ChUnvani ?? "",
                Tutar    = Math.Abs(h.Havale != 0m ? h.Havale : h.Diger),
                Tarih    = h.Tarih,
            })
            .Where(x => x.Tutar > 0m)
            .ToList();

        sonuc.LogoSatirSayisi = odemeler.Count;
        sonuc.LogoToplam      = odemeler.Sum(x => x.Tutar);

        // 3) Logo ödemelerini hem CH_KOD (TC bazlı) hem AdSoyadNormal key'iyle indexle
        var logoByKod = odemeler
            .GroupBy(o => (o.CariKod ?? "").Trim().ToUpperInvariant())
            .ToDictionary(g => g.Key, g => g.ToList());
        var logoByAd = odemeler
            .GroupBy(o => NormalizeAd(o.Unvan))
            .ToDictionary(g => g.Key, g => g.ToList());

        // 4) Dosyadaki her satırı eşleştir
        foreach (var d in dosyaSatirlari)
        {
            var satir = new OdemeKarsilastirSatir
            {
                Banka      = d.Banka,
                AdSoyad    = d.AdSoyad,
                IBAN       = d.IBAN,
                TcKimlikNo = d.TcKimlikNo,
                DosyaTutar = d.Tutar
            };

            // EŞLEŞTİRME ÖNCELİĞİ:
            //   1) TC bazlı (dosyada TC varsa) → Logo CH_KOD = "S" + TC (müstahsil cariler)
            //   2) Ad Soyad normalize
            List<LogoOdeme>? adayLar = null;
            string? eslesenKaynak = null;

            if (!string.IsNullOrWhiteSpace(d.TcKimlikNo))
            {
                var tcKey = ("S" + d.TcKimlikNo.Trim()).ToUpperInvariant();
                if (logoByKod.TryGetValue(tcKey, out var tcAdayLar) && tcAdayLar.Count > 0)
                {
                    adayLar = tcAdayLar;
                    eslesenKaynak = "TC";
                }
            }

            if (adayLar == null)
            {
                var key = NormalizeAd(d.AdSoyad);
                if (!string.IsNullOrEmpty(key) && logoByAd.TryGetValue(key, out var adAdayLar) && adAdayLar.Count > 0)
                {
                    adayLar = adAdayLar;
                    eslesenKaynak = "Ad";
                }
            }

            if (adayLar == null || adayLar.Count == 0)
            {
                satir.Aciklama = "Logo'da bu kişiye bu dönemde ödeme yapılmamış (IBAN/ad/TC hatası olabilir).";
                sonuc.Odenmedi.Add(satir);
                continue;
            }

            // Tutarı en yakın olan adayı seç
            var tutarEslesen = adayLar
                .Where(a => Math.Abs(a.Tutar - d.Tutar) <= tutarToleransi)
                .ToList();

            if (tutarEslesen.Count == 0)
            {
                // Ad eşleşti ama tutar eşleşmedi → muhtemelen ödenmedi/farklı bir ödeme
                satir.LogoTutar = adayLar.Sum(a => a.Tutar);
                satir.LogoAdet  = adayLar.Count;
                satir.LogoCariKod = adayLar[0].CariKod;
                satir.LogoTarih = adayLar[0].Tarih;
                satir.Aciklama = $"Logo'da {adayLar.Count} farklı tutarlı ödeme bulundu (toplam {satir.LogoTutar:N2}), beklenen {d.Tutar:N2}.";
                sonuc.Odenmedi.Add(satir);
                continue;
            }

            // Tutar eşleşen aday var
            satir.LogoTutar   = tutarEslesen[0].Tutar;
            satir.LogoAdet    = tutarEslesen.Count;
            satir.LogoCariKod = tutarEslesen[0].CariKod;
            satir.LogoTarih   = tutarEslesen[0].Tarih;

            if (tutarEslesen.Count >= 2)
            {
                satir.Aciklama = $"Aynı kişiye Logo'da {tutarEslesen.Count} kez aynı tutarda ödeme yapılmış (mükerrer şüphesi). [{eslesenKaynak}]";
                sonuc.Mukerrer.Add(satir);
            }
            else
            {
                satir.Aciklama = $"Eşleşti ({eslesenKaynak}) — Logo: {tutarEslesen[0].Tarih:dd.MM.yyyy} — Cari: {tutarEslesen[0].CariKod}";
                sonuc.Eslesen.Add(satir);
            }
        }

        return sonuc;
    }

    // ----------------------------------------------------------------------
    // SABNET ↔ LOGO KARŞILAŞTIRMA
    // ----------------------------------------------------------------------
    /// <summary>
    /// Sabnet'teki seçilen avans listesini Logo'da gerçekleşen ödemelerle karşılaştırır.
    /// "Ödenmedi" = Sabnet'te var ama Logo'da yok (Sabnette fazla görünenler — IBAN/ad
    /// hatasıyla geri çevrilmiş ya da hiç gönderilmemiş).
    /// </summary>
    public async Task<OdemeKarsilastirSonuc> SabnetIleKarsilastirAsync(
        int avansNo, int sozlesmeYili,
        DateTime logoBaslangic, DateTime logoBitis,
        decimal tutarToleransi = 0.01m,
        CancellationToken ct = default)
    {
        var sonuc = new OdemeKarsilastirSonuc
        {
            LogoBaslangic = logoBaslangic,
            LogoBitis     = logoBitis
        };

        // 1) Sabnet avans listesini çek
        if (!OdemeService.AvansAdlari.TryGetValue(avansNo, out var avansAdi))
        {
            sonuc.Hatalar.Add($"Tanımsız AvansNo: {avansNo}");
            return sonuc;
        }
        var avansSatirlari = await _odeme.AvansVerileriniGetirAsync(avansNo, sozlesmeYili);

        // TC bazlı tekilleştir (aynı kişiye birden fazla form girilmiş olabilir → topla)
        var sabnetByTc = avansSatirlari
            .Where(x => !string.IsNullOrWhiteSpace(x.TcKimlikNo))
            .GroupBy(x => x.TcKimlikNo.Trim())
            .Select(g => new DosyaKayit
            {
                Banka      = $"Sabnet — {avansAdi}",
                AdSoyad    = g.First().AdSoyad,
                IBAN       = g.First().IBAN,
                TcKimlikNo = g.Key,
                Tutar      = g.Sum(x => x.Tutar)
            })
            .ToList();

        sonuc.DosyaSatirSayisi = sabnetByTc.Count;
        sonuc.DosyaToplam      = sabnetByTc.Sum(x => x.Tutar);

        if (sabnetByTc.Count == 0)
        {
            sonuc.Hatalar.Add($"Sabnet'te {avansAdi} ({sozlesmeYili}) için kayıt bulunamadı.");
            return sonuc;
        }

        // 2) Logo ödemeleri (sadece "Gönderilen Havale"/"Tediye" — Gelen Havale dışlanır)
        var logoHareketler = await _finans.GetHareketlerAsync(logoBaslangic, logoBitis, ct);
        var odemeler = logoHareketler
            .Where(h => IsCikisOdeme(h.FisTuru))
            .Select(h => new LogoOdeme
            {
                CariKod = h.ChKod ?? "",
                Unvan   = h.ChUnvani ?? "",
                Tutar   = Math.Abs(h.Havale != 0m ? h.Havale : h.Diger),
                Tarih   = h.Tarih
            })
            .Where(x => x.Tutar > 0m)
            .ToList();

        sonuc.LogoSatirSayisi = odemeler.Count;
        sonuc.LogoToplam      = odemeler.Sum(x => x.Tutar);

        var logoByKod = odemeler
            .GroupBy(o => (o.CariKod ?? "").Trim().ToUpperInvariant())
            .ToDictionary(g => g.Key, g => g.ToList());
        var logoByAd = odemeler
            .GroupBy(o => NormalizeAd(o.Unvan))
            .ToDictionary(g => g.Key, g => g.ToList());

        // Logo ödemelerinin Sabnet ile eşleşip eşleşmediği takibi — sonunda eşleşmemişler
        // "Logo'da fazla (Sabnet'te yok)" listesine gider.
        var kullanilanLogo = new HashSet<LogoOdeme>(ReferenceEqualityComparer.Instance);

        // 3) Karşılaştır
        foreach (var d in sabnetByTc)
        {
            var satir = new OdemeKarsilastirSatir
            {
                Banka      = d.Banka,
                AdSoyad    = d.AdSoyad,
                IBAN       = d.IBAN,
                TcKimlikNo = d.TcKimlikNo,
                DosyaTutar = d.Tutar
            };

            List<LogoOdeme>? adayLar = null;
            string? eslesenKaynak = null;

            if (!string.IsNullOrWhiteSpace(d.TcKimlikNo))
            {
                var tcKey = ("S" + d.TcKimlikNo.Trim()).ToUpperInvariant();
                if (logoByKod.TryGetValue(tcKey, out var tcAdayLar) && tcAdayLar.Count > 0)
                {
                    adayLar = tcAdayLar;
                    eslesenKaynak = "TC";
                }
            }

            if (adayLar == null)
            {
                var key = NormalizeAd(d.AdSoyad);
                if (!string.IsNullOrEmpty(key) && logoByAd.TryGetValue(key, out var adAdayLar) && adAdayLar.Count > 0)
                {
                    adayLar = adAdayLar;
                    eslesenKaynak = "Ad";
                }
            }

            if (adayLar == null || adayLar.Count == 0)
            {
                satir.Aciklama = "Sabnet'te kayıtlı ama Logo'da bu kişiye bu dönemde ödeme yok.";
                sonuc.Odenmedi.Add(satir);
                continue;
            }

            var tutarEslesen = adayLar
                .Where(a => Math.Abs(a.Tutar - d.Tutar) <= tutarToleransi)
                .ToList();

            if (tutarEslesen.Count == 0)
            {
                satir.LogoTutar   = adayLar.Sum(a => a.Tutar);
                satir.LogoAdet    = adayLar.Count;
                satir.LogoCariKod = adayLar[0].CariKod;
                satir.LogoTarih   = adayLar[0].Tarih;
                satir.Aciklama    = $"Logo'da {adayLar.Count} farklı tutarlı ödeme bulundu (toplam {satir.LogoTutar:N2}), Sabnet bekleneni {d.Tutar:N2}.";
                sonuc.Odenmedi.Add(satir);
                // Tutar tutmasa da TC/ad eşleşti — bu Logo kayıtları "sahipsiz fazla" değil; işaretle.
                foreach (var a in adayLar) kullanilanLogo.Add(a);
                continue;
            }

            satir.LogoTutar   = tutarEslesen[0].Tutar;
            satir.LogoAdet    = tutarEslesen.Count;
            satir.LogoCariKod = tutarEslesen[0].CariKod;
            satir.LogoTarih   = tutarEslesen[0].Tarih;

            if (tutarEslesen.Count >= 2)
            {
                satir.Aciklama = $"Logo'da {tutarEslesen.Count} kez aynı tutarda ödeme — mükerrer şüphesi. [{eslesenKaynak}]";
                sonuc.Mukerrer.Add(satir);
            }
            else
            {
                satir.Aciklama = $"Eşleşti ({eslesenKaynak}) — Logo: {tutarEslesen[0].Tarih:dd.MM.yyyy} — Cari: {tutarEslesen[0].CariKod}";
                sonuc.Eslesen.Add(satir);
            }
            // Eşleşmiş Logo kayıtlarını işaretle (tutar uyan tüm adaylar — mükerrer dahil)
            foreach (var a in adayLar) kullanilanLogo.Add(a);
        }

        // 4) Logo'da var ama Sabnet ile hiç eşleşmemiş MÜSTAHSİL ödemeleri.
        //    Sadece müstahsil cari kodları (Sxxxxxxxxxx[x] formatı: S + 10 hane VKN veya 11 hane TC)
        //    listeye alınır — diğer ödemeler (banka, market, kira vb.) elenir.
        //    Sahipsiz Logo ödemeleri: yetkili numara değişikliği, iptal/manuel ödeme,
        //    başka bir avans formundan ödenmiş vb.
        foreach (var o in odemeler)
        {
            if (kullanilanLogo.Contains(o)) continue;
            if (!IsMustahsilCariKodu(o.CariKod)) continue;

            var tcVkn = TcVknKodundanCikar(o.CariKod);
            sonuc.LogoFazlasi.Add(new OdemeKarsilastirSatir
            {
                Banka       = "Logo",
                AdSoyad     = o.Unvan,
                TcKimlikNo  = tcVkn,
                LogoTutar   = o.Tutar,
                LogoAdet    = 1,
                LogoCariKod = o.CariKod,
                LogoTarih   = o.Tarih,
                Aciklama    = $"Logo'da ödeme var ({o.Tarih:dd.MM.yyyy}) ama bu kişi Sabnet'teki '{avansAdi} {sozlesmeYili}' listesinde yok.",
            });
        }

        return sonuc;
    }

    /// <summary>
    /// Müstahsil cari kodu mu? S + 10 hane (VKN) veya S + 11 hane (TC) deseni.
    /// </summary>
    private static bool IsMustahsilCariKodu(string? cariKod)
    {
        if (string.IsNullOrWhiteSpace(cariKod)) return false;
        var k = cariKod.Trim().ToUpperInvariant();
        if (!k.StartsWith("S")) return false;
        var rest = k.Substring(1);
        return (rest.Length == 10 || rest.Length == 11) && rest.All(char.IsDigit);
    }

    /// <summary>
    /// Logo cari kodundan TC (11) veya VKN (10) çıkarır: "S12345678901" → "12345678901".
    /// </summary>
    private static string TcVknKodundanCikar(string? cariKod)
    {
        if (!IsMustahsilCariKodu(cariKod)) return "";
        return cariKod!.Trim().Substring(1);
    }

    // ----------------------------------------------------------------------
    // EXCEL EXPORT
    // ----------------------------------------------------------------------
    public byte[] ExportToExcel(OdemeKarsilastirSonuc s)
    {
        using var wb = new XLWorkbook();

        // ---- Özet
        var ws = wb.Worksheets.Add("Özet");
        ws.Cell(1, 1).Value = "Banka Ödeme Dosyaları ↔ Logo Finans Karşılaştırma";
        ws.Range(1, 1, 1, 4).Merge().Style.Font.SetBold().Font.SetFontSize(14);
        ws.Cell(3, 1).Value = "Logo Tarama Aralığı";
        ws.Cell(3, 2).Value = $"{s.LogoBaslangic:dd.MM.yyyy} – {s.LogoBitis:dd.MM.yyyy}";
        ws.Cell(4, 1).Value = "Rapor Tarihi";
        ws.Cell(4, 2).Value = DateTime.Now.ToString("dd.MM.yyyy HH:mm");

        ws.Cell(6, 1).Value = "Kategori";
        ws.Cell(6, 2).Value = "Kayıt";
        ws.Cell(6, 3).Value = "Tutar (₺)";
        ws.Range(6, 1, 6, 3).Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.LightGray);

        ws.Cell(7, 1).Value = "Banka Dosyaları (Gönderilen)";
        ws.Cell(7, 2).Value = s.DosyaSatirSayisi;
        ws.Cell(7, 3).Value = s.DosyaToplam;
        ws.Cell(8, 1).Value = "Logo Finans (Gerçekleşen)";
        ws.Cell(8, 2).Value = s.LogoSatirSayisi;
        ws.Cell(8, 3).Value = s.LogoToplam;
        ws.Cell(9, 1).Value = "ÖDENMEDİ";
        ws.Cell(9, 2).Value = s.Odenmedi.Count;
        ws.Cell(9, 3).Value = s.OdenmediToplam;
        ws.Range(9, 1, 9, 3).Style.Fill.SetBackgroundColor(XLColor.LightPink);
        ws.Cell(10, 1).Value = "MÜKERRER";
        ws.Cell(10, 2).Value = s.Mukerrer.Count;
        ws.Cell(10, 3).Value = s.MukerrerToplam;
        ws.Range(10, 1, 10, 3).Style.Fill.SetBackgroundColor(XLColor.LightYellow);
        ws.Cell(11, 1).Value = "LOGO'DA FAZLA";
        ws.Cell(11, 2).Value = s.LogoFazlasi.Count;
        ws.Cell(11, 3).Value = s.LogoFazlasiToplam;
        ws.Range(11, 1, 11, 3).Style.Fill.SetBackgroundColor(XLColor.FromHtml("#FCE4EC"));
        ws.Cell(12, 1).Value = "EŞLEŞEN";
        ws.Cell(12, 2).Value = s.Eslesen.Count;
        ws.Cell(12, 3).Value = s.Eslesen.Sum(x => x.DosyaTutar);
        ws.Range(12, 1, 12, 3).Style.Fill.SetBackgroundColor(XLColor.LightGreen);

        ws.Range(7, 3, 12, 3).Style.NumberFormat.Format = "#,##0.00";
        ws.Columns(1, 3).AdjustToContents();

        if (s.Hatalar.Count > 0)
        {
            int r = 14;
            ws.Cell(r, 1).Value = "Uyarılar/Hatalar";
            ws.Cell(r, 1).Style.Font.SetBold();
            r++;
            foreach (var h in s.Hatalar) { ws.Cell(r++, 1).Value = h; }
        }

        // ---- ÖDENMEDİ
        var wsO = wb.Worksheets.Add("ÖDENMEDİ");
        SatirSheetYaz(wsO, s.Odenmedi, dosyaTutarliSheet: true);
        wsO.TabColor = XLColor.Red;

        // ---- MÜKERRER
        var wsM = wb.Worksheets.Add("MÜKERRER");
        SatirSheetYaz(wsM, s.Mukerrer, dosyaTutarliSheet: false);
        wsM.TabColor = XLColor.Orange;

        // ---- LOGO'DA FAZLA (Sabnet'te yok)
        if (s.LogoFazlasi.Count > 0)
        {
            var wsLF = wb.Worksheets.Add("LOGODA FAZLA");
            SatirSheetYaz(wsLF, s.LogoFazlasi, dosyaTutarliSheet: false);
            wsLF.TabColor = XLColor.FromHtml("#AD1457");
        }

        // ---- EŞLEŞEN
        var wsE = wb.Worksheets.Add("Eşleşen");
        SatirSheetYaz(wsE, s.Eslesen, dosyaTutarliSheet: false);
        wsE.TabColor = XLColor.Green;

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    private static void SatirSheetYaz(IXLWorksheet ws, List<OdemeKarsilastirSatir> liste, bool dosyaTutarliSheet)
    {
        string[] hdr = dosyaTutarliSheet
            ? new[] { "Banka", "Ad Soyad", "TC", "IBAN", "Dosya Tutar", "Açıklama" }
            : new[] { "Banka", "Ad Soyad", "TC", "Logo Cari", "Logo Tarih", "Logo Tutar", "Adet", "Açıklama" };
        for (int i = 0; i < hdr.Length; i++) ws.Cell(1, i + 1).Value = hdr[i];
        ws.Range(1, 1, 1, hdr.Length).Style.Font.SetBold().Fill.SetBackgroundColor(XLColor.LightGray);

        int r = 2;
        foreach (var x in liste)
        {
            if (dosyaTutarliSheet)
            {
                ws.Cell(r, 1).Value = x.Banka;
                ws.Cell(r, 2).Value = x.AdSoyad;
                ws.Cell(r, 3).SetValue("'" + x.TcKimlikNo);
                ws.Cell(r, 4).SetValue("'" + x.IBAN);
                ws.Cell(r, 5).Value = x.DosyaTutar;
                ws.Cell(r, 5).Style.NumberFormat.Format = "#,##0.00";
                ws.Cell(r, 6).Value = x.Aciklama;
            }
            else
            {
                ws.Cell(r, 1).Value = x.Banka;
                ws.Cell(r, 2).Value = x.AdSoyad;
                ws.Cell(r, 3).SetValue("'" + x.TcKimlikNo);
                ws.Cell(r, 4).Value = x.LogoCariKod;
                if (x.LogoTarih.HasValue) ws.Cell(r, 5).Value = x.LogoTarih.Value;
                ws.Cell(r, 5).Style.DateFormat.Format = "dd.MM.yyyy";
                ws.Cell(r, 6).Value = x.LogoTutar;
                ws.Cell(r, 6).Style.NumberFormat.Format = "#,##0.00";
                ws.Cell(r, 7).Value = x.LogoAdet;
                ws.Cell(r, 8).Value = x.Aciklama;
            }
            r++;
        }
        if (liste.Count > 0)
        {
            ws.Range(1, 1, r - 1, hdr.Length).CreateTable();
        }
        ws.Columns().AdjustToContents();
    }

    // ----------------------------------------------------------------------
    // Yardımcılar
    // ----------------------------------------------------------------------
    private static bool IsCikisOdeme(string? fisTuru)
    {
        if (string.IsNullOrWhiteSpace(fisTuru)) return false;
        var ft = fisTuru.Trim();
        // Gelen / Tahsilat içeren her şeyi dışla (banka, kasa, çek)
        if (ft.Contains("Gelen", StringComparison.OrdinalIgnoreCase)) return false;
        if (ft.Contains("Tahsilat", StringComparison.OrdinalIgnoreCase)) return false;
        // Çıkan ödemeler — havale/tediye
        if (ft.Contains("Gönderilen", StringComparison.OrdinalIgnoreCase)) return true;
        if (ft.Contains("Tediye", StringComparison.OrdinalIgnoreCase)) return true;
        return false;
    }

    private static string NormalizeAd(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        // Logo cari ünvanlarında [18575-1] veya (sözleşme no) gibi ekler olabilir → temizle
        s = System.Text.RegularExpressions.Regex.Replace(s, @"\[[^\]]*\]", " ");
        s = System.Text.RegularExpressions.Regex.Replace(s, @"\([^\)]*\)", " ");
        var sb = new StringBuilder(s.Length);
        foreach (var ch in s.Trim().ToUpper(new CultureInfo("tr-TR")))
        {
            if (char.IsLetter(ch)) sb.Append(ch);
            else if (ch == ' ') sb.Append(' ');
            // sayı/noktalama at — sadece harfler ve boşluk kalsın
        }
        var u = sb.ToString();
        // Türkçe karakterlerin ASCII karşılığı — banka dosyaları farklı olabilir
        u = u.Replace('İ','I').Replace('Ş','S').Replace('Ğ','G').Replace('Ü','U').Replace('Ö','O').Replace('Ç','C');
        // Çoklu boşlukları tek boşluk
        while (u.Contains("  ")) u = u.Replace("  ", " ");
        return u.Trim();
    }

    private static string Norm(string? s) => (s ?? "").Trim().Replace(" ", "").TrimStart('\'');
    private static string NormIban(string? s) => (s ?? "").Trim().Replace(" ", "").ToUpperInvariant();
    private static decimal NormDecimal(IXLCell c)
    {
        if (c.IsEmpty()) return 0m;
        if (c.DataType == XLDataType.Number) return (decimal)c.GetDouble();
        var s = c.GetString().Trim().Replace(".", "").Replace(",", ".");
        return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0m;
    }

    private class DosyaKayit
    {
        public string Banka      { get; set; } = "";
        public string AdSoyad    { get; set; } = "";
        public string IBAN       { get; set; } = "";
        public string TcKimlikNo { get; set; } = "";
        public decimal Tutar     { get; set; }
    }

    private class LogoOdeme
    {
        public string CariKod { get; set; } = "";
        public string Unvan   { get; set; } = "";
        public decimal Tutar  { get; set; }
        public DateTime Tarih { get; set; }
    }

    // Ziraat: DATA, satir 14+; F=IBAN, G=AdSoyad, I=TCKN, J=Tutar
    private static IEnumerable<DosyaKayit> ZiraatOku(byte[] bytes, List<string> hata)
    {
        var list = new List<DosyaKayit>();
        try
        {
            using var ms = new MemoryStream(bytes);
            using var wb = new XLWorkbook(ms);
            var ws = wb.Worksheets.FirstOrDefault(w => w.Name.Equals("DATA", StringComparison.OrdinalIgnoreCase))
                     ?? wb.Worksheet(1);
            int son = ws.LastRowUsed()?.RowNumber() ?? 14;
            for (int r = 14; r <= son; r++)
            {
                var iban = NormIban(ws.Cell(r, 6).GetString());
                if (!iban.StartsWith("TR")) continue;
                list.Add(new DosyaKayit
                {
                    Banka      = "Ziraat",
                    IBAN       = iban,
                    AdSoyad    = ws.Cell(r, 7).GetString().Trim(),
                    TcKimlikNo = Norm(ws.Cell(r, 9).GetString()),
                    Tutar      = NormDecimal(ws.Cell(r, 10))
                });
            }
        }
        catch (Exception ex) { hata.Add("Ziraat parse: " + ex.Message); }
        return list;
    }

    // Garanti: tek sayfa, satir 6+; D=IBAN, H=AdSoyad, I=TCKN, J=Tutar
    private static IEnumerable<DosyaKayit> GarantiOku(byte[] bytes, List<string> hata)
    {
        var list = new List<DosyaKayit>();
        try
        {
            using var ms = new MemoryStream(bytes);
            using var wb = new XLWorkbook(ms);
            var ws = wb.Worksheet(1);
            int son = ws.LastRowUsed()?.RowNumber() ?? 6;
            for (int r = 6; r <= son; r++)
            {
                var iban = NormIban(ws.Cell(r, 4).GetString());
                if (!iban.StartsWith("TR")) continue;
                list.Add(new DosyaKayit
                {
                    Banka      = "Garanti",
                    IBAN       = iban,
                    AdSoyad    = ws.Cell(r, 8).GetString().Trim(),
                    TcKimlikNo = Norm(ws.Cell(r, 9).GetString()),
                    Tutar      = NormDecimal(ws.Cell(r, 10))
                });
            }
        }
        catch (Exception ex) { hata.Add("Garanti parse: " + ex.Message); }
        return list;
    }

    // İş: ÇALIŞMA SAYFASI, satir 4+; F=AdSoyad, G=Tutar, N=IBAN, U=TCKN (yeni format)
    private static IEnumerable<DosyaKayit> IsOku(byte[] bytes, List<string> hata)
    {
        var list = new List<DosyaKayit>();
        try
        {
            using var ms = new MemoryStream(bytes);
            using var wb = new XLWorkbook(ms);
            var ws = wb.Worksheets.FirstOrDefault(w => w.Name.Contains("ÇALIŞMA", StringComparison.OrdinalIgnoreCase))
                     ?? wb.Worksheet(1);
            int son = ws.LastRowUsed()?.RowNumber() ?? 4;
            for (int r = 4; r <= son; r++)
            {
                var iban = NormIban(ws.Cell(r, 14).GetString());
                if (!iban.StartsWith("TR")) continue;
                list.Add(new DosyaKayit
                {
                    Banka      = "İş Bankası",
                    IBAN       = iban,
                    AdSoyad    = ws.Cell(r, 6).GetString().Trim(),
                    TcKimlikNo = Norm(ws.Cell(r, 21).GetString()), // U sütunu (yeni eklenen)
                    Tutar      = NormDecimal(ws.Cell(r, 7))
                });
            }
        }
        catch (Exception ex) { hata.Add("İş Bankası parse: " + ex.Message); }
        return list;
    }
}
