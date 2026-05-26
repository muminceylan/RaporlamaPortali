using ClosedXML.Excel;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

public class OdemeKontrolService
{
    private readonly OdemeService _odeme;

    public OdemeKontrolService(OdemeService odeme) => _odeme = odeme;

    public async Task<KontrolSonuc> KontrolEtAsync(int avansNo,
        byte[]? ziraatBytes, byte[]? garantiBytes, byte[]? isBytes,
        int sozlesmeYili = 2026)
    {
        var sonuc = new KontrolSonuc { AvansNo = avansNo };
        if (!OdemeService.AvansAdlari.TryGetValue(avansNo, out var avansAdi))
        {
            sonuc.Hatalar.Add($"Tanımsız AvansNo: {avansNo}");
            return sonuc;
        }
        sonuc.AvansAdi = avansAdi;

        // DB verisi
        var ham = await _odeme.AvansVerileriniGetirAsync(avansNo, sozlesmeYili);

        var dbZiraat  = new List<OdemeAvansSatiri>();
        var dbGaranti = new List<OdemeAvansSatiri>();
        var dbIs      = new List<OdemeAvansSatiri>();
        foreach (var s in ham)
        {
            if (string.IsNullOrWhiteSpace(s.IBAN) || s.IBAN.Length != 26) continue;
            switch (OdemeService.BankaTespit(s.IBAN))
            {
                case OdemeBankasi.Ziraat:    dbZiraat.Add(s);  break;
                case OdemeBankasi.Garanti:   dbGaranti.Add(s); break;
                case OdemeBankasi.IsBankasi: dbIs.Add(s);      break;
            }
        }

        if (ziraatBytes != null)
            sonuc.Bankalar.Add(Karsilastir("Ziraat", ZiraatOku(ziraatBytes, sonuc.Hatalar), dbZiraat, OdemeBankasi.Ziraat));
        if (garantiBytes != null)
            sonuc.Bankalar.Add(Karsilastir("Garanti", GarantiOku(garantiBytes, sonuc.Hatalar), dbGaranti, OdemeBankasi.Garanti));
        if (isBytes != null)
            sonuc.Bankalar.Add(Karsilastir("İş Bankası", IsOku(isBytes, sonuc.Hatalar), dbIs, OdemeBankasi.IsBankasi));

        return sonuc;
    }

    // ----------------------------------------------------------------------
    // PARSER (ClosedXML — sadece okuma)
    // ----------------------------------------------------------------------
    private static string Norm(string? s) => (s ?? "").Trim().Replace(" ", "").TrimStart('\'');
    private static string NormIban(string? s)
    {
        var n = (s ?? "").Trim().Replace(" ", "").ToUpperInvariant();
        return n;
    }
    private static decimal NormDecimal(IXLCell c)
    {
        if (c.IsEmpty()) return 0m;
        if (c.DataType == XLDataType.Number) return (decimal)c.GetDouble();
        var s = c.GetString().Trim().Replace(".", "").Replace(",", ".");
        return decimal.TryParse(s, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out var d) ? d : 0m;
    }

    // Ziraat: DATA, ilk satir 14, kolonlar: A=Sira,B=İşlemTürü,C=ÖdemeAmaci,D=BankaKodu,E=ŞubeKodu,
    //                                       F=IBAN,G=AdSoyad,H=Soyad,I=TCKN,J=Tutar,K=Açıklama
    private static List<KontrolDosyaSatiri> ZiraatOku(byte[] bytes, List<string> hata)
    {
        var liste = new List<KontrolDosyaSatiri>();
        try
        {
            using var ms = new MemoryStream(bytes);
            using var wb = new XLWorkbook(ms);
            var ws = wb.Worksheets.FirstOrDefault(w => w.Name.Equals("DATA", StringComparison.OrdinalIgnoreCase))
                     ?? wb.Worksheet(1);
            int r = 14;
            int son = ws.LastRowUsed()?.RowNumber() ?? r;
            for (; r <= son; r++)
            {
                var iban = NormIban(ws.Cell(r, 6).GetString());
                if (string.IsNullOrEmpty(iban)) continue;
                if (!iban.StartsWith("TR")) continue;
                liste.Add(new KontrolDosyaSatiri
                {
                    SatirNo    = r,
                    IBAN       = iban,
                    AdSoyad    = ws.Cell(r, 7).GetString().Trim(),
                    TcKimlikNo = Norm(ws.Cell(r, 9).GetString()),
                    Tutar      = NormDecimal(ws.Cell(r, 10)),
                });
            }
        }
        catch (Exception ex) { hata.Add($"Ziraat parse hatası: {ex.Message}"); }
        return liste;
    }

    // Garanti: tek sayfa, ilk satir 6, kolonlar: A=BorcluSube,B=BorcluHesap,C=TT,D=IBAN,E-G=...,
    //                                            H=AdSoyad,I=TCKN,J=Tutar,K=Doviz,L=Tarih,M=Borc,N=Alacak
    private static List<KontrolDosyaSatiri> GarantiOku(byte[] bytes, List<string> hata)
    {
        var liste = new List<KontrolDosyaSatiri>();
        try
        {
            using var ms = new MemoryStream(bytes);
            using var wb = new XLWorkbook(ms);
            var ws = wb.Worksheet(1);
            int r = 6;
            int son = ws.LastRowUsed()?.RowNumber() ?? r;
            for (; r <= son; r++)
            {
                var iban = NormIban(ws.Cell(r, 4).GetString());
                if (string.IsNullOrEmpty(iban)) continue;
                if (!iban.StartsWith("TR")) continue;
                liste.Add(new KontrolDosyaSatiri
                {
                    SatirNo    = r,
                    IBAN       = iban,
                    AdSoyad    = ws.Cell(r, 8).GetString().Trim(),
                    TcKimlikNo = Norm(ws.Cell(r, 9).GetString()),
                    Tutar      = NormDecimal(ws.Cell(r, 10)),
                });
            }
        }
        catch (Exception ex) { hata.Add($"Garanti parse hatası: {ex.Message}"); }
        return liste;
    }

    // İş: ÇALIŞMA SAYFASI, ilk satir 4, kolonlar: A=Sira,B=Tarih,...,F=AliciAdi,G=Tutar,H=PB,
    //                                              N=AliciIBAN,Q=Aciklama. TCKN YOK.
    private static List<KontrolDosyaSatiri> IsOku(byte[] bytes, List<string> hata)
    {
        var liste = new List<KontrolDosyaSatiri>();
        try
        {
            using var ms = new MemoryStream(bytes);
            using var wb = new XLWorkbook(ms);
            var ws = wb.Worksheets.FirstOrDefault(w => w.Name.Contains("ÇALIŞMA", StringComparison.OrdinalIgnoreCase))
                     ?? wb.Worksheet(1);
            int r = 4;
            int son = ws.LastRowUsed()?.RowNumber() ?? r;
            for (; r <= son; r++)
            {
                var iban = NormIban(ws.Cell(r, 14).GetString());
                if (string.IsNullOrEmpty(iban)) continue;
                if (!iban.StartsWith("TR")) continue;
                liste.Add(new KontrolDosyaSatiri
                {
                    SatirNo = r,
                    IBAN    = iban,
                    AdSoyad = ws.Cell(r, 6).GetString().Trim(),
                    Tutar   = NormDecimal(ws.Cell(r, 7)),
                    TcKimlikNo = ""   // İş dosyasında yok
                });
            }
        }
        catch (Exception ex) { hata.Add($"İş parse hatası: {ex.Message}"); }
        return liste;
    }

    // ----------------------------------------------------------------------
    // KARŞILAŞTIRMA (IBAN bazında, çoklu kayıt destekli)
    // ----------------------------------------------------------------------
    private static KontrolBankaSonuc Karsilastir(
        string banka, List<KontrolDosyaSatiri> dosya, List<OdemeAvansSatiri> db, OdemeBankasi beklenenBanka)
    {
        var sonuc = new KontrolBankaSonuc
        {
            Banka            = banka,
            DosyaSatirSayisi = dosya.Count,
            DbSatirSayisi    = db.Count,
            DosyaToplam      = dosya.Sum(x => x.Tutar),
            DbToplam         = db.Sum(x => x.Tutar)
        };

        // IBAN bazında grupla — aynı IBAN'da birden fazla kayıt toplanır
        var dosyaByIban = dosya.GroupBy(x => x.IBAN).ToDictionary(
            g => g.Key,
            g => new {
                Adet  = g.Count(),
                Tutar = g.Sum(x => x.Tutar),
                AdSoyad = g.First().AdSoyad,
                Tckler  = g.Select(x => x.TcKimlikNo).Where(t => !string.IsNullOrEmpty(t)).Distinct().ToList()
            });
        var dbByIban = db.GroupBy(x => x.IBAN).ToDictionary(
            g => g.Key,
            g => new {
                Adet  = g.Count(),
                Tutar = g.Sum(x => x.Tutar),
                AdSoyad = g.First().AdSoyad,
                Tckler  = g.Select(x => x.TcKimlikNo).Where(t => !string.IsNullOrEmpty(t)).Distinct().ToList()
            });

        var tumIbanlar = new HashSet<string>(dosyaByIban.Keys.Concat(dbByIban.Keys));

        foreach (var iban in tumIbanlar)
        {
            var inDosya = dosyaByIban.TryGetValue(iban, out var d) ? d : null;
            var inDb    = dbByIban.TryGetValue(iban, out var x) ? x : null;

            // Yanlış banka? (IBAN bu bankaya ait değil)
            var ibanBankasi = OdemeService.BankaTespit(iban);
            if (ibanBankasi != beklenenBanka && ibanBankasi != OdemeBankasi.Bilinmeyen && inDosya != null)
            {
                sonuc.Uyumsuzluklar.Add(new KontrolUyumsuzluk
                {
                    Banka = banka, Kategori = "YanlisBanka", IBAN = iban,
                    AdSoyad = inDosya.AdSoyad, DosyaTutar = inDosya.Tutar,
                    Aciklama = $"IBAN {ibanBankasi} bankasına ait ama {banka} dosyasında"
                });
                continue;
            }

            if (inDosya != null && inDb == null)
            {
                sonuc.Uyumsuzluklar.Add(new KontrolUyumsuzluk
                {
                    Banka = banka, Kategori = "SadeceDosyada", IBAN = iban,
                    AdSoyad = inDosya.AdSoyad, DosyaTutar = inDosya.Tutar,
                    DosyaTck = string.Join(",", inDosya.Tckler),
                    Aciklama = "Dosyada var, DB'de yok — fazladan ödeme riski"
                });
                continue;
            }
            if (inDosya == null && inDb != null)
            {
                sonuc.Uyumsuzluklar.Add(new KontrolUyumsuzluk
                {
                    Banka = banka, Kategori = "SadeceDbde", IBAN = iban,
                    AdSoyad = inDb.AdSoyad, DbTutar = inDb.Tutar,
                    DbTck = string.Join(",", inDb.Tckler),
                    Aciklama = "DB'de var, dosyaya yazılmamış — eksik ödeme"
                });
                continue;
            }
            // ikisi de var — karşılaştır
            if (inDosya!.Tutar != inDb!.Tutar)
            {
                sonuc.Uyumsuzluklar.Add(new KontrolUyumsuzluk
                {
                    Banka = banka, Kategori = "TutarFarkli", IBAN = iban,
                    AdSoyad = inDosya.AdSoyad,
                    DosyaTutar = inDosya.Tutar, DbTutar = inDb.Tutar,
                    DosyaTck = string.Join(",", inDosya.Tckler),
                    DbTck    = string.Join(",", inDb.Tckler),
                    Aciklama = $"Fark: {(inDosya.Tutar - inDb.Tutar):N2} TL"
                });
            }
            // TCKN kontrolu (dosyada varsa)
            if (inDosya.Tckler.Count > 0 && inDb.Tckler.Count > 0)
            {
                var ortak = inDosya.Tckler.Intersect(inDb.Tckler).ToList();
                if (ortak.Count == 0)
                {
                    sonuc.Uyumsuzluklar.Add(new KontrolUyumsuzluk
                    {
                        Banka = banka, Kategori = "TckFarkli", IBAN = iban,
                        AdSoyad = inDosya.AdSoyad,
                        DosyaTck = string.Join(",", inDosya.Tckler),
                        DbTck    = string.Join(",", inDb.Tckler),
                        DosyaTutar = inDosya.Tutar, DbTutar = inDb.Tutar,
                        Aciklama = "Aynı IBAN'da TCKN uyuşmuyor — KRİTİK"
                    });
                }
            }

            if (inDosya.Tutar == inDb.Tutar
                && (inDosya.Tckler.Count == 0 || inDosya.Tckler.Intersect(inDb.Tckler).Any()))
            {
                sonuc.EslesenSayisi += inDosya.Adet;
            }
        }

        // Sıralama: kritikler önce
        var oncelik = new Dictionary<string, int> {
            ["TckFarkli"]=0, ["TutarFarkli"]=1, ["YanlisBanka"]=2, ["SadeceDosyada"]=3, ["SadeceDbde"]=4
        };
        sonuc.Uyumsuzluklar = sonuc.Uyumsuzluklar
            .OrderBy(u => oncelik.TryGetValue(u.Kategori, out var p) ? p : 99)
            .ThenBy(u => u.IBAN)
            .ToList();

        return sonuc;
    }
}
