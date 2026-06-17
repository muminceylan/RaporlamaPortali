using System.Data;
using Microsoft.Data.SqlClient;
using System.Runtime.InteropServices;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

public class OdemeService
{
    private const string SabNetConnStr =
        "Server=192.168.77.7;Database=SabNetPMHS;User Id=reportuser;" +
        "Password=reportuser;TrustServerCertificate=True;Connect Timeout=30;";

    private const string SablonKlasor = @"C:\Users\muminceylan\Desktop\Ödeme Dosyaları";
    private const string ZiraatSablon  = "Ziraat Bankası Ödeme Dosyası.xlsm";
    private const string GarantiSablon = "Garanti Bankası Ödeme Dosyası.xlsx";
    private const string IsSablon      = "İş Bankası Ödeme Dosyası.xlsx";  // banka şablonu xlsx (makro/format aynen taşınır)

    private static readonly string CiktiKokKlasor = Path.Combine(AppDataPaths.DataRoot, "OdemeDosyalari");

    // AvansNo → görünür ad
    public static readonly Dictionary<int, string> AvansAdlari = new()
    {
        { 37,  "1. Avans" },
        { 40,  "2. Avans" },
        { 46,  "3. Avans" },
        { 58,  "4. Avans" },
        { 59,  "5. Avans" },
        { 132, "6. Avans" }
    };

    public async Task<List<OdemeAvansSatiri>> AvansVerileriniGetirAsync(int avansNo, int sozlesmeYili = 2026)
    {
        var liste = new List<OdemeAvansSatiri>();
        await using var conn = new SqlConnection(SabNetConnStr);
        await conn.OpenAsync();
        var sql = @"
SELECT AF.FormNo,
       ISNULL(AF.TcKimlikNo,'') AS TcKimlikNo,
       ISNULL(CK.AdiSoyadi,'')  AS AdSoyad,
       ISNULL(CK.IBAN,'')       AS IBAN,
       ISNULL(CK.BankaAdi,'')   AS BankaAdi,
       AF.Tutar,
       ISNULL(AF.KaynakBolge,'') AS KaynakBolge
FROM PMHS_AvansFormu AF
LEFT JOIN PMHS_CiftciKarti CK ON CK.TcKimlikNo = AF.TcKimlikNo
WHERE AF.SozlesmeYili = @yil
  AND AF.AvansNo = @av
  AND AF.KaynakBolge <> ''
ORDER BY CK.AdiSoyadi, AF.FormNo";
        await using var cmd = new SqlCommand(sql, conn);
        cmd.Parameters.AddWithValue("@yil", sozlesmeYili);
        cmd.Parameters.AddWithValue("@av", avansNo);
        cmd.CommandTimeout = 120;
        await using var rd = await cmd.ExecuteReaderAsync();
        while (await rd.ReadAsync())
        {
            liste.Add(new OdemeAvansSatiri
            {
                FormNo      = rd["FormNo"]?.ToString() ?? "",
                TcKimlikNo  = rd["TcKimlikNo"]?.ToString() ?? "",
                AdSoyad     = rd["AdSoyad"]?.ToString() ?? "",
                IBAN        = (rd["IBAN"]?.ToString() ?? "").Replace(" ", "").ToUpperInvariant(),
                BankaAdi    = rd["BankaAdi"]?.ToString() ?? "",
                Tutar       = rd["Tutar"] == DBNull.Value ? 0m : Convert.ToDecimal(rd["Tutar"]),
                KaynakBolge = rd["KaynakBolge"]?.ToString() ?? ""
            });
        }
        return liste;
    }

    public static OdemeBankasi BankaTespit(string iban)
    {
        if (string.IsNullOrWhiteSpace(iban) || iban.Length != 26 || !iban.StartsWith("TR")) return OdemeBankasi.Bilinmeyen;
        if (!iban.Substring(2).All(char.IsDigit)) return OdemeBankasi.Bilinmeyen;
        var kod = iban.Substring(4, 5);
        return kod switch
        {
            "00010" => OdemeBankasi.Ziraat,
            "00012" => OdemeBankasi.Ziraat,
            "00062" => OdemeBankasi.Garanti,
            "00064" => OdemeBankasi.IsBankasi,
            _       => OdemeBankasi.Bilinmeyen
        };
    }

    private static bool IbanGecerliMi(string iban) =>
        !string.IsNullOrWhiteSpace(iban) && iban.Length == 26 && iban.StartsWith("TR")
        && iban.Substring(2).All(char.IsDigit);

    public async Task<OdemeHazirlamaSonuc> HazirlaAsync(int avansNo, DateTime odemeTarihi, int sozlesmeYili = 2026)
    {
        if (!AvansAdlari.TryGetValue(avansNo, out var avansAdi))
            throw new InvalidOperationException($"Tanimsiz AvansNo: {avansNo}");

        // "1. Avans" → "1. AVANS ÖDEMESİ"
        var aciklama = $"{avansAdi.Replace(". Avans", ". AVANS")} ÖDEMESİ";

        var sonuc = new OdemeHazirlamaSonuc
        {
            AvansNo     = avansNo,
            AvansAdi    = avansAdi,
            OdemeTarihi = odemeTarihi
        };

        var ham = await AvansVerileriniGetirAsync(avansNo, sozlesmeYili);

        var ziraat   = new List<OdemeAvansSatiri>();
        var garanti  = new List<OdemeAvansSatiri>();
        var isBank   = new List<OdemeAvansSatiri>();

        foreach (var s in ham)
        {
            if (!IbanGecerliMi(s.IBAN))
            {
                sonuc.IbanEksikUyarilari.Add(new OdemeUyariSatiri
                {
                    FormNo = s.FormNo, TcKimlikNo = s.TcKimlikNo, AdSoyad = s.AdSoyad,
                    IBAN = s.IBAN, BankaAdi = s.BankaAdi, Tutar = s.Tutar,
                    KaynakBolge = s.KaynakBolge,
                    Sebep = string.IsNullOrWhiteSpace(s.IBAN) ? "IBAN boş" : "IBAN geçersiz"
                });
                continue;
            }
            s.Banka = BankaTespit(s.IBAN);
            switch (s.Banka)
            {
                case OdemeBankasi.Ziraat:    ziraat.Add(s);  break;
                case OdemeBankasi.Garanti:   garanti.Add(s); break;
                case OdemeBankasi.IsBankasi: isBank.Add(s);  break;
                default:
                    sonuc.DigerBankaUyarilari.Add(new OdemeUyariSatiri
                    {
                        FormNo = s.FormNo, TcKimlikNo = s.TcKimlikNo, AdSoyad = s.AdSoyad,
                        IBAN = s.IBAN, BankaAdi = s.BankaAdi, Tutar = s.Tutar,
                        KaynakBolge = s.KaynakBolge,
                        Sebep = $"Bilinmeyen banka (IBAN kodu {s.IBAN.Substring(4,5)})"
                    });
                    break;
            }
        }

        // Çıktı klasörü
        var damga = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var alt   = Path.Combine(CiktiKokKlasor, $"Avans{avansNo}_{damga}");
        Directory.CreateDirectory(alt);
        sonuc.CiktiKlasor = alt;

        sonuc.ZiraatAdet  = ziraat.Count;
        sonuc.ZiraatTutar = ziraat.Sum(x => x.Tutar);
        sonuc.GarantiAdet = garanti.Count;
        sonuc.GarantiTutar= garanti.Sum(x => x.Tutar);
        sonuc.IsAdet      = isBank.Count;
        sonuc.IsTutar     = isBank.Sum(x => x.Tutar);

        // Banka x bölge özeti (büyükten küçüğe sıralı)
        OdemeBankaOzeti OzetUret(string ad, OdemeBankasi b, List<OdemeAvansSatiri> liste) => new()
        {
            Banka = b,
            Ad    = ad,
            Bolgeler = liste
                .GroupBy(x => x.KaynakBolge)
                .Select(g => new BankaBolgeKayit
                {
                    Bolge = g.Key,
                    Adet  = g.Count(),
                    Tutar = g.Sum(x => x.Tutar)
                })
                .OrderByDescending(x => x.Tutar)
                .ToList()
        };
        sonuc.BankaBolgeOzeti.Add(OzetUret("Ziraat Bankası", OdemeBankasi.Ziraat, ziraat));
        sonuc.BankaBolgeOzeti.Add(OzetUret("Garanti BBVA",   OdemeBankasi.Garanti, garanti));
        sonuc.BankaBolgeOzeti.Add(OzetUret("İş Bankası",      OdemeBankasi.IsBankasi, isBank));

        // Kullanıcı ayarları (banka hesap bilgileri) — JSON dosyasından oku
        var bankaAyari = BankaHesapAyarlari.Yukle();

        // Excel COM ile dosyaları üret (her birini ayrı COM oturumunda; sorun çıkarsa devam et)
        if (ziraat.Count > 0)
        {
            try { sonuc.ZiraatDosya = ZiraatDosyasiUret(ziraat, alt, avansNo, avansAdi, odemeTarihi, aciklama, bankaAyari); }
            catch (Exception ex) { sonuc.Hatalar.Add($"Ziraat dosyası üretilemedi: {ex.Message}"); }
        }
        if (garanti.Count > 0)
        {
            try { sonuc.GarantiDosya = GarantiDosyasiUret(garanti, alt, avansNo, avansAdi, odemeTarihi, aciklama, bankaAyari); }
            catch (Exception ex) { sonuc.Hatalar.Add($"Garanti dosyası üretilemedi: {ex.Message}"); }
        }
        if (isBank.Count > 0)
        {
            try { sonuc.IsDosya = IsDosyasiUret(isBank, alt, avansNo, avansAdi, odemeTarihi, aciklama, bankaAyari); }
            catch (Exception ex) { sonuc.Hatalar.Add($"İş Bankası dosyası üretilemedi: {ex.Message}"); }
        }

        return sonuc;
    }

    // ------------------------------------------------------------------
    // Excel COM yardımcıları
    // ------------------------------------------------------------------
    private static void ComBosalt(dynamic? obj)
    {
        if (obj == null) return;
        try { Marshal.ReleaseComObject(obj); } catch { }
    }

    private static void ExcelCom(string srcPath, string dstPath, Action<dynamic> doldur, int fileFormat)
    {
        if (File.Exists(dstPath)) File.Delete(dstPath);
        File.Copy(srcPath, dstPath, overwrite: false);

        var excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null) throw new InvalidOperationException("Excel kurulu değil veya COM erişimi yok.");

        dynamic excel = Activator.CreateInstance(excelType)!;
        dynamic? wb = null;
        try
        {
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            excel.EnableEvents = false;
            // Calculation manual: -4135 = xlCalculationManual
            try { excel.Calculation = -4135; } catch { }

            wb = excel.Workbooks.Open(dstPath, false, false);   // UpdateLinks=false, ReadOnly=false
            doldur(wb);

            // Geri ayarla
            try { excel.Calculation = -4105; } catch { } // xlCalculationAutomatic

            // Save() mevcut format ile üzerine yazar — File.Copy ile dst yerinde olduğu için
            // format zaten doğru (.xls için .xls, .xlsm için .xlsm). SaveAs "üzerine yazılsın mı?"
            // popup'ı veriyor DisplayAlerts'a rağmen. fileFormat parametresi sadece referans amaçlı.
            _ = fileFormat;
            try { excel.DisplayAlerts = false; } catch { }
            wb.Save();
            wb.Close(false);
        }
        finally
        {
            ComBosalt(wb);
            try { excel.Quit(); } catch { }
            ComBosalt(excel);
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    // Bir 2D dizi olusturup tek COM cagrisiyla Range'e atar — hucre hucre yazimdan 20-50x hizli
    private static void BulkYaz(dynamic ws, int ilkSatir, int ilkSutun, object[,] veri)
    {
        int satirSayisi = veri.GetLength(0);
        int sutunSayisi = veri.GetLength(1);
        if (satirSayisi == 0 || sutunSayisi == 0) return;
        dynamic ilk = ws.Cells[ilkSatir, ilkSutun];
        dynamic son = ws.Cells[ilkSatir + satirSayisi - 1, ilkSutun + sutunSayisi - 1];
        dynamic rng = ws.Range[ilk, son];
        rng.Value2 = veri;
    }

    // Ziraat: DATA sayfası, başlık D7=ödeme tarihi, D8=bordro açıklaması, veri satır 14'ten itibaren
    private static string ZiraatDosyasiUret(List<OdemeAvansSatiri> liste, string klasor,
        int avansNo, string avansAdi, DateTime odemeTarihi, string aciklama, BankaHesapAyarlari bankaAyari)
    {
        var dst = Path.Combine(klasor, $"Ziraat_Bankasi_{avansNo}_{odemeTarihi:yyyyMMdd}.xlsm");
        var src = Path.Combine(SablonKlasor, ZiraatSablon);

        ExcelCom(src, dst, wb =>
        {
            dynamic ws = wb.Worksheets["DATA"];
            // Header bilgileri
            ws.Range["D5"].Value2  = bankaAyari.ZiraatKurumKodu;     // Kurum Kodu (değiştirilebilir)
            ws.Range["D7"].Value2  = odemeTarihi.ToOADate();
            ws.Cells[7, 4].NumberFormatLocal = "gg.aa.yyyy";
            ws.Range["D8"].Value2  = aciklama;
            ws.Range["D11"].Value2 = bankaAyari.ZiraatKurumHesapNo;  // Kurum Hesap No (değiştirilebilir)

            const int ILK = 14;
            // Mevcut tüm veri satırlarını temizle (şablonun örnek kayıtları varsa)
            ws.Range[ws.Cells[ILK, 1], ws.Cells[70000, 11]].ClearContents();

            int n = liste.Count;
            if (n == 0) return;
            var veri = new object[n, 11];
            for (int i = 0; i < n; i++)
            {
                var s = liste[i];
                veri[i, 0]  = i + 1;
                veri[i, 1]  = "M01_HesabaÖdeme";
                veri[i, 2]  = "99_Diğer";
                veri[i, 3]  = "";
                veri[i, 4]  = "";
                veri[i, 5]  = s.IBAN;
                veri[i, 6]  = s.AdSoyad;
                veri[i, 7]  = "";
                veri[i, 8]  = "'" + s.TcKimlikNo; // text olarak yaz
                veri[i, 9]  = (double)s.Tutar;
                veri[i, 10] = aciklama;
            }
            BulkYaz(ws, ILK, 1, veri);
        }, fileFormat: 52);
        return dst;
    }

    // Garanti: tek sayfa, B2=toplam adet (otomatik), B3=toplam tutar (otomatik). Veri satır 6'dan itibaren.
    private static string GarantiDosyasiUret(List<OdemeAvansSatiri> liste, string klasor,
        int avansNo, string avansAdi, DateTime odemeTarihi, string aciklama, BankaHesapAyarlari bankaAyari)
    {
        var dst = Path.Combine(klasor, $"Garanti_Bankasi_{avansNo}_{odemeTarihi:yyyyMMdd}.xlsx");
        var src = Path.Combine(SablonKlasor, GarantiSablon);

        // Banka şube/hesap değişebilir → Ayarlar > Banka Hesap Ayarları'ndan
        string borcluSube  = bankaAyari.GarantiBorcluSube;
        string borcluHesap = bankaAyari.GarantiBorcluHesap;
        var odemeTarihStr = odemeTarihi.ToString("ddMMyyyy"); // Garanti formatı

        ExcelCom(src, dst, wb =>
        {
            dynamic ws = wb.Worksheets[1];
            const int ILK = 6;
            // Şablondaki örnek 254 satırı temizle
            ws.Range[ws.Cells[ILK, 1], ws.Cells[70000, 18]].ClearContents();

            int n = liste.Count;
            if (n == 0) return;
            var veri = new object[n, 14];
            for (int i = 0; i < n; i++)
            {
                var s = liste[i];
                veri[i, 0]  = borcluSube;
                veri[i, 1]  = borcluHesap;
                veri[i, 2]  = "";
                veri[i, 3]  = s.IBAN;
                veri[i, 4]  = "";
                veri[i, 5]  = "";
                veri[i, 6]  = "";
                veri[i, 7]  = s.AdSoyad;
                veri[i, 8]  = "'" + s.TcKimlikNo;
                veri[i, 9]  = (double)s.Tutar;
                veri[i, 10] = "TRY";
                veri[i, 11] = odemeTarihStr;
                veri[i, 12] = aciklama;
                veri[i, 13] = aciklama;
            }
            BulkYaz(ws, ILK, 1, veri);
            // B2 ve B3 otomatik formül — dokunmuyoruz.
        }, fileFormat: 51);
        return dst;
    }

    // İş Bankası: ÇALIŞMA SAYFASI, veri satır 4'ten itibaren. Gönderen şube/hesap/IBAN Ayarlar'dan.
    private static string IsDosyasiUret(List<OdemeAvansSatiri> liste, string klasor,
        int avansNo, string avansAdi, DateTime odemeTarihi, string aciklama, BankaHesapAyarlari bankaAyari)
    {
        // Banka portali xlsx (OpenXML) formatını bekliyor — şablon makroları/formatları aynen
        // taşınır; format/formül/data-validation değiştirilirse banka iade gönderiyor.
        var dst = Path.Combine(klasor, $"Is_Bankasi_{avansNo}_{odemeTarihi:yyyyMMdd}.xlsx");
        var src = Path.Combine(SablonKlasor, IsSablon);

        string gondSube  = bankaAyari.IsGondernSube;
        string gondHesap = bankaAyari.IsGondernHesap;
        string gondIban  = bankaAyari.IsGondernIban;

        ExcelCom(src, dst, wb =>
        {
            dynamic ws = wb.Worksheets["ÇALIŞMA SAYFASI"];
            const int ILK = 4;
            // Şablon xlsx — banka 63.460 satır pre-format etmiş. ClearContents sadece
            // hücre value'larını siler; format/formül/data-validation/makrolar korunur.
            ws.Range[ws.Cells[ILK, 1], ws.Cells[63460, 23]].ClearContents();

            int n = liste.Count;
            if (n == 0) return;

            // B2: DOSYA TARİHİ (header) — ödeme tarihi yaz
            ws.Cells[2, 2].Value2 = odemeTarihi.ToOADate();
            ws.Cells[2, 2].NumberFormatLocal = "gg.aa.yyyy";

            // Veri kolonu B: İşlem Tarihi
            dynamic tarihRng = ws.Range[ws.Cells[ILK, 2], ws.Cells[ILK + n - 1, 2]];
            tarihRng.NumberFormatLocal = "gg.aa.yyyy";

            // A-U arası 21 sütun. Şablon header (satır 3):
            //   A SIRANO | B İŞLEMTARİHİ | C GÖNDERENŞUBEKODU | D GÖNDERENHESAPNO | E GÖNDERENIBAN
            //   F ALICIADI | G TUTAR | H PARABİRİMİ | I ALICIBANKAKODU | J ALICIBANKAADI
            //   K ALICIŞUBEKODU | L ALICISUBEADI | M ALICIHESAPNO | N ALICIIBAN | O ALICIADRES
            //   P ALICIŞEHİR | Q AÇIKLAMA | R GÖNDERENREFERANS | S ALICIREFERANS
            //   T ALICIVERGİDAİRESİ | U ALICIVERGİNO  ← TC Kimlik buraya
            var veri = new object[n, 21];
            double tarihSerial = odemeTarihi.ToOADate();
            for (int i = 0; i < n; i++)
            {
                var s = liste[i];
                veri[i, 0]  = i + 1;            // A SIRA NO
                veri[i, 1]  = tarihSerial;       // B İŞLEM TARİHİ
                veri[i, 2]  = gondSube;          // C
                veri[i, 3]  = gondHesap;         // D
                veri[i, 4]  = gondIban;          // E
                veri[i, 5]  = s.AdSoyad;         // F ALICIADI
                veri[i, 6]  = (double)s.Tutar;   // G TUTAR
                veri[i, 7]  = "TRY";             // H
                // I-M boş (alıcı banka şube hesap — IBAN'dan zaten anlaşılır)
                veri[i, 8]  = ""; veri[i, 9]  = ""; veri[i, 10] = ""; veri[i, 11] = ""; veri[i, 12] = "";
                veri[i, 13] = s.IBAN;            // N ALICI IBAN
                veri[i, 14] = ""; veri[i, 15] = "";   // O,P alıcı adres/şehir
                veri[i, 16] = aciklama;          // Q AÇIKLAMA
                veri[i, 17] = ""; veri[i, 18] = ""; veri[i, 19] = "";   // R,S,T
                veri[i, 20] = "'" + (s.TcKimlikNo ?? ""); // U ALICIVERGİNO (TC) — text olsun diye apostrof
            }
            BulkYaz(ws, ILK, 1, veri);
        }, fileFormat: 51);   // 51 = xlsx (OpenXML) — banka portali bu formatı bekliyor
        return dst;
    }

    // BANKAYA GÖRE BÖLGE ÖDEME RAPORU — ekran tasarımıyla uyumlu Excel
    public byte[] BankaBolgeOzetExcel(OdemeHazirlamaSonuc sonuc)
    {
        using var wb = new ClosedXML.Excel.XLWorkbook();
        var ws = wb.Worksheets.Add("Banka Bolge Ozet");

        // Sütun genişlikleri
        ws.Column(1).Width = 30;
        ws.Column(2).Width = 22;

        // Başlık (A1:B1) — koyu mavi
        var bas = ws.Range(1, 1, 1, 2).Merge();
        bas.Value = "BANKAYA GÖRE BÖLGE ÖDEME RAPORU";
        bas.Style.Font.Bold = true;
        bas.Style.Font.FontSize = 14;
        bas.Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
        bas.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#1565C0");
        bas.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Left;
        bas.Style.Alignment.Vertical   = ClosedXML.Excel.XLAlignmentVerticalValues.Center;
        ws.Row(1).Height = 24;

        // Rapor Tarihi (A2:B2) — açık mavi, ortalı
        var tar = ws.Range(2, 1, 2, 2).Merge();
        tar.Value = $"Rapor Tarihi: {sonuc.RaporZamani:dd.MM.yyyy HH:mm}";
        tar.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#E3F2FD");
        tar.Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Center;
        tar.Style.Font.Italic = true;

        // Tablo başlığı (A3:B3) — açık mavi
        ws.Cell(3, 1).Value = "BANKA";
        ws.Cell(3, 2).Value = "ÖDEME TUTARI";
        var thRng = ws.Range(3, 1, 3, 2);
        thRng.Style.Font.Bold = true;
        thRng.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#BBDEFB");
        ws.Cell(3, 2).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Right;

        int row = 4;
        foreach (var b in sonuc.BankaBolgeOzeti)
        {
            // Banka adı satırı (mavi vurgu)
            var bRng = ws.Range(row, 1, row, 2).Merge();
            bRng.Value = b.Ad;
            bRng.Style.Font.Bold = true;
            bRng.Style.Font.FontColor = ClosedXML.Excel.XLColor.FromHtml("#0D47A1");
            bRng.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#E3F2FD");
            row++;

            foreach (var bk in b.Bolgeler)
            {
                ws.Cell(row, 1).Value = bk.Bolge;
                ws.Cell(row, 1).Style.Alignment.Indent = 2;
                ws.Cell(row, 2).Value = bk.Tutar;
                ws.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00\\ \\T\\L";
                ws.Cell(row, 2).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Right;
                row++;
            }

            // Banka TOPLAM (koyu mavi)
            ws.Cell(row, 1).Value = $"{b.Ad} TOPLAM";
            ws.Cell(row, 2).Value = b.Toplam;
            ws.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00\\ \\T\\L";
            ws.Cell(row, 2).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Right;
            var topRng = ws.Range(row, 1, row, 2);
            topRng.Style.Font.Bold = true;
            topRng.Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
            topRng.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#1565C0");
            row++;
        }

        // GENEL TOPLAM
        ws.Cell(row, 1).Value = "GENEL TOPLAM";
        ws.Cell(row, 2).Value = sonuc.GenelToplam;
        ws.Cell(row, 2).Style.NumberFormat.Format = "#,##0.00\\ \\T\\L";
        ws.Cell(row, 2).Style.Alignment.Horizontal = ClosedXML.Excel.XLAlignmentHorizontalValues.Right;
        var gtRng = ws.Range(row, 1, row, 2);
        gtRng.Style.Font.Bold = true;
        gtRng.Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
        gtRng.Style.Font.FontSize  = 12;
        gtRng.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#0D47A1");
        ws.Row(row).Height = 22;

        // Tüm tablo etrafına kenarlık
        var tum = ws.Range(1, 1, row, 2);
        tum.Style.Border.OutsideBorder = ClosedXML.Excel.XLBorderStyleValues.Medium;
        tum.Style.Border.OutsideBorderColor = ClosedXML.Excel.XLColor.FromHtml("#1565C0");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    // Uyarı listelerini xlsx olarak üret (basit ClosedXML)
    public byte[] UyariListesiExcel(List<OdemeUyariSatiri> uyarilar, string baslik)
    {
        using var wb = new ClosedXML.Excel.XLWorkbook();
        var ws = wb.Worksheets.Add(baslik.Length > 30 ? baslik.Substring(0, 30) : baslik);
        string[] h = { "Form No", "TCKN", "Ad Soyad", "IBAN", "Banka", "Tutar", "Bölge", "Sebep" };
        for (int i = 0; i < h.Length; i++) { ws.Cell(1, i + 1).Value = h[i]; ws.Cell(1, i + 1).Style.Font.Bold = true; }
        int r = 2;
        foreach (var u in uyarilar)
        {
            ws.Cell(r, 1).Value = u.FormNo;
            ws.Cell(r, 2).Value = u.TcKimlikNo;
            ws.Cell(r, 3).Value = u.AdSoyad;
            ws.Cell(r, 4).Value = u.IBAN;
            ws.Cell(r, 5).Value = u.BankaAdi;
            ws.Cell(r, 6).Value = u.Tutar;
            ws.Cell(r, 6).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(r, 7).Value = u.KaynakBolge;
            ws.Cell(r, 8).Value = u.Sebep;
            r++;
        }
        ws.Columns().AdjustToContents();
        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }
}
