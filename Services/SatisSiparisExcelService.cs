using System.Globalization;
using ClosedXML.Excel;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Pendik/Adana benzeri Excel formatını parse eder. Format YATAY MATRİS:
///   - Satırlar:  R1..R13 → sipariş başlık alanları (Fabrika, Sevk Tarihi, Müşteri Kodu vb.)
///                R14    → sütun başlıkları (Araç No, Squ Kodu = malzeme kodu, vb.)
///                R15+   → her satır bir malzeme
///   - Sütunlar:  A..H  → malzeme bilgileri (Squ Kodu = D, Malzeme adı = H)
///                I, J, K, ... → her sütun bir SİPARİŞ (farklı müşteri / sevk adresi)
///                Hücredeki sayı → o malzemenin o siparişte KOLİ miktarı
///
/// Bizim parse mantığımız:
///   1. Sheet seç (varsayılan: tüm görünür sheet'leri tara)
///   2. R14'te B (Squ Kodu) ve H (malzeme adı) sütunlarını sabit al
///   3. Sütun I'dan itibaren her sütun için:
///        - Müşteri kodu boş ise sütunu ATLA (gerçek bir sipariş değil — başlık/özet)
///        - Bir SatisSiparisFis oluştur
///        - R15+ satırlardan o sütunun değerini oku; >0 ise satır ekle
///   4. Satış elemanı + teslim tarihi sayfada hesaplanır (mevcut tarihe göre).
/// </summary>
public class SatisSiparisExcelService
{
    private readonly ILogger<SatisSiparisExcelService> _log;
    public SatisSiparisExcelService(ILogger<SatisSiparisExcelService> log) => _log = log;

    /// <summary>Header satır indeksleri (1-tabanlı). Excel'de fix.</summary>
    private const int ROW_FABRIKA      = 1;
    private const int ROW_SEVK_TARIHI  = 2;
    private const int ROW_MUSTERI_KODU = 3;
    private const int ROW_DEPO_KODU    = 4;
    private const int ROW_MUSTERI      = 5;
    private const int ROW_SEVK_DEPO    = 6;
    private const int ROW_ODEME        = 8;
    private const int ROW_ARAC_NO      = 14;   // "2026-25--10" gibi sütun başlığı
    private const int ROW_ILK_MALZEME  = 15;

    /// <summary>Squ Kodu sütunu (malzeme kodu) — sabit D kolonu.</summary>
    private const int COL_MALZEME_KODU = 4;    // D
    /// <summary>Malzeme adı (önizleme) — sabit H kolonu (15+ satırlarında).</summary>
    private const int COL_MALZEME_ADI  = 8;    // H
    /// <summary>Sipariş sütunları buradan itibaren (I).</summary>
    private const int COL_ILK_SIPARIS  = 9;    // I

    public SatisSiparisExcelSonuc Parse(byte[] dosyaBytes, string dosyaAdi, DateTime referansTarihi)
    {
        var sonuc = new SatisSiparisExcelSonuc { DosyaAdi = dosyaAdi };

        using var ms = new MemoryStream(dosyaBytes);
        using var wb = new XLWorkbook(ms);

        var satisElemani = HaftalikSatisElemaniKodu(referansTarihi);
        var teslimTarihi = SonrakiPazartesi(referansTarihi);

        // CİPS + İÇECEK sheet'lerini işle — başka sheet'ler (ÖZET vb.) atlanır.
        bool islendi = false;
        foreach (var ws in wb.Worksheets)
        {
            if (ws.Visibility != XLWorksheetVisibility.Visible) continue;
            var ad = (ws.Name ?? "").Trim();
            if (string.IsNullOrEmpty(SheetTipi(ad))) continue;

            int eklenenSheet = ParseSheet(ws, sonuc, satisElemani, teslimTarihi);
            _log.LogInformation("Sheet '{Sheet}' işlendi: {Adet} sipariş", ad, eklenenSheet);
            islendi = true;
        }
        if (!islendi)
            sonuc.Uyarilar.Add("Dosyada 'CİPS' veya 'İÇECEK' adında sheet bulunamadı.");

        return sonuc;
    }

    /// <summary>
    /// Sheet adından kanonik tip döner: "CIPS", "ICECEK" veya "" (işlenmeyecek).
    /// Transfer'de TRADING_GRP, PESIN ödeme kodu ve satış elemanı SC prefix'i bu tipe göre seçilir.
    /// </summary>
    public static string SheetTipi(string ad)
    {
        if (string.IsNullOrWhiteSpace(ad)) return "";
        var n = ad.Trim().ToUpperInvariant()
                  .Replace("İ", "I").Replace("Ç", "C").Replace("Ğ", "G");
        if (n == "CIPS" || n == "CHIPS") return "CIPS";
        if (n == "ICECEK" || n == "ICECEKLER") return "ICECEK";
        return "";
    }

    private int ParseSheet(IXLWorksheet ws, SatisSiparisExcelSonuc sonuc,
                           string satisElemani, DateTime teslimTarihi)
    {
        int eklenen = 0;
        var sonKolon = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
        if (sonKolon < COL_ILK_SIPARIS) return 0;
        var sonSatir = ws.LastRowUsed()?.RowNumber() ?? 0;
        if (sonSatir < ROW_ILK_MALZEME) return 0;

        for (int col = COL_ILK_SIPARIS; col <= sonKolon; col++)
        {
            var cariKodu = ws.Cell(ROW_MUSTERI_KODU, col).GetString().Trim();
            if (string.IsNullOrWhiteSpace(cariKodu)) continue; // boş sütun (özet/başlık) atla
            // 120.xx.xx... gibi gerçek bir kod mu?
            if (!cariKodu.Contains('.')) continue;

            var fis = new SatisSiparisFis
            {
                ExcelSutunIndex  = col,
                Sheet            = ws.Name ?? "",
                FabrikaCikis     = ws.Cell(ROW_FABRIKA,      col).GetString().Trim(),
                CariKodu         = cariKodu,
                SevkiyatAdresi   = ws.Cell(ROW_DEPO_KODU,    col).GetString().Trim(),
                CariUnvani       = ws.Cell(ROW_MUSTERI,      col).GetString().Trim(),
                SevkDepoAdi      = ws.Cell(ROW_SEVK_DEPO,    col).GetString().Trim(),
                SevkTarihi       = TarihOku(ws.Cell(ROW_SEVK_TARIHI, col)),
                OdemeTuru        = ws.Cell(ROW_ODEME, col).GetString().Trim(),
                AracNo           = ws.Cell(ROW_ARAC_NO,      col).GetString().Trim(),
                // İÇECEK sayfasında satış elemanı koduna "SC" prefix eklenir (örn 26W25 → SC26W25)
                SatisElemani     = SheetTipi(ws.Name ?? "") == "ICECEK"
                                     ? "SC" + satisElemani
                                     : satisElemani,
                TeslimTarihi     = teslimTarihi,
            };

            // Satırları doldur
            for (int r = ROW_ILK_MALZEME; r <= sonSatir; r++)
            {
                var miktarHucresi = ws.Cell(r, col);
                if (miktarHucresi.IsEmpty()) continue;
                if (!TrySayi(miktarHucresi, out decimal koli) || koli <= 0) continue;

                var malzemeKodu = ws.Cell(r, COL_MALZEME_KODU).GetString().Trim();
                if (string.IsNullOrWhiteSpace(malzemeKodu)) continue;
                // Yazı şeklinde sayı veya bozuk satır
                if (!malzemeKodu.Any(char.IsDigit)) continue;

                // H sütununda malzeme adı genelde "KOD-AÇIKLAMA" formatında — açıklama kısmını al
                var malzemeAdiHam = ws.Cell(r, COL_MALZEME_ADI).GetString().Trim();
                var malzemeAdi = malzemeAdiHam.Contains('-')
                    ? malzemeAdiHam.Substring(malzemeAdiHam.IndexOf('-') + 1).Trim()
                    : malzemeAdiHam;

                fis.Satirlar.Add(new SatisSiparisSatir
                {
                    MalzemeKodu = malzemeKodu,
                    MalzemeAdi  = malzemeAdi,
                    KoliMiktari = koli,
                });
            }

            if (fis.Satirlar.Count == 0)
            {
                sonuc.Uyarilar.Add($"Sheet '{ws.Name}' sütun {col} ({fis.CariKodu}): hiç satır okunamadı, atlandı.");
                continue;
            }
            sonuc.Fisler.Add(fis);
            eklenen++;
        }
        return eklenen;
    }

    // ── Yardımcılar ───────────────────────────────────────────────────────────

    /// <summary>
    /// Satış elemanı kodu — "{YY}W{HH}" formatı.
    /// İçinde bulunduğumuz haftada yapılan iş aslında BİR SONRAKİ haftanın siparişi olduğu için
    /// referans tarihe 7 gün eklenir; ISO hafta numarası ve yıl o tarih üzerinden hesaplanır.
    /// Örn: referans 12.06.2026 (ISO hafta 24) → +7 gün → 19.06.2026 (ISO hafta 25) → "26W25".
    /// Yıl sınırı doğru: 28.12.2025 (hafta 52/53) → +7 → 04.01.2026 (hafta 1) → "26W01".
    /// </summary>
    public static string HaftalikSatisElemaniKodu(DateTime referans)
    {
        var hedef = referans.AddDays(7);
        int hafta = ISOWeek.GetWeekOfYear(hedef);
        int yilSon2 = hedef.Year % 100;
        return $"{yilSon2:00}W{hafta:00}";
    }

    /// <summary>Verilen tarihten sonraki ilk Pazartesi (kendisi pazartesi ise +7).</summary>
    public static DateTime SonrakiPazartesi(DateTime referans)
    {
        var t = referans.Date;
        // Pazartesi=1, Salı=2, ..., Pazar=0 (C# DayOfWeek)
        int gun = (int)t.DayOfWeek;
        // (1 - gun + 7) % 7 = sonraki pazartesiye kaç gün; 0 olursa zaten pazartesi → 7
        int eklenecek = ((int)DayOfWeek.Monday - gun + 7) % 7;
        if (eklenecek == 0) eklenecek = 7;
        return t.AddDays(eklenecek);
    }

    private static DateTime TarihOku(IXLCell cell)
    {
        if (cell.IsEmpty()) return DateTime.MinValue;
        if (cell.DataType == XLDataType.DateTime) return cell.GetDateTime();
        var s = cell.GetString().Trim();
        if (DateTime.TryParseExact(s,
                new[] { "dd/MM/yyyy", "dd.MM.yyyy", "dd-MM-yyyy", "yyyy-MM-dd" },
                CultureInfo.InvariantCulture, DateTimeStyles.None, out var dt))
            return dt;
        if (DateTime.TryParse(s, new CultureInfo("tr-TR"), DateTimeStyles.None, out dt))
            return dt;
        return DateTime.MinValue;
    }

    private static bool TrySayi(IXLCell c, out decimal v)
    {
        v = 0m;
        if (c.IsEmpty()) return false;
        if (c.DataType == XLDataType.Number) { v = (decimal)c.GetDouble(); return true; }
        var s = c.GetString().Trim().Replace(".", "").Replace(",", ".");
        return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out v);
    }
}
