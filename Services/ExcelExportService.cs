using ClosedXML.Excel;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Excel export servisi - Raporları Excel dosyasına aktarır
/// </summary>
public class ExcelExportService
{
    /// <summary>
    /// Yan Ürünler raporunu Excel'e aktarır
    /// </summary>
    public byte[] ExportYanUrunlerRaporu(
        List<YanUrunOzet> yanUrunler,
        List<AlkolOzet> alkoller,
        decimal alkolIcinMelas,
        DateTime baslangic,
        DateTime bitis)
    {
        using var workbook = new XLWorkbook();
        
        // Yan Ürünler Sayfası
        var wsYanUrun = workbook.Worksheets.Add("Yan Ürünler");
        CreateYanUrunlerSheet(wsYanUrun, yanUrunler, baslangic, bitis);
        
        // Alkol Sayfası
        var wsAlkol = workbook.Worksheets.Add("Etil Alkol");
        CreateAlkolSheet(wsAlkol, alkoller, alkolIcinMelas, baslangic, bitis);

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    private void CreateYanUrunlerSheet(IXLWorksheet ws, List<YanUrunOzet> veriler, DateTime baslangic, DateTime bitis)
    {
        // Başlık
        ws.Cell("A1").Value = "YAN ÜRÜNLER SATIŞ RAPORU";
        ws.Range("A1:J1").Merge();
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 16;
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Cell("A2").Value = $"Tarih Aralığı: {baslangic:dd.MM.yyyy} - {bitis:dd.MM.yyyy}";
        ws.Range("A2:J2").Merge();
        ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Cell("A3").Value = "Birim: TON";
        ws.Range("A3:J3").Merge();
        ws.Cell("A3").Style.Font.Italic = true;

        // Başlık satırı
        int row = 5;
        string[] basliklar = { "Malzeme", "Devir Stok", "Üretim", "Satın Alma", "Satış", "İade", "Tüketim", "STOK", "Satış Tutarı (TL)", "Ort. Fiyat" };
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(row, i + 1).Value = basliklar[i];
            ws.Cell(row, i + 1).Style.Font.Bold = true;
            ws.Cell(row, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
            ws.Cell(row, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        // STOK sütunu vurgulu
        ws.Cell(row, 8).Style.Fill.BackgroundColor = XLColor.FromHtml("#C8E6C9");

        row++;

        // Kategorilere göre grupla
        var kategoriler = new[] { "MELAS", "YAS_KUSPE", "KURU_KUSPE", "DIGER" };
        
        foreach (var kategori in kategoriler)
        {
            var kategoridekiler = veriler.Where(x => x.Kategori == kategori).ToList();
            if (!kategoridekiler.Any()) continue;

            foreach (var urun in kategoridekiler)
            {
                ws.Cell(row, 1).Value = urun.MalzemeAdi;
                ws.Cell(row, 2).Value = urun.DevirStokTon;
                ws.Cell(row, 3).Value = urun.UretimTon;
                ws.Cell(row, 4).Value = urun.SatinAlmaTon;
                ws.Cell(row, 5).Value = urun.SatisTon;
                ws.Cell(row, 6).Value = urun.IadeTon;
                ws.Cell(row, 7).Value = urun.TuketimTon;
                ws.Cell(row, 8).Value = urun.StokTon;
                ws.Cell(row, 9).Value = urun.SatisTutari;
                ws.Cell(row, 10).Value = urun.OrtalamaFiyat;
                
                // STOK sütunu vurgulu
                ws.Cell(row, 8).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
                ws.Cell(row, 8).Style.Font.Bold = true;
                row++;
            }

            // Kategori toplam satırı
            if (kategori == "YAS_KUSPE" || kategori == "KURU_KUSPE")
            {
                var toplam = new YanUrunOzet
                {
                    DevirStok = kategoridekiler.Sum(x => x.DevirStok),
                    SatinAlmaMiktari = kategoridekiler.Sum(x => x.SatinAlmaMiktari),
                    UretimMiktari = kategoridekiler.Sum(x => x.UretimMiktari),
                    SatisMiktari = kategoridekiler.Sum(x => x.SatisMiktari),
                    SatisTutari = kategoridekiler.Sum(x => x.SatisTutari),
                    IadeMiktari = kategoridekiler.Sum(x => x.IadeMiktari),
                    IadeTutari = kategoridekiler.Sum(x => x.IadeTutari),
                    TuketimMiktari = kategoridekiler.Sum(x => x.TuketimMiktari)
                };

                string toplamAdi = kategori switch
                {
                    "YAS_KUSPE" => "YAŞ KÜSPE TOPLAM",
                    "KURU_KUSPE" => "KURU KÜSPE TOPLAM",
                    _ => $"{kategori} TOPLAM"
                };

                ws.Cell(row, 1).Value = toplamAdi;
                ws.Cell(row, 2).Value = toplam.DevirStokTon;
                ws.Cell(row, 3).Value = toplam.UretimTon;
                ws.Cell(row, 4).Value = toplam.SatinAlmaTon;
                ws.Cell(row, 5).Value = toplam.SatisTon;
                ws.Cell(row, 6).Value = toplam.IadeTon;
                ws.Cell(row, 7).Value = toplam.TuketimTon;
                ws.Cell(row, 8).Value = toplam.StokTon;
                ws.Cell(row, 9).Value = toplam.SatisTutari;
                ws.Cell(row, 10).Value = toplam.OrtalamaFiyat;

                ws.Range(row, 1, row, 10).Style.Font.Bold = true;
                var bgColor = kategori == "YAS_KUSPE" 
                    ? XLColor.FromHtml("#90EE90") 
                    : XLColor.FromHtml("#FFDAB9");
                ws.Range(row, 1, row, 10).Style.Fill.BackgroundColor = bgColor;
                row++;
            }
        }

        // Tablo kenarlıkları
        var dataRange = ws.Range(5, 1, row - 1, 10);
        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.OutsideBorderColor = XLColor.Black;
        dataRange.Style.Border.InsideBorderColor = XLColor.Gray;

        // Sayı formatları
        ws.Range(6, 2, row, 10).Style.NumberFormat.Format = "#,##0.00";
        ws.Range(6, 2, row, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Columns().AdjustToContents();
    }

    private void CreateAlkolSheet(IXLWorksheet ws, List<AlkolOzet> veriler, decimal tuketilenMelas, DateTime baslangic, DateTime bitis)
    {
        // Başlık
        ws.Cell("A1").Value = "ETİL ALKOL SATIŞ RAPORU";
        ws.Range("A1:J1").Merge();
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 16;
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Cell("A2").Value = $"Tarih Aralığı: {baslangic:dd.MM.yyyy} - {bitis:dd.MM.yyyy}";
        ws.Range("A2:J2").Merge();
        ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Cell("A3").Value = "Birim: TON (Lt/1000)";
        ws.Range("A3:J3").Merge();
        ws.Cell("A3").Style.Font.Italic = true;

        // Başlık satırı
        int row = 5;
        string[] basliklar = { "Alkol Türü", "Devir Stok", "Üretim", "Satın Alma", "Satış", "İade", "STOK", "Satış Tutarı (TL)", "Ort. Fiyat" };
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(row, i + 1).Value = basliklar[i];
            ws.Cell(row, i + 1).Style.Font.Bold = true;
            ws.Cell(row, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");
            ws.Cell(row, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        // STOK sütunu vurgulu
        ws.Cell(row, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#BBDEFB");

        row++;
        
        foreach (var alkol in veriler)
        {
            ws.Cell(row, 1).Value = alkol.MalzemeAdi;
            ws.Cell(row, 2).Value = alkol.DevirStokTon;
            ws.Cell(row, 3).Value = alkol.UretimTon;
            ws.Cell(row, 4).Value = alkol.SatinAlmaTon;
            ws.Cell(row, 5).Value = alkol.SatisTon;
            ws.Cell(row, 6).Value = alkol.IadeTon;
            ws.Cell(row, 7).Value = alkol.StokTon;
            ws.Cell(row, 8).Value = alkol.SatisTutari;
            ws.Cell(row, 9).Value = alkol.OrtalamaFiyat;

            // STOK sütunu vurgulu
            ws.Cell(row, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");
            ws.Cell(row, 7).Style.Font.Bold = true;
            row++;
        }

        // Alkol Toplam
        ws.Cell(row, 1).Value = "ALKOL TOPLAMI";
        ws.Cell(row, 2).Value = veriler.Sum(x => x.DevirStokTon);
        ws.Cell(row, 3).Value = veriler.Sum(x => x.UretimTon);
        ws.Cell(row, 4).Value = veriler.Sum(x => x.SatinAlmaTon);
        ws.Cell(row, 5).Value = veriler.Sum(x => x.SatisTon);
        ws.Cell(row, 6).Value = veriler.Sum(x => x.IadeTon);
        ws.Cell(row, 7).Value = veriler.Sum(x => x.StokTon);
        ws.Cell(row, 8).Value = veriler.Sum(x => x.SatisTutari);
        var toplamSatis = veriler.Sum(x => x.SatisMiktari);
        ws.Cell(row, 9).Value = toplamSatis > 0 ? veriler.Sum(x => x.SatisTutari) / toplamSatis : 0;
        ws.Range(row, 1, row, 9).Style.Font.Bold = true;
        ws.Range(row, 1, row, 9).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFB6C1");

        int lastDataRow = row;
        row += 2;

        // Tüketilen Melas
        ws.Cell(row, 1).Value = "Alkol Üretimi için Tüketilen Melas (Ton)";
        ws.Cell(row, 3).Value = tuketilenMelas / 1000; // TON'a çevir
        ws.Range(row, 1, row, 9).Style.Font.Bold = true;
        ws.Range(row, 1, row, 9).Style.Fill.BackgroundColor = XLColor.FromHtml("#ADD8E6");

        // Tablo kenarlıkları
        var dataRange = ws.Range(5, 1, lastDataRow, 9);
        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.OutsideBorderColor = XLColor.Black;
        dataRange.Style.Border.InsideBorderColor = XLColor.Gray;

        // Sayı formatları
        ws.Range(6, 2, row, 9).Style.NumberFormat.Format = "#,##0.00";
        ws.Range(6, 2, row, 7).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Columns().AdjustToContents();
    }

    /// <summary>
    /// Detay hareketlerini Excel'e aktarır
    /// </summary>
    public byte[] ExportStokHareketleri(List<StokHareket> hareketler, string malzemeAdi, DateTime baslangic, DateTime bitis)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Stok Hareketleri");

        // Başlık
        ws.Cell("A1").Value = $"{malzemeAdi} - Stok Hareketleri";
        ws.Range("A1:J1").Merge();
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 14;

        ws.Cell("A2").Value = $"Tarih Aralığı: {baslangic:dd.MM.yyyy} - {bitis:dd.MM.yyyy}";
        ws.Range("A2:J2").Merge();

        // Başlık satırı
        int row = 4;
        string[] basliklar = { "Tarih", "Fiş Türü", "Fiş No", "Cari Kodu", "Cari Adı", "Malzeme", "Giriş Miktar", "Giriş Tutar", "Çıkış Miktar", "Çıkış Tutar" };
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(row, i + 1).Value = basliklar[i];
            ws.Cell(row, i + 1).Style.Font.Bold = true;
            ws.Cell(row, i + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
        }

        row++;
        foreach (var hareket in hareketler)
        {
            ws.Cell(row, 1).Value = hareket.Tarih;
            ws.Cell(row, 1).Style.DateFormat.Format = "dd.MM.yyyy";
            ws.Cell(row, 2).Value = hareket.FisTuru;
            ws.Cell(row, 3).Value = hareket.FisNo;
            ws.Cell(row, 4).Value = hareket.CariKodu;
            ws.Cell(row, 5).Value = hareket.CariAdi;
            ws.Cell(row, 6).Value = hareket.MalzemeAdi;
            ws.Cell(row, 7).Value = hareket.GirisMiktari;
            ws.Cell(row, 8).Value = hareket.GirisTutari;
            ws.Cell(row, 9).Value = hareket.CikisMiktari;
            ws.Cell(row, 10).Value = hareket.CikisTutari;
            row++;
        }

        // Tablo kenarlıkları
        var dataRange = ws.Range(4, 1, row - 1, 10);
        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        // Format
        ws.Range(5, 7, row, 10).Style.NumberFormat.Format = "#,##0.00";
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Şeker Satış raporunu Excel'e aktarır
    /// VBA formatında 10 sütunlu tablo
    /// </summary>
    public byte[] ExportSekerSatisRaporu(List<SekerSatisOzet> sekerler, DateTime baslangic, DateTime bitis)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Şeker Satış");

        // Başlık
        ws.Cell("A1").Value = "ŞEKER ÜRETİM - SATIŞ - STOK TABLOSU";
        ws.Range("A1:J1").Merge();
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 16;
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromHtml("#059669");
        ws.Cell("A1").Style.Font.FontColor = XLColor.White;

        ws.Cell("A2").Value = $"Tarih Aralığı: {baslangic:dd.MM.yyyy} - {bitis:dd.MM.yyyy}";
        ws.Range("A2:J2").Merge();
        ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Cell("A3").Value = "Birim: TON";
        ws.Range("A3:J3").Merge();
        ws.Cell("A3").Style.Font.Italic = true;

        // Başlık satırı - VBA formatında 10 sütun
        int row = 5;
        string[] basliklar = { "KATEGORİ", "Devir (Ton)", "Üretim (Ton)", "Satın Alma (Ton)", 
                               "Satıştan İade (Ton)", "Satınalma İade (Ton)", "Satış (Ton)", 
                               "Promosyon (Ton)", "Sarf (Ton)", "Stok (Ton)" };
        
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(row, i + 1).Value = basliklar[i];
            ws.Cell(row, i + 1).Style.Font.Bold = true;
            ws.Cell(row, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#4a5568");
            ws.Cell(row, i + 1).Style.Font.FontColor = XLColor.White;
            ws.Cell(row, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        // STOK sütunu vurgulu
        ws.Cell(row, 10).Style.Fill.BackgroundColor = XLColor.FromHtml("#059669");

        row++;

        foreach (var seker in sekerler)
        {
            var bgColor = (row - 6) % 2 == 0 ? "#ffffff" : "#f8fafc";
            
            ws.Cell(row, 1).Value = seker.KategoriAdi;
            ws.Cell(row, 1).Style.Font.Bold = true;
            ws.Cell(row, 2).Value = seker.DevirStokTon;
            ws.Cell(row, 3).Value = seker.UretimTon;
            ws.Cell(row, 4).Value = seker.SatinAlmaTon;
            ws.Cell(row, 5).Value = seker.IadeTon;
            ws.Cell(row, 6).Value = seker.SatinAlmaIadeTon;
            ws.Cell(row, 7).Value = seker.SatisTon;
            ws.Cell(row, 8).Value = seker.PromosyonTon;
            ws.Cell(row, 9).Value = seker.SarfTon;
            ws.Cell(row, 10).Value = seker.StokTon;

            // Arka plan rengi
            ws.Range(row, 1, row, 10).Style.Fill.BackgroundColor = XLColor.FromHtml(bgColor);
            
            // Negatif stok kırmızı
            if (seker.StokTon < 0)
            {
                ws.Cell(row, 10).Style.Font.FontColor = XLColor.FromHtml("#dc2626");
                ws.Cell(row, 10).Style.Font.Bold = true;
            }
            
            row++;
        }

        // TOPLAM Satırı
        ws.Cell(row, 1).Value = "TOPLAM";
        ws.Cell(row, 2).Value = sekerler.Sum(x => x.DevirStokTon);
        ws.Cell(row, 3).Value = sekerler.Sum(x => x.UretimTon);
        ws.Cell(row, 4).Value = sekerler.Sum(x => x.SatinAlmaTon);
        ws.Cell(row, 5).Value = sekerler.Sum(x => x.IadeTon);
        ws.Cell(row, 6).Value = sekerler.Sum(x => x.SatinAlmaIadeTon);
        ws.Cell(row, 7).Value = sekerler.Sum(x => x.SatisTon);
        ws.Cell(row, 8).Value = sekerler.Sum(x => x.PromosyonTon);
        ws.Cell(row, 9).Value = sekerler.Sum(x => x.SarfTon);
        ws.Cell(row, 10).Value = sekerler.Sum(x => x.StokTon);
        
        ws.Range(row, 1, row, 10).Style.Font.Bold = true;
        ws.Range(row, 1, row, 10).Style.Fill.BackgroundColor = XLColor.FromHtml("#fef08a");
        ws.Range(row, 1, row, 10).Style.Font.FontColor = XLColor.FromHtml("#059669");

        // Tablo kenarlıkları
        var dataRange = ws.Range(5, 1, row, 10);
        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        dataRange.Style.Border.OutsideBorderColor = XLColor.Black;
        dataRange.Style.Border.InsideBorderColor = XLColor.Gray;

        // Sayı formatları
        ws.Range(6, 2, row, 10).Style.NumberFormat.Format = "#,##0.00";
        ws.Range(6, 2, row, 10).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    public byte[] ExportKasaHareketleri(List<KasaHareketi> veriler, DateTime? baslangic, DateTime? bitis)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Kasa İşlemleri");

        string[] basliklar = {
            "TARIH","İŞLEM_NO","BELGE_NO","CARI_UNVANI","CARI_KODU","SATIR_ACIKLAMASI",
            "OZEL_KODU","TICARI_ISLEM_GRUBU","FIS_TURU","IPTAL","MUHASEBELESTIRME","TUTAR",
            "IS_YERI","BOLUM","DOVIZ_TURU","KUR","ISLEM_DOVIZI_TUTARI",
            "RAPORLAMA_DOVIZI_TUTARI","RAPORLAMA_DOVIZI_KURU"
        };
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(1, i + 1).Value = basliklar[i];
            ws.Cell(1, i + 1).Style.Font.Bold = true;
            ws.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
            ws.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        int row = 2;
        foreach (var k in veriler)
        {
            ws.Cell(row, 1).Value = k.Tarih;
            ws.Cell(row, 1).Style.DateFormat.Format = "dd.MM.yyyy";
            ws.Cell(row, 2).Value = k.IslemNo;
            ws.Cell(row, 3).Value = k.BelgeNo;
            ws.Cell(row, 4).Value = k.CariUnvani;
            ws.Cell(row, 5).Value = k.CariKodu;
            ws.Cell(row, 6).Value = k.SatirAciklamasi;
            ws.Cell(row, 7).Value = k.OzelKodu;
            ws.Cell(row, 8).Value = k.TicariIslemGrubu;
            ws.Cell(row, 9).Value = k.FisTuru;
            ws.Cell(row, 10).Value = k.Iptal ? "EVET" : "HAYIR";
            ws.Cell(row, 11).Value = k.Muhasebelesti ? "EVET" : "HAYIR";
            ws.Cell(row, 12).Value = k.Tutar;
            ws.Cell(row, 12).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 13).Value = k.IsYeri;
            ws.Cell(row, 14).Value = k.Bolum;
            ws.Cell(row, 15).Value = k.DovizTuru;
            ws.Cell(row, 16).Value = k.Kur;
            ws.Cell(row, 16).Style.NumberFormat.Format = "#,##0.0000";
            ws.Cell(row, 17).Value = k.IslemDoviziTutari;
            ws.Cell(row, 17).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 18).Value = k.RaporlamaDoviziTutari;
            ws.Cell(row, 18).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 19).Value = k.RaporlamaDoviziKuru;
            ws.Cell(row, 19).Style.NumberFormat.Format = "#,##0.0000";
            row++;
        }

        ws.SheetView.FreezeRows(1);
        ws.RangeUsed()?.SetAutoFilter();
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    public byte[] ExportCariBakiye(List<CariBakiye> veriler, DateTime baslangic, DateTime bitis)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Cari Bakiye");

        ws.Cell(1, 1).Value = $"Cari Hesap Bakiyesi — {baslangic:dd.MM.yyyy} / {bitis:dd.MM.yyyy}";
        ws.Range(1, 1, 1, 12).Merge().Style.Font.Bold = true;

        string[] basliklar = {
            "CARI_HESAP_KODU","CARI_HESAP_UNVANI","TC_KIMLIK_NO","VERGI_NO",
            "BORC","ALACAK","BAKIYE","DURUM",
            "OZEL_KOD","OZEL_KOD2","OZEL_KOD3","OZEL_KOD4","OZEL_KOD5"
        };
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(3, i + 1).Value = basliklar[i];
            ws.Cell(3, i + 1).Style.Font.Bold = true;
            ws.Cell(3, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
            ws.Cell(3, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        int row = 4;
        foreach (var c in veriler)
        {
            ws.Cell(row, 1).Value = c.CariHesapKodu;
            ws.Cell(row, 2).Value = c.CariHesapUnvani;
            ws.Cell(row, 3).Value = c.TcKimlikNo;
            ws.Cell(row, 4).Value = c.VergiNo;
            ws.Cell(row, 5).Value = c.Borc;
            ws.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 6).Value = c.Alacak;
            ws.Cell(row, 6).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 7).Value = c.Bakiye;
            ws.Cell(row, 7).Style.NumberFormat.Format = "#,##0.00";
            string durum = c.Bakiye < 0 ? "FIRMA ALACAKLI" : (c.Bakiye > 0 ? "FIRMA BORÇLU" : "EŞİT");
            ws.Cell(row, 8).Value = durum;
            if (c.Bakiye < 0)
            {
                ws.Cell(row, 7).Style.Font.FontColor = XLColor.FromHtml("#2E7D32");
                ws.Cell(row, 8).Style.Font.FontColor = XLColor.FromHtml("#2E7D32");
            }
            else if (c.Bakiye > 0)
            {
                ws.Cell(row, 7).Style.Font.FontColor = XLColor.Red;
                ws.Cell(row, 8).Style.Font.FontColor = XLColor.Red;
            }
            ws.Cell(row, 9).Value = c.OzelKod;
            ws.Cell(row, 10).Value = c.OzelKod2;
            ws.Cell(row, 11).Value = c.OzelKod3;
            ws.Cell(row, 12).Value = c.OzelKod4;
            ws.Cell(row, 13).Value = c.OzelKod5;
            row++;
        }

        ws.SheetView.FreezeRows(3);
        ws.Range(3, 1, row - 1, 13).SetAutoFilter();
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    public byte[] ExportCariHareket(List<CariHareket> veriler, DateTime baslangic, DateTime bitis, string cariEtiket)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Cari Hareket");

        const int colCount = 14;
        ws.Cell(1, 1).Value = $"Cari Hesap Hareketleri — {baslangic:dd.MM.yyyy} / {bitis:dd.MM.yyyy}";
        ws.Range(1, 1, 1, colCount).Merge().Style.Font.Bold = true;
        if (!string.IsNullOrWhiteSpace(cariEtiket))
        {
            ws.Cell(2, 1).Value = cariEtiket;
            ws.Range(2, 1, 2, colCount).Merge().Style.Font.Italic = true;
        }

        string[] basliklar = {
            "TARIH","ODEMEPLANI","FISTURU","FISNO","ACIKLAMA",
            "CARIKODU","CARIUNVAN","BORC","ALACAK","BAKIYE",
            "DOVIZ","KUR","DOVIZ_BORC","DOVIZ_ALACAK"
        };
        int hdr = 4;
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(hdr, i + 1).Value = basliklar[i];
            ws.Cell(hdr, i + 1).Style.Font.Bold = true;
            ws.Cell(hdr, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
            ws.Cell(hdr, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        int row = hdr + 1;
        foreach (var c in veriler)
        {
            ws.Cell(row, 1).Value = c.Tarih;
            ws.Cell(row, 1).Style.DateFormat.Format = "dd.MM.yyyy";
            ws.Cell(row, 2).Value = c.OdemePlani;
            ws.Cell(row, 3).Value = c.FisTuru;
            ws.Cell(row, 4).Value = c.FisNo;
            ws.Cell(row, 5).Value = c.Aciklama;
            ws.Cell(row, 6).Value = c.CariKodu;
            ws.Cell(row, 7).Value = c.CariUnvan;
            ws.Cell(row, 8).Value = c.Borc;
            ws.Cell(row, 8).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 9).Value = c.Alacak;
            ws.Cell(row, 9).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 10).Value = c.Bakiye;
            ws.Cell(row, 10).Style.NumberFormat.Format = "#,##0.00";
            if (c.Bakiye < 0)
                ws.Cell(row, 10).Style.Font.FontColor = XLColor.FromHtml("#2E7D32");
            else if (c.Bakiye > 0)
                ws.Cell(row, 10).Style.Font.FontColor = XLColor.Red;

            ws.Cell(row, 11).Value = c.DovizKodu;
            ws.Cell(row, 11).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            if (c.TrCurr != 0 && c.DovizKur != 0)
            {
                ws.Cell(row, 12).Value = c.DovizKur;
                ws.Cell(row, 12).Style.NumberFormat.Format = "#,##0.0000";
            }
            if (c.DovizBorc != 0)
            {
                ws.Cell(row, 13).Value = c.DovizBorc;
                ws.Cell(row, 13).Style.NumberFormat.Format = "#,##0.00";
                ws.Cell(row, 13).Style.Font.FontColor = XLColor.FromHtml("#2E7D32");
            }
            if (c.DovizAlacak != 0)
            {
                ws.Cell(row, 14).Value = c.DovizAlacak;
                ws.Cell(row, 14).Style.NumberFormat.Format = "#,##0.00";
                ws.Cell(row, 14).Style.Font.FontColor = XLColor.Red;
            }
            row++;
        }

        ws.SheetView.FreezeRows(hdr);
        if (row > hdr + 1)
            ws.Range(hdr, 1, row - 1, colCount).SetAutoFilter();
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    public byte[] ExportStokDurumu(List<StokSatiri> veriler)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Stok Durumu");

        string[] basliklar = { "AMBAR_NO", "AMBAR", "MALZEME_KODU", "MALZEME_ADI", "STOK" };
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(1, i + 1).Value = basliklar[i];
            ws.Cell(1, i + 1).Style.Font.Bold = true;
            ws.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
            ws.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        int row = 2;
        foreach (var s in veriler)
        {
            ws.Cell(row, 1).Value = s.AmbarNo;
            ws.Cell(row, 2).Value = s.Ambar;
            ws.Cell(row, 3).Value = s.MalzemeKodu;
            ws.Cell(row, 4).Value = s.MalzemeAdi;
            ws.Cell(row, 5).Value = s.Stok;
            ws.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00";
            row++;
        }

        ws.SheetView.FreezeRows(1);
        ws.RangeUsed()?.SetAutoFilter();
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    public byte[] ExportGubreStok(List<StokSatiri> veriler)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Gubre Stok");

        string[] basliklar = { "AMBAR_NO", "AMBAR", "MALZEME_KODU", "MALZEME_ADI", "STOK" };
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(1, i + 1).Value = basliklar[i];
            ws.Cell(1, i + 1).Style.Font.Bold = true;
            ws.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
            ws.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        int row = 2;
        foreach (var s in veriler)
        {
            ws.Cell(row, 1).Value = s.AmbarNo;
            ws.Cell(row, 2).Value = s.Ambar;
            ws.Cell(row, 3).Value = s.MalzemeKodu;
            ws.Cell(row, 4).Value = s.MalzemeAdi;
            ws.Cell(row, 5).Value = s.Stok;
            ws.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00";
            row++;
        }

        // Veri sonuna toplam satırı — AutoFilter aralığından önce belirlenmesi için row değişkeni
        int veriSonRow = row - 1;
        ws.Cell(row, 4).Value = "TOPLAM";
        ws.Cell(row, 4).Style.Font.Bold = true;
        ws.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Cell(row, 4).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
        ws.Cell(row, 5).Value = veriler.Sum(s => s.Stok);
        ws.Cell(row, 5).Style.Font.Bold = true;
        ws.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00";
        ws.Cell(row, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");

        ws.SheetView.FreezeRows(1);
        // AutoFilter sadece veri aralığında — toplam satırı filtreye dahil edilmez
        if (veriSonRow >= 1)
            ws.Range(1, 1, veriSonRow, basliklar.Length).SetAutoFilter();
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    public byte[] ExportCayDurumu(List<StokSatiri> veriler)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Cay Durum");

        string[] basliklar = { "AMBAR_NO", "AMBAR", "MALZEME_KODU", "MALZEME_ADI", "STOK" };
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(1, i + 1).Value = basliklar[i];
            ws.Cell(1, i + 1).Style.Font.Bold = true;
            ws.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#EFEBE9");
            ws.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        int row = 2;
        foreach (var s in veriler)
        {
            ws.Cell(row, 1).Value = s.AmbarNo;
            ws.Cell(row, 2).Value = s.Ambar;
            ws.Cell(row, 3).Value = s.MalzemeKodu;
            ws.Cell(row, 4).Value = s.MalzemeAdi;
            ws.Cell(row, 5).Value = s.Stok;
            ws.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00";
            row++;
        }

        int veriSonRow = row - 1;
        ws.Cell(row, 4).Value = "TOPLAM";
        ws.Cell(row, 4).Style.Font.Bold = true;
        ws.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Cell(row, 4).Style.Fill.BackgroundColor = XLColor.FromHtml("#EFEBE9");
        ws.Cell(row, 5).Value = veriler.Sum(s => s.Stok);
        ws.Cell(row, 5).Style.Font.Bold = true;
        ws.Cell(row, 5).Style.NumberFormat.Format = "#,##0.00";
        ws.Cell(row, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#EFEBE9");

        ws.SheetView.FreezeRows(1);
        if (veriSonRow >= 1)
            ws.Range(1, 1, veriSonRow, basliklar.Length).SetAutoFilter();
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    public byte[] ExportGubreCiro(
        List<GubreCiroMalzeme> malzeme,
        List<GubreCiroMusteriAy> musteriAy,
        List<GubreAlimFaturaSatiri> alim)
    {
        using var workbook = new XLWorkbook();

        // ---- Sayfa 1: Malzeme Bazlı Ciro ----
        var ws1 = workbook.Worksheets.Add("Malzeme Ciro");
        string[] b1 = { "MALZEME_KODU", "MALZEME_ADI", "NET_MIKTAR", "NET_CIRO (TL)" };
        for (int i = 0; i < b1.Length; i++)
        {
            ws1.Cell(1, i + 1).Value = b1[i];
            ws1.Cell(1, i + 1).Style.Font.Bold = true;
            ws1.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
            ws1.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        int r = 2;
        foreach (var m in malzeme)
        {
            ws1.Cell(r, 1).Value = m.MalzemeKodu;
            ws1.Cell(r, 2).Value = m.MalzemeAdi;
            ws1.Cell(r, 3).Value = m.NetMiktar;
            ws1.Cell(r, 3).Style.NumberFormat.Format = "#,##0.00";
            ws1.Cell(r, 4).Value = m.NetCiro;
            ws1.Cell(r, 4).Style.NumberFormat.Format = "#,##0.00";
            r++;
        }
        int s1 = r - 1;
        ws1.Cell(r, 2).Value = "TOPLAM";
        ws1.Cell(r, 2).Style.Font.Bold = true;
        ws1.Cell(r, 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws1.Cell(r, 2).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
        ws1.Cell(r, 3).Value = malzeme.Sum(x => x.NetMiktar);
        ws1.Cell(r, 4).Value = malzeme.Sum(x => x.NetCiro);
        ws1.Range(r, 3, r, 4).Style.NumberFormat.Format = "#,##0.00";
        ws1.Range(r, 1, r, 4).Style.Font.Bold = true;
        ws1.Range(r, 1, r, 4).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
        ws1.SheetView.FreezeRows(1);
        if (s1 >= 1) ws1.Range(1, 1, s1, b1.Length).SetAutoFilter();
        ws1.Columns().AdjustToContents();

        // ---- Sayfa 2: Müşteri + Ay Kırılımı ----
        var ws2 = workbook.Worksheets.Add("Musteri-Ay Ciro");
        string[] b2 = { "MALZEME_KODU", "MALZEME_ADI", "MUSTERI_KODU", "MUSTERI_UNVAN", "DONEM (YYYY-AA)", "NET_MIKTAR", "NET_CIRO (TL)" };
        for (int i = 0; i < b2.Length; i++)
        {
            ws2.Cell(1, i + 1).Value = b2[i];
            ws2.Cell(1, i + 1).Style.Font.Bold = true;
            ws2.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
            ws2.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        r = 2;
        foreach (var m in musteriAy)
        {
            ws2.Cell(r, 1).Value = m.MalzemeKodu;
            ws2.Cell(r, 2).Value = m.MalzemeAdi;
            ws2.Cell(r, 3).Value = m.MusteriKodu;
            ws2.Cell(r, 4).Value = m.MusteriUnvan;
            ws2.Cell(r, 5).Value = m.Donem;
            ws2.Cell(r, 6).Value = m.NetMiktar;
            ws2.Cell(r, 6).Style.NumberFormat.Format = "#,##0.00";
            ws2.Cell(r, 7).Value = m.NetCiro;
            ws2.Cell(r, 7).Style.NumberFormat.Format = "#,##0.00";
            r++;
        }
        int s2 = r - 1;
        ws2.Cell(r, 5).Value = "TOPLAM";
        ws2.Cell(r, 5).Style.Font.Bold = true;
        ws2.Cell(r, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws2.Cell(r, 6).Value = musteriAy.Sum(x => x.NetMiktar);
        ws2.Cell(r, 7).Value = musteriAy.Sum(x => x.NetCiro);
        ws2.Range(r, 6, r, 7).Style.NumberFormat.Format = "#,##0.00";
        ws2.Range(r, 1, r, 7).Style.Font.Bold = true;
        ws2.Range(r, 1, r, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
        ws2.SheetView.FreezeRows(1);
        if (s2 >= 1) ws2.Range(1, 1, s2, b2.Length).SetAutoFilter();
        ws2.Columns().AdjustToContents();

        // ---- Sayfa 3: Alım Faturaları ----
        var ws3 = workbook.Worksheets.Add("Alim Faturalari");
        string[] b3 = { "FATURA_NO", "TARIH", "TEDARIKCI_KOD", "TEDARIKCI_UNVAN", "MALZEME_KODU", "MALZEME_ADI", "MIKTAR", "NET_TUTAR (TL)" };
        for (int i = 0; i < b3.Length; i++)
        {
            ws3.Cell(1, i + 1).Value = b3[i];
            ws3.Cell(1, i + 1).Style.Font.Bold = true;
            ws3.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
            ws3.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
        r = 2;
        foreach (var a in alim)
        {
            ws3.Cell(r, 1).Value = a.FaturaNo;
            ws3.Cell(r, 2).Value = a.Tarih;
            ws3.Cell(r, 2).Style.DateFormat.Format = "dd.MM.yyyy";
            ws3.Cell(r, 3).Value = a.TedarikciKodu;
            ws3.Cell(r, 4).Value = a.TedarikciUnvan;
            ws3.Cell(r, 5).Value = a.MalzemeKodu;
            ws3.Cell(r, 6).Value = a.MalzemeAdi;
            ws3.Cell(r, 7).Value = a.Miktar;
            ws3.Cell(r, 7).Style.NumberFormat.Format = "#,##0.00";
            ws3.Cell(r, 8).Value = a.NetTutar;
            ws3.Cell(r, 8).Style.NumberFormat.Format = "#,##0.00";
            r++;
        }
        int s3 = r - 1;
        ws3.Cell(r, 6).Value = "TOPLAM";
        ws3.Cell(r, 6).Style.Font.Bold = true;
        ws3.Cell(r, 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws3.Cell(r, 7).Value = alim.Sum(x => x.Miktar);
        ws3.Cell(r, 8).Value = alim.Sum(x => x.NetTutar);
        ws3.Range(r, 7, r, 8).Style.NumberFormat.Format = "#,##0.00";
        ws3.Range(r, 1, r, 8).Style.Font.Bold = true;
        ws3.Range(r, 1, r, 8).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
        ws3.SheetView.FreezeRows(1);
        if (s3 >= 1) ws3.Range(1, 1, s3, b3.Length).SetAutoFilter();
        ws3.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    public byte[] ExportKurBilgileri(List<KurBilgisi> veriler)
    {
        using var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add("Kurlar");

        string[] basliklar = { "EDATE", "DOVIZ_KODU", "DOVIZ_ADI", "TUR1", "TUR2", "TUR3", "TUR4" };
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(1, i + 1).Value = basliklar[i];
            ws.Cell(1, i + 1).Style.Font.Bold = true;
            ws.Cell(1, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");
            ws.Cell(1, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        int row = 2;
        foreach (var k in veriler)
        {
            ws.Cell(row, 1).Value = k.Tarih;
            ws.Cell(row, 1).Style.DateFormat.Format = "dd.MM.yyyy";
            ws.Cell(row, 2).Value = k.DovizKodu;
            ws.Cell(row, 3).Value = k.DovizAdi;
            ws.Cell(row, 4).Value = k.Rate1;
            ws.Cell(row, 5).Value = k.Rate2;
            ws.Cell(row, 6).Value = k.Rate3;
            ws.Cell(row, 7).Value = k.Rate4;
            ws.Range(row, 4, row, 7).Style.NumberFormat.Format = "#,##0.0000";
            row++;
        }

        ws.SheetView.FreezeRows(1);
        ws.RangeUsed()?.SetAutoFilter();
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Tarım Kredi Raporu: her bölge ayrı sayfa, firma alt toplamları,
    /// iade satırları kırmızı/negatif. En sonda "Rapor Özeti" sayfası.
    /// </summary>
    public byte[] ExportTarimKrediRaporu(List<TarimKrediBolgeRapor> bolgeler, DateTime bas, DateTime bit)
    {
        using var wb = new XLWorkbook();

        // Özet sayfası (sonradan doldurulacak ama başa ekleyelim)
        var ozet = wb.Worksheets.Add("Rapor Özeti");

        foreach (var r in bolgeler)
        {
            var safeName = r.Bolge.Replace("/", "-").Replace("\\", "-");
            if (safeName.Length > 28) safeName = safeName[..28];
            var ws = wb.Worksheets.Add(safeName);
            TarimKrediBolgeSayfasiDoldur(ws, r, bas, bit);
        }

        // Özet sayfası doldur
        ozet.Cell("A1").Value = "TARIM KREDİ BÖLGE RAPORU ÖZETİ";
        ozet.Range("A1:E1").Merge();
        ozet.Cell("A1").Style.Font.Bold = true;
        ozet.Cell("A1").Style.Font.FontSize = 14;
        ozet.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ozet.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
        ozet.Cell("A1").Style.Font.FontColor = XLColor.White;

        ozet.Cell("A2").Value = $"Tarih Aralığı: {bas:dd.MM.yyyy} - {bit:dd.MM.yyyy}";
        ozet.Range("A2:E2").Merge();
        ozet.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        string[] bas2 = { "#", "Bölge", "Firma", "Hareket", "Net Miktar (KG)", "Net Tutar (TL)" };
        int r0 = 4;
        for (int i = 0; i < bas2.Length; i++)
        {
            ozet.Cell(r0, i + 1).Value = bas2[i];
            ozet.Cell(r0, i + 1).Style.Font.Bold = true;
            ozet.Cell(r0, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");
            ozet.Cell(r0, i + 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        int rr = r0 + 1;
        int sira = 1;
        foreach (var b in bolgeler)
        {
            ozet.Cell(rr, 1).Value = sira++;
            ozet.Cell(rr, 2).Value = b.Bolge;
            try
            {
                ozet.Cell(rr, 2).SetHyperlink(new XLHyperlink($"'{(b.Bolge.Length > 28 ? b.Bolge[..28] : b.Bolge)}'!A1"));
            }
            catch { }
            ozet.Cell(rr, 3).Value = b.FirmaSayisi;
            ozet.Cell(rr, 4).Value = b.HareketSayisi;
            ozet.Cell(rr, 5).Value = Math.Abs(b.ToplamMiktar);
            ozet.Cell(rr, 6).Value = Math.Abs(b.ToplamTutar);
            ozet.Range(rr, 5, rr, 6).Style.NumberFormat.Format = "#,##0.00";
            rr++;
        }

        // Genel toplam
        ozet.Cell(rr, 2).Value = "GENEL TOPLAM";
        ozet.Cell(rr, 3).Value = bolgeler.Sum(b => b.FirmaSayisi);
        ozet.Cell(rr, 4).Value = bolgeler.Sum(b => b.HareketSayisi);
        ozet.Cell(rr, 5).Value = Math.Abs(bolgeler.Sum(b => b.ToplamMiktar));
        ozet.Cell(rr, 6).Value = Math.Abs(bolgeler.Sum(b => b.ToplamTutar));
        ozet.Range(rr, 1, rr, 6).Style.Font.Bold = true;
        ozet.Range(rr, 1, rr, 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF59D");
        ozet.Range(rr, 5, rr, 6).Style.NumberFormat.Format = "#,##0.00";

        ozet.Columns().AdjustToContents();
        ozet.SheetView.FreezeRows(r0);
        YazdirmaAyariUygula(ozet, basliksatiri: r0);

        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Yatay yön + 1 sayfa genişliğine sığdırma + dar kenar boşlukları.
    /// Yazdırma anında kullanıcının el ayarı yapması gerekmez.
    /// </summary>
    private static void YazdirmaAyariUygula(IXLWorksheet ws, int? basliksatiri = null)
    {
        ws.PageSetup.PageOrientation = XLPageOrientation.Landscape;
        ws.PageSetup.PaperSize = XLPaperSize.A4Paper;
        ws.PageSetup.PagesWide = 1;     // Genişlikte 1 sayfaya sıkıştır
        ws.PageSetup.PagesTall = 0;     // Yükseklikte sınır yok (taşarsa alt sayfaya geçsin)
        ws.PageSetup.CenterHorizontally = true;
        ws.PageSetup.Margins.Top = 0.5;
        ws.PageSetup.Margins.Bottom = 0.5;
        ws.PageSetup.Margins.Left = 0.25;
        ws.PageSetup.Margins.Right = 0.25;
        ws.PageSetup.Margins.Header = 0.3;
        ws.PageSetup.Margins.Footer = 0.3;
        if (basliksatiri.HasValue)
        {
            // Çok sayfaya bölünürse her sayfanın üstünde başlık satırı tekrarlansın
            ws.PageSetup.SetRowsToRepeatAtTop(1, basliksatiri.Value);
        }
    }

    private void TarimKrediBolgeSayfasiDoldur(IXLWorksheet ws, TarimKrediBolgeRapor r, DateTime bas, DateTime bit)
    {
        // Başlık
        ws.Cell("A1").Value = $"TARIM KREDİ KOOPERATİFLERİ {r.Bolge} BÖLGESİ SATIŞ RAPORU";
        ws.Range("A1:H1").Merge();
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 13;
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Cell("A2").Value = $"Rapor Tarihi: {DateTime.Now:dd.MM.yyyy}";
        ws.Range("A2:B2").Merge();
        ws.Cell("A2").Style.Font.Bold = true;
        ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

        ws.Cell("A3").Value = $"Tarih Aralığı: {bas:dd.MM.yyyy} - {bit:dd.MM.yyyy}";
        ws.Range("A3:B3").Merge();
        ws.Cell("A3").Style.Font.Bold = true;
        ws.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

        ws.Cell("A4").Value = "Not: Kırmızı renkli satırlar 'Toptan Satış İade İrsaliyesi' içeren kayıtları gösterir.";
        ws.Range("A4:H4").Merge();
        ws.Cell("A4").Style.Font.Italic = true;
        ws.Cell("A4").Style.Font.FontSize = 10;
        ws.Cell("A4").Style.Alignment.WrapText = false;
        ws.Cell("A4").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

        // Başlık satırı
        string[] basliklar = { "Cari Hesap Kodu", "Tarih", "Belge Tipi", "Firma Adı", "Malzeme Adı", "Miktar", "Tutar", "Fatura Numarası" };
        int hr = 5;
        for (int i = 0; i < basliklar.Length; i++)
        {
            ws.Cell(hr, i + 1).Value = basliklar[i];
            ws.Cell(hr, i + 1).Style.Font.Bold = true;
            ws.Cell(hr, i + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#D9D9D9");
            ws.Cell(hr, i + 1).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Cell(hr, i + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        // Excel tema renkleri
        var altToplamRengi   = XLColor.FromHtml("#DDEBF7"); // Mavi, Vurgu 1, Daha Açık %80
        var genelToplamRengi = XLColor.FromHtml("#D9D9D9"); // Beyaz, Arka Plan 1, Daha Koyu %15
        var dataRengi        = XLColor.White;

        int row = hr + 1;
        int iadeSayisi = 0;
        foreach (var f in r.Firmalar)
        {
            foreach (var h in f.Hareketler)
            {
                ws.Cell(row, 1).Value = string.IsNullOrWhiteSpace(h.CariHesapKodu) ? f.CariHesapKodu : h.CariHesapKodu;
                ws.Cell(row, 2).Value = h.Tarih;
                ws.Cell(row, 2).Style.DateFormat.Format = "dd.MM.yyyy";
                ws.Cell(row, 3).Value = h.FisTuruGorunen;
                ws.Cell(row, 4).Value = f.CariHesapUnvani;
                ws.Cell(row, 5).Value = h.MalzemeAciklamasi;
                ws.Cell(row, 6).Value = h.MiktarGorunen;
                ws.Cell(row, 7).Value = h.TutarGorunen;
                ws.Cell(row, 8).Value = h.FaturaNo;
                ws.Range(row, 6, row, 7).Style.NumberFormat.Format = "#,##0.00";
                ws.Range(row, 1, row, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(row, 1, row, 8).Style.Fill.BackgroundColor = dataRengi;

                if (h.Iade)
                {
                    ws.Range(row, 1, row, 8).Style.Font.FontColor = XLColor.FromHtml("#C62828");
                    ws.Range(row, 1, row, 8).Style.Font.Bold = true;
                    iadeSayisi++;
                }
                row++;
            }

            // Firma alt toplam — italic + bold, Mavi Vurgu 1 Daha Açık %80
            ws.Cell(row, 1).Value = f.CariHesapKodu;
            ws.Cell(row, 4).Value = $"{f.CariHesapUnvani} - ALT TOPLAM";
            ws.Cell(row, 6).Value = Math.Abs(f.ToplamMiktar);
            ws.Cell(row, 7).Value = Math.Abs(f.ToplamTutar);
            ws.Range(row, 1, row, 8).Style.Font.Italic = true;
            ws.Range(row, 1, row, 8).Style.Font.Bold = true;
            ws.Range(row, 6, row, 7).Style.NumberFormat.Format = "#,##0.00";
            ws.Range(row, 1, row, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            ws.Range(row, 1, row, 8).Style.Fill.BackgroundColor = altToplamRengi;
            row++;
        }

        // Alt bilgi: İade Sayısı + GENEL TOPLAM — Beyaz Arka Plan 1 Daha Koyu %15
        ws.Cell(row, 1).Value = $"İade Sayısı: {iadeSayisi}";
        ws.Cell(row, 2).Value = "GENEL TOPLAM:";
        ws.Cell(row, 6).Value = Math.Abs(r.ToplamMiktar);
        ws.Cell(row, 7).Value = Math.Abs(r.ToplamTutar);
        ws.Range(row, 1, row, 8).Style.Font.Bold = true;
        ws.Range(row, 6, row, 7).Style.NumberFormat.Format = "#,##0.00";
        ws.Range(row, 1, row, 8).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.Range(row, 1, row, 8).Style.Fill.BackgroundColor = genelToplamRengi;

        // Tüm tabloya tam grid — her hücreye 4 kenar ayrı ayrı uygulanır (sütun araları net görünür)
        var tumTablo = ws.Range(hr, 1, row, 8);
        tumTablo.Style.Border.TopBorder = XLBorderStyleValues.Thin;
        tumTablo.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        tumTablo.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
        tumTablo.Style.Border.RightBorder = XLBorderStyleValues.Thin;
        tumTablo.Style.Border.TopBorderColor = XLColor.Black;
        tumTablo.Style.Border.BottomBorderColor = XLColor.Black;
        tumTablo.Style.Border.LeftBorderColor = XLColor.Black;
        tumTablo.Style.Border.RightBorderColor = XLColor.Black;

        ws.SheetView.FreezeRows(hr);
        ws.Columns().AdjustToContents();
        YazdirmaAyariUygula(ws, basliksatiri: hr);
    }

    /// <summary>
    /// Malzeme Hareket Listesi — VBA'daki "Yan Ürünler Satış Hareketleri" sayfasıyla
    /// aynı yerleşim: başlık/markalama ilk 18 satır, 19. satırda sütun başlıkları,
    /// 20. satırdan itibaren veri. Sütunlar: YIL | AY | TARIH | FIS_TURU | FIS_NUMARASI |
    /// CARI_HESAP_KODU | CARI_HESAP_UNVANI | MALZEME_KODU | MALZEME_ACIKLAMASI |
    /// GIRIS_MIKTARI | GIRIS_FIYATI | GIRIS_TUTARI | CIKIS_MIKTARI | CIKIS_FIYATI |
    /// CIKIS_TUTARI | Sütun1 | Sütun2 (=CIKIS_MIKTARI*CIKIS_FIYATI) | Sütun3.
    /// </summary>
    public byte[] ExportMalzemeHareketleri(
        List<MalzemeHareketSatiri> satirlar,
        List<string> malzemeKodlari,
        DateTime baslangic,
        DateTime bitis,
        string? listeAdi = null)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Malzeme Hareket Listesi");

        // --- 1-17: Marka/başlık bloğu ---
        ws.Cell(1, 1).Value = "DOĞUŞ ÇAY ve GIDA A.Ş. — AFYON ŞEKER FABRİKASI";
        ws.Range(1, 1, 1, 18).Merge();
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 14;
        ws.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#0D47A1");
        ws.Cell(1, 1).Style.Font.FontColor = XLColor.White;
        ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Cell(2, 1).Value = "MALZEME HAREKET LİSTESİ";
        ws.Range(2, 1, 2, 18).Merge();
        ws.Cell(2, 1).Style.Font.Bold = true;
        ws.Cell(2, 1).Style.Font.FontSize = 12;
        ws.Cell(2, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2");
        ws.Cell(2, 1).Style.Font.FontColor = XLColor.White;
        ws.Cell(2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ws.Cell(4, 1).Value = "Liste Adı:";
        ws.Cell(4, 2).Value = listeAdi ?? "(kaydedilmedi)";
        ws.Cell(5, 1).Value = "Tarih Aralığı:";
        ws.Cell(5, 2).Value = $"{baslangic:dd.MM.yyyy} — {bitis:dd.MM.yyyy}";
        ws.Cell(6, 1).Value = "Rapor Tarihi:";
        ws.Cell(6, 2).Value = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
        ws.Cell(7, 1).Value = "Toplam Satır:";
        ws.Cell(7, 2).Value = satirlar.Count;
        ws.Cell(4, 1).Style.Font.Bold = true;
        ws.Cell(5, 1).Style.Font.Bold = true;
        ws.Cell(6, 1).Style.Font.Bold = true;
        ws.Cell(7, 1).Style.Font.Bold = true;

        ws.Cell(9, 1).Value = "Malzeme Kodları:";
        ws.Cell(9, 1).Style.Font.Bold = true;
        ws.Cell(9, 2).Value = string.Join(", ", malzemeKodlari);
        ws.Range(9, 2, 9, 18).Merge();
        ws.Cell(9, 2).Style.Alignment.WrapText = true;

        // Notlar
        ws.Cell(12, 1).Value = "Açıklama:";
        ws.Cell(12, 1).Style.Font.Bold = true;
        ws.Cell(13, 1).Value = "• Giriş hareketleri: Satınalma İrsaliyesi, Toptan Satış İade İrsaliyesi, Üretimden Giriş Fişi.";
        ws.Cell(14, 1).Value = "• Çıkış hareketleri: Toptan Satış İrsaliyesi (eksi işaretli).";
        ws.Cell(15, 1).Value = "• Miktarlar ana birimdedir (AMOUNT × UINFO2 / UINFO1). Tutarlar VATMATRAH'tır.";
        ws.Range(13, 1, 15, 18).Style.Font.Italic = true;

        // --- 19: Sütun başlıkları (xlsm ile aynı sıra) ---
        int hr = 19;
        string[] basliklar = {
            "YIL","AY","TARIH","FIS_TURU","FIS_NUMARASI",
            "CARI_HESAP_KODU","CARI_HESAP_UNVANI","MALZEME_KODU","MALZEME_ACIKLAMASI",
            "GIRIS_MIKTARI","GIRIS_FIYATI","GIRIS_TUTARI",
            "CIKIS_MIKTARI","CIKIS_FIYATI","CIKIS_TUTARI",
            "Sütun1","Sütun2","Sütun3"
        };
        for (int i = 0; i < basliklar.Length; i++)
        {
            var c = ws.Cell(hr, i + 1);
            c.Value = basliklar[i];
            c.Style.Font.Bold = true;
            c.Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF9C4");
            c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            c.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        // --- 20+: Veri satırları ---
        int r = hr + 1;
        foreach (var s in satirlar)
        {
            ws.Cell(r,  1).Value = s.Yil;
            ws.Cell(r,  2).Value = s.Ay;
            ws.Cell(r,  3).Value = s.Tarih;
            ws.Cell(r,  3).Style.DateFormat.Format = "dd.MM.yyyy";
            ws.Cell(r,  4).Value = s.FisTuru;
            ws.Cell(r,  5).Value = s.FisNumarasi;
            ws.Cell(r,  6).Value = s.CariHesapKodu;
            ws.Cell(r,  7).Value = s.CariHesapUnvani;
            ws.Cell(r,  8).Value = s.MalzemeKodu;
            ws.Cell(r,  9).Value = s.MalzemeAciklamasi;
            ws.Cell(r, 10).Value = s.GirisMiktari;
            ws.Cell(r, 11).Value = s.GirisFiyati;
            ws.Cell(r, 12).Value = s.GirisTutari;
            ws.Cell(r, 13).Value = s.CikisMiktari;
            ws.Cell(r, 14).Value = s.CikisFiyati;
            ws.Cell(r, 15).Value = s.CikisTutari;
            // Sütun2: xlsm'deki hesap — CIKIS_MIKTARI * CIKIS_FIYATI
            ws.Cell(r, 17).FormulaA1 = $"M{r}*N{r}";
            r++;
        }

        // Sayı formatı — miktar/fiyat/tutar sütunları
        if (r > hr + 1)
        {
            ws.Range(hr + 1, 10, r - 1, 15).Style.NumberFormat.Format = "#,##0.00";
            ws.Range(hr + 1, 17, r - 1, 17).Style.NumberFormat.Format = "#,##0.00";

            // Toplam satırı
            int totR = r;
            ws.Cell(totR, 1).Value = "TOPLAM";
            ws.Range(totR, 1, totR, 9).Merge();
            ws.Cell(totR, 10).FormulaA1 = $"SUM(J{hr + 1}:J{r - 1})";
            ws.Cell(totR, 12).FormulaA1 = $"SUM(L{hr + 1}:L{r - 1})";
            ws.Cell(totR, 13).FormulaA1 = $"SUM(M{hr + 1}:M{r - 1})";
            ws.Cell(totR, 15).FormulaA1 = $"SUM(O{hr + 1}:O{r - 1})";
            ws.Cell(totR, 17).FormulaA1 = $"SUM(Q{hr + 1}:Q{r - 1})";
            ws.Range(totR, 1, totR, 18).Style.Font.Bold = true;
            ws.Range(totR, 1, totR, 18).Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");
            ws.Range(totR, 10, totR, 17).Style.NumberFormat.Format = "#,##0.00";
        }

        // Kenarlıklar
        if (r > hr + 1)
        {
            var dataRange = ws.Range(hr, 1, r - 1, 18);
            dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
        }

        ws.SheetView.FreezeRows(hr);
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        return stream.ToArray();
    }

    private static readonly DateTime _delphiBase = new(1899, 12, 30);
    private static DateTime? DelphiToDate(int? d) =>
        (d == null || d <= 0) ? null : _delphiBase.AddDays(d.Value);
    private static TimeSpan? DelphiToTime(int? sn) =>
        (sn == null || sn <= 0) ? null : TimeSpan.FromSeconds(sn.Value);

    public byte[] ExportSabNetKantarHareketleri(
        List<SabNetKantarHareketi> rows,
        Dictionary<string, string> firmaAdlari,
        Dictionary<string, string> urunAdlari,
        string? sozlesmeYili,
        string? kantar)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Kantar Hareketleri");

        ws.Cell("A1").Value = "SABNET KANTAR HAREKETLERİ";
        ws.Range("A1:K1").Merge();
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 14;
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        var filtre = $"Sözleşme: {(string.IsNullOrEmpty(sozlesmeYili) ? "Tümü" : sozlesmeYili)}  |  " +
                     $"Kantar: {(string.IsNullOrEmpty(kantar) ? "Tümü" : kantar)}  |  " +
                     $"Toplam: {rows.Count:N0} kayıt  |  " +
                     $"Oluşturma: {DateTime.Now:dd.MM.yyyy HH:mm}";
        ws.Cell("A2").Value = filtre;
        ws.Range("A2:K2").Merge();
        ws.Cell("A2").Style.Font.Italic = true;
        ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        string[] basliklar = {
            "Tarih", "Fiş No", "İşlem Tipi", "Sözleşme Yılı", "TC Kimlik No",
            "Hesap Kodu", "Firma Adı", "Ürün Kodu", "Ürün Adı", "Plaka No",
            "Şoför", "Açıklama2", "Birim Fiyat", "Brüt", "Dara", "Net",
            "Sevk", "Fark", "Fire %", "Polar %",
            "Nakit", "Kredi Kartı", "Cari", "Havale",
            "Kayıt Tarihi", "Kayıt Saati", "Çıkış Tarihi", "Çıkış Saati",
            "Boşaltma Yeri", "Açıklama", "Kantar", "Row_ID"
        };

        int hr = 4;
        for (int i = 0; i < basliklar.Length; i++)
        {
            var c = ws.Cell(hr, i + 1);
            c.Value = basliklar[i];
            c.Style.Font.Bold = true;
            c.Style.Fill.BackgroundColor = XLColor.FromHtml("#343A40");
            c.Style.Font.FontColor = XLColor.White;
            c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        int r = hr + 1;
        foreach (var h in rows)
        {
            int c = 1;
            ws.Cell(r, c++).Value = DelphiToDate(h.Tarih);
            ws.Cell(r, c++).Value = h.FisNo;
            ws.Cell(r, c++).Value = h.IslemTipi;
            ws.Cell(r, c++).Value = h.SozlesmeYili;
            ws.Cell(r, c++).Value = h.TcKimlikNo;
            ws.Cell(r, c++).Value = h.HesapKodu;
            ws.Cell(r, c++).Value = !string.IsNullOrEmpty(h.HesapKodu) && firmaAdlari.TryGetValue(h.HesapKodu, out var fa) ? fa : "";
            ws.Cell(r, c++).Value = h.UrunKodu;
            ws.Cell(r, c++).Value = !string.IsNullOrEmpty(h.UrunKodu) && urunAdlari.TryGetValue(h.UrunKodu, out var ua) ? ua : "";
            ws.Cell(r, c++).Value = h.PlakaNo;
            ws.Cell(r, c++).Value = h.SoforAdiSoyadi;
            ws.Cell(r, c++).Value = h.Kod5;
            ws.Cell(r, c++).Value = h.BirimFiyat;
            ws.Cell(r, c++).Value = h.Brut;
            ws.Cell(r, c++).Value = h.Dara;
            ws.Cell(r, c++).Value = h.Net;
            ws.Cell(r, c++).Value = h.Sevk;
            ws.Cell(r, c++).Value = h.Fark;
            ws.Cell(r, c++).Value = h.FireOrani;
            ws.Cell(r, c++).Value = h.PolarOrani;
            ws.Cell(r, c++).Value = h.Nakit;
            ws.Cell(r, c++).Value = h.KrediKarti;
            ws.Cell(r, c++).Value = h.Cari;
            ws.Cell(r, c++).Value = h.Havale;
            ws.Cell(r, c++).Value = DelphiToDate(h.KayitTarihi);
            ws.Cell(r, c++).Value = DelphiToTime(h.KayitSaati)?.ToString(@"hh\:mm\:ss");
            ws.Cell(r, c++).Value = DelphiToDate(h.CikisTarihi);
            ws.Cell(r, c++).Value = DelphiToTime(h.CikisSaati)?.ToString(@"hh\:mm\:ss");
            ws.Cell(r, c++).Value = h.BosaltmaYeri;
            ws.Cell(r, c++).Value = h.Aciklama;
            ws.Cell(r, c++).Value = h.KantarKodu;
            ws.Cell(r, c++).Value = h.RowId;
            r++;
        }

        if (rows.Count > 0)
        {
            ws.Range(hr + 1, 1, r - 1, 1).Style.NumberFormat.Format = "dd.mm.yyyy";
            ws.Range(hr + 1, 25, r - 1, 25).Style.NumberFormat.Format = "dd.mm.yyyy";
            ws.Range(hr + 1, 27, r - 1, 27).Style.NumberFormat.Format = "dd.mm.yyyy";
            ws.Range(hr + 1, 13, r - 1, 13).Style.NumberFormat.Format = "#,##0.0000";
            for (int col = 14; col <= 24; col++)
                ws.Range(hr + 1, col, r - 1, col).Style.NumberFormat.Format = "#,##0.00";

            var dataRange = ws.Range(hr, 1, r - 1, basliklar.Length);
            dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.RangeUsed()!.SetAutoFilter();
        }

        ws.SheetView.FreezeRows(hr);
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        return stream.ToArray();
    }

    public byte[] ExportSabNetKantarHareketleriLog(
        List<SabNetKantarHareketiLog> rows,
        Dictionary<string, string> firmaAdlari,
        Dictionary<string, string> urunAdlari,
        string? sozlesmeYili,
        string? kantar)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Kantar Log");

        ws.Cell("A1").Value = "SABNET KANTAR LOG KAYITLARI";
        ws.Range("A1:K1").Merge();
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 14;
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        var filtre = $"Sözleşme: {(string.IsNullOrEmpty(sozlesmeYili) ? "Tümü" : sozlesmeYili)}  |  " +
                     $"Kantar: {(string.IsNullOrEmpty(kantar) ? "Tümü" : kantar)}  |  " +
                     $"Toplam: {rows.Count:N0} kayıt  |  " +
                     $"Oluşturma: {DateTime.Now:dd.MM.yyyy HH:mm}";
        ws.Cell("A2").Value = filtre;
        ws.Range("A2:K2").Merge();
        ws.Cell("A2").Style.Font.Italic = true;
        ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        string[] basliklar = {
            "Log İşlem", "Log Kaydeden", "Log Tarih", "Log Saat",
            "Tarih", "Fiş No", "İşlem Tipi", "Sözleşme Yılı", "TC Kimlik No",
            "Hesap Kodu", "Firma Adı", "Ürün Kodu", "Ürün Adı", "Plaka No",
            "Şoför", "Birim Fiyat", "Brüt", "Dara", "Net",
            "Kayıt Tarihi", "Kayıt Saati", "Çıkış Tarihi", "Çıkış Saati",
            "Kantar", "Row_ID"
        };

        int hr = 4;
        for (int i = 0; i < basliklar.Length; i++)
        {
            var c = ws.Cell(hr, i + 1);
            c.Value = basliklar[i];
            c.Style.Font.Bold = true;
            c.Style.Fill.BackgroundColor = XLColor.FromHtml("#343A40");
            c.Style.Font.FontColor = XLColor.White;
            c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }

        int r = hr + 1;
        foreach (var l in rows)
        {
            int c = 1;
            ws.Cell(r, c++).Value = l.LogIslemTipi;
            ws.Cell(r, c++).Value = l.LogKaydeden;
            ws.Cell(r, c++).Value = DelphiToDate(l.LogKayitTarihi);
            ws.Cell(r, c++).Value = DelphiToTime(l.LogKayitSaati)?.ToString(@"hh\:mm\:ss");
            ws.Cell(r, c++).Value = DelphiToDate(l.Tarih);
            ws.Cell(r, c++).Value = l.FisNo;
            ws.Cell(r, c++).Value = l.IslemTipi;
            ws.Cell(r, c++).Value = l.SozlesmeYili;
            ws.Cell(r, c++).Value = l.TcKimlikNo;
            ws.Cell(r, c++).Value = l.HesapKodu;
            ws.Cell(r, c++).Value = !string.IsNullOrEmpty(l.HesapKodu) && firmaAdlari.TryGetValue(l.HesapKodu, out var fa) ? fa : "";
            ws.Cell(r, c++).Value = l.UrunKodu;
            ws.Cell(r, c++).Value = !string.IsNullOrEmpty(l.UrunKodu) && urunAdlari.TryGetValue(l.UrunKodu, out var ua) ? ua : "";
            ws.Cell(r, c++).Value = l.PlakaNo;
            ws.Cell(r, c++).Value = l.SoforAdiSoyadi;
            ws.Cell(r, c++).Value = l.BirimFiyat;
            ws.Cell(r, c++).Value = l.Brut;
            ws.Cell(r, c++).Value = l.Dara;
            ws.Cell(r, c++).Value = l.Net;
            ws.Cell(r, c++).Value = DelphiToDate(l.KayitTarihi);
            ws.Cell(r, c++).Value = DelphiToTime(l.KayitSaati)?.ToString(@"hh\:mm\:ss");
            ws.Cell(r, c++).Value = DelphiToDate(l.CikisTarihi);
            ws.Cell(r, c++).Value = DelphiToTime(l.CikisSaati)?.ToString(@"hh\:mm\:ss");
            ws.Cell(r, c++).Value = l.KantarKodu;
            ws.Cell(r, c++).Value = l.RowId;
            r++;
        }

        if (rows.Count > 0)
        {
            ws.Range(hr + 1, 3, r - 1, 3).Style.NumberFormat.Format = "dd.mm.yyyy";
            ws.Range(hr + 1, 5, r - 1, 5).Style.NumberFormat.Format = "dd.mm.yyyy";
            ws.Range(hr + 1, 20, r - 1, 20).Style.NumberFormat.Format = "dd.mm.yyyy";
            ws.Range(hr + 1, 22, r - 1, 22).Style.NumberFormat.Format = "dd.mm.yyyy";
            ws.Range(hr + 1, 16, r - 1, 16).Style.NumberFormat.Format = "#,##0.0000";
            for (int col = 17; col <= 19; col++)
                ws.Range(hr + 1, col, r - 1, col).Style.NumberFormat.Format = "#,##0.00";

            var dataRange = ws.Range(hr, 1, r - 1, basliklar.Length);
            dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            ws.RangeUsed()!.SetAutoFilter();
        }

        ws.SheetView.FreezeRows(hr);
        ws.Columns().AdjustToContents();

        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Etil Alkol Carileri (CH 120.99.10.*) için gelen havale tahsilatlarını Excel'e aktarır.
    /// </summary>
    public byte[] ExportEtilAlkolGelenBedeller(List<FinansRaporSatiri> veriler, DateTime baslangic, DateTime bitis)
        => ExportCariGelenBedeller(veriler, baslangic, bitis,
            sheetAdi: "Etil Alkol Gelen Bedel",
            baslik: "Etil Alkol Carileri — Gelen Bedeller (Havale)",
            filtreAciklama: "Filtre: CH Kodu 120.99.10 ile başlayan, Hareket Türü = TAHSİLAT, Havale > 0");

    /// <summary>
    /// Cari Gelen Bedeller (Havale tahsilat) için genel Excel dışa aktarımı.
    /// </summary>
    public byte[] ExportCariGelenBedeller(
        List<FinansRaporSatiri> veriler, DateTime baslangic, DateTime bitis,
        string sheetAdi, string baslik, string filtreAciklama)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add(sheetAdi);

        const int colCount = 9;
        ws.Cell(1, 1).Value = $"{baslik}   {baslangic:dd.MM.yyyy} / {bitis:dd.MM.yyyy}";
        ws.Range(1, 1, 1, colCount).Merge();
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 13;
        ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");

        ws.Cell(2, 1).Value = filtreAciklama;
        ws.Range(2, 1, 2, colCount).Merge();
        ws.Cell(2, 1).Style.Font.Italic = true;
        ws.Cell(2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        string[] basliklar = { "Tarih", "CH Kodu", "CH Ünvanı", "Banka Kodu",
                               "Banka Açıklaması", "Proje Kodu", "İşlem No", "Fiş Türü", "Havale (₺)" };
        const int hr = 4;
        for (int i = 0; i < basliklar.Length; i++)
        {
            var c = ws.Cell(hr, i + 1);
            c.Value = basliklar[i];
            c.Style.Font.Bold = true;
            c.Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2");
            c.Style.Font.FontColor = XLColor.White;
            c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            c.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        int row = hr + 1;
        decimal toplam = 0;
        foreach (var x in veriler)
        {
            ws.Cell(row, 1).Value = x.Tarih;
            ws.Cell(row, 1).Style.NumberFormat.Format = "dd.MM.yyyy";
            ws.Cell(row, 2).Value = x.ChKod;
            ws.Cell(row, 3).Value = x.ChUnvani;
            ws.Cell(row, 4).Value = x.BankaHesapKodu;
            ws.Cell(row, 5).Value = x.BankaHesapAciklamasi;
            ws.Cell(row, 6).Value = x.ProjeKodu;
            ws.Cell(row, 7).Value = x.IslemNo;
            ws.Cell(row, 8).Value = x.FisTuru;
            ws.Cell(row, 9).Value = x.Havale;
            ws.Cell(row, 9).Style.NumberFormat.Format = "#,##0.00";
            toplam += x.Havale;
            row++;
        }

        // Toplam satırı
        ws.Cell(row, 1).Value = $"TOPLAM ({veriler.Count} satır)";
        ws.Range(row, 1, row, 8).Merge();
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF59D");
        ws.Cell(row, 9).Value = toplam;
        ws.Cell(row, 9).Style.NumberFormat.Format = "#,##0.00";
        ws.Cell(row, 9).Style.Font.Bold = true;
        ws.Cell(row, 9).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF59D");
        ws.Range(row, 1, row, 9).Style.Border.TopBorder = XLBorderStyleValues.Medium;

        // Sütun genişlikleri (manuel ayar — AdjustToContents bazen başlığa göre dar bırakıyor)
        ws.Column(1).Width = 12;   // Tarih
        ws.Column(2).Width = 18;   // CH Kodu
        ws.Column(3).Width = 45;   // CH Ünvanı
        ws.Column(4).Width = 14;   // Banka Kodu
        ws.Column(5).Width = 30;   // Banka Açıklaması
        ws.Column(6).Width = 14;   // Proje Kodu
        ws.Column(7).Width = 16;   // İşlem No
        ws.Column(8).Width = 12;   // Fiş Türü
        ws.Column(9).Width = 16;   // Havale

        // Otomatik filtre + üst satır dondurma
        if (row > hr + 1)
            ws.Range(hr, 1, row - 1, colCount).SetAutoFilter();
        ws.SheetView.FreezeRows(hr);

        // Yazdırma ayarı: yatay + 1 sayfa eni + başlık satırı her sayfada
        YazdirmaAyariUygula(ws, basliksatiri: hr);

        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// 120.04.* carileri için TAHSİLAT (Havale + Kredi Kartı + diğer) tüm tahsilat tiplerini Excel'e aktarır.
    /// </summary>
    public byte[] ExportCariTahsilatlari(
        List<FinansRaporSatiri> veriler, DateTime baslangic, DateTime bitis,
        string sheetAdi, string baslik, string filtreAciklama)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add(sheetAdi);

        const int colCount = 10;
        ws.Cell(1, 1).Value = $"{baslik}   {baslangic:dd.MM.yyyy} / {bitis:dd.MM.yyyy}";
        ws.Range(1, 1, 1, colCount).Merge();
        ws.Cell(1, 1).Style.Font.Bold = true;
        ws.Cell(1, 1).Style.Font.FontSize = 13;
        ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E8F5E9");

        ws.Cell(2, 1).Value = filtreAciklama;
        ws.Range(2, 1, 2, colCount).Merge();
        ws.Cell(2, 1).Style.Font.Italic = true;
        ws.Cell(2, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        string[] basliklar = { "Tarih", "CH Kodu", "CH Ünvanı", "Modül",
                               "Banka Kodu", "Banka Açıklaması", "Proje Kodu",
                               "İşlem No", "Fiş Türü", "Tutar (₺)" };
        const int hr = 4;
        for (int i = 0; i < basliklar.Length; i++)
        {
            var c = ws.Cell(hr, i + 1);
            c.Value = basliklar[i];
            c.Style.Font.Bold = true;
            c.Style.Fill.BackgroundColor = XLColor.FromHtml("#2E7D32");
            c.Style.Font.FontColor = XLColor.White;
            c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            c.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        int row = hr + 1;
        decimal toplam = 0;
        foreach (var x in veriler)
        {
            var tutar = x.Havale + x.Cek + x.Devir + x.Diger;
            ws.Cell(row, 1).Value = x.Tarih;
            ws.Cell(row, 1).Style.NumberFormat.Format = "dd.MM.yyyy";
            ws.Cell(row, 2).Value = x.ChKod;
            ws.Cell(row, 3).Value = x.ChUnvani;
            ws.Cell(row, 4).Value = x.Modul;
            ws.Cell(row, 5).Value = x.BankaHesapKodu;
            ws.Cell(row, 6).Value = x.BankaHesapAciklamasi;
            ws.Cell(row, 7).Value = x.ProjeKodu;
            ws.Cell(row, 8).Value = x.IslemNo;
            ws.Cell(row, 9).Value = x.FisTuru;
            ws.Cell(row, 10).Value = tutar;
            ws.Cell(row, 10).Style.NumberFormat.Format = "#,##0.00";
            toplam += tutar;
            row++;
        }

        // Toplam satırı
        ws.Cell(row, 1).Value = $"TOPLAM ({veriler.Count} satır)";
        ws.Range(row, 1, row, 9).Merge();
        ws.Cell(row, 1).Style.Font.Bold = true;
        ws.Cell(row, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
        ws.Cell(row, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF59D");
        ws.Cell(row, 10).Value = toplam;
        ws.Cell(row, 10).Style.NumberFormat.Format = "#,##0.00";
        ws.Cell(row, 10).Style.Font.Bold = true;
        ws.Cell(row, 10).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF59D");
        ws.Range(row, 1, row, 10).Style.Border.TopBorder = XLBorderStyleValues.Medium;

        ws.Column(1).Width = 12;
        ws.Column(2).Width = 18;
        ws.Column(3).Width = 45;
        ws.Column(4).Width = 14;
        ws.Column(5).Width = 14;
        ws.Column(6).Width = 30;
        ws.Column(7).Width = 14;
        ws.Column(8).Width = 16;
        ws.Column(9).Width = 22;
        ws.Column(10).Width = 16;

        if (row > hr + 1)
            ws.Range(hr, 1, row - 1, colCount).SetAutoFilter();
        ws.SheetView.FreezeRows(hr);

        YazdirmaAyariUygula(ws, basliksatiri: hr);

        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Cari mutabakat sonucunu üç sayfalık Excel'e aktarır:
    /// 1) Eşleşenler  2) Sadece Bizde  3) Sadece Onlarda  +  4) Özet (üst sayfa)
    /// </summary>
    public byte[] MutabakatExportEt(MutabakatSonuc s)
    {
        using var wb = new XLWorkbook();

        // ---- Sayfa 1: ÖZET ----
        var ozet = wb.Worksheets.Add("Özet");
        ozet.Cell(1, 1).Value = "CARİ MUTABAKAT RAPORU";
        ozet.Cell(1, 1).Style.Font.Bold = true;
        ozet.Cell(1, 1).Style.Font.FontSize = 16;
        ozet.Range(1, 1, 1, 4).Merge();
        ozet.Range(1, 1, 1, 4).Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2");
        ozet.Range(1, 1, 1, 4).Style.Font.FontColor = XLColor.White;
        ozet.Range(1, 1, 1, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ozet.Cell(3, 1).Value = "Vergi/TC No:";
        ozet.Cell(3, 2).Value = s.AranılanVergiNo;
        ozet.Cell(4, 1).Value = "Tarih Aralığı:";
        ozet.Cell(4, 2).Value = $"{s.Baslangic:dd.MM.yyyy} - {s.Bitis:dd.MM.yyyy}";
        ozet.Cell(5, 1).Value = "Bizdeki Cari Sayısı:";
        ozet.Cell(5, 2).Value = s.EslesenCariler.Count;
        ozet.Range(3, 1, 5, 1).Style.Font.Bold = true;

        ozet.Cell(7, 1).Value = "TOPLAMLAR";
        ozet.Cell(7, 1).Style.Font.Bold = true;
        ozet.Cell(7, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");
        ozet.Range(7, 1, 7, 4).Merge();

        ozet.Cell(8, 1).Value = "";
        ozet.Cell(8, 2).Value = "Borç";
        ozet.Cell(8, 3).Value = "Alacak";
        ozet.Cell(8, 4).Value = "Net (Borç-Alacak)";
        ozet.Range(8, 1, 8, 4).Style.Font.Bold = true;

        ozet.Cell(9, 1).Value = "Bizdeki";
        ozet.Cell(9, 2).Value = s.BizimToplamBorc;
        ozet.Cell(9, 3).Value = s.BizimToplamAlacak;
        ozet.Cell(9, 4).Value = s.BizimNet;

        ozet.Cell(10, 1).Value = "Onlarda";
        ozet.Cell(10, 2).Value = s.OnlarToplamBorc;
        ozet.Cell(10, 3).Value = s.OnlarToplamAlacak;
        ozet.Cell(10, 4).Value = s.OnlarNet;

        ozet.Cell(11, 1).Value = "Mutabakat Farkı";
        ozet.Cell(11, 4).Value = s.MutabakatFarki;
        ozet.Cell(11, 1).Style.Font.Bold = true;
        ozet.Cell(11, 4).Style.Font.Bold = true;
        ozet.Cell(11, 4).Style.Fill.BackgroundColor =
            Math.Abs(s.MutabakatFarki) < 0.01m ? XLColor.FromHtml("#C8E6C9") : XLColor.FromHtml("#FFCDD2");

        ozet.Range(9, 2, 11, 4).Style.NumberFormat.Format = "#,##0.00";

        ozet.Cell(13, 1).Value = "Eşleşen Hareketler:";
        ozet.Cell(13, 2).Value = s.Eslesenler.Count;
        ozet.Cell(14, 1).Value = "Sadece Bizde:";
        ozet.Cell(14, 2).Value = s.SadeceBizde.Count;
        ozet.Cell(15, 1).Value = "Sadece Onlarda:";
        ozet.Cell(15, 2).Value = s.SadeceOnlarda.Count;
        ozet.Range(13, 1, 15, 1).Style.Font.Bold = true;

        if (s.EslesenCariler.Count > 0)
        {
            int r = 17;
            ozet.Cell(r, 1).Value = "BİZDEKİ CARİ KARTLARI";
            ozet.Cell(r, 1).Style.Font.Bold = true;
            ozet.Cell(r, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");
            ozet.Range(r, 1, r, 4).Merge();
            r++;
            ozet.Cell(r, 1).Value = "Kod";
            ozet.Cell(r, 2).Value = "Ünvan";
            ozet.Cell(r, 3).Value = "Vergi/TC";
            ozet.Cell(r, 4).Value = "Hareket";
            ozet.Range(r, 1, r, 4).Style.Font.Bold = true;
            r++;
            foreach (var c in s.EslesenCariler)
            {
                ozet.Cell(r, 1).Value = c.Kod;
                ozet.Cell(r, 2).Value = c.Unvan;
                ozet.Cell(r, 3).Value = c.VergiNo;
                ozet.Cell(r, 4).Value = c.HareketSayisi;
                r++;
            }
        }
        ozet.Column(1).Width = 22;
        ozet.Column(2).Width = 38;
        ozet.Column(3).Width = 18;
        ozet.Column(4).Width = 18;
        YazdirmaAyariUygula(ozet);

        // ---- Sayfa 2: Eşleşenler ----
        var es = wb.Worksheets.Add("Eşleşenler");
        var esBaslik = new[] { "Bizim Tarih", "Bizim Cari", "Bizim Cari Ünvan", "Bizim Fiş",
            "Bizim Borç", "Bizim Alacak",
            "Onlar Tarih", "Onlar Açıklama", "Onlar Fiş", "Onlar Borç", "Onlar Alacak",
            "Gün Farkı", "Tutar Farkı", "Yöntem" };
        for (int i = 0; i < esBaslik.Length; i++)
        {
            var c = es.Cell(1, i + 1);
            c.Value = esBaslik[i];
            c.Style.Font.Bold = true;
            c.Style.Fill.BackgroundColor = XLColor.FromHtml("#2E7D32");
            c.Style.Font.FontColor = XLColor.White;
        }
        int row1 = 2;
        foreach (var e in s.Eslesenler)
        {
            es.Cell(row1, 1).Value = e.Bizim.Tarih;
            es.Cell(row1, 1).Style.DateFormat.Format = "dd.MM.yyyy";
            es.Cell(row1, 2).Value = e.Bizim.CariKodu;
            es.Cell(row1, 3).Value = e.Bizim.CariUnvan;
            es.Cell(row1, 4).Value = e.Bizim.FisNo;
            es.Cell(row1, 5).Value = e.Bizim.Borc;
            es.Cell(row1, 6).Value = e.Bizim.Alacak;
            es.Cell(row1, 7).Value = e.Onlar.Tarih;
            es.Cell(row1, 7).Style.DateFormat.Format = "dd.MM.yyyy";
            es.Cell(row1, 8).Value = e.Onlar.Aciklama;
            es.Cell(row1, 9).Value = e.Onlar.FisNo;
            es.Cell(row1, 10).Value = e.Onlar.Borc;
            es.Cell(row1, 11).Value = e.Onlar.Alacak;
            es.Cell(row1, 12).Value = e.GunFarki;
            es.Cell(row1, 13).Value = e.TutarFarki;
            es.Cell(row1, 14).Value = e.EslesmeYontemi;
            row1++;
        }
        es.Range(2, 5, Math.Max(2, row1 - 1), 6).Style.NumberFormat.Format = "#,##0.00";
        es.Range(2, 10, Math.Max(2, row1 - 1), 11).Style.NumberFormat.Format = "#,##0.00";
        es.Range(2, 13, Math.Max(2, row1 - 1), 13).Style.NumberFormat.Format = "#,##0.00";
        es.Columns().AdjustToContents(1, 60);
        if (row1 > 2) es.Range(1, 1, row1 - 1, esBaslik.Length).SetAutoFilter();
        es.SheetView.FreezeRows(1);
        YazdirmaAyariUygula(es, basliksatiri: 1);

        // ---- Sayfa 3: Sadece Bizde ----
        var sb = wb.Worksheets.Add("Sadece Bizde");
        TekTaraflıYaz(sb, s.SadeceBizde, bizim: true);
        YazdirmaAyariUygula(sb, basliksatiri: 1);

        // ---- Sayfa 4: Sadece Onlarda ----
        var so = wb.Worksheets.Add("Sadece Onlarda");
        TekTaraflıYaz(so, s.SadeceOnlarda, bizim: false);
        YazdirmaAyariUygula(so, basliksatiri: 1);

        using var stream = new MemoryStream();
        wb.SaveAs(stream);
        return stream.ToArray();
    }

    private static void TekTaraflıYaz(IXLWorksheet ws, List<MutabakatKayit> liste, bool bizim)
    {
        var basliklar = bizim
            ? new[] { "Tarih", "Cari Kodu", "Cari Ünvanı", "Fiş No", "Fiş Türü", "Açıklama", "Borç", "Alacak" }
            : new[] { "Tarih", "Açıklama", "Fiş No", "Borç", "Alacak" };

        var baslikRengi = bizim ? "#F57C00" : "#C62828";
        for (int i = 0; i < basliklar.Length; i++)
        {
            var c = ws.Cell(1, i + 1);
            c.Value = basliklar[i];
            c.Style.Font.Bold = true;
            c.Style.Fill.BackgroundColor = XLColor.FromHtml(baslikRengi);
            c.Style.Font.FontColor = XLColor.White;
        }

        int r = 2;
        foreach (var k in liste)
        {
            int col = 1;
            ws.Cell(r, col).Value = k.Tarih;
            ws.Cell(r, col).Style.DateFormat.Format = "dd.MM.yyyy";
            col++;
            if (bizim)
            {
                ws.Cell(r, col++).Value = k.CariKodu;
                ws.Cell(r, col++).Value = k.CariUnvan;
                ws.Cell(r, col++).Value = k.FisNo;
                ws.Cell(r, col++).Value = k.FisTuru;
                ws.Cell(r, col++).Value = k.Aciklama;
            }
            else
            {
                ws.Cell(r, col++).Value = k.Aciklama;
                ws.Cell(r, col++).Value = k.FisNo;
            }
            ws.Cell(r, col++).Value = k.Borc;
            ws.Cell(r, col++).Value = k.Alacak;
            r++;
        }
        int borcAlacakStart = bizim ? 7 : 4;
        if (r > 2)
            ws.Range(2, borcAlacakStart, r - 1, borcAlacakStart + 1).Style.NumberFormat.Format = "#,##0.00";
        ws.Columns().AdjustToContents(1, 60);
        if (r > 2) ws.Range(1, 1, r - 1, basliklar.Length).SetAutoFilter();
        ws.SheetView.FreezeRows(1);
    }

    public byte[] KurFarkiExportEt(KurFarkiSonuc s)
    {
        using var wb = new XLWorkbook();

        // ---- Özet ----
        var ozet = wb.Worksheets.Add("Özet");
        ozet.Cell(1, 1).Value = "KARŞI FİRMA KUR FARKI RAPORU (FIFO)";
        ozet.Cell(1, 1).Style.Font.Bold = true;
        ozet.Cell(1, 1).Style.Font.FontSize = 16;
        ozet.Range(1, 1, 1, 5).Merge();
        ozet.Range(1, 1, 1, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2");
        ozet.Range(1, 1, 1, 5).Style.Font.FontColor = XLColor.White;
        ozet.Range(1, 1, 1, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        ozet.Cell(3, 1).Value = "Vergi/TC No:";
        ozet.Cell(3, 2).Value = s.AranılanVergiNo;
        ozet.Cell(4, 1).Value = "Tarih Aralığı:";
        ozet.Cell(4, 2).Value = $"{s.Baslangic:dd.MM.yyyy} - {s.Bitis:dd.MM.yyyy}";
        ozet.Cell(5, 1).Value = "Döviz Cinsi:";
        ozet.Cell(5, 2).Value = s.Doviz;
        ozet.Cell(6, 1).Value = "Toplam Fatura (Döviz):";
        ozet.Cell(6, 2).Value = s.ToplamFaturaDoviz;
        ozet.Cell(7, 1).Value = "Toplam Ödeme (Döviz):";
        ozet.Cell(7, 2).Value = s.ToplamOdemeDoviz;
        ozet.Cell(8, 1).Value = "① FIFO Toplam Kur Farkı (TL):";
        ozet.Cell(8, 2).Value = s.ToplamKurFarki;
        ozet.Cell(8, 1).Style.Font.Bold = true;
        ozet.Cell(8, 2).Style.Font.Bold = true;
        ozet.Cell(8, 2).Style.Fill.BackgroundColor =
            s.ToplamKurFarki >= 0 ? XLColor.FromHtml("#C8E6C9") : XLColor.FromHtml("#FFCDD2");
        ozet.Cell(9, 1).Value = "② Belge Bazlı Toplam TL Farkı:";
        ozet.Cell(9, 2).Value = s.BelgeBazliToplamTLFarki;
        ozet.Cell(9, 1).Style.Font.Bold = true;
        ozet.Cell(9, 2).Style.Font.Bold = true;
        ozet.Cell(9, 2).Style.Fill.BackgroundColor =
            s.BelgeBazliToplamTLFarki >= 0 ? XLColor.FromHtml("#C8E6C9") : XLColor.FromHtml("#FFCDD2");
        ozet.Range(3, 1, 9, 1).Style.Font.Bold = true;
        ozet.Range(6, 2, 9, 2).Style.NumberFormat.Format = "#,##0.00";
        ozet.Column(1).Width = 32;
        ozet.Column(2).Width = 20;
        YazdirmaAyariUygula(ozet);

        // ---- ① FIFO Eşleşmeler ----
        var es = wb.Worksheets.Add("① FIFO Eşleşmeleri");
        var basliklar = new[]
        {
            "Fatura Tarih","Fatura Açıklama","Fatura Fiş",
            "Fatura Döviz","Fatura TL","Fatura Kur",
            "Ödeme Tarih","Ödeme Açıklama","Ödeme Fiş",
            "Ödeme Döviz","Ödeme TL","Ödeme Kur",
            "Eşleşen Döviz","Kur Farkı (TL)"
        };
        for (int i = 0; i < basliklar.Length; i++)
        {
            var c = es.Cell(1, i + 1);
            c.Value = basliklar[i];
            c.Style.Font.Bold = true;
            c.Style.Fill.BackgroundColor = XLColor.FromHtml("#1565C0");
            c.Style.Font.FontColor = XLColor.White;
        }
        int r1 = 2;
        foreach (var e in s.Eslesmeler)
        {
            es.Cell(r1, 1).Value = e.FaturaTarihi;  es.Cell(r1, 1).Style.DateFormat.Format = "dd.MM.yyyy";
            es.Cell(r1, 2).Value = e.FaturaAciklama;
            es.Cell(r1, 3).Value = e.FaturaFisNo;
            es.Cell(r1, 4).Value = e.FaturaDoviz;
            es.Cell(r1, 5).Value = e.FaturaTL;
            es.Cell(r1, 6).Value = e.FaturaKur;
            es.Cell(r1, 7).Value = e.OdemeTarihi;   es.Cell(r1, 7).Style.DateFormat.Format = "dd.MM.yyyy";
            es.Cell(r1, 8).Value = e.OdemeAciklama;
            es.Cell(r1, 9).Value = e.OdemeFisNo;
            es.Cell(r1, 10).Value = e.OdemeDoviz;
            es.Cell(r1, 11).Value = e.OdemeTL;
            es.Cell(r1, 12).Value = e.OdemeKur;
            es.Cell(r1, 13).Value = e.EslesenDoviz;
            es.Cell(r1, 14).Value = e.KurFarki;
            r1++;
        }
        if (r1 > 2)
        {
            es.Range(2, 4, r1 - 1, 5).Style.NumberFormat.Format = "#,##0.00";
            es.Range(2, 6, r1 - 1, 6).Style.NumberFormat.Format = "#,##0.0000";
            es.Range(2, 10, r1 - 1, 11).Style.NumberFormat.Format = "#,##0.00";
            es.Range(2, 12, r1 - 1, 12).Style.NumberFormat.Format = "#,##0.0000";
            es.Range(2, 13, r1 - 1, 13).Style.NumberFormat.Format = "#,##0.00";
            es.Range(2, 14, r1 - 1, 14).Style.NumberFormat.Format = "#,##0.00";
            es.Range(1, 1, r1 - 1, basliklar.Length).SetAutoFilter();
        }
        es.Columns().AdjustToContents(1, 60);
        es.SheetView.FreezeRows(1);
        YazdirmaAyariUygula(es, basliksatiri: 1);

        // ---- ② Belge Bazlı Kur Farkı (Bizim Logo ↔ Onlar) ----
        if (s.BelgeBazliFarklar.Count > 0)
        {
            var bb = wb.Worksheets.Add("② Belge Bazlı Kur Farkı");
            var bbBaslik = new[]
            {
                "Tarih","Yön","Cari","Fiş No","Açıklama",
                "Döviz Tutarı","Döviz Cinsi",
                "Bizim TL","Bizim Kur",
                "Onların TL","Onların Kur",
                "TL Farkı","Kur Farkı","Yöntem"
            };
            for (int i = 0; i < bbBaslik.Length; i++)
            {
                var c = bb.Cell(1, i + 1);
                c.Value = bbBaslik[i];
                c.Style.Font.Bold = true;
                c.Style.Fill.BackgroundColor = XLColor.FromHtml("#EF6C00");
                c.Style.Font.FontColor = XLColor.White;
            }
            int rb = 2;
            foreach (var b in s.BelgeBazliFarklar)
            {
                bb.Cell(rb, 1).Value = b.Tarih; bb.Cell(rb, 1).Style.DateFormat.Format = "dd.MM.yyyy";
                bb.Cell(rb, 2).Value = b.Yon;
                bb.Cell(rb, 3).Value = b.BizimCariKodu;
                bb.Cell(rb, 4).Value = b.FisNo;
                bb.Cell(rb, 5).Value = b.Aciklama;
                bb.Cell(rb, 6).Value = b.DovizTutar;
                bb.Cell(rb, 7).Value = b.Doviz;
                bb.Cell(rb, 8).Value = b.BizimTL;
                bb.Cell(rb, 9).Value = b.BizimKur;
                bb.Cell(rb, 10).Value = b.OnlarTL;
                bb.Cell(rb, 11).Value = b.OnlarKur;
                bb.Cell(rb, 12).Value = b.TLFarki;
                bb.Cell(rb, 13).Value = b.KurFarki;
                bb.Cell(rb, 14).Value = b.EslesmeYontemi;
                bb.Cell(rb, 12).Style.Fill.BackgroundColor = b.TLFarki > 0
                    ? XLColor.FromHtml("#C8E6C9")
                    : (b.TLFarki < 0 ? XLColor.FromHtml("#FFCDD2") : XLColor.NoColor);
                rb++;
            }
            if (rb > 2)
            {
                bb.Range(2, 6, rb - 1, 6).Style.NumberFormat.Format = "#,##0.00";
                bb.Range(2, 8, rb - 1, 8).Style.NumberFormat.Format = "#,##0.00";
                bb.Range(2, 9, rb - 1, 9).Style.NumberFormat.Format = "#,##0.0000";
                bb.Range(2, 10, rb - 1, 10).Style.NumberFormat.Format = "#,##0.00";
                bb.Range(2, 11, rb - 1, 11).Style.NumberFormat.Format = "#,##0.0000";
                bb.Range(2, 12, rb - 1, 12).Style.NumberFormat.Format = "#,##0.00";
                bb.Range(2, 13, rb - 1, 13).Style.NumberFormat.Format = "#,##0.0000";
                bb.Range(1, 1, rb - 1, bbBaslik.Length).SetAutoFilter();
            }
            bb.Columns().AdjustToContents(1, 60);
            bb.SheetView.FreezeRows(1);
            YazdirmaAyariUygula(bb, basliksatiri: 1);
        }

        // ---- Açık Faturalar ----
        if (s.AcikFaturalar.Count > 0)
        {
            var af = wb.Worksheets.Add("Açık Faturalar");
            af.Cell(1, 1).Value = "Tarih";
            af.Cell(1, 2).Value = "Açıklama";
            af.Cell(1, 3).Value = "Fiş";
            af.Cell(1, 4).Value = "Kalan Döviz";
            af.Cell(1, 5).Value = "Kalan TL (yaklaşık)";
            af.Range(1, 1, 1, 5).Style.Font.Bold = true;
            af.Range(1, 1, 1, 5).Style.Fill.BackgroundColor = XLColor.FromHtml("#EF6C00");
            af.Range(1, 1, 1, 5).Style.Font.FontColor = XLColor.White;
            int r2 = 2;
            foreach (var f in s.AcikFaturalar)
            {
                af.Cell(r2, 1).Value = f.Tarih; af.Cell(r2, 1).Style.DateFormat.Format = "dd.MM.yyyy";
                af.Cell(r2, 2).Value = f.Aciklama;
                af.Cell(r2, 3).Value = f.FisNo;
                af.Cell(r2, 4).Value = f.DovizAlacak;
                af.Cell(r2, 5).Value = f.Alacak;
                r2++;
            }
            if (r2 > 2) af.Range(2, 4, r2 - 1, 5).Style.NumberFormat.Format = "#,##0.00";
            af.Columns().AdjustToContents(1, 60);
            af.SheetView.FreezeRows(1);
            YazdirmaAyariUygula(af, basliksatiri: 1);
        }

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    public byte[] ExportEFaturaListesi(
        List<EFaturaListItem> kayitlar,
        EFaturaFiltre filtre)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("e-Fatura Listesi");

        // Başlık bandı
        ws.Cell("A1").Value = "e-FATURA LİSTESİ";
        ws.Range("A1:P1").Merge();
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 14;
        ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromHtml("#1976D2");
        ws.Cell("A1").Style.Font.FontColor = XLColor.White;
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        // Filtre özeti
        var ozet = $"Yön: {filtre.Yon} | Tarih: {filtre.Baslangic:dd.MM.yyyy} - {filtre.Bitis:dd.MM.yyyy} | Tip: {(string.IsNullOrWhiteSpace(filtre.Tip) ? "Hepsi" : filtre.Tip)} | Kayıt: {kayitlar.Count}";
        ws.Cell("A2").Value = ozet;
        ws.Range("A2:P2").Merge();
        ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell("A2").Style.Font.Italic = true;
        ws.Cell("A2").Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");

        // Başlık satırı (4. satır)
        string[] basliklar = {
            "Tarih", "Saat", "Yön", "Tip", "Senaryo",
            "Fatura No", "VKN/TCKN", "Karşı Taraf", "İlk Kalem",
            "Logo Cari Kod", "Logo Cari Ünvan",
            "Döviz", "Matrah", "KDV", "Fatura Tutarı", "Tevkifat"
        };
        int hdrRow = 4;
        for (int i = 0; i < basliklar.Length; i++)
        {
            var c = ws.Cell(hdrRow, i + 1);
            c.Value = basliklar[i];
            c.Style.Font.Bold = true;
            c.Style.Fill.BackgroundColor = XLColor.FromHtml("#0D47A1");
            c.Style.Font.FontColor = XLColor.White;
            c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            c.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }

        int row = hdrRow + 1;
        foreach (var r in kayitlar)
        {
            ws.Cell(row, 1).Value  = r.Tarih.Date;
            ws.Cell(row, 1).Style.DateFormat.Format = "dd.MM.yyyy";
            ws.Cell(row, 2).Value  = r.Tarih.ToString("HH:mm");
            ws.Cell(row, 3).Value  = r.Yon;
            ws.Cell(row, 4).Value  = r.Tip;
            ws.Cell(row, 5).Value  = r.ProfileId;
            ws.Cell(row, 6).Value  = r.FaturaNo;
            ws.Cell(row, 7).Value  = r.KarsiVKN;
            ws.Cell(row, 8).Value  = r.KarsiUnvan;
            ws.Cell(row, 9).Value  = r.IlkKalem;
            ws.Cell(row, 10).Value = r.LogoCariKod;
            ws.Cell(row, 11).Value = r.LogoCariUnvan;
            ws.Cell(row, 12).Value = r.Doviz;
            ws.Cell(row, 13).Value = r.Matrah;
            ws.Cell(row, 14).Value = r.KdvTutar;
            ws.Cell(row, 15).Value = r.ToplamTutar;
            ws.Cell(row, 16).Value = r.TevkifatTutar;
            ws.Cell(row, 13).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 14).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 15).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 16).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 15).Style.Font.Bold = true;
            ws.Cell(row, 15).Style.Font.FontColor = XLColor.FromHtml("#0D47A1");
            if (r.TevkifatTutar != 0)
                ws.Cell(row, 16).Style.Font.FontColor = XLColor.FromHtml("#C62828");

            // Logo eşleşmesi olanları hafifçe sarımsı vurgula
            if (!string.IsNullOrEmpty(r.LogoCariKod))
            {
                ws.Cell(row, 10).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF8E1");
                ws.Cell(row, 11).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF8E1");
                ws.Cell(row, 10).Style.Font.Bold = true;
            }
            row++;
        }

        // Tablo aralığını çerçevele + auto-filter
        var dataRange = ws.Range(hdrRow, 1, row - 1, basliklar.Length);
        dataRange.Style.Border.InsideBorder  = XLBorderStyleValues.Thin;
        dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        ws.RangeUsed()!.SetAutoFilter();
        ws.SheetView.FreezeRows(hdrRow);

        // Toplam satırı
        if (kayitlar.Count > 0)
        {
            var sumRow = row;
            ws.Cell(sumRow, 1).Value = "TOPLAM";
            ws.Range(sumRow, 1, sumRow, 12).Merge();
            ws.Cell(sumRow, 1).Style.Font.Bold = true;
            ws.Cell(sumRow, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            ws.Cell(sumRow, 13).FormulaA1 = $"SUM(M{hdrRow + 1}:M{sumRow - 1})";
            ws.Cell(sumRow, 14).FormulaA1 = $"SUM(N{hdrRow + 1}:N{sumRow - 1})";
            ws.Cell(sumRow, 15).FormulaA1 = $"SUM(O{hdrRow + 1}:O{sumRow - 1})";
            ws.Cell(sumRow, 16).FormulaA1 = $"SUM(P{hdrRow + 1}:P{sumRow - 1})";
            for (int col = 13; col <= 16; col++)
            {
                ws.Cell(sumRow, col).Style.NumberFormat.Format = "#,##0.00";
                ws.Cell(sumRow, col).Style.Font.Bold = true;
                ws.Cell(sumRow, col).Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");
            }
            ws.Cell(sumRow, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");
            ws.Range(sumRow, 1, sumRow, 16).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        }

        ws.Columns().AdjustToContents();
        // Aşırı geniş ünvan/kalem için tavan koy
        if (ws.Column(8).Width  > 60) ws.Column(8).Width  = 60;
        if (ws.Column(9).Width  > 50) ws.Column(9).Width  = 50;
        if (ws.Column(11).Width > 50) ws.Column(11).Width = 50;

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    /// <summary>
    /// e-Fatura Logo İşleme Kontrolü ekranındaki listeyi Excel'e aktarır.
    /// Durum bazlı renklendirme ve özet satır içerir.
    /// </summary>
    public byte[] EFaturaLogoKontrolExportEt(
        List<EFaturaListItem> kayitlar,
        DateTime baslangic,
        DateTime bitis,
        int tarihToleransGun)
    {
        using var wb = new XLWorkbook();

        // Tüm faturalar (özet sayfa)
        SayfaDolur(wb, "Tümü", "GELEN e-FATURA · LOGO İŞLEME KONTROLÜ · TÜMÜ",
                   kayitlar, baslangic, bitis, tarihToleransGun, "#1976D2");

        // İşlenen = Kesin + Modifiye (red hariç)
        var islenen = kayitlar.Where(x => !x.RedEdildi &&
                                          (x.LogoDurum == LogoIslemDurumu.Kesin ||
                                           x.LogoDurum == LogoIslemDurumu.Modifiye)).ToList();
        SayfaDolur(wb, "İşlenen", "LOGO'YA İŞLENEN FATURALAR (Kesin + Modifiye)",
                   islenen, baslangic, bitis, tarihToleransGun, "#2E7D32");

        // İade Kesilenler / Şüpheli
        var supheli = kayitlar.Where(x => !x.RedEdildi && x.LogoDurum == LogoIslemDurumu.Supheli).ToList();
        SayfaDolur(wb, "İade Kesilenler-Şüpheli",
                   "İADE KESİLENLER / ŞÜPHELİ FATURALAR (Tutar+Tarih Tutuyor, FICHENO Tutmuyor)",
                   supheli, baslangic, bitis, tarihToleransGun, "#E65100");

        // İşlenmeyenler = Yok (red hariç)
        var islenmeyen = kayitlar.Where(x => !x.RedEdildi && x.LogoDurum == LogoIslemDurumu.Yok).ToList();
        SayfaDolur(wb, "İşlenmeyenler", "LOGO'YA İŞLENMEMİŞ FATURALAR (Uyarı Gerektiren)",
                   islenmeyen, baslangic, bitis, tarihToleransGun, "#B71C1C");

        // Red / İptal
        var redler = kayitlar.Where(x => x.RedEdildi).ToList();
        SayfaDolur(wb, "Red-İptal", "RED EDİLEN / İPTAL EDİLEN FATURALAR",
                   redler, baslangic, bitis, tarihToleransGun, "#424242");

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }

    private static void SayfaDolur(
        XLWorkbook wb,
        string sayfaAdi,
        string baslik,
        List<EFaturaListItem> kayitlar,
        DateTime baslangic,
        DateTime bitis,
        int tarihToleransGun,
        string baslikRengi)
    {
        var ws = wb.Worksheets.Add(sayfaAdi);
        const int sonSutun = 18;   // A..R

        ws.Cell("A1").Value = baslik;
        ws.Range("A1:R1").Merge();
        ws.Cell("A1").Style.Font.Bold = true;
        ws.Cell("A1").Style.Font.FontSize = 14;
        ws.Cell("A1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell("A1").Style.Fill.BackgroundColor = XLColor.FromHtml(baslikRengi);
        ws.Cell("A1").Style.Font.FontColor = XLColor.White;

        ws.Cell("A2").Value = $"Tarih: {baslangic:dd.MM.yyyy} - {bitis:dd.MM.yyyy}  ·  Tolerans: {tarihToleransGun} gün  ·  Bu sayfa: {kayitlar.Count} fatura";
        ws.Range("A2:R2").Merge();
        ws.Cell("A2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell("A2").Style.Font.Italic = true;

        int kesinSay    = kayitlar.Count(x => x.LogoDurum == LogoIslemDurumu.Kesin && !x.RedEdildi);
        int modifiyeSay = kayitlar.Count(x => x.LogoDurum == LogoIslemDurumu.Modifiye && !x.RedEdildi);
        int supheliSay  = kayitlar.Count(x => x.LogoDurum == LogoIslemDurumu.Supheli && !x.RedEdildi);
        int yokSay      = kayitlar.Count(x => x.LogoDurum == LogoIslemDurumu.Yok && !x.RedEdildi);
        int redSay      = kayitlar.Count(x => x.RedEdildi);

        ws.Cell("A3").Value = $"✅ Kesin: {kesinSay}    ✅ Modifiye: {modifiyeSay}    ⚠️ İade Kesilenler/Şüpheli: {supheliSay}    ❌ İşlenmedi: {yokSay}    🚫 Red/İptal: {redSay}";
        ws.Range("A3:R3").Merge();
        ws.Cell("A3").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        ws.Cell("A3").Style.Font.Bold = true;

        int row = 5;
        string[] basliklar = {
            "Tarih", "e-Fatura No", "VKN/TCKN", "Gönderen",
            "Logo Cari Kod", "Logo Cari Ünvan", "Logo Cari Özel Kod",
            "Matrah", "KDV", "Tevkifat", "Tutar (KDV Dahil)",
            "Logo Durumu", "Logo FICHENO", "Logo Tarih", "Logo Tutar",
            "Red/İptal", "Red Tarihi", "Red Açıklama"
        };
        for (int i = 0; i < basliklar.Length; i++)
        {
            var c = ws.Cell(row, i + 1);
            c.Value = basliklar[i];
            c.Style.Font.Bold = true;
            c.Style.Fill.BackgroundColor = XLColor.FromHtml("#E3F2FD");
            c.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            c.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        }
        row++;

        foreach (var k in kayitlar)
        {
            ws.Cell(row, 1).Value  = k.Tarih;
            ws.Cell(row, 1).Style.DateFormat.Format = "dd.MM.yyyy";
            ws.Cell(row, 2).Value  = k.FaturaNo;
            ws.Cell(row, 3).Value  = k.KarsiVKN;
            ws.Cell(row, 4).Value  = k.KarsiUnvan;
            ws.Cell(row, 5).Value  = k.LogoCariKod;
            ws.Cell(row, 6).Value  = k.LogoCariUnvan;
            ws.Cell(row, 7).Value  = k.LogoCariOzelKod;
            ws.Cell(row, 8).Value  = k.Matrah;
            ws.Cell(row, 9).Value  = k.KdvTutar;
            ws.Cell(row, 10).Value = k.TevkifatTutar;
            ws.Cell(row, 11).Value = EFaturaService.GosterilecekTutar(k);
            ws.Cell(row, 12).Value = DurumMetni(k);
            ws.Cell(row, 13).Value = k.LogoFicheno;
            if (k.LogoTarih.HasValue)
            {
                ws.Cell(row, 14).Value = k.LogoTarih.Value;
                ws.Cell(row, 14).Style.DateFormat.Format = "dd.MM.yyyy";
            }
            ws.Cell(row, 15).Value = k.LogoTutar;

            if (k.RedEdildi)
            {
                ws.Cell(row, 16).Value = "🚫 RED/İPTAL";
                ws.Cell(row, 16).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFCDD2");
                ws.Cell(row, 16).Style.Font.FontColor = XLColor.FromHtml("#B71C1C");
                ws.Cell(row, 16).Style.Font.Bold = true;
                ws.Cell(row, 16).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            }
            if (k.RedTarihi.HasValue)
            {
                ws.Cell(row, 17).Value = k.RedTarihi.Value;
                ws.Cell(row, 17).Style.DateFormat.Format = "dd.MM.yyyy";
            }
            ws.Cell(row, 18).Value = k.RedAciklama;

            for (int col = 8; col <= 11; col++)
                ws.Cell(row, col).Style.NumberFormat.Format = "#,##0.00";
            ws.Cell(row, 15).Style.NumberFormat.Format = "#,##0.00";

            var (bg, fg) = DurumRengi(k);
            ws.Cell(row, 12).Style.Fill.BackgroundColor = bg;
            ws.Cell(row, 12).Style.Font.FontColor = fg;
            ws.Cell(row, 12).Style.Font.Bold = true;
            ws.Cell(row, 12).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            for (int col = 1; col <= sonSutun; col++)
                ws.Cell(row, col).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

            row++;
        }

        ws.SheetView.FreezeRows(5);
        if (ws.RangeUsed() is { } kullanilan) kullanilan.SetAutoFilter();
        ws.Columns().AdjustToContents();
        if (ws.Column(4).Width  > 45) ws.Column(4).Width  = 45;
        if (ws.Column(6).Width  > 45) ws.Column(6).Width  = 45;

        // Boş sayfaları da koru ama "Veri yok" notu ekle
        if (kayitlar.Count == 0)
        {
            ws.Cell(6, 1).Value = "(Bu kategoride kayıt bulunmuyor)";
            ws.Range(6, 1, 6, sonSutun).Merge();
            ws.Cell(6, 1).Style.Font.Italic = true;
            ws.Cell(6, 1).Style.Font.FontColor = XLColor.Gray;
            ws.Cell(6, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
        }
    }

    // RED/İptal edilen fatura, LogoDurum'u "Yok" olsa bile asla "İşlenmedi" gösterilmez.
    private static string DurumMetni(EFaturaListItem k)
    {
        if (k.RedEdildi) return "🚫 Red/İptal";
        return k.LogoDurum switch
        {
            LogoIslemDurumu.Kesin    => "✅ Kesin",
            LogoIslemDurumu.Modifiye => "✅ Modifiye",
            LogoIslemDurumu.Supheli  => "⚠️ İade Kesilenler/Şüpheli",
            LogoIslemDurumu.Yok      => "❌ İşlenmedi",
            _                        => "?"
        };
    }

    private static (XLColor Bg, XLColor Fg) DurumRengi(EFaturaListItem k)
    {
        if (k.RedEdildi) return (XLColor.FromHtml("#ECEFF1"), XLColor.FromHtml("#263238"));
        return k.LogoDurum switch
        {
            LogoIslemDurumu.Kesin    => (XLColor.FromHtml("#C8E6C9"), XLColor.FromHtml("#1B5E20")),
            LogoIslemDurumu.Modifiye => (XLColor.FromHtml("#E3F2FD"), XLColor.FromHtml("#0D47A1")),
            LogoIslemDurumu.Supheli  => (XLColor.FromHtml("#FFF8E1"), XLColor.FromHtml("#E65100")),
            LogoIslemDurumu.Yok      => (XLColor.FromHtml("#FFEBEE"), XLColor.FromHtml("#B71C1C")),
            _                        => (XLColor.White, XLColor.Black)
        };
    }
}
