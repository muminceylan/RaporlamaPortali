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
}
