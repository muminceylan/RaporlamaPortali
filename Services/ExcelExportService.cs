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
}
