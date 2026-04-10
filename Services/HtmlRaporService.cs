using RaporlamaPortali.Models;
using System.Text;
using System.Globalization;

namespace RaporlamaPortali.Services;

/// <summary>
/// HTML formatında mail içeriği oluşturan servis
/// </summary>
public class HtmlRaporService
{
    /// <summary>
    /// Logo dosyasını base64 olarak okur. Dosya bulunamazsa boş string döner.
    /// Logoyu publish/wwwroot/images/dogus-logo.png konumuna kaydedin.
    /// </summary>
    private static string LogoBase64Oku()
    {
        try
        {
            var klasor = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "images");
            if (Directory.Exists(klasor))
            {
                // Klasördeki ilk PNG dosyasını kullan (dosya adından bağımsız)
                var png = Directory.GetFiles(klasor, "*.png").FirstOrDefault();
                if (png != null)
                    return Convert.ToBase64String(File.ReadAllBytes(png));
            }
        }
        catch { }
        return "";
    }

    private static string LogoImgHtml()
    {
        var b64 = LogoBase64Oku();
        if (!string.IsNullOrEmpty(b64))
            return $"<img src='data:image/png;base64,{b64}' alt='Dogus' style='height:36px;width:auto;vertical-align:middle;margin-right:8px;'>";
        return "<span style='font-weight:900;color:#fff;vertical-align:middle;margin-right:8px;font-size:14px;'>DOĞUŞ</span>";
    }

    /// <summary>
    /// Yan Ürünler + Şeker raporlarını birleşik, mobil uyumlu HTML olarak oluşturur (mail için).
    /// Tüm Yan Ürünler TEK tabloda — sütunlar garantili hizalı.
    /// </summary>
    public string BirlesikRaporHtmlOlustur(
        List<SekerSatisOzet> sekerVerileri,
        List<YanUrunOzet> yanUrunVerileri,
        DateTime baslangic,
        DateTime bitis)
    {
        var sb = new StringBuilder();
        var logo = LogoImgHtml();

        // Tablo CSS stilleri (sadece veri tabloları için - dış wrapper HTML table kullanır)
        sb.AppendLine($@"<!DOCTYPE html>
<html lang='tr'>
<head>
  <meta charset='UTF-8'>
  <meta name='viewport' content='width=device-width,initial-scale=1'>
  <style>
    body{{font-family:'Segoe UI',Arial,sans-serif;background:#fff;padding:8px;color:#222;margin:0}}
    .dt{{border-collapse:collapse;font-size:11px;table-layout:fixed}}
    .dt th{{background:#37474f;color:#fff;padding:4px 6px;text-align:right;border:1px solid #263238;font-size:10px;white-space:nowrap}}
    .dt th.L{{text-align:left}}
    .dt th.G{{background:#2e7d32}}
    .dt th.C{{background:#00838f}}
    .dt th.M{{background:#6a1b9a}}
    .dt .colhdr th{{background:#455a64}}
    .dt td{{padding:3px 6px;border:1px solid #ddd;text-align:right;font-size:11px;white-space:nowrap}}
    .dt td.L{{text-align:left;font-weight:600}}
    .dt tr:nth-child(even){{background:#f9f9f9}}
    .dt .grp td{{font-weight:700;font-size:11px;color:#fff;padding:5px 8px;text-align:left;border:none}}
    .dt .tot{{background:#fff9c4!important;font-weight:700}}
    .dt .tot td{{color:#1b5e20;border-color:#f9a825}}
    .cv{{color:#00838f}}
    .mv{{color:#6a1b9a}}
    .neg{{color:#c62828;font-weight:700}}
  </style>
</head>
<body>
<table cellpadding='0' cellspacing='0' border='0' style='border-radius:4px;overflow:hidden;background:#fff'>
  <tr>
    <td style='background:#b71c1c;padding:10px 14px'>
      <table cellpadding='0' cellspacing='0' border='0'>
        <tr>
          <td style='vertical-align:middle;padding-right:10px'>{logo}</td>
          <td style='vertical-align:middle'>
            <div style='color:#fff;font-size:13px;font-weight:700;line-height:1.4'>YAN ÜRÜNLER ve ŞEKER ÜRETİM–SATIŞ–STOK RAPORU</div>
            <div style='color:rgba(255,255,255,0.82);font-size:10px;margin-top:2px'>Doğuş Çay – Afyon Şeker Fabrikası &nbsp;|&nbsp; {baslangic:dd.MM.yyyy} – {bitis:dd.MM.yyyy} &nbsp;|&nbsp; Oluşturma: {DateTime.Now:dd.MM.yyyy HH:mm}</div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td style='background:#fff8e1;border-left:4px solid #f9a825;padding:6px 12px;font-size:11px;color:#666'>
      Bu mail <strong>Mümin CEYLAN</strong> tarafından geliştirilen otomasyon ile otomatik olarak gönderilmiştir.
    </td>
  </tr>
  <tr>
    <td style='background:#1b5e20;padding:7px 0;text-align:center;color:#fff;font-size:13px;font-weight:700;letter-spacing:0.5px'>
      🌿 &nbsp; YAN ÜRÜNLER
    </td>
  </tr>
  <tr>
    <td>");

        // ── YAN ÜRÜNLER — TEK TABLO ───────────────────────────
        sb.AppendLine(@"<table class='dt' cellpadding='0' cellspacing='0'>
  <colgroup>
    <col style='width:150px'><col style='width:72px'><col style='width:72px'><col style='width:72px'>
    <col style='width:72px'><col style='width:62px'><col style='width:62px'><col style='width:72px'>
    <col style='width:118px'><col style='width:82px'><col style='width:72px'>
  </colgroup>");

        var yanGruplar = new (string Baslik, string Renk, string Filtre)[]
        {
            ("🍯 MELAS",       "#78350f", "MELAS"),
            ("🌿 YAŞ KÜSPE",   "#166534", "YAS_KUSPE"),
            ("🌾 KURU KÜSPE",  "#92400e", "KURU_KUSPE"),
            ("🧪 ETİL ALKOL",  "#1e3a5f", "ALKOL"),
            ("📦 DİĞER",       "#374151", "DIGER"),
        };

        foreach (var (baslik, renk, filtre) in yanGruplar)
        {
            var grup = yanUrunVerileri.Where(x => x.Kategori == filtre && x.MalzemeKodu != "Y_100153").ToList();
            if (!grup.Any()) continue;

            sb.AppendLine($"  <tr class='grp'><td colspan='11' style='background:{renk}'>{baslik}</td></tr>");
            sb.AppendLine("  <tr class='colhdr'><th class='L'>Ürün</th><th>Devir (T)</th><th>Üretim (T)</th><th>S.Alma (T)</th><th>Satış (T)</th><th>İade (T)</th><th>SA.İade (T)</th><th>Tüketim (T)</th><th class='C'>Tutar (₺)</th><th class='M'>Ort.Fyt (₺/kg)</th><th class='G'>Stok (T)</th></tr>");
            foreach (var u in grup)
                sb.AppendLine($"  <tr><td class='L'>{u.MalzemeAdi}</td><td>{u.DevirStokTon:N2}</td><td>{u.UretimTon:N2}</td><td>{u.SatinAlmaTon:N2}</td><td>{u.SatisTon:N2}</td><td>{u.IadeTon:N2}</td><td>{u.SatinAlmaIadeTon:N2}</td><td>{u.TuketimTon:N2}</td><td class='cv'>{u.SatisTutari:N2}</td><td class='mv'>{(u.OrtalamaFiyat>0?u.OrtalamaFiyat.ToString("N2"):"–")}</td><td class='{(u.StokTon<0?"neg":"")}'>{u.StokTon:N2}</td></tr>");
            sb.AppendLine($"  <tr class='tot'><td class='L'>TOPLAM</td><td>{grup.Sum(x=>x.DevirStokTon):N2}</td><td>{grup.Sum(x=>x.UretimTon):N2}</td><td>{grup.Sum(x=>x.SatinAlmaTon):N2}</td><td>{grup.Sum(x=>x.SatisTon):N2}</td><td>{grup.Sum(x=>x.IadeTon):N2}</td><td>{grup.Sum(x=>x.SatinAlmaIadeTon):N2}</td><td>{grup.Sum(x=>x.TuketimTon):N2}</td><td class='cv'>{grup.Sum(x=>x.SatisTutari):N2}</td><td>–</td><td>{grup.Sum(x=>x.StokTon):N2}</td></tr>");
        }

        sb.AppendLine("</table>");
        sb.AppendLine("    </td>");
        sb.AppendLine("  </tr>");

        // ── ŞEKER AYRAÇ + BAŞLIK ──────────────────────────────
        sb.AppendLine(@"  <tr><td style='height:12px;background:#fff;border-top:3px solid #0d47a1'></td></tr>");
        sb.AppendLine(@"  <tr>
    <td style='background:#0d47a1;padding:8px 0;text-align:center;color:#fff;font-size:13px;font-weight:700;letter-spacing:0.5px'>
      🏭 &nbsp; ŞEKER ÜRETİM–SATIŞ–STOK TABLOSU
    </td>
  </tr>
  <tr><td>");

        // ── ŞEKER TABLOSU ─────────────────────────────────────
        sb.AppendLine(@"<table class='dt' cellpadding='0' cellspacing='0'>
  <colgroup>
    <col style='width:150px'><col style='width:72px'><col style='width:72px'><col style='width:62px'>
    <col style='width:62px'><col style='width:62px'><col style='width:72px'><col style='width:62px'>
    <col style='width:62px'><col style='width:72px'><col style='width:118px'><col style='width:82px'>
  </colgroup>");
        sb.AppendLine("  <tr class='colhdr'><th class='L'>Kategori</th><th>Devir (T)</th><th>Üretim (T)</th><th>S.Alma (T)</th><th>S.İade (T)</th><th>SA.İade (T)</th><th>Satış (T)</th><th>Promo (T)</th><th>Sarf (T)</th><th class='G'>Stok (T)</th><th class='C'>Tutar (₺)</th><th class='M'>Ort.Fyt (₺/kg)</th></tr>");

        foreach (var s in sekerVerileri)
            sb.AppendLine($"  <tr><td class='L'>{s.KategoriAdi}</td><td>{s.DevirStokTon:N2}</td><td>{s.UretimTon:N2}</td><td>{s.SatinAlmaTon:N2}</td><td>{s.IadeTon:N2}</td><td>{s.SatinAlmaIadeTon:N2}</td><td>{s.SatisTon:N2}</td><td>{s.PromosyonTon:N2}</td><td>{s.SarfTon:N2}</td><td class='{(s.StokTon<0?"neg":"")}'>{s.StokTon:N2}</td><td class='cv'>{s.SatisTutari:N2}</td><td class='mv'>{s.OrtalamaFiyat:N2}</td></tr>");

        var tD=sekerVerileri.Sum(x=>x.DevirStokTon); var tU=sekerVerileri.Sum(x=>x.UretimTon);
        var tSA=sekerVerileri.Sum(x=>x.SatinAlmaTon); var tI=sekerVerileri.Sum(x=>x.IadeTon);
        var tSAI=sekerVerileri.Sum(x=>x.SatinAlmaIadeTon); var tS=sekerVerileri.Sum(x=>x.SatisTon);
        var tP=sekerVerileri.Sum(x=>x.PromosyonTon); var tSarf=sekerVerileri.Sum(x=>x.SarfTon);
        var tSt=sekerVerileri.Sum(x=>x.StokTon); var tT=sekerVerileri.Sum(x=>x.SatisTutari);
        var tM=sekerVerileri.Sum(x=>x.SatisMiktari); var tF=tM>0?tT/tM:0;
        sb.AppendLine($"  <tr class='tot'><td class='L'>TOPLAM</td><td>{tD:N2}</td><td>{tU:N2}</td><td>{tSA:N2}</td><td>{tI:N2}</td><td>{tSAI:N2}</td><td>{tS:N2}</td><td>{tP:N2}</td><td>{tSarf:N2}</td><td>{tSt:N2}</td><td class='cv'>{tT:N2}</td><td class='mv'>{tF:N2}</td></tr>");
        sb.AppendLine("</table>");

        sb.AppendLine($@"    </td>
  </tr>
  <tr>
    <td style='padding:5px 12px;font-size:10px;color:#aaa;border-top:1px solid #eee;text-align:center'>
      Stok = Devir + Üretim + S.Alma + İade – SA.İade – Satış – Promo – Sarf &nbsp;|&nbsp; Devir: 01.09.2025 &nbsp;|&nbsp; &copy; {DateTime.Now.Year} Doğuş Çay
    </td>
  </tr>
</table>
</body>
</html>");

        return sb.ToString();
    }

    /// <summary>
    /// Sadece şeker satış raporunu HTML tablo olarak oluşturur (web arayüzü için)
    /// </summary>
    public string SekerSatisHtmlOlustur(List<SekerSatisOzet> veriler, DateTime baslangic, DateTime bitis)
    {
        var sb = new StringBuilder();

        sb.AppendLine(@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); max-width: 1400px; }
        h1 { color: #059669; margin-bottom: 5px; font-size: 24px; }
        h2 { color: #4a5568; margin-top: 0; font-size: 16px; font-weight: normal; }
        .tarih-bilgi { color: #666; font-size: 14px; margin-bottom: 20px; padding: 10px; background-color: #f0fdf4; border-radius: 4px; border-left: 4px solid #059669; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 13px; }
        th { background-color: #4a5568; color: white; padding: 12px 8px; text-align: right; border: 1px solid #2d3748; }
        th:first-child { text-align: left; }
        td { padding: 10px 8px; border: 1px solid #e2e8f0; text-align: right; font-family: Consolas, monospace; }
        td:first-child { text-align: left; font-weight: 600; font-family: 'Segoe UI', sans-serif; }
        tr:nth-child(even) { background-color: #f8fafc; }
        tr:hover { background-color: #f0f9ff; }
        .toplam-satir { background-color: #fef08a !important; font-weight: 700; }
        .toplam-satir td { border-color: #ca8a04; color: #059669; }
        .stok-header { background-color: #059669 !important; }
        .tutar-header { background-color: #0284c7 !important; }
        .fiyat-header { background-color: #7c3aed !important; }
        .negatif { color: #dc2626; font-weight: 600; }
        .tutar-deger { color: #0284c7; }
        .fiyat-deger { color: #7c3aed; }
        .footer { margin-top: 20px; padding-top: 15px; border-top: 1px solid #e2e8f0; color: #666; font-size: 12px; }
    </style>
</head>
<body>
    <div class='container'>");

        sb.AppendLine($@"
        <h1>SEKER URETIM - SATIS - STOK RAPORU</h1>
        <h2>Dogus Cay - Afyon Seker Fabrikasi</h2>

        <div class='tarih-bilgi'>
            <strong>Rapor Donemi:</strong> {baslangic:dd.MM.yyyy} - {bitis:dd.MM.yyyy}<br>
            <strong>Olusturulma Zamani:</strong> {DateTime.Now:dd.MM.yyyy HH:mm}
        </div>");

        // şeker tablosu satırları
        var toplamDevir = veriler.Sum(x => x.DevirStokTon);
        var toplamUretim = veriler.Sum(x => x.UretimTon);
        var toplamSatinAlma = veriler.Sum(x => x.SatinAlmaTon);
        var toplamIade = veriler.Sum(x => x.IadeTon);
        var toplamSatinAlmaIade = veriler.Sum(x => x.SatinAlmaIadeTon);
        var toplamSatis = veriler.Sum(x => x.SatisTon);
        var toplamPromosyon = veriler.Sum(x => x.PromosyonTon);
        var toplamSarf = veriler.Sum(x => x.SarfTon);
        var toplamStok = veriler.Sum(x => x.StokTon);
        var toplamSatisTutari = veriler.Sum(x => x.SatisTutari);

        sb.AppendLine(@"<table><thead><tr>
            <th style='text-align:left'>KATEGORİ</th><th>Devir</th><th>Üretim</th><th>S.Alma</th><th>S.İade</th><th>SA.İade</th><th>Satış</th><th>Promo</th><th>Sarf</th>
            <th class='stok-header'>Stok</th><th class='tutar-header'>Tutar (₺)</th><th class='fiyat-header'>Ort.Fyt (₺/kg)</th>
        </tr></thead><tbody>");

        foreach (var s in veriler)
        {
            var cls = s.StokTon < 0 ? "negatif" : "";
            sb.AppendLine($"<tr><td style='text-align:left;font-weight:600'>{s.KategoriAdi}</td><td>{s.DevirStokTon:N2}</td><td>{s.UretimTon:N2}</td><td>{s.SatinAlmaTon:N2}</td><td>{s.IadeTon:N2}</td><td>{s.SatinAlmaIadeTon:N2}</td><td>{s.SatisTon:N2}</td><td>{s.PromosyonTon:N2}</td><td>{s.SarfTon:N2}</td><td class='{cls}'>{s.StokTon:N2}</td><td class='tutar-deger'>{s.SatisTutari:N2}</td><td class='fiyat-deger'>{s.OrtalamaFiyat:N2}</td></tr>");
        }

        sb.AppendLine($"<tr class='toplam-satir'><td style='text-align:left'>TOPLAM</td><td>{toplamDevir:N2}</td><td>{toplamUretim:N2}</td><td>{toplamSatinAlma:N2}</td><td>{toplamIade:N2}</td><td>{toplamSatinAlmaIade:N2}</td><td>{toplamSatis:N2}</td><td>{toplamPromosyon:N2}</td><td>{toplamSarf:N2}</td><td>{toplamStok:N2}</td><td class='tutar-deger'>{toplamSatisTutari:N2}</td><td>-</td></tr>");
        sb.AppendLine("</tbody></table>");

        sb.AppendLine(@"
        <div class='footer'>
            <em>Bu rapor otomatik olarak olusturulmustur.</em>
        </div>
    </div>
</body>
</html>");

        return sb.ToString();
    }

    /// <summary>
    /// Yan ürünler raporunu HTML tablo olarak oluşturur (web arayüzü için)
    /// </summary>
    public string YanUrunlerHtmlOlustur(List<YanUrunOzet> veriler, DateTime baslangic, DateTime bitis)
    {
        var sb = new StringBuilder();

        sb.AppendLine(@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); max-width: 1200px; }
        h1 { color: #059669; margin-bottom: 5px; font-size: 24px; }
        h2 { color: #4a5568; margin-top: 0; font-size: 16px; font-weight: normal; }
        .tarih-bilgi { color: #666; font-size: 14px; margin-bottom: 20px; padding: 10px; background-color: #f0fdf4; border-radius: 4px; border-left: 4px solid #059669; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 13px; }
        th { background-color: #4a5568; color: white; padding: 10px 8px; text-align: right; border: 1px solid #2d3748; }
        th:first-child { text-align: left; }
        td { padding: 8px; border: 1px solid #e2e8f0; text-align: right; font-family: Consolas, monospace; }
        td:first-child { text-align: left; font-weight: 600; font-family: 'Segoe UI', sans-serif; }
        tr:nth-child(even) { background-color: #f8fafc; }
        .toplam-satir { background-color: #fef08a !important; font-weight: 700; }
        .toplam-satir td { border-color: #ca8a04; color: #059669; }
        .footer { margin-top: 20px; padding-top: 15px; border-top: 1px solid #e2e8f0; color: #666; font-size: 12px; }
    </style>
</head>
<body>
    <div class='container'>");

        sb.AppendLine($@"
        <h1>YAN URUNLER RAPORU</h1>
        <h2>Dogus Cay - Afyon Seker Fabrikasi</h2>

        <div class='tarih-bilgi'>
            <strong>Rapor Donemi:</strong> {baslangic:dd.MM.yyyy} - {bitis:dd.MM.yyyy}<br>
            <strong>Olusturulma Zamani:</strong> {DateTime.Now:dd.MM.yyyy HH:mm}
        </div>");

        // kategorilere göre grupla
        var gruplar = new (string Baslik, string Filtre)[]
        {
            ("MELAS", "MELAS"), ("YAŞ KÜSPE", "YAS_KUSPE"), ("KURU KÜSPE", "KURU_KUSPE"),
            ("ETİL ALKOL", "ALKOL"), ("DİĞER", "DIGER")
        };

        foreach (var (baslik, filtre) in gruplar)
        {
            var grup = veriler.Where(x => x.Kategori == filtre).ToList();
            if (!grup.Any()) continue;

            sb.AppendLine($"<h3 style='color:#059669;border-bottom:2px solid #059669;padding-bottom:4px;margin:20px 0 8px'>{baslik}</h3>");
            sb.AppendLine("<table><thead><tr>");
            sb.AppendLine("<th style='text-align:left'>Malzeme</th><th>Devir (T)</th><th>Üretim (T)</th><th>S.Alma (T)</th><th>Satış (T)</th><th>İade (T)</th><th>SA.İade (T)</th><th>Tüketim (T)</th><th style='background:#0284c7'>Tutar (₺)</th><th style='background:#7c3aed'>Ort.Fyt (₺/kg)</th><th style='background:#059669'>Stok (T)</th>");
            sb.AppendLine("</tr></thead><tbody>");

            foreach (var u in grup)
            {
                var cls = u.StokTon < 0 ? "color:#dc2626;font-weight:700" : "";
                sb.AppendLine($"<tr><td style='text-align:left;font-weight:600'>{u.MalzemeAdi}</td><td>{u.DevirStokTon:N2}</td><td>{u.UretimTon:N2}</td><td>{u.SatinAlmaTon:N2}</td><td>{u.SatisTon:N2}</td><td>{u.IadeTon:N2}</td><td>{u.SatinAlmaIadeTon:N2}</td><td>{u.TuketimTon:N2}</td><td style='color:#0284c7'>{u.SatisTutari:N2}</td><td style='color:#7c3aed'>{(u.OrtalamaFiyat>0?u.OrtalamaFiyat.ToString("N2"):"–")}</td><td style='{cls}'>{u.StokTon:N2}</td></tr>");
            }

            sb.AppendLine($"<tr style='background:#fef08a;font-weight:700'><td style='text-align:left'>TOPLAM</td><td>{grup.Sum(x=>x.DevirStokTon):N2}</td><td>{grup.Sum(x=>x.UretimTon):N2}</td><td>{grup.Sum(x=>x.SatinAlmaTon):N2}</td><td>{grup.Sum(x=>x.SatisTon):N2}</td><td>{grup.Sum(x=>x.IadeTon):N2}</td><td>{grup.Sum(x=>x.SatinAlmaIadeTon):N2}</td><td>{grup.Sum(x=>x.TuketimTon):N2}</td><td style='color:#0284c7'>{grup.Sum(x=>x.SatisTutari):N2}</td><td>–</td><td>{grup.Sum(x=>x.StokTon):N2}</td></tr>");
            sb.AppendLine("</tbody></table>");
        }

        sb.AppendLine(@"
        <div class='footer'>
            <em>Bu rapor otomatik olarak olusturulmustur.</em>
        </div>
    </div>
</body>
</html>");

        return sb.ToString();
    }

    // =========================================================
    // PANCAR RAPORU HTML
    // =========================================================

    public string PancarRaporHtmlOlustur(
        List<PancarIcmalKayit>  icmal,
        List<PancarCiftciDetay> ciftciler,
        DateTime                tarih,
        List<PancarAvansKayit>? avans  = null,
        PancarFinansOzet?       finans = null)
    {
        var tr   = new CultureInfo("tr-TR");
        var logo = LogoImgHtml();
        var sb   = new StringBuilder();

        sb.AppendLine($@"<!DOCTYPE html>
<html lang='tr'>
<head>
<meta charset='UTF-8'>
<meta name='viewport' content='width=device-width,initial-scale=1'>
<title>Pancar Raporu {tarih:dd.MM.yyyy}</title>
<style>
  body {{ font-family:Arial,sans-serif; font-size:12px; margin:0; padding:8px; background:#f5f5f5; }}
  table {{ border-collapse:collapse; width:100%; table-layout:fixed; }}
  th {{ background:#1a237e; color:#fff; padding:6px 8px; text-align:center; font-size:11px; border:1px solid #0d1457; }}
  td {{ padding:5px 8px; border:1px solid #ccc; text-align:center; }}
  tr:nth-child(even) {{ background:#e8eaf6; }}
  tr:hover {{ background:#c5cae9; }}
  .baslik-bar {{ background:#1a237e; color:#fff; padding:8px 12px; display:inline-block; width:100%; box-sizing:border-box; }}
  .section-title {{ background:#283593; color:#fff; padding:5px 10px; font-weight:bold; font-size:13px; text-align:center; margin:12px 0 0 0; }}
  .ozet-kart {{ display:inline-block; background:#fff; border:2px solid #1a237e; border-radius:6px; padding:8px 16px; margin:4px; text-align:center; min-width:130px; }}
  .ozet-kart .deger {{ font-size:16px; font-weight:bold; color:#1a237e; }}
  .ozet-kart .etiket {{ font-size:10px; color:#666; }}
  .tip-ciftci {{ background:#e8f5e9; }}
  .tip-muteahhit {{ background:#e3f2fd; }}
  .tip-mouse {{ background:#fff3e0; }}
  .tip-kepce {{ background:#fce4ec; }}
  .info {{ font-size:10px; color:#888; margin-top:10px; text-align:right; }}
  .avans-tbl {{ width:480px; border-collapse:collapse; margin:0; }}
  .avans-tbl td {{ padding:5px 12px; border-bottom:1px solid #eee; font-size:12px; }}
  .avans-tbl td:last-child {{ text-align:right; white-space:nowrap; }}
</style>
</head>
<body>
<table style='width:100%;max-width:700px;'><tr><td>");

        // Başlık
        sb.AppendLine($@"<div class='baslik-bar'>
  {logo}
  <span style='font-size:15px;font-weight:bold;'>PANCAR RAPORU — {tarih:dd.MM.yyyy}</span>
  <span style='font-size:11px;margin-left:12px;opacity:.8;'>Afyon Şeker Fabrikası / Kampanya {PancarRaporService.KampanyaYili()}</span>
</div>");

        // İCMAL tablosu
        sb.AppendLine("<div class='section-title'>KANTAR HAREKETLERİ İCMAL</div>");
        sb.AppendLine("<table><thead><tr>");
        sb.AppendLine("<th style='width:15%'>TİP</th><th style='width:45%'>AÇIKLAMA</th><th style='width:20%'>NET (kg)</th><th style='width:20%'>TUTAR (₺)</th>");
        sb.AppendLine("</tr></thead><tbody>");

        if (icmal.Count == 0)
        {
            sb.AppendLine("<tr><td colspan='4' style='text-align:center;color:#888;'>Henüz veri yok</td></tr>");
        }
        else
        {
            foreach (var k in icmal)
            {
                var tipClass = k.Tip switch
                {
                    var t when t.Contains("iftci",   StringComparison.OrdinalIgnoreCase) => "tip-ciftci",
                    var t when t.Contains("teahhit", StringComparison.OrdinalIgnoreCase) => "tip-muteahhit",
                    var t when t.Contains("ouse",    StringComparison.OrdinalIgnoreCase) => "tip-mouse",
                    var t when t.Contains("ep",      StringComparison.OrdinalIgnoreCase) => "tip-kepce",
                    _ => ""
                };
                sb.AppendLine($"<tr class='{tipClass}'>");
                sb.AppendLine($"<td style='font-weight:bold'>{k.Tip}</td>");
                sb.AppendLine($"<td style='text-align:left'>{k.Aciklama}</td>");
                sb.AppendLine($"<td>{k.Net.ToString("N0", tr)}</td>");
                sb.AppendLine($"<td>{(k.Tutar > 0 ? k.Tutar.ToString("N2", tr) + " ₺" : "—")}</td>");
                sb.AppendLine("</tr>");
            }
            var topNet   = icmal.Sum(x => x.Net);
            var topTutar = icmal.Sum(x => x.Tutar);
            sb.AppendLine($@"<tr style='font-weight:bold;background:#1a237e;color:#fff;'>
              <td colspan='2'>GENEL TOPLAM</td>
              <td>{topNet.ToString("N0", tr)}</td>
              <td>{(topTutar > 0 ? topTutar.ToString("N2", tr) + " ₺" : "—")}</td>
            </tr>");
        }
        sb.AppendLine("</tbody></table>");

        // Çiftçi özet istatistik
        if (ciftciler.Count > 0)
        {
            var topTaahhut = ciftciler.Sum(x => x.TaahhutTon);
            var topNet2    = ciftciler.Sum(x => x.NetMiktar);
            var ortPolar   = ciftciler.Where(x => x.OrtalamaPolar > 0).Select(x => (double)x.OrtalamaPolar).DefaultIfEmpty(0).Average();

            sb.AppendLine("<div class='section-title'>ÇİFTÇİ ÖZETİ</div>");
            sb.AppendLine($@"<div style='padding:8px;background:#fff;'>
              <div class='ozet-kart'><div class='deger'>{ciftciler.Count:N0}</div><div class='etiket'>Toplam Çiftçi</div></div>
              <div class='ozet-kart'><div class='deger'>{topTaahhut / 1000:N0} ton</div><div class='etiket'>Taahhüt Tonu</div></div>
              <div class='ozet-kart'><div class='deger'>{topNet2 / 1000:N1} ton</div><div class='etiket'>Gerçekleşen Net</div></div>
              <div class='ozet-kart'><div class='deger'>{ortPolar:N2}</div><div class='etiket'>Ort. Polar</div></div>
            </div>");
        }

        // AVANS TABLOSU
        if (avans != null && finans != null)
        {
            sb.AppendLine("<div class='section-title'>AVANS RAPORU</div>");
            sb.AppendLine("<div style='background:#fff;padding:8px;'>");
            sb.AppendLine("<table class='avans-tbl'>");

            // Başlık satırı
            sb.AppendLine("<tr style='background:#C62828;color:white;font-weight:bold;'><td>AVANS ADI</td><td>TUTAR</td></tr>");

            var nakdiDict = avans.Where(x => x.KaynakEvrak.Contains("NAKD"))
                                 .ToDictionary(x => x.AvansGrubu.Trim(), x => x.TutarToplami);
            var ayniDict  = avans.Where(x => x.KaynakEvrak.Contains("AYN"))
                                 .ToDictionary(x => x.AvansGrubu.Trim(), x => x.TutarToplami);

            decimal Get(Dictionary<string, decimal> d, string k) =>
                d.TryGetValue(k, out var v) ? v : 0m;

            // NAKDİ satırları
            (string Grup, string Ad)[] nakdiSira = {
                ("Pancar Avansı","Pancar Avansı"),("Hasat Makinesi Avansı","Hasat Makinesi Avansı"),
                ("1. Avans","1. Avans"),("2. Avans","2. Avans"),("3. Avans","3. Avans"),
                ("4. Avans","4. Avans"),("5. Avans","5. Avans"),("6. Avans","6. Avans"),
                ("Küspe","Küspe Avansı"),("Söküm Avansı","Söküm Avansı"),
            };
            foreach (var (grup, ad) in nakdiSira)
                sb.AppendLine($"<tr><td>{ad}</td><td>{Get(nakdiDict, grup).ToString("N2", tr)}</td></tr>");
            foreach (var kv in nakdiDict.Where(kv => !nakdiSira.Any(n => n.Grup == kv.Key)))
                sb.AppendLine($"<tr><td>{kv.Key}</td><td>{kv.Value.ToString("N2", tr)}</td></tr>");

            decimal nakdiToplam = nakdiDict.Values.Sum();
            sb.AppendLine($"<tr style='background:#FDD835;font-weight:bold;'><td>NAKDİ AVANS TOPLAMI</td><td>{nakdiToplam.ToString("N2", tr)}</td></tr>");

            // AYNİ satırları
            (string Grup, string Ad)[] ayniSira = {
                ("Gübre","Gübre Avansı"),("İlaç","İlaç Avansı"),("Tohum","Tohum Avansı"),
                ("Çay","Çay Avansı"),("Şeker","Şeker Avansı"),("Küspe","Küspe Avansı"),
                ("Fatura Edilen Söküm Avansı","Fatura Edilen Söküm Avansı"),("Söküm Avansı","Söküm Avansı"),
            };
            foreach (var (grup, ad) in ayniSira)
                sb.AppendLine($"<tr><td>{ad}</td><td>{Get(ayniDict, grup).ToString("N2", tr)}</td></tr>");
            foreach (var kv in ayniDict.Where(kv => !ayniSira.Any(n => n.Grup == kv.Key)))
                sb.AppendLine($"<tr><td>{kv.Key}</td><td>{kv.Value.ToString("N2", tr)}</td></tr>");

            decimal ayniToplam = ayniDict.Values.Sum();
            sb.AppendLine($"<tr style='background:#388E3C;color:white;font-weight:bold;'><td>AYNİ AVANS TOPLAMI</td><td>{ayniToplam.ToString("N2", tr)}</td></tr>");

            // PANCAR BEDELİ ÖDEMESİ
            sb.AppendLine($"<tr><td>PANCAR BEDELİ ÖDEMESİ</td><td>0,00</td></tr>");
            sb.AppendLine($"<tr><td>Ödenen Kota Fazlası Bedeli</td><td>{finans.KotaFazlasi.ToString("N2", tr)}</td></tr>");
            sb.AppendLine($"<tr><td>Ödenen C Pancar Bedeli</td><td>{finans.CPancari.ToString("N2", tr)}</td></tr>");
            decimal pancarBedeli = finans.KotaFazlasi + finans.CPancari;
            sb.AppendLine($"<tr style='background:#E53935;color:white;font-weight:bold;'><td>PANCAR BEDELİ ÖDEMESİ</td><td>{pancarBedeli.ToString("N2", tr)}</td></tr>");

            // AVANS TOPLAMI
            decimal avansToplami = nakdiToplam + ayniToplam + pancarBedeli;
            sb.AppendLine($"<tr style='background:#B71C1C;color:white;font-weight:bold;font-size:1.05em;'><td>AVANS TOPLAMI</td><td>{avansToplami.ToString("N2", tr)}</td></tr>");

            // Finansal satırlar — KotaCezası = 0
            sb.AppendLine($"<tr style='background:#fffde7;'><td>AVANS KDV Sİ</td><td>{finans.AvansKdv.ToString("N2", tr)}</td></tr>");
            sb.AppendLine($"<tr><td>STOPAJ</td><td>{finans.AlimStopaji.ToString("N2", tr)}</td></tr>");
            sb.AppendLine($"<tr><td>Ödenen Nakliye Primi</td><td>{finans.NakliyePrimi.ToString("N2", tr)}</td></tr>");
            sb.AppendLine($"<tr><td>Kota Cezası</td><td>0,00</td></tr>");
            sb.AppendLine($"<tr><td>Ödenen Bağkur Primi</td><td>{finans.BagkurBorcu.ToString("N2", tr)}</td></tr>");
            sb.AppendLine($"<tr><td>Borsa Tescil</td><td>{finans.BorsaTescil.ToString("N2", tr)}</td></tr>");

            // GENEL TOPLAM — KotaCezası dahil edilmiyor
            decimal genelToplam = avansToplami + finans.AvansKdv + finans.AlimStopaji
                                  + finans.NakliyePrimi + finans.BagkurBorcu + finans.BorsaTescil;
            sb.AppendLine($"<tr style='background:#B71C1C;color:white;font-weight:bold;font-size:1.08em;'><td>AVANSLAR GENEL TOPLAMI</td><td>{genelToplam.ToString("N2", tr)}</td></tr>");

            sb.AppendLine("</table></div>");
        }

        sb.AppendLine($"<p class='info'>Bu mail <strong>Mümin CEYLAN</strong> tarafından geliştirilen otomasyon ile otomatik olarak gönderilmiştir. | {DateTime.Now:dd.MM.yyyy HH:mm}</p>");
        sb.AppendLine("</td></tr></table></body></html>");

        return sb.ToString();
    }
}
