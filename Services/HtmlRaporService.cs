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
        DateTime bitis,
        bool bulanik = false)
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

        return UygulaBlur(sb.ToString(), bulanik);
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
        List<PancarIcmalKayit>   icmal,
        List<PancarCiftciDetay>  ciftciler,
        DateTime                 tarih,
        List<PancarAvansKayit>?  avans       = null,
        PancarFinansOzet?        finans      = null,
        PancarIcmalDetay?        icmalDetay  = null,
        PancarOzetIstatistik?    ozet        = null,
        bool                     bulanik     = false)
    {
        var tr   = new CultureInfo("tr-TR");
        var logo = LogoImgHtml();
        var sb   = new StringBuilder();

        // ── Hesaplamalar ──────────────────────────────────────
        var nakdiDict = avans?.Where(x => x.KaynakEvrak.Contains("NAKD"))
                              .ToDictionary(x => x.AvansGrubu.Trim(), x => x.TutarToplami)
                       ?? new Dictionary<string, decimal>();
        var ayniDict  = avans?.Where(x => x.KaynakEvrak.Contains("AYN"))
                              .ToDictionary(x => x.AvansGrubu.Trim(), x => x.TutarToplami)
                       ?? new Dictionary<string, decimal>();

        decimal GetD(Dictionary<string, decimal> d, string k) =>
            d.TryGetValue(k, out var v) ? v : 0m;

        decimal nakdiToplam  = nakdiDict.Values.Sum();
        decimal ayniToplam   = ayniDict.Values.Sum();
        decimal pancarBedeli = finans != null ? finans.KotaFazlasi + finans.CPancari : 0m;
        decimal avansToplami = nakdiToplam + ayniToplam + pancarBedeli;
        decimal genelToplam  = finans != null
            ? avansToplami + finans.AvansKdv + finans.AlimStopaji + finans.NakliyePrimi + finans.BagkurBorcu + finans.BorsaTescil
            : avansToplami;
        decimal hakedis      = icmalDetay?.HakedisToplamı ?? 0m;
        decimal borcAlacak   = genelToplam - hakedis;

        (string Grup, string Ad)[] nakdiSira = {
            ("Pancar Avansı","Pancar Avansı"),("Hasat Makinesi Avansı","Hasat Makinesi Avansı"),
            ("1. Avans","1. Avans"),("2. Avans","2. Avans"),("3. Avans","3. Avans"),
            ("4. Avans","4. Avans"),("5. Avans","5. Avans"),("6. Avans","6. Avans"),
            ("Küspe","Küspe Avansı"),("Söküm Avansı","Söküm Avansı"),
        };
        (string Grup, string Ad)[] ayniSira = {
            ("Gübre","Gübre Avansı"),("İlaç","İlaç Avansı"),("Tohum","Tohum Avansı"),
            ("Çay","Çay Avansı"),("Şeker","Şeker Avansı"),("Küspe","Küspe Avansı"),
            ("Fatura Edilen Söküm Avansı","Fatura Edilen Söküm Avansı"),("Söküm Avansı","Söküm Avansı"),
        };

        // ── HTML ──────────────────────────────────────────────
        sb.AppendLine($@"<!DOCTYPE html>
<html lang='tr'>
<head>
<meta charset='UTF-8'>
<title>Pancar Raporu {tarih:dd.MM.yyyy}</title>
<style>
  body {{ font-family:Arial,sans-serif; font-size:12px; margin:0; padding:6px; background:#f5f5f5; }}
  .baslik-bar {{ background:#1a237e; color:#fff; padding:8px 12px; width:100%; box-sizing:border-box; margin-bottom:8px; }}
  .section-title {{ background:#283593; color:#fff; padding:4px 10px; font-weight:bold; font-size:12px; margin:10px 0 0 0; }}
  .ozet-wrap {{ display:flex; flex-wrap:wrap; gap:6px; padding:6px; background:#fff; margin-bottom:4px; }}
  .ozet-kart {{ background:#e8eaf6; border-radius:4px; padding:6px 12px; text-align:center; min-width:100px; }}
  .ozet-kart .deger {{ font-size:14px; font-weight:bold; color:#1a237e; }}
  .ozet-kart .etiket {{ font-size:10px; color:#555; }}
  .iki-sutun {{ display:flex; gap:8px; align-items:flex-start; }}
  .sutun {{ flex:1; min-width:0; }}
  .at {{ border-collapse:collapse; width:100%; font-size:11px; }}
  .at td, .at th {{ padding:4px 8px; border:1px solid #bbb; }}
  .at td:last-child, .at th:last-child {{ text-align:right; white-space:nowrap; }}
  .info {{ font-size:10px; color:#888; margin-top:8px; text-align:right; }}
</style>
</head>
<body>");

        // Başlık
        sb.AppendLine($@"<div class='baslik-bar'>
  {logo}
  <span style='font-size:14px;font-weight:bold;'>PANCAR RAPORU — {tarih:dd.MM.yyyy}</span>
  <span style='font-size:11px;margin-left:12px;opacity:.8;'>Afyon Şeker Fabrikası / Kampanya {PancarRaporService.KampanyaYili()}</span>
</div>");

        // Özet kartlar
        if (ozet != null || icmalDetay != null)
        {
            sb.AppendLine("<div class='ozet-wrap'>");
            if (ozet != null)
            {
                sb.AppendLine($"<div class='ozet-kart'><div class='deger'>{ozet.ToplamCiftci:N0}</div><div class='etiket'>Toplam Çiftçi</div></div>");
                sb.AppendLine($"<div class='ozet-kart'><div class='deger'>{ozet.ToplamTaahhut/1000:N0} ton</div><div class='etiket'>Taahhüt</div></div>");
            }
            if (icmalDetay != null)
                sb.AppendLine($"<div class='ozet-kart'><div class='deger'>{icmalDetay.NetMiktarTon:N1} ton</div><div class='etiket'>Gelen Net</div></div>");
            if (ozet != null)
            {
                sb.AppendLine($"<div class='ozet-kart'><div class='deger'>%{ozet.OrtFireOrani:N2}</div><div class='etiket'>Fire Oranı</div></div>");
                sb.AppendLine($"<div class='ozet-kart'><div class='deger'>%{ozet.OrtPolar:N2}</div><div class='etiket'>Polar</div></div>");
            }
            sb.AppendLine($"<div class='ozet-kart'><div class='deger' style='color:{(borcAlacak > 0 ? "#B71C1C" : "#1B5E20")}'>{Math.Abs(borcAlacak):N0} ₺</div><div class='etiket'>{(borcAlacak > 0 ? "Müstahsil Borçlu" : "Müstahsil Alacaklı")}</div></div>");
            sb.AppendLine("</div>");
        }

        // İki sütun yan yana
        sb.AppendLine("<div class='iki-sutun'>");

        // SOL: İCMAL (avans tablosu)
        sb.AppendLine("<div class='sutun'>");
        sb.AppendLine("<div class='section-title'>İCMAL</div>");
        sb.AppendLine("<table class='at'>");
        sb.AppendLine("<tr style='background:#C62828;color:white;font-weight:bold;'><td>AVANS ADI</td><td>TUTAR (₺)</td></tr>");

        foreach (var (grup, ad) in nakdiSira)
            sb.AppendLine($"<tr><td>{ad}</td><td>{GetD(nakdiDict, grup).ToString("N2", tr)}</td></tr>");
        foreach (var kv in nakdiDict.Where(kv => !nakdiSira.Any(n => n.Grup == kv.Key)))
            sb.AppendLine($"<tr><td>{kv.Key}</td><td>{kv.Value.ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr style='background:#FDD835;font-weight:bold;'><td>NAKDİ AVANS TOPLAMI</td><td>{nakdiToplam.ToString("N2", tr)}</td></tr>");

        foreach (var (grup, ad) in ayniSira)
            sb.AppendLine($"<tr><td>{ad}</td><td>{GetD(ayniDict, grup).ToString("N2", tr)}</td></tr>");
        foreach (var kv in ayniDict.Where(kv => !ayniSira.Any(n => n.Grup == kv.Key)))
            sb.AppendLine($"<tr><td>{kv.Key}</td><td>{kv.Value.ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr style='background:#388E3C;color:white;font-weight:bold;'><td>AYNİ AVANS TOPLAMI</td><td>{ayniToplam.ToString("N2", tr)}</td></tr>");

        sb.AppendLine($"<tr><td>PANCAR BEDELİ ÖDEMESİ</td><td>0,00</td></tr>");
        sb.AppendLine($"<tr><td>Ödenen Kota Fazlası Bedeli</td><td>{(finans?.KotaFazlasi ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>Ödenen C Pancar Bedeli</td><td>{(finans?.CPancari ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr style='background:#E53935;color:white;font-weight:bold;'><td>PANCAR BEDELİ ÖDEMESİ</td><td>{pancarBedeli.ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr style='background:#B71C1C;color:white;font-weight:bold;'><td>AVANS TOPLAMI</td><td>{avansToplami.ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr style='background:#fffde7;'><td>AVANS KDV Sİ</td><td>{(finans?.AvansKdv ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>STOPAJ</td><td>{(finans?.AlimStopaji ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>Ödenen Nakliye Primi</td><td>{(finans?.NakliyePrimi ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>Kota Cezası</td><td>0,00</td></tr>");
        sb.AppendLine($"<tr><td>Ödenen Bağkur Primi</td><td>{(finans?.BagkurBorcu ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>Borsa Tescil</td><td>{(finans?.BorsaTescil ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr style='background:#B71C1C;color:white;font-weight:bold;'><td>TOPLAM MÜSTAHSİL BORCU</td><td>{genelToplam.ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr style='background:#e8eaf6;'><td colspan='2' style='font-size:10px;color:#666;padding:3px 8px;'>▼ Hakediş Özeti</td></tr>");
        sb.AppendLine($"<tr><td>PANCAR BEDELİ</td><td>{(icmalDetay?.PancarBedeliToplam ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>MÜSTAHSİL NAKLİYE</td><td>{(icmalDetay?.MustahsilNakliye ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>KÜSPE PRİMİ</td><td>{(icmalDetay?.KuspePrimi ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>KOTA TAMAMLAMA PRİMİ</td><td>{(icmalDetay?.KotaTamamlamaPrimi ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr style='background:#1565C0;color:white;font-weight:bold;'><td>MÜSTAHSİL HAKEDİŞ TOPLAMI</td><td>{hakedis.ToString("N2", tr)}</td></tr>");
        var borcRenk = borcAlacak > 0 ? "#B71C1C" : "#1B5E20";
        sb.AppendLine($"<tr style='background:{borcRenk};color:white;font-weight:bold;'><td>{(borcAlacak > 0 ? "MÜSTAHSİL BORÇLU" : "MÜSTAHSİL ALACAKLI")}</td><td>{Math.Abs(borcAlacak).ToString("N2", tr)}</td></tr>");
        sb.AppendLine("</table>");
        sb.AppendLine("</div>"); // sol sutun bitti

        // SAĞ: GENEL İCMAL + PANCAR TÜRLERİ + NAKLİYE/MOUSE/KEPÇE
        sb.AppendLine("<div class='sutun'>");

        // Genel İcmal
        sb.AppendLine("<div class='section-title'>GENEL İCMAL</div>");
        sb.AppendLine("<table class='at'>");
        sb.AppendLine("<tr style='background:#283593;color:white;font-weight:bold;'><td>AÇIKLAMA</td><td>MİKTAR (ton)</td><td>TUTAR (₺)</td></tr>");
        sb.AppendLine($"<tr><td>Pancar Bedeli</td><td>{(icmalDetay?.NetMiktarTon ?? 0).ToString("N3", tr)}</td><td>{(icmalDetay?.PancarBedeliToplam ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>Müstahsil Nakliye</td><td>—</td><td>{(icmalDetay?.MustahsilNakliye ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>Küspe Primi</td><td>—</td><td>{(icmalDetay?.KuspePrimi ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>Kota Tamamlama Primi</td><td>—</td><td>{(icmalDetay?.KotaTamamlamaPrimi ?? 0).ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr style='background:#e8eaf6;font-weight:bold;'><td>Toplam Maliyet</td><td>—</td><td>{hakedis.ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>Müstahsile Ödenen</td><td>—</td><td>{genelToplam.ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>Müstahsil Hakedişi</td><td>—</td><td>{hakedis.ToString("N2", tr)}</td></tr>");
        sb.AppendLine($"<tr style='background:{borcRenk};color:white;font-weight:bold;'><td>{(borcAlacak > 0 ? "Müstahsil Borçlu" : "Müstahsil Alacaklı")}</td><td>—</td><td>{Math.Abs(borcAlacak).ToString("N2", tr)}</td></tr>");
        sb.AppendLine("</table>");

        // Pancar Türleri
        sb.AppendLine("<div class='section-title'>PANCAR TÜRLERİ</div>");
        sb.AppendLine("<table class='at'>");
        sb.AppendLine("<tr style='background:#2E7D32;color:white;font-weight:bold;'><td>TÜR</td><td>MİKTAR (ton)</td><td>TUTAR (₺)</td><td>BİRİM FİYAT</td></tr>");
        sb.AppendLine($"<tr><td>A Pancarı</td><td>{(icmalDetay?.APancariTon ?? 0).ToString("N3", tr)}</td><td>{(icmalDetay?.APancariBedeli ?? 0).ToString("N2", tr)}</td><td>{(icmalDetay?.ABirimFiyati ?? 0).ToString("N4", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>C Pancarı</td><td>{(icmalDetay?.CPancariTon ?? 0).ToString("N3", tr)}</td><td>{(icmalDetay?.CPancariBedeli ?? 0).ToString("N2", tr)}</td><td>{(icmalDetay?.CBirimFiyati ?? 0).ToString("N4", tr)}</td></tr>");
        sb.AppendLine($"<tr><td>Kota Fazlası</td><td>{(icmalDetay?.KotaFazlasiTon ?? 0).ToString("N3", tr)}</td><td>{(icmalDetay?.KotaFazlasiBedeli ?? 0).ToString("N2", tr)}</td><td>{(icmalDetay?.KFBirimFiyati ?? 0).ToString("N4", tr)}</td></tr>");
        var topPancarTon = (icmalDetay?.APancariTon ?? 0) + (icmalDetay?.CPancariTon ?? 0) + (icmalDetay?.KotaFazlasiTon ?? 0);
        sb.AppendLine($"<tr style='background:#1B5E20;color:white;font-weight:bold;'><td>TOPLAM</td><td>{topPancarTon.ToString("N3", tr)}</td><td>{(icmalDetay?.PancarBedeliToplam ?? 0).ToString("N2", tr)}</td><td>—</td></tr>");
        sb.AppendLine("</table>");

        // Nakliye / Mouse / Kepçe
        if (icmal.Count > 0)
        {
            sb.AppendLine("<div class='section-title'>NAKLİYE / MOUSE / KEPÇE</div>");
            sb.AppendLine("<table class='at'>");
            sb.AppendLine("<tr style='background:#4527A0;color:white;font-weight:bold;'><td>TİP / AÇIKLAMA</td><td>NET (kg)</td><td>TUTAR (₺)</td><td>ORT. (₺/ton)</td></tr>");
            string? sonTip = null;
            foreach (var k in icmal)
            {
                if (k.Tip != sonTip)
                {
                    sb.AppendLine($"<tr style='background:#ede7f6;font-weight:bold;'><td colspan='4'>{k.Tip}</td></tr>");
                    sonTip = k.Tip;
                }
                var ort = k.Net > 0 ? (k.Tutar / k.Net * 1000).ToString("N2", tr) : "—";
                sb.AppendLine($"<tr><td style='padding-left:16px;'>{k.Aciklama}</td><td>{k.Net.ToString("N0", tr)}</td><td>{k.Tutar.ToString("N2", tr)}</td><td>{ort}</td></tr>");
            }
            sb.AppendLine("</table>");
        }

        sb.AppendLine("</div>"); // sağ sutun bitti
        sb.AppendLine("</div>"); // iki-sutun bitti

        sb.AppendLine($"<p class='info'>Bu mail <strong>Mümin CEYLAN</strong> tarafından geliştirilen otomasyon ile otomatik olarak gönderilmiştir. | {DateTime.Now:dd.MM.yyyy HH:mm}</p>");
        sb.AppendLine("</body></html>");

        return UygulaBlur(sb.ToString(), bulanik);
    }

    /// <summary>
    /// Şeker Kategorisi Bazlı Analiz tablosunu HTML olarak oluşturur (üst tablo – ham LOGO verisi).
    /// </summary>
    public string SekerAnalizHtmlOlustur(
        List<SekerKategoriAnaliz> analiz,
        DateTime baslangic,
        DateTime bitis,
        bool bulanik = false)
    {
        var tr = new CultureInfo("tr-TR");
        string N0(decimal v) => v.ToString("N0", tr);
        string N2(decimal v) => v.ToString("N2", tr);
        string Fiyat(decimal m, decimal t) => m > 0 ? (t / m).ToString("N2", tr) : "–";

        var logo = LogoImgHtml();

        var sb = new StringBuilder();
        sb.AppendLine($@"<!DOCTYPE html>
<html lang='tr'>
<head>
  <meta charset='UTF-8'>
  <style>
    body{{font-family:'Segoe UI',Arial,sans-serif;background:#f5f5f5;padding:10px;margin:0;color:#222}}
    .wrap{{background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.15);width:1580px}}
    .hdr{{background:#2e7d32;padding:10px 14px;color:#fff}}
    .hdr-title{{font-size:13px;font-weight:700}}
    .hdr-sub{{font-size:10px;color:rgba(255,255,255,.8);margin-top:2px}}
    table{{border-collapse:collapse;width:100%;font-size:10px}}
    th{{padding:5px 4px;border:1px solid #263238;text-align:right;white-space:nowrap;font-size:9px}}
    th.L{{text-align:left}}
    th.G{{background:#2e7d32;color:#fff}}
    th.B{{background:#0D47A1;color:#fff}}
    th.R{{background:#b71c1c;color:#fff}}
    th.M{{background:#880E4F;color:#fff}}
    th.P{{background:#7B1FA2;color:#fff}}
    th.O{{background:#E65100;color:#fff}}
    td{{padding:4px;border:1px solid #ddd;text-align:right;font-family:Consolas,monospace;font-size:10px;white-space:nowrap}}
    td.L{{text-align:left;font-weight:600;font-family:'Segoe UI',Arial,sans-serif;white-space:nowrap}}
    td.neg{{color:#c62828;font-weight:700}}
    td.blue{{color:#0D47A1;font-weight:700}}
    td.purple{{color:#7B1FA2;font-weight:700}}
    td.orange{{color:#E65100;font-weight:700}}
    td.red{{color:#b71c1c}}
    td.pink{{color:#880E4F}}
    .alt{{background:#f9f9f9}}
    .tot td{{background:#fff9c4;font-weight:700;color:#1b5e20;border-color:#f9a825}}
    .footer{{padding:5px 10px;font-size:9px;color:#aaa;border-top:1px solid #eee;text-align:center}}
  </style>
</head>
<body>
<div class='wrap'>
  <div class='hdr'>
    <table cellpadding='0' cellspacing='0' border='0'><tr>
      <td style='padding-right:10px;vertical-align:middle'>{logo}</td>
      <td style='vertical-align:middle'>
        <div class='hdr-title'>ŞEKER KATEGORİSİ BAZLI ANALİZ (Ham LOGO Verisi)</div>
        <div class='hdr-sub'>Doğuş Çay – Afyon Şeker Fabrikası &nbsp;|&nbsp; {baslangic:dd.MM.yyyy} – {bitis:dd.MM.yyyy} &nbsp;|&nbsp; {DateTime.Now:dd.MM.yyyy HH:mm}</div>
      </td>
    </tr></table>
  </div>
  <table>
    <colgroup>
      <col style='width:145px'>
      <col style='width:75px'><col style='width:75px'><col style='width:70px'><col style='width:70px'><col style='width:68px'><col style='width:80px'>
      <col style='width:75px'><col style='width:88px'><col style='width:88px'>
      <col style='width:70px'><col style='width:62px'><col style='width:72px'><col style='width:68px'><col style='width:72px'><col style='width:80px'>
      <col style='width:82px'><col style='width:72px'>
    </colgroup>
    <thead>
      <tr>
        <th class='L G'>KATEGORİ</th>
        <th class='G'>Dönem Başı (Kg)</th>
        <th class='B'>Üretim (Kg)</th>
        <th class='B'>Satın Alma (Kg)</th>
        <th class='B'>Satış İade (Kg)</th>
        <th class='B'>Diğer Giriş (Kg)</th>
        <th class='B' style='background:#0D47A1;font-weight:900'>Top. Giriş (Kg)</th>
        <th class='R'>Satış (Kg)</th>
        <th class='M'>Satış Tutarı (₺)</th>
        <th class='M'>Ort. Fiyat (₺/Kg)</th>
        <th class='R'>Sarf (Kg)</th>
        <th class='R'>Fire (Kg)</th>
        <th class='R'>Yemekhane (Kg)</th>
        <th class='R'>PROMS (Kg)</th>
        <th class='R'>Diğer Çıkış (Kg)</th>
        <th class='P' style='font-weight:900'>Top. Çıkış (Kg)</th>
        <th class='O'>Dönem Sonu (Kg)</th>
        <th class='O'>Dönem Sonu (Ton)</th>
      </tr>
    </thead>
    <tbody>");

        int rowIdx = 0;
        foreach (var k in analiz)
        {
            var alt = rowIdx++ % 2 == 1 ? " class='alt'" : "";
            var stokCls = k.DonemSonuMiktar < 0 ? "neg" : "orange";
            sb.AppendLine($@"<tr{alt}>
        <td class='L'>{k.KategoriAdi}</td>
        <td>{N0(k.DonemBasiMiktar)}</td>
        <td>{N0(k.UretimMiktar)}</td>
        <td>{N0(k.SatinAlmaMiktar)}</td>
        <td>{N0(k.SatisIadeMiktar)}</td>
        <td>{N0(k.ReceteFarkMiktar + k.SayimFazlasiMiktar)}</td>
        <td class='blue'>{N0(k.ToplamGirisMiktar)}</td>
        <td class='red'>{N0(k.SatisMiktar)}</td>
        <td class='pink'>{N2(k.SatisTutar)}</td>
        <td class='pink'>{Fiyat(k.SatisMiktar, k.SatisTutar)}</td>
        <td class='red'>{N0(k.SarfMiktar)}</td>
        <td class='red'>{N0(k.FireMiktar)}</td>
        <td class='red'>{N0(k.YemekhaneMiktar)}</td>
        <td class='red'>{N0(k.PromsMiktar)}</td>
        <td class='red'>{N0(k.SatinAlmaIadeMiktar)}</td>
        <td class='purple'>{N0(k.ToplamCikisMiktar)}</td>
        <td class='{stokCls}'>{N0(k.DonemSonuMiktar)}</td>
        <td class='{stokCls}'>{(k.DonemSonuMiktar / 1000m).ToString("N3", tr)}</td>
      </tr>");
        }

        // Toplam satırı
        var tGiris  = analiz.Sum(x => x.ToplamGirisMiktar);
        var tCikis  = analiz.Sum(x => x.ToplamCikisMiktar);
        var tSatis  = analiz.Sum(x => x.SatisMiktar);
        var tTutar  = analiz.Sum(x => x.SatisTutar);
        var tSonu   = analiz.Sum(x => x.DonemSonuMiktar);
        sb.AppendLine($@"<tr class='tot'>
        <td class='L'>TOPLAM</td>
        <td>{N0(analiz.Sum(x => x.DonemBasiMiktar))}</td>
        <td>{N0(analiz.Sum(x => x.UretimMiktar))}</td>
        <td>{N0(analiz.Sum(x => x.SatinAlmaMiktar))}</td>
        <td>{N0(analiz.Sum(x => x.SatisIadeMiktar))}</td>
        <td>{N0(analiz.Sum(x => x.ReceteFarkMiktar + x.SayimFazlasiMiktar))}</td>
        <td style='color:#0D47A1'>{N0(tGiris)}</td>
        <td style='color:#b71c1c'>{N0(tSatis)}</td>
        <td style='color:#880E4F'>{N2(tTutar)}</td>
        <td style='color:#880E4F'>{Fiyat(tSatis, tTutar)}</td>
        <td style='color:#b71c1c'>{N0(analiz.Sum(x => x.SarfMiktar))}</td>
        <td style='color:#b71c1c'>{N0(analiz.Sum(x => x.FireMiktar))}</td>
        <td style='color:#b71c1c'>{N0(analiz.Sum(x => x.YemekhaneMiktar))}</td>
        <td style='color:#b71c1c'>{N0(analiz.Sum(x => x.PromsMiktar))}</td>
        <td style='color:#b71c1c'>{N0(analiz.Sum(x => x.SatinAlmaIadeMiktar))}</td>
        <td style='color:#7B1FA2'>{N0(tCikis)}</td>
        <td style='color:#E65100'>{N0(tSonu)}</td>
        <td style='color:#E65100'>{(tSonu / 1000m).ToString("N3", tr)}</td>
      </tr>");

        sb.AppendLine($@"    </tbody>
  </table>
  <div class='footer'>{baslangic:dd.MM.yyyy} – {bitis:dd.MM.yyyy} &nbsp;|&nbsp; Ham LOGO verisi, düzeltme uygulanmamış &nbsp;|&nbsp; © {DateTime.Now.Year} Doğuş Çay</div>
</div>
</body>
</html>");
        return UygulaBlur(sb.ToString(), bulanik);
    }

    /// <summary>
    /// Şeker Dairesi Başkanlığı raporunu HTML olarak oluşturur (WhatsApp PNG ekranı için).
    /// Başkanlık tablosu hesaplama mantığı SekerDairesiRaporu.razor'daki HesaplaBaskanlikTablosu() ile aynıdır.
    /// </summary>
    public string SekerRaporHtmlOlustur(
        List<SekerKategoriAnaliz> analiz,
        List<SatisIadeDipnot> dipnotlar,
        DateTime baslangic,
        DateTime bitis,
        Dictionary<string, decimal>? baskanlikDonemBasi = null,
        Dictionary<string, decimal>? baskanlikDonemSonu = null,
        bool bulanik = false)
    {
        var tr = new CultureInfo("tr-TR");
        string N0(decimal v) => v.ToString("N0", tr);
        string N2(decimal v) => v.ToString("N2", tr);

        SekerKategoriAnaliz? GetKat(string key) => analiz.FirstOrDefault(x => x.Kategori == key);

        var a    = GetKat("A_KOTASI");
        var b    = GetKat("B_KOTASI");
        var c    = GetKat("C_KOTASI");
        var tk   = GetKat("TICARI_KRISTAL");
        var knya = GetKat("KONYA_TICARI");
        var pak  = GetKat("PAKETLI");
        var tpak = GetKat("TICARI_PAKET");

        decimal pakUretim  = pak?.UretimMiktar  ?? 0m;
        decimal tpakUretim = tpak?.UretimMiktar ?? 0m;

        decimal aSarfExtra  = Math.Max(0m, (a?.SarfMiktar  ?? 0m) - pakUretim);
        decimal tkSarfExtra = Math.Max(0m, (tk?.SarfMiktar ?? 0m) - tpakUretim);

        decimal aOrtFiyat   = (a?.SatisMiktar   ?? 0m) > 0 ? (a?.SatisTutar   ?? 0m) / a!.SatisMiktar   : 0m;
        // Konya Şeker satışı Ticari Kristal'e ekleniyor — ortalama fiyat her ikisi dahil hesaplanır
        decimal tkToplamSatisMiktar = (tk?.SatisMiktar ?? 0m) + (knya?.SatisMiktar ?? 0m);
        decimal tkToplamSatisTutar  = (tk?.SatisTutar  ?? 0m) + (knya?.SatisTutar  ?? 0m);
        decimal tkOrtFiyat  = tkToplamSatisMiktar > 0 ? tkToplamSatisTutar / tkToplamSatisMiktar : 0m;
        decimal pakOrtFiyat = (pak?.SatisMiktar  ?? 0m) > 0 ? (pak?.SatisTutar  ?? 0m) / pak!.SatisMiktar  : 0m;
        decimal tpOrtFiyat  = (tpak?.SatisMiktar ?? 0m) > 0 ? (tpak?.SatisTutar ?? 0m) / tpak!.SatisMiktar : 0m;

        decimal aSarfExtraTutar  = aOrtFiyat  * aSarfExtra;
        decimal tkSarfExtraTutar = tkOrtFiyat * tkSarfExtra;
        decimal pakSarfTutar     = pakOrtFiyat * (pak?.SarfMiktar  ?? 0m);
        decimal tpSarfTutar      = tpOrtFiyat  * (tpak?.SarfMiktar ?? 0m);

        decimal aYemPromsLocal   = (a?.YemekhaneMiktar   ?? 0m) + (a?.PromsMiktar   ?? 0m);
        decimal pakYemPromsLocal = (pak?.YemekhaneMiktar  ?? 0m) + (pak?.PromsMiktar  ?? 0m);
        decimal aYemPromsTutar   = aOrtFiyat   * aYemPromsLocal;
        decimal pakYemPromsTutar = pakOrtFiyat * pakYemPromsLocal;

        decimal aSatisBrut    = (a?.SatisMiktar   ?? 0m) + aSarfExtra   + aYemPromsLocal;
        decimal aSatisBrutTl  = (a?.SatisTutar    ?? 0m) + aSarfExtraTutar + aYemPromsTutar;
        decimal tkSatisBrut   = (tk?.SatisMiktar  ?? 0m) + tkSarfExtra + (knya?.SatisMiktar ?? 0m);
        decimal tkSatisBrutTl = (tk?.SatisTutar   ?? 0m) + tkSarfExtraTutar + (knya?.SatisTutar ?? 0m);
        decimal pakSatisBrut  = (pak?.SatisMiktar  ?? 0m) + (pak?.SarfMiktar  ?? 0m) + pakYemPromsLocal;
        decimal pakSatisBrutTl= (pak?.SatisTutar   ?? 0m) + pakSarfTutar + pakYemPromsTutar;
        decimal tpSatisBrut   = (tpak?.SatisMiktar ?? 0m) + (tpak?.SarfMiktar ?? 0m);
        decimal tpSatisBrutTl = (tpak?.SatisTutar  ?? 0m) + tpSarfTutar;

        decimal IadeHesapla(string kat, decimal brut) =>
            Math.Min(brut, Math.Max(0m, dipnotlar.Where(d => d.HedefKategori == kat).Sum(d => d.Miktar)));
        decimal IadeTutar(decimal iade, decimal brut, decimal brutTl) =>
            brut > 0 ? brutTl * iade / brut : 0m;

        decimal aIade   = IadeHesapla("A_KOTASI",       aSatisBrut);
        // Konya Şeker iadeleri de Ticari Kristal toplamına dahil edilir
        decimal tkIade  = Math.Min(tkSatisBrut, Math.Max(0m,
            dipnotlar.Where(d => d.HedefKategori == "TICARI_KRISTAL" || d.HedefKategori == "KONYA_TICARI").Sum(d => d.Miktar)));
        decimal pakIade = IadeHesapla("PAKETLI",        pakSatisBrut);
        decimal tpIade  = IadeHesapla("TICARI_PAKET",   tpSatisBrut);

        decimal aIadeTl   = IadeTutar(aIade,  aSatisBrut,  aSatisBrutTl);
        decimal tkIadeTl  = IadeTutar(tkIade, tkSatisBrut, tkSatisBrutTl);
        decimal pakIadeTl = IadeTutar(pakIade,pakSatisBrut,pakSatisBrutTl);
        decimal tpIadeTl  = IadeTutar(tpIade, tpSatisBrut, tpSatisBrutTl);

        decimal aSatMiktar    = aSatisBrut   - aIade;
        decimal aSatTutar     = aSatisBrutTl - aIadeTl;
        decimal tkSatMiktar   = tkSatisBrut  - tkIade;
        decimal tkSatTutar    = tkSatisBrutTl- tkIadeTl;
        decimal pakSatMiktar  = pakSatisBrut  - pakIade;
        decimal pakSatTutar   = pakSatisBrutTl- pakIadeTl;
        decimal tpakSatMiktar = tpSatisBrut   - tpIade;
        decimal tpakSatTutar  = tpSatisBrutTl - tpIadeTl;

        decimal bDigerK    = (b?.YemekhaneMiktar    ?? 0m) + (b?.PromsMiktar    ?? 0m);
        decimal cDigerK    = (c?.YemekhaneMiktar    ?? 0m) + (c?.PromsMiktar    ?? 0m);
        decimal tkDigerK   = (tk?.YemekhaneMiktar   ?? 0m) + (tk?.PromsMiktar   ?? 0m);
        decimal tpakDigerP = (tpak?.YemekhaneMiktar ?? 0m) + (tpak?.PromsMiktar ?? 0m);
        decimal digCikisK  = bDigerK + cDigerK + tkDigerK;
        decimal digCikisP  = tpakDigerP;

        decimal satisK = aSatMiktar + (b?.SatisMiktar ?? 0m) + (c?.SatisMiktar ?? 0m) + tkSatMiktar;
        decimal satisP = pakSatMiktar + tpakSatMiktar;
        decimal ciroK  = aSatTutar  + (b?.SatisTutar ?? 0m) + (c?.SatisTutar ?? 0m) + tkSatTutar;
        decimal ciroP  = pakSatTutar + tpakSatTutar;

        decimal aKFiyat  = aSatMiktar    > 0 ? aSatTutar    / aSatMiktar    : 0m;
        decimal aPFiyat  = pakSatMiktar  > 0 ? pakSatTutar  / pakSatMiktar  : 0m;
        decimal aFiyatT  = (aSatMiktar + pakSatMiktar) > 0 ? (aSatTutar + pakSatTutar) / (aSatMiktar + pakSatMiktar) : 0m;
        decimal bFiyat   = (b?.SatisMiktar ?? 0m) > 0 ? (b?.SatisTutar ?? 0m) / b!.SatisMiktar : 0m;
        decimal cFiyat   = (c?.SatisMiktar ?? 0m) > 0 ? (c?.SatisTutar ?? 0m) / c!.SatisMiktar : 0m;
        decimal tkKFiyat = tkSatMiktar   > 0 ? tkSatTutar   / tkSatMiktar   : 0m;
        decimal tkPFiyat = tpakSatMiktar > 0 ? tpakSatTutar / tpakSatMiktar : 0m;
        decimal tkFiyatT = (tkSatMiktar + tpakSatMiktar) > 0 ? (tkSatTutar + tpakSatTutar) / (tkSatMiktar + tpakSatMiktar) : 0m;

        var sb2 = new StringBuilder();
        var logo = LogoImgHtml();
        var ayAdi = baslangic.ToString("MMMM yyyy", tr);

        sb2.AppendLine($@"<!DOCTYPE html>
<html lang='tr'>
<head>
  <meta charset='UTF-8'>
  <style>
    body{{font-family:'Segoe UI',Arial,sans-serif;background:#f5f5f5;padding:10px;margin:0;color:#222}}
    .wrap{{background:#fff;border-radius:6px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.15);max-width:920px}}
    .hdr{{background:#1a237e;padding:10px 14px;color:#fff}}
    .hdr-title{{font-size:13px;font-weight:700}}
    .hdr-sub{{font-size:10px;color:rgba(255,255,255,.8);margin-top:2px}}
    table{{border-collapse:collapse;width:100%;font-size:11px}}
    td,th{{padding:4px 8px;border:1px solid #ddd}}
    td.L{{text-align:left}}
    td.R{{text-align:right;font-family:Consolas,monospace}}
    th{{background:#37474f;color:#fff;text-align:right;font-size:10px;white-space:nowrap}}
    th.L{{text-align:left}}
    .grp td{{background:#1a237e;color:#fff;font-weight:700;font-size:11px;border:none;padding:5px 8px}}
    .tot td{{background:#fff9c4;font-weight:700;color:#1b5e20;border-color:#f9a825}}
    .tl td{{background:#e3f2fd;font-weight:600;color:#0d47a1;border-color:#90caf9}}
    .fyt td{{background:#f3e5f5;font-weight:600;color:#6a1b9a;border-color:#ce93d8}}
    .alt td{{background:#f9f9f9}}
    .footer{{padding:5px 10px;font-size:9px;color:#aaa;border-top:1px solid #eee;text-align:center}}
    .dipnot{{background:#fff3e0;border-left:3px solid #ff9800;padding:6px 10px;font-size:10px;color:#555;margin:0 0 0 0}}
  </style>
</head>
<body>
<div class='wrap'>
  <div class='hdr'>
    <table cellpadding='0' cellspacing='0' border='0'><tr>
      <td style='padding-right:10px;vertical-align:middle'>{logo}</td>
      <td style='vertical-align:middle'>
        <div class='hdr-title'>ŞEKER ÜRETİM, SATIŞ VE STOK BİLGİLERİ</div>
        <div class='hdr-sub'>Doğuş Çay – Afyon Şeker Fabrikası &nbsp;|&nbsp; {baslangic:dd.MM.yyyy} – {bitis:dd.MM.yyyy} &nbsp;|&nbsp; {DateTime.Now:dd.MM.yyyy HH:mm}</div>
      </td>
    </tr></table>
  </div>
  <table>
    <colgroup><col style='width:460px'><col style='width:140px'><col style='width:140px'><col style='width:140px'></colgroup>
    <thead>
      <tr>
        <th class='L'>AÇIKLAMA</th>
        <th>Kristal Şeker (Kg)</th>
        <th>Küp/Paket Şeker (Kg)</th>
        <th>Toplam (Kg)</th>
      </tr>
    </thead>
    <tbody>");

        void GrupBaslik(string ad) =>
            sb2.AppendLine($"<tr class='grp'><td colspan='4'>{ad}</td></tr>");
        int _rowIdx = 0;
        void Veri(string ad, decimal k = 0, decimal p = 0)
        {
            var cls = _rowIdx++ % 2 == 1 ? " class='alt'" : "";
            sb2.AppendLine($"<tr{cls}><td class='L'>{ad}</td><td class='R'>{N0(k)}</td><td class='R'>{N0(p)}</td><td class='R'>{N0(k+p)}</td></tr>");
        }
        void Toplam(string ad, decimal k = 0, decimal p = 0) =>
            sb2.AppendLine($"<tr class='tot'><td class='L'>{ad}</td><td class='R'>{N0(k)}</td><td class='R'>{N0(p)}</td><td class='R'>{N0(k+p)}</td></tr>");
        void Tutar(string ad, decimal k = 0, decimal p = 0, decimal? tot = null) =>
            sb2.AppendLine($"<tr class='tl'><td class='L'>{ad}</td><td class='R'>{N2(k)} ₺</td><td class='R'>{N2(p)} ₺</td><td class='R'>{N2(tot ?? k+p)} ₺</td></tr>");
        void Fiyat(string ad, decimal k = 0, decimal p = 0, decimal? tot = null) =>
            sb2.AppendLine($"<tr class='fyt'><td class='L'>{ad}</td><td class='R'>{N2(k)} ₺/Kg</td><td class='R'>{N2(p)} ₺/Kg</td><td class='R'>{N2(tot ?? 0m)} ₺/Kg</td></tr>");

        // AY BAŞI STOK
        // AY BAŞI: Başkanlık zincirinden gelen değerler kullanılır (LOGO başı değil)
        decimal basA   = baskanlikDonemBasi?.GetValueOrDefault("A_KOTASI")       ?? (a?.DonemBasiMiktar    ?? 0m);
        decimal basPak = baskanlikDonemBasi?.GetValueOrDefault("PAKETLI")        ?? (pak?.DonemBasiMiktar   ?? 0m);
        decimal basB   = baskanlikDonemBasi?.GetValueOrDefault("B_KOTASI")       ?? (b?.DonemBasiMiktar    ?? 0m);
        decimal basC   = baskanlikDonemBasi?.GetValueOrDefault("C_KOTASI")       ?? (c?.DonemBasiMiktar    ?? 0m);
        decimal basTK  = baskanlikDonemBasi?.GetValueOrDefault("TICARI_KRISTAL") ?? (tk?.DonemBasiMiktar   ?? 0m);
        decimal basTP  = baskanlikDonemBasi?.GetValueOrDefault("TICARI_PAKET")   ?? (tpak?.DonemBasiMiktar  ?? 0m);

        GrupBaslik("AY BAŞI STOK"); _rowIdx = 0;
        Veri("Şirkete Ait Stok (Şeker Türleri İçin A Kotası)", basA, basPak);
        Veri("Şirkete Ait Stok (Şeker Türleri İçin B Kotası)", basB);
        Veri("Şirkete Ait Stok (Şeker Türleri İçin C Şekeri)", basC);
        Veri("Ticari Mal Stok", basTK, basTP);
        Toplam("Aylık \"Ay Başı Stok\" Toplam",
            basA + basB + basC + basTK,
            basPak + basTP);

        // SATINALMA
        GrupBaslik("SATINALMA / SATIN ALINAN"); _rowIdx = 0;
        // Konya Şeker satın alması da Ticari Kristal satırına eklenir (satışları da oraya ekleniyor)
        decimal tkSatinAlma = (tk?.SatinAlmaMiktar ?? 0m) + (knya?.SatinAlmaMiktar ?? 0m);
        Veri("Yurtiçinden Satın Alınan (Ticari Mal)", tkSatinAlma, tpak?.SatinAlmaMiktar ?? 0m);
        Veri("Yurtdışından Satın Alınan");
        Veri("Fason Üretim İçin Teslim Alınan");
        Toplam("Aylık \"Satın Alınan\" Toplam", tkSatinAlma, tpak?.SatinAlmaMiktar ?? 0m);

        // İŞLENEN
        GrupBaslik("İŞLENEN"); _rowIdx = 0;
        Veri("Şirkete Ait İşlenen A Kotası", pakUretim);
        Veri("Şirkete Ait İşlenen B Kotası");
        Veri("Şirkete Ait İşlenen C Şekeri");
        Veri("Ticari Mal İşlenen", tpakUretim);
        Veri("Fason İşlenen");
        Toplam("Aylık \"İşlenen\" Toplam", pakUretim + tpakUretim);

        // ÜRETİM
        GrupBaslik("ÜRETİM"); _rowIdx = 0;
        Veri("Şirkete Ait Üretim (Şeker Türleri İçin A Kotası)", a?.UretimMiktar ?? 0m, pakUretim);
        Veri("Şirkete Ait Üretim (Şeker Türleri İçin B Kotası)", b?.UretimMiktar ?? 0m);
        Veri("Şirkete Ait Üretim (Şeker Türleri İçin C Şekeri)", c?.UretimMiktar ?? 0m);
        Veri("Fason Üretim");
        Veri("Ticari Mal Üretimi", tk?.UretimMiktar ?? 0m, tpakUretim);
        Toplam("Aylık \"Üretim\" Toplam",
            (a?.UretimMiktar ?? 0m) + (b?.UretimMiktar ?? 0m) + (c?.UretimMiktar ?? 0m) + (tk?.UretimMiktar ?? 0m),
            pakUretim + tpakUretim);

        // DİĞER YOLLARLA GİRİŞ
        GrupBaslik("DİĞER YOLLARLA GİRİŞ"); _rowIdx = 0;
        Veri("Giriş (A Kotası)"); Veri("Giriş (B Kotası)"); Veri("Giriş (C Şekeri)");
        Veri("Giriş Fason"); Veri("Giriş Ticari Mal");
        Toplam("Aylık \"Diğer Yollarla Giriş\" Toplam");

        // DİĞER YOLLARLA ÇIKIŞ
        GrupBaslik("DİĞER YOLLARLA ÇIKIŞ"); _rowIdx = 0;
        Veri("Çıkış (A Kotası)");
        Veri("Çıkış (B Kotası)", bDigerK);
        Veri("Çıkış (C Şekeri)", cDigerK);
        Veri("Çıkış Fason");
        Veri("Çıkış Ticari Mal", tkDigerK, tpakDigerP);
        Toplam("Aylık \"Diğer Yollarla Çıkış\" Toplam", digCikisK, digCikisP);

        // SATIŞ
        GrupBaslik("SATIŞ (Satış ve/veya Teslim Eden)"); _rowIdx = 0;
        Veri("Yurtiçi Satış (Şeker Türleri İçin A Kotası)", aSatMiktar, pakSatMiktar);
        Veri("Yurtiçi Satış (Şeker Türleri İçin B Kotası)", b?.SatisMiktar ?? 0m);
        Veri("C Şekeri Satışı (Toplam)", c?.SatisMiktar ?? 0m);
        Veri("Yurtiçi Ticari Mal Satışı", tkSatMiktar, tpakSatMiktar);
        Veri("Yurtdışı Ticari Mal Satışı");
        Toplam("Aylık \"Satış ve/veya Teslim Edilen\" Toplam", satisK, satisP);
        Tutar("A Kotası Satış Hâsılatı (₺)", aSatTutar, pakSatTutar);
        Tutar("B Kotası Satış Hâsılatı (₺)", b?.SatisTutar ?? 0m);
        Tutar("C Şekeri Satış Hâsılatı (₺)", c?.SatisTutar ?? 0m);
        Tutar("Ticari Mal Satış Hâsılatı (₺)", tkSatTutar, tpakSatTutar);
        Tutar("Toplam Ciro", ciroK, ciroP);
        Fiyat("Aylık Fiyat Ortalaması – A Kotası (KDV hariç)", aKFiyat, aPFiyat, aFiyatT);
        Fiyat("Aylık Fiyat Ortalaması – B Kotası (KDV hariç)", bFiyat, 0m, bFiyat);
        Fiyat("Aylık Fiyat Ortalaması – C Şekeri (KDV hariç)", cFiyat, 0m, cFiyat);
        Fiyat("Aylık Fiyat Ortalaması – Ticari Mal (KDV hariç)", tkKFiyat, tkPFiyat, tkFiyatT);

        // AY SONU STOK (Başkanlık Portalı ile Uyumlu)
        // baskanlikDonemSonu mevcutsa doğrudan zincir sonucunu kullan (tüm kurallar uygulanmış)
        // Yoksa eski delta formülüne dön
        GrupBaslik("AY SONU STOK (Portal ile Uyumlu)"); _rowIdx = 0;
        decimal bIade  = IadeHesapla("B_KOTASI", b?.SatisMiktar ?? 0m);
        decimal cIade  = IadeHesapla("C_KOTASI", c?.SatisMiktar ?? 0m);
        decimal aKristalSonu, aPaketSonu, bSonu, cSonu, tkKristalSonu, tpakSonu;
        if (baskanlikDonemSonu != null)
        {
            aKristalSonu  = baskanlikDonemSonu.GetValueOrDefault("A_KOTASI");
            aPaketSonu    = baskanlikDonemSonu.GetValueOrDefault("PAKETLI");
            bSonu         = baskanlikDonemSonu.GetValueOrDefault("B_KOTASI");
            cSonu         = baskanlikDonemSonu.GetValueOrDefault("C_KOTASI");
            tkKristalSonu = baskanlikDonemSonu.GetValueOrDefault("TICARI_KRISTAL");
            tpakSonu      = baskanlikDonemSonu.GetValueOrDefault("TICARI_PAKET");
        }
        else
        {
            decimal bkA2   = baskanlikDonemBasi?.GetValueOrDefault("A_KOTASI")       ?? (a?.DonemBasiMiktar   ?? 0m);
            decimal bkPak2 = baskanlikDonemBasi?.GetValueOrDefault("PAKETLI")        ?? (pak?.DonemBasiMiktar  ?? 0m);
            decimal bkB2   = baskanlikDonemBasi?.GetValueOrDefault("B_KOTASI")       ?? (b?.DonemBasiMiktar   ?? 0m);
            decimal bkC2   = baskanlikDonemBasi?.GetValueOrDefault("C_KOTASI")       ?? (c?.DonemBasiMiktar   ?? 0m);
            decimal bkTK2  = baskanlikDonemBasi?.GetValueOrDefault("TICARI_KRISTAL") ?? (tk?.DonemBasiMiktar  ?? 0m);
            decimal bkTP2  = baskanlikDonemBasi?.GetValueOrDefault("TICARI_PAKET")   ?? (tpak?.DonemBasiMiktar ?? 0m);
            aKristalSonu  = bkA2   + (a?.DonemSonuMiktar   ?? 0m) - (a?.DonemBasiMiktar   ?? 0m) + aIade;
            aPaketSonu    = bkPak2 + (pak?.DonemSonuMiktar  ?? 0m) - (pak?.DonemBasiMiktar  ?? 0m) + pakIade;
            bSonu         = bkB2   + (b?.DonemSonuMiktar   ?? 0m) - (b?.DonemBasiMiktar   ?? 0m) + bIade;
            cSonu         = bkC2   + (c?.DonemSonuMiktar   ?? 0m) - (c?.DonemBasiMiktar   ?? 0m) + cIade;
            tkKristalSonu = bkTK2  + (tk?.DonemSonuMiktar  ?? 0m) - (tk?.DonemBasiMiktar  ?? 0m)
                                   + (knya?.DonemSonuMiktar ?? 0m) - (knya?.DonemBasiMiktar ?? 0m) + tkIade;
            tpakSonu      = bkTP2  + (tpak?.DonemSonuMiktar ?? 0m) - (tpak?.DonemBasiMiktar ?? 0m) + tpIade;
        }
        Veri("Ay Sonu Stok – A Kotası (Kristal / Paket)", aKristalSonu, aPaketSonu);
        Veri("Ay Sonu Stok – B Kotası",                   bSonu);
        Veri("Ay Sonu Stok – C Şekeri",                   cSonu);
        Veri("Ay Sonu Stok – Ticari Mal (Kristal / Paket)", tkKristalSonu, tpakSonu);
        decimal aysonuK = aKristalSonu + bSonu + cSonu + tkKristalSonu;
        decimal aysonuP = aPaketSonu + tpakSonu;
        Toplam("Aylık \"Ay Sonu Stok\" Toplam", aysonuK, aysonuP);

        sb2.AppendLine("    </tbody>\n  </table>");

        if (dipnotlar.Count > 0)
        {
            sb2.AppendLine("  <div class='dipnot'><strong>Dipnot:</strong>");
            foreach (var d in dipnotlar)
            {
                var yonStr = d.Yonlendirildi ? $" ({d.KaynakKategoriAdi} → {d.HedefKategoriAdi})" : "";
                sb2.AppendLine($"<br>• {d.SonrakiAyAdi} ayı {d.KaynakKategoriAdi} satış iadesi{yonStr}: {N0(d.Miktar)} Kg ({N2(d.Tutar)} ₺) {d.HedefKategoriAdi} satışından düşüldü.");
            }
            sb2.AppendLine("  </div>");
        }

        sb2.AppendLine($"  <div class='footer'>{ayAdi} &nbsp;|&nbsp; Oluşturma: {DateTime.Now:dd.MM.yyyy HH:mm} &nbsp;|&nbsp; © {DateTime.Now.Year} Doğuş Çay – Afyon Şeker Fabrikası</div>\n</div>\n</body>\n</html>");
        return UygulaBlur(sb2.ToString(), bulanik);
    }

    // =========================================================
    // BUZLAMA YARDIMCISI
    // =========================================================

    /// <summary>
    /// Yetkisiz WhatsApp kullanıcıları için rapor verilerini bulanıklaştırır.
    /// Sadece veri hücrelerini etkiler; satır başlıkları (td:first-child, td.L) açık kalır.
    /// </summary>
    private static string UygulaBlur(string html, bool bulanik)
    {
        if (!bulanik) return html;

        // CSS filter doğrudan <td>'de Puppeteer/Chromium'da güvenilir çalışmaz.
        // Çözüm: JS ile td içeriğini display:inline-block span içine sar → filter garantili çalışır.
        const string blurScript = @"
<style>
  .bz-span { filter:blur(5px); display:inline-block; user-select:none; pointer-events:none; }
  .bz-wm {
    position:fixed; top:50%; left:50%;
    transform:translate(-50%,-50%) rotate(-30deg);
    font-size:26px; font-weight:900; color:rgba(160,0,0,0.22);
    z-index:9999; pointer-events:none; white-space:nowrap; font-family:sans-serif;
    letter-spacing:2px;
  }
</style>
<script>
(function(){
  function blurHucreler() {
    var hedefler = [];
    // 1. SekerRaporu: td.R
    document.querySelectorAll('td.R').forEach(function(td){ hedefler.push(td); });
    // 2. BirlesikRapor / YanUrunler: .dt tablosundaki veri hücreleri
    document.querySelectorAll('.dt td').forEach(function(td){
      if (!td.classList.contains('L')) {
        var tr = td.parentElement;
        if (tr && !tr.classList.contains('grp')) hedefler.push(td);
      }
    });
    // 3. Genel tablolar (SekerAnaliz, PancarRapor, SekerSatis):
    //    tbody içindeki, ilk sütun olmayan, L sınıfı olmayan hücreler
    document.querySelectorAll('tbody tr td:not(:first-child):not(.L)').forEach(function(td){
      if (!td.classList.contains('R')) hedefler.push(td); // R zaten eklendi
    });

    // Tekrarları kaldır ve sar
    var gorulenler = new Set();
    hedefler.forEach(function(td){
      if (gorulenler.has(td)) return;
      gorulenler.add(td);
      if (td.querySelector('.bz-span')) return; // zaten sarilmis
      var s = document.createElement('span');
      s.className = 'bz-span';
      s.innerHTML = td.innerHTML;
      td.innerHTML = '';
      td.appendChild(s);
    });

    // Filigran
    if (!document.querySelector('.bz-wm')) {
      var wm = document.createElement('div');
      wm.className = 'bz-wm';
      wm.textContent = '\uD83D\uDD12 ORNEK - YETKISIZ ERISIM';
      document.body.appendChild(wm);
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', blurHucreler);
  } else {
    blurHucreler();
  }
})();
</script>";

        return html.Replace("</head>", blurScript + "\n</head>", StringComparison.OrdinalIgnoreCase);
    }
}
