using System.Globalization;
using System.Net;
using System.Text;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

// e-Fatura UBL XML'den çıkarılmış EFaturaDetay'i yazdırılabilir HTML'e dönüştürür.
// Kullanıcı tarayıcıda Ctrl+P veya "Yazdır / PDF" butonuyla PDF kaydedebilir.
public static class EFaturaHtmlBuilder
{
    private static readonly CultureInfo TR = new("tr-TR");

    public static string Build(EFaturaDetay d)
    {
        string N(decimal v) => v.ToString("N2", TR);
        string E(string? s) => WebUtility.HtmlEncode(s ?? "");

        var dov = E(d.Doviz);
        var sb = new StringBuilder(32_000);

        sb.Append(@"<!doctype html>
<html lang=""tr""><head>
<meta charset=""utf-8"" />
<title>e-Fatura ").Append(E(d.FaturaNo)).Append(@"</title>
<style>
  *{box-sizing:border-box}
  body{font-family:'Segoe UI',Arial,sans-serif;font-size:11px;color:#222;margin:0;background:#eef2f7}
  .page{background:#fff;max-width:210mm;min-height:297mm;margin:8mm auto;padding:14mm 12mm;box-shadow:0 2px 14px rgba(0,0,0,.14)}
  .bar{position:sticky;top:0;background:#0d47a1;color:#fff;padding:8px 14px;display:flex;align-items:center;gap:10px;box-shadow:0 2px 6px rgba(0,0,0,.2);z-index:50}
  .bar h1{font-size:14px;margin:0;font-weight:600;flex:1}
  .btn{background:#fff;color:#0d47a1;border:0;border-radius:4px;padding:6px 12px;font-weight:700;font-size:12px;cursor:pointer}
  .btn:hover{background:#bbdefb}
  .btn.sec{background:transparent;color:#fff;border:1px solid #fff}
  .btn.sec:hover{background:rgba(255,255,255,.15)}

  h2.title{font-size:16px;color:#0d47a1;border-bottom:2px solid #0d47a1;padding-bottom:4px;margin:0 0 10px 0;display:flex;justify-content:space-between;align-items:flex-end}
  h2.title small{font-weight:normal;font-size:11px;color:#666}

  .parties{display:flex;gap:10px;margin-bottom:10px}
  .party{flex:1;border:1px solid #cfd8dc;border-radius:4px;padding:8px 10px}
  .party.satici{border-left:4px solid #2e7d32}
  .party.alici{border-left:4px solid #1976d2}
  .party .label{font-size:9px;color:#888;text-transform:uppercase;letter-spacing:.5px;font-weight:700}
  .party .unvan{font-size:12px;font-weight:700;margin:2px 0 4px}
  .party .row{font-size:10px;margin:1px 0;color:#444}
  .cari-badge{display:inline-block;background:#fff3cd;border:1px solid #f6c243;color:#7a5a00;padding:2px 6px;border-radius:3px;font-family:Consolas,monospace;font-weight:700;font-size:10px;margin-top:3px}

  .meta{display:flex;flex-wrap:wrap;gap:6px;border:1px solid #cfd8dc;border-radius:4px;padding:6px 8px;background:#f5f7fa;margin-bottom:10px}
  .meta .cell{flex:1;min-width:120px}
  .meta .cell .k{font-size:9px;color:#888;text-transform:uppercase;letter-spacing:.5px}
  .meta .cell .v{font-size:12px;font-weight:700;color:#0d47a1}

  table.items{width:100%;border-collapse:collapse;margin-bottom:10px;font-size:10px}
  table.items th{background:#0d47a1;color:#fff;padding:5px 4px;font-weight:600;text-align:left;border:1px solid #0d47a1}
  table.items td{padding:4px;border:1px solid #cfd8dc;vertical-align:top}
  table.items td.num{text-align:right;font-variant-numeric:tabular-nums}
  table.items td.ctr{text-align:center}
  table.items tr:nth-child(even) td{background:#f9fafb}
  .ist{display:inline-block;margin-top:2px;padding:1px 4px;border:1px solid #f57c00;color:#e65100;border-radius:2px;font-size:9px}

  .bottom{display:flex;gap:10px;margin-top:8px}
  .notes{flex:1;border:1px solid #cfd8dc;border-radius:4px;padding:6px 8px;background:#faf5ff;font-size:10px;border-left:4px solid #6a1b9a}
  .notes .h{font-weight:700;color:#4a148c;margin-bottom:4px;font-size:10px}
  .notes li{margin:1px 0}
  .totals{width:42%;border:2px solid #0d47a1;border-radius:4px;background:linear-gradient(135deg,#f8fbff,#e3f2fd)}
  .totals table{width:100%;border-collapse:collapse;font-size:10px}
  .totals td{padding:4px 8px;border-bottom:1px dashed #cfd8dc}
  .totals td.num{text-align:right;font-variant-numeric:tabular-nums;font-weight:500}
  .totals tr.pay td{background:#0d47a1;color:#fff;font-weight:700;font-size:11px;border-bottom:0}
  .totals tr.tev td{color:#b71c1c}

  .panel{margin-top:6px;padding:5px 8px;border-radius:3px;font-size:10px}
  .panel.tev{background:#fff3e0;border-left:4px solid #f57c00}
  .panel.kdv{background:#fce4ec;border-left:4px solid #c2185b}
  .panel .h{font-weight:700;margin-bottom:2px}

  .footer{margin-top:14px;padding-top:6px;border-top:1px dashed #b0bec5;font-size:9px;color:#888;text-align:center}

  @media print{
    body{background:#fff}
    .bar{display:none}
    .page{box-shadow:none;margin:0;max-width:none;min-height:auto;padding:6mm 8mm}
    @page{size:A4;margin:8mm}
  }
</style></head>
<body>
  <div class=""bar"">
    <h1>e-Fatura — ").Append(E(d.FaturaNo)).Append(@"</h1>
    <button class=""btn"" onclick=""window.print()"">🖨️ Yazdır / PDF olarak kaydet</button>
    <button class=""btn sec"" onclick=""window.close()"">Kapat</button>
  </div>
  <div class=""page"">
    <h2 class=""title"">
      <span>e-FATURA</span>
      <small>ETTN: ").Append(E(d.Uuid)).Append(@"</small>
    </h2>

    <div class=""parties"">
      <div class=""party satici"">
        <div class=""label"">SATICI</div>
        <div class=""unvan"">").Append(E(d.SaticiUnvan)).Append(@"</div>
        <div class=""row""><b>VKN/TCKN:</b> ").Append(E(d.SaticiVkn));

        if (!string.IsNullOrWhiteSpace(d.SaticiVergiDairesi))
            sb.Append(@" &nbsp; <b>V.D.:</b> ").Append(E(d.SaticiVergiDairesi));

        sb.Append(@"</div>");
        if (!string.IsNullOrWhiteSpace(d.SaticiAdres))
            sb.Append(@"<div class=""row"">").Append(E(d.SaticiAdres)).Append(@"</div>");
        if (!string.IsNullOrWhiteSpace(d.SaticiTelefon) || !string.IsNullOrWhiteSpace(d.SaticiEposta))
            sb.Append(@"<div class=""row"">").Append(E(d.SaticiTelefon))
              .Append(string.IsNullOrWhiteSpace(d.SaticiEposta) ? "" : " · " + E(d.SaticiEposta))
              .Append(@"</div>");
        if (!string.IsNullOrWhiteSpace(d.LogoCariKod))
            sb.Append(@"<div class=""cari-badge"" title=""Logo Tiger LG_211_CLCARD"">Logo Cari: ")
              .Append(E(d.LogoCariKod))
              .Append(string.IsNullOrWhiteSpace(d.LogoCariUnvan) ? "" : " — " + E(d.LogoCariUnvan))
              .Append(@"</div>");

        sb.Append(@"
      </div>
      <div class=""party alici"">
        <div class=""label"">ALICI</div>
        <div class=""unvan"">").Append(E(d.AliciUnvan)).Append(@"</div>
        <div class=""row""><b>VKN/TCKN:</b> ").Append(E(d.AliciVkn));
        if (!string.IsNullOrWhiteSpace(d.AliciVergiDairesi))
            sb.Append(@" &nbsp; <b>V.D.:</b> ").Append(E(d.AliciVergiDairesi));
        sb.Append(@"</div>");
        if (!string.IsNullOrWhiteSpace(d.AliciAdres))
            sb.Append(@"<div class=""row"">").Append(E(d.AliciAdres)).Append(@"</div>");
        sb.Append(@"
      </div>
    </div>

    <div class=""meta"">
      <div class=""cell""><div class=""k"">Fatura No</div><div class=""v"">").Append(E(d.FaturaNo)).Append(@"</div></div>
      <div class=""cell""><div class=""k"">Tarih</div><div class=""v"">").Append(d.Tarih?.ToString("dd.MM.yyyy") ?? "-")
          .Append(string.IsNullOrEmpty(d.Saat) ? "" : " <span style='font-weight:400;font-size:10px;color:#666'>" + E(d.Saat) + "</span>")
          .Append(@"</div></div>
      <div class=""cell""><div class=""k"">Tipi</div><div class=""v"">").Append(E(d.Tip)).Append(@"</div></div>
      <div class=""cell""><div class=""k"">Senaryo</div><div class=""v"">").Append(E(d.ProfileId)).Append(@"</div></div>
      <div class=""cell""><div class=""k"">Döviz</div><div class=""v"">").Append(dov);
        if (d.DovizKuru > 0) sb.Append(" <span style='font-weight:400;font-size:10px;color:#666'>· kur ").Append(d.DovizKuru.ToString("N4", TR)).Append("</span>");
        sb.Append(@"</div></div>
    </div>

    <table class=""items"">
      <thead>
        <tr>
          <th style=""width:24px"">#</th>
          <th>Mal / Hizmet</th>
          <th style=""width:90px"">Stok Kodu</th>
          <th style=""width:70px;text-align:right"">Miktar</th>
          <th style=""width:30px"">Brm</th>
          <th style=""width:80px;text-align:right"">Birim Fiyat</th>
          <th style=""width:90px;text-align:right"">Net Tutar</th>
          <th style=""width:38px;text-align:center"">KDV %</th>
          <th style=""width:80px;text-align:right"">KDV Tutar</th>
          <th style=""width:38px;text-align:center"">Tvk %</th>
          <th style=""width:80px;text-align:right"">Tvk Tutar</th>
        </tr>
      </thead>
      <tbody>");

        foreach (var k in d.Kalemler)
        {
            sb.Append(@"
        <tr>
          <td class=""ctr"">").Append(k.Sira).Append(@"</td>
          <td><b>").Append(E(k.Aciklama)).Append("</b>");
            if (!string.IsNullOrWhiteSpace(k.StokAdi) && k.StokAdi != k.Aciklama)
                sb.Append(@"<div style=""color:#888;font-size:9px"">↳ ").Append(E(k.StokAdi)).Append("</div>");
            if (!string.IsNullOrWhiteSpace(k.IstisnaKodu) || !string.IsNullOrWhiteSpace(k.IstisnaSebep))
                sb.Append(@"<span class=""ist"">İstisna: ").Append(E(k.IstisnaKodu)).Append(' ').Append(E(k.IstisnaSebep)).Append("</span>");
            sb.Append(@"</td>
          <td style=""font-family:Consolas,monospace;font-size:9px"">").Append(E(k.StokKodu)).Append(@"</td>
          <td class=""num"">").Append(N(k.Miktar)).Append(@"</td>
          <td class=""ctr"">").Append(E(k.Birim)).Append(@"</td>
          <td class=""num"">").Append(N(k.BirimFiyat)).Append(@"</td>
          <td class=""num""><b>").Append(N(k.NetTutar)).Append(@"</b></td>
          <td class=""ctr"">").Append(k.KdvOran == 0 ? "" : "%" + k.KdvOran.ToString("N0", TR)).Append(@"</td>
          <td class=""num"">").Append(k.KdvTutar == 0 ? "" : N(k.KdvTutar)).Append(@"</td>
          <td class=""ctr"">").Append(k.TevkifatOran == 0 ? "" : "%" + k.TevkifatOran.ToString("N0", TR)).Append(@"</td>
          <td class=""num"" style=""color:#b71c1c;font-weight:600"">").Append(k.TevkifatTutar == 0 ? "" : N(k.TevkifatTutar)).Append(@"</td>
        </tr>");
        }

        sb.Append(@"
      </tbody>
    </table>

    <div class=""bottom"">
      <div class=""notes"">");

        if (d.FaturaNotlari.Count > 0)
        {
            sb.Append(@"<div class=""h"">Fatura Açıklamaları</div><ul style=""margin:0;padding-left:14px"">");
            foreach (var n in d.FaturaNotlari)
                sb.Append("<li>").Append(E(n)).Append("</li>");
            sb.Append("</ul>");
        }
        else
        {
            sb.Append(@"<div class=""h"" style=""color:#aaa"">— Açıklama yok —</div>");
        }

        sb.Append(@"
      </div>
      <div class=""totals"">
        <table>
          <tr><td>Mal/Hizmet Toplamı</td><td class=""num"">").Append(N(d.LineExtensionAmount)).Append(' ').Append(dov).Append("</td></tr>");

        if (d.AllowanceTotalAmount > 0)
            sb.Append(@"<tr><td>Toplam İskonto</td><td class=""num"">").Append(N(d.AllowanceTotalAmount)).Append(' ').Append(dov).Append("</td></tr>");

        foreach (var v in d.VergiOzetleri)
        {
            sb.Append("<tr><td>").Append(E(v.Adi));
            if (v.Oran > 0) sb.Append(" (%").Append(v.Oran.ToString("N0", TR)).Append(')');
            sb.Append(@"</td><td class=""num"">").Append(N(v.Tutar)).Append(' ').Append(dov).Append("</td></tr>");
        }

        foreach (var t in d.TevkifatOzetleri)
        {
            sb.Append(@"<tr class=""tev""><td>↓ Tevkifat");
            if (t.Oran > 0) sb.Append(" (%").Append(t.Oran.ToString("N0", TR)).Append(')');
            sb.Append(@"</td><td class=""num"">-").Append(N(t.Tutar)).Append(' ').Append(dov).Append("</td></tr>");
        }

        sb.Append("<tr><td>Vergiler Dahil Toplam</td><td class=\"num\">").Append(N(d.TaxInclusiveAmount)).Append(' ').Append(dov).Append("</td></tr>");
        sb.Append(@"<tr class=""pay""><td>ÖDENECEK TUTAR</td><td class=""num"">").Append(N(d.PayableAmount)).Append(' ').Append(dov).Append(@"</td></tr>
        </table>
      </div>
    </div>");

        // Tevkifat kodları paneli
        var tevkifatKodlari = d.Kalemler
            .Where(k => !string.IsNullOrWhiteSpace(k.TevkifatKodu))
            .GroupBy(k => new { k.TevkifatKodu, k.TevkifatAdi })
            .ToList();
        if (tevkifatKodlari.Count > 0)
        {
            sb.Append(@"
    <div class=""panel tev"">
      <div class=""h"">Tevkifat Kodları</div>");
            foreach (var grp in tevkifatKodlari)
                sb.Append("<div><b>").Append(E(grp.Key.TevkifatKodu)).Append("</b> — ").Append(E(grp.Key.TevkifatAdi)).Append("</div>");
            sb.Append(@"</div>");
        }

        // KDV istisnaları
        var istisnali = d.VergiOzetleri
            .Where(v => !string.IsNullOrWhiteSpace(v.IstisnaKodu) || !string.IsNullOrWhiteSpace(v.IstisnaSebep))
            .ToList();
        if (istisnali.Count > 0)
        {
            sb.Append(@"
    <div class=""panel kdv"">
      <div class=""h"">KDV İstisna Açıklamaları</div>");
            foreach (var v in istisnali)
                sb.Append("<div><b>").Append(E(v.IstisnaKodu)).Append("</b> — ").Append(E(v.IstisnaSebep)).Append("</div>");
            sb.Append(@"</div>");
        }

        sb.Append(@"
    <div class=""footer"">Raporlama Portalı · e-Fatura Önizleme · ").Append(DateTime.Now.ToString("dd.MM.yyyy HH:mm")).Append(@"</div>
  </div>
</body></html>");

        return sb.ToString();
    }
}
