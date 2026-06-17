namespace RaporlamaPortali.Models;

/// <summary>
/// Excel'den okunup Logo Unity COM ile aktarılacak satış siparişi fişi.
/// Her Excel sütunu (CIPS/İÇECEK sheet'lerinde 9. sütundan başlayarak) bir fişi temsil eder.
/// </summary>
public class SatisSiparisFis
{
    /// <summary>Excel'deki sütun indeksi (debug/önizleme için).</summary>
    public int ExcelSutunIndex { get; set; }

    /// <summary>Hangi sheet'ten geldi (CIPS / İÇECEK).</summary>
    public string Sheet { get; set; } = "";

    /// <summary>
    /// R1: "Fabrika Çıkışı" — PENDİK / KARACABEY / AKSARAY / ŞANLIURFA / ÖDEMİŞ.
    /// FabrikaCikisAyari.Coz(...) ile BRANCH/DEPARTMENT/FACTORY/SOURCE_WH değerlerine çevrilir.
    /// </summary>
    public string FabrikaCikis { get; set; } = "";

    /// <summary>
    /// Logo cari kartında tanımlı ödeme planı KODU (örn. "47"). Boş = cari'de plan yok →
    /// Excel "Peşin/Vadeli" kuralı devreye girer. Dolu ise PAYMENT_CODE bu değer olarak gönderilir
    /// (Logo Unity COM cari default'unu otomatik uygulamadığı için elle gönderiyoruz).
    /// </summary>
    public string CariOdemePlanKodu { get; set; } = "";

    // ── Excel'den okunan başlık alanları ──────────────────────────────────────
    /// <summary>R3: 120.00.02.34.015 — Logo cari kodu (ARP_CODE).</summary>
    public string CariKodu { get; set; } = "";

    /// <summary>R4: 120.00.02.34.015.001 — Sevkiyat adresi kodu (SHIPLOC_CODE).</summary>
    public string SevkiyatAdresi { get; set; } = "";

    /// <summary>R5: Müşteri unvanı (önizleme amaçlı).</summary>
    public string CariUnvani { get; set; } = "";

    /// <summary>R6: Sevk depo adı (önizleme amaçlı).</summary>
    public string SevkDepoAdi { get; set; } = "";

    /// <summary>R2: Sevk Tarihi (15/06/2026 vb.).</summary>
    public DateTime SevkTarihi { get; set; }

    /// <summary>
    /// R8: Excel'deki ödeme türü metni — "Vadeli" / "Peşin" / "P" / boş gibi.
    /// Logo PAYMENT_CODE eşleştirmesi:
    ///   - Vadeli  → "40"        (cariden gelen 40 günlük vade)
    ///   - Peşin/P → "PESIN CIPS" (peşin ödeme kodu)
    ///   - Boş     → boş (cari kartının ödeme planı uygulanır)
    /// </summary>
    public string OdemeTuru { get; set; } = "";

    /// <summary>R14 (sütun başlığı): "2026-25--10" gibi araç no → DOC_NUMBER.</summary>
    public string AracNo { get; set; } = "";

    // ── Hesaplanan değerler ──────────────────────────────────────────────────
    /// <summary>26W{hafta} formatında (yıl son 2 hane + W + ISO hafta).</summary>
    public string SatisElemani { get; set; } = "";

    /// <summary>Sipariş tarihinden sonraki ilk Pazartesi (siparişin teslim tarihi).</summary>
    public DateTime TeslimTarihi { get; set; }

    /// <summary>Logo'dan çekilecek — cari kartının muhasebe hesap kodu (GL_CODE).</summary>
    public string MuhasebeHesabi { get; set; } = "";

    // ── Sipariş satırları ────────────────────────────────────────────────────
    public List<SatisSiparisSatir> Satirlar { get; set; } = new();

    // ── Aktarım sonucu (Logo'ya gönderildikten sonra dolar) ──────────────────
    public bool LogoyaGonderildi { get; set; }
    public string? LogoFisNo { get; set; }
    public string? LogoHata { get; set; }
}

/// <summary>
/// Satış siparişinin bir kalemi (malzeme satırı).
/// </summary>
public class SatisSiparisSatir
{
    /// <summary>Excel R14+ sütunundan (D) "Squ Kodu" — malzeme kodu (MASTER_CODE).</summary>
    public string MalzemeKodu { get; set; } = "";

    /// <summary>Excel R15+ sütundan H ya da malzeme açıklaması (önizleme).</summary>
    public string MalzemeAdi { get; set; } = "";

    /// <summary>Excel'deki hücre değeri — koli miktarı (QUANTITY, UNIT=KL).</summary>
    public decimal KoliMiktari { get; set; }

    /// <summary>
    /// PRCLIST'ten çekilen genel satış birim fiyatı (cariye özel olmayan, en yeni tarihli).
    /// 0 ise lookup başarısız oldu (malzeme için tanım yok).
    /// </summary>
    public decimal BirimFiyat { get; set; }
}

/// <summary>
/// Excel parse sonucu — Logo'ya gönderilmeye hazır fişlerin listesi.
/// </summary>
public class SatisSiparisExcelSonuc
{
    public List<SatisSiparisFis> Fisler { get; set; } = new();
    public List<string> Uyarilar { get; set; } = new();
    public string DosyaAdi { get; set; } = "";
}
