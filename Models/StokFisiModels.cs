namespace RaporlamaPortali.Models;

// Logo Unity COM API'sine doMaterialSlip (DataObjectType=1) ile gönderilen
// stok hareketi fişi. XML referansı: <MATERIAL_SLIPS / SLIP / TRANSACTIONS>.
//
// Aynı modeli iki sayfa kullanır:
//   - Üretimden Giriş Fişi   → Tip = 13
//   - Promosyon - Çıkış Fişi → Tip = 22
public class StokFisiModel
{
    /// <summary>Logo TRCODE: 13=Üretimden Giriş, 22=Sarf/Çıkış</summary>
    public int Tip { get; set; }

    /// <summary>Fiş tarihi</summary>
    public DateTime Tarih { get; set; } = DateTime.Today;

    /// <summary>Ambar no (SOURCE_WH). L_CAPIWHOUSE.NR.</summary>
    public int Ambar { get; set; }

    /// <summary>İşyeri/Fabrika no (SOURCE_FACTORY_NR). L_CAPIFIRM.NR.</summary>
    public int Fabrika { get; set; }

    /// <summary>Fiş açıklaması (GENEXP1)</summary>
    public string? Aciklama1 { get; set; }
    public string? Aciklama2 { get; set; }

    /// <summary>Belge no (manuel/elle girilen referans, opsiyonel)</summary>
    public string? BelgeNo { get; set; }

    public List<StokFisiSatir> Satirlar { get; set; } = new();
}

public class StokFisiSatir
{
    /// <summary>Malzeme kodu (ITEM_CODE) — Logo ITEMS.CODE</summary>
    public string MalzemeKodu { get; set; } = "";

    /// <summary>Malzeme adı (read-only, sadece UI için)</summary>
    public string MalzemeAdi { get; set; } = "";

    /// <summary>Birim kodu (UNIT_CODE) — örn KG, LT, AD</summary>
    public string Birim { get; set; } = "";

    public decimal Miktar { get; set; }

    /// <summary>Satır açıklaması</summary>
    public string? Aciklama { get; set; }
}

// Aktarım sonucu
public class StokFisiAktarimSonuc
{
    public bool    LoginBasarili { get; set; }
    public string? LoginHata     { get; set; }
    public bool    Basarili      { get; set; }
    public string? Hata          { get; set; }
    /// <summary>Logo'nun verdiği fiş numarası (yeni oluşturulan)</summary>
    public string? LogoFisNo     { get; set; }
    /// <summary>Detay tanı bilgileri</summary>
    public string  DiagLog       { get; set; } = "";
}
