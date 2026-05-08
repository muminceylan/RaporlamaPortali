namespace RaporlamaPortali.Models;

/// <summary>
/// Cari mutabakatı için "tek" hareket satırı (bizden veya karşı firmadan).
/// Bizdeki kayıtlar: CLFLINE üzerinden BORÇ/ALACAK Logo konvansiyonuyla.
/// Karşı firma kayıtları: yüklenen ekstreden parse edilir; ekstrede karşı tarafın
/// gözünden BORÇ/ALACAK olur — eşleştirmede TARAFIN BORC = BIZIM ALACAK gibi
/// "ters yön" mantığı uygulanır.
/// </summary>
public class MutabakatKayit
{
    public DateTime Tarih { get; set; }
    public decimal Borc { get; set; }
    public decimal Alacak { get; set; }
    public string Aciklama { get; set; } = "";
    public string FisNo { get; set; } = "";

    // Sadece "biz" tarafı için anlamlı
    public string CariKodu { get; set; } = "";
    public string CariUnvan { get; set; } = "";
    public int TrCode { get; set; }
    public string FisTuru { get; set; } = "";

    // Kullanıcıya net görünen tutar (mutlak değer + yön etiketi)
    public decimal Tutar => Borc != 0 ? Borc : Alacak;
    public string Yon => Borc != 0 ? "BORÇ" : (Alacak != 0 ? "ALACAK" : "-");
}

/// <summary>
/// Eşleşen iki kayıt (bizden + onlardan).
/// </summary>
public class MutabakatEslesme
{
    public MutabakatKayit Bizim { get; set; } = new();
    public MutabakatKayit Onlar { get; set; } = new();
    public int GunFarki { get; set; }      // |bizim.Tarih - onlar.Tarih| (gün)
    public decimal TutarFarki { get; set; } // (bizim.Borç-Alacak) - (onlar.Alacak-Borç) → 0 ideal
    public string EslesmeYontemi { get; set; } = ""; // "Aynı tarih + tutar", "±N gün + tutar", "Belge no"
}

/// <summary>
/// Mutabakat sonucu — 3 grup: Eşleşen / Sadece Bizde / Sadece Onlarda.
/// </summary>
public class MutabakatSonuc
{
    public string AranılanVergiNo { get; set; } = "";
    public DateTime Baslangic { get; set; }
    public DateTime Bitis { get; set; }
    public List<EslesenCari> EslesenCariler { get; set; } = new();

    public List<MutabakatEslesme> Eslesenler { get; set; } = new();
    public List<MutabakatKayit> SadeceBizde { get; set; } = new();
    public List<MutabakatKayit> SadeceOnlarda { get; set; } = new();

    /// <summary>AI'nın ekstreden ham olarak çıkardığı tüm satırlar (tarih filtresi öncesi). Debug için.</summary>
    public List<MutabakatKayit> AIHamKayitlar { get; set; } = new();

    // Toplamlar (Logo konvansiyonu: BORÇ pozitif, ALACAK pozitif; net = BORÇ - ALACAK)
    public decimal BizimToplamBorc { get; set; }
    public decimal BizimToplamAlacak { get; set; }
    public decimal OnlarToplamBorc { get; set; }
    public decimal OnlarToplamAlacak { get; set; }

    public decimal BizimNet => BizimToplamBorc - BizimToplamAlacak;
    public decimal OnlarNet => OnlarToplamBorc - OnlarToplamAlacak;

    // Karşı firmadan baktığında bizim BORÇ → onların ALACAK olur. Mutabakat fark:
    //   BIZIM_NET + ONLAR_NET → 0 ideal (biz X borçlu isek onlar X alacaklı olmalı).
    public decimal MutabakatFarki => BizimNet + OnlarNet;
}

public class EslesenCari
{
    public string Kod { get; set; } = "";
    public string Unvan { get; set; } = "";
    public string VergiNo { get; set; } = "";
    public int HareketSayisi { get; set; }
}
