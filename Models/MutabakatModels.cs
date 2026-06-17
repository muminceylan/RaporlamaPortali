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

    // Döviz alanları (karşı firma ekstresinde TL ile birlikte döviz kolonları varsa)
    public string Doviz { get; set; } = "";       // "USD", "EUR", "" (TL)
    public decimal DovizBorc { get; set; }
    public decimal DovizAlacak { get; set; }

    // Kullanıcıya net görünen tutar (mutlak değer + yön etiketi)
    public decimal Tutar => Borc != 0 ? Borc : Alacak;
    public string Yon => Borc != 0 ? "BORÇ" : (Alacak != 0 ? "ALACAK" : "-");

    // Döviz tutarı (örtük kur hesaplamaları için)
    public decimal DovizTutar => DovizBorc != 0 ? DovizBorc : DovizAlacak;
    // Örtük kur: TL / Döviz (sıfır kontrolü dışarda yapılır)
    public decimal? OrtukKur =>
        DovizTutar != 0 && Tutar != 0 ? Math.Round(Tutar / DovizTutar, 4) : null;
}

/// <summary>
/// FIFO yöntemiyle eşleştirilmiş bir kur farkı kaydı.
/// Fatura (alacak) ile karşılığında yapılan ödeme/iade (borç) arasındaki kur farkı.
/// </summary>
public class KurFarkiEslesme
{
    public DateTime FaturaTarihi { get; set; }
    public string FaturaAciklama { get; set; } = "";
    public string FaturaFisNo { get; set; } = "";
    public decimal FaturaDoviz { get; set; }
    public decimal FaturaTL { get; set; }
    public decimal FaturaKur { get; set; }

    public DateTime OdemeTarihi { get; set; }
    public string OdemeAciklama { get; set; } = "";
    public string OdemeFisNo { get; set; } = "";
    public decimal OdemeDoviz { get; set; }
    public decimal OdemeTL { get; set; }
    public decimal OdemeKur { get; set; }

    public decimal EslesenDoviz { get; set; }
    public decimal KurFarki { get; set; } // (OdemeKur - FaturaKur) * EslesenDoviz; (+) lehe, (-) aleyhe
    public string Doviz { get; set; } = "";
}

/// <summary>
/// Aynı belgenin (fatura/ödeme) bizdeki Logo TL'si ile karşı firma ekstresindeki TL'si
/// arasındaki kur farkı. Belge no veya tarih+döviz tutarı üzerinden eşleştirilir.
/// </summary>
public class BelgeBazliKurFarki
{
    public DateTime Tarih { get; set; }
    public string FisNo { get; set; } = "";
    public string Aciklama { get; set; } = "";
    public string Yon { get; set; } = "";       // "FATURA" / "ÖDEME"
    public string Doviz { get; set; } = "";
    public decimal DovizTutar { get; set; }

    public decimal BizimTL { get; set; }
    public decimal BizimKur { get; set; }       // BizimTL / DovizTutar
    public string BizimCariKodu { get; set; } = "";

    public decimal OnlarTL { get; set; }
    public decimal OnlarKur { get; set; }       // OnlarTL / DovizTutar

    public decimal TLFarki { get; set; }        // BizimTL - OnlarTL  (+) bizde fazla, (-) onlarda fazla
    public decimal KurFarki { get; set; }       // BizimKur - OnlarKur
    public string EslesmeYontemi { get; set; } = ""; // "Fiş No", "Tarih + Döviz"
}

/// <summary>
/// Karşı firma ekstresinden AI ile döviz okuyup FIFO ile kur farkı hesaplaması sonucu.
/// İki tip kur farkı raporlanır:
///   1) Karşı firma içi (FIFO): kendi faturaları ↔ kendi ödemeleri arası kur değişimi
///   2) Belge bazlı: aynı belgenin bizdeki Logo TL'si ↔ onların TL'si arası fark
/// </summary>
public class KurFarkiSonuc
{
    public string AranılanVergiNo { get; set; } = "";
    public DateTime Baslangic { get; set; }
    public DateTime Bitis { get; set; }
    public string Doviz { get; set; } = "";

    // (1) Karşı firma içi FIFO
    public List<KurFarkiEslesme> Eslesmeler { get; set; } = new();
    public List<MutabakatKayit> AcikFaturalar { get; set; } = new();
    public List<MutabakatKayit> AcikOdemeler { get; set; } = new();
    public decimal ToplamKurFarki { get; set; }
    public decimal ToplamFaturaDoviz { get; set; }
    public decimal ToplamOdemeDoviz { get; set; }

    // (2) Belge bazlı (bizim Logo ↔ onların ekstresi)
    public List<BelgeBazliKurFarki> BelgeBazliFarklar { get; set; } = new();
    public List<EslesenCari> BizdekiCariler { get; set; } = new();
    public decimal BelgeBazliToplamTLFarki { get; set; }

    // Ham veriler (debug)
    public List<MutabakatKayit> AIHamKayitlar { get; set; } = new();
    public List<MutabakatKayit> BizimDovizliKayitlar { get; set; } = new();
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

/// <summary>
/// Mutabakatta bizdeki carileri filtreleme tipi.
/// Logo hesap planında 12x = Alıcılar (satış yaptığımız), 32x = Satıcılar (mal aldığımız).
/// Karşı firma ekstresi tek bir grup için geldiyse (örn sadece satıcı olarak alış faturaları)
/// bizdeki 120 grubu karıştırmasın diye filtre uygulanır.
/// </summary>
public enum MutabakatCariTuru
{
    /// <summary>Hem 12x hem 32x — eski davranış (default).</summary>
    Hepsi,
    /// <summary>Sadece 12x (Alıcılar). Onlar bize fatura kesmiş olabilir ama bizim sattıklarımız.</summary>
    Alici120,
    /// <summary>Sadece 32x (Satıcılar). Bizim aldığımız mal/hizmet için kestikleri faturalar.</summary>
    Satici320,
}
