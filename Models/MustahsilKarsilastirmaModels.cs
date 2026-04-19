namespace RaporlamaPortali.Models;

public class TcOzet
{
    public string TcKimlikNo { get; set; } = "";
    public string AdiSoyadi { get; set; } = "";

    // Logo (Excel MUSTAHSIL sheet) — hareket türü bazlı toplamlar
    public decimal LogoToptanSatisFaturasi { get; set; }  // = SabNet Avans karşılığı
    public decimal LogoMustahsilMakbuzu { get; set; }     // = SabNet Makbuz karşılığı
    public decimal LogoVirmanBorc { get; set; }           // = SabNet Cari BORÇ karşılığı
    public decimal LogoVirmanAlacak { get; set; }         // = SabNet Cari ALACAK karşılığı
    public int LogoSatirSayisi { get; set; }

    // SabNet — 3 tablodan toplamlar
    public decimal SabAvans { get; set; }                 // PMHS_AvansFormu toplamı
    public decimal SabMakbuz { get; set; }                // PMHS_MustahsilMakbuzu NetHakedis
    public decimal SabCariBorc { get; set; }              // PMHS_CariHareketler BA='BORÇ'
    public decimal SabCariAlacak { get; set; }            // PMHS_CariHareketler BA='ALACAK'

    // Farklar (Logo - SabNet)
    public decimal AvansFarki => LogoToptanSatisFaturasi - SabAvans;
    public decimal MakbuzFarki => LogoMustahsilMakbuzu - SabMakbuz;
    public decimal VirmanBorcFarki => LogoVirmanBorc - SabCariBorc;
    public decimal VirmanAlacakFarki => LogoVirmanAlacak - SabCariAlacak;

    // Makbuz için ±15 TL tolerans (küçük yuvarlama farkları göz ardı edilir)
    public const decimal MakbuzTolerans = 15m;

    public bool Eslesiyor =>
        Math.Abs(AvansFarki) < 0.01m &&
        Math.Abs(MakbuzFarki) <= MakbuzTolerans &&
        Math.Abs(VirmanBorcFarki) < 0.01m &&
        Math.Abs(VirmanAlacakFarki) < 0.01m;

    public string Durum => Eslesiyor ? "EŞLEŞME TAMAM" : "FARK VAR";
}

public class MustahsilKarsilastirmaSonuc
{
    public List<TcOzet> Kayitlar { get; set; } = new();
    public int LogoSatirSayisi { get; set; }
    public int SabAvansSatirSayisi { get; set; }
    public int SabMakbuzSatirSayisi { get; set; }
    public int SabCariSatirSayisi { get; set; }
    public DateTime BaslangicTarihi { get; set; }
    public DateTime BitisTarihi { get; set; }
    public string KampanyaYili { get; set; } = "";
    public string? Hata { get; set; }
}

/// <summary>
/// Bir TC için satır-seviyesi eşleştirme sonucu.
/// Her kategori (Avans / Makbuz / Virman) için Logo ve SabNet satırları Tutar üzerinden
/// greedy olarak eşleştirilir; eşleşmeyen satırlar "LOGO'DA YOK" / "SABNET'TE YOK" olur.
/// </summary>
public class MustahsilTcDetay
{
    public string TcKimlikNo { get; set; } = "";
    public string AdiSoyadi { get; set; } = "";
    public List<MustahsilDetayEslesme> Avans { get; set; } = new();
    public List<MustahsilDetayEslesme> Makbuz { get; set; } = new();
    public List<MustahsilDetayEslesme> Virman { get; set; } = new();
}

public class MustahsilDetayEslesme
{
    // Logo tarafı
    public DateTime? LogoTarihi { get; set; }
    public string? LogoIslemNo { get; set; }
    public string? LogoBelgeNo { get; set; }
    public decimal? LogoTutar { get; set; }

    // SabNet tarafı
    public DateTime? SabNetTarihi { get; set; }
    public string? SabNetNo { get; set; }
    public decimal? SabNetTutar { get; set; }
    public string? SabNetAciklama { get; set; }

    public string Durum { get; set; } = "";   // EŞLEŞTİ / LOGO'DA YOK / SABNET'TE YOK
    public bool Eslesti => Durum == "EŞLEŞTİ";
}
