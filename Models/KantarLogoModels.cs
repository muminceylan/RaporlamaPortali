namespace RaporlamaPortali.Models;

public class KantarLogoEslesme
{
    public string KantarKodu { get; set; } = "";
    public string LogoKodu { get; set; } = "";
    public string MalzemeAdi { get; set; } = "";
    public bool KdvDus { get; set; }
}

public class KarsilastirmaSatiri
{
    public DateTime Tarih { get; set; }
    public string LogoKodu { get; set; } = "";
    public string MalzemeAdi { get; set; } = "";
    public decimal KantarMiktar { get; set; }
    public decimal KantarBirimFiyat { get; set; }
    public decimal KantarTutar { get; set; }
    public decimal LogoMiktar { get; set; }
    public decimal LogoBirimFiyat { get; set; }
    public decimal LogoTutar { get; set; }

    public decimal MiktarFarki => KantarMiktar - LogoMiktar;
    public decimal TutarFarki => KantarTutar - LogoTutar;
    public bool Eslesiyor => Math.Abs(MiktarFarki) < 0.01m && Math.Abs(TutarFarki) < 0.01m;
    public string Durum => Eslesiyor ? "EŞLEŞME TAMAM" : "FARK VAR";
}

public class TopluKarsilastirma
{
    public string LogoKodu { get; set; } = "";
    public string MalzemeAdi { get; set; } = "";
    public decimal KantarMiktar { get; set; }
    public decimal KantarTutar { get; set; }
    public decimal LogoMiktar { get; set; }
    public decimal LogoTutar { get; set; }
    public decimal MiktarFarki => KantarMiktar - LogoMiktar;
    public decimal TutarFarki => KantarTutar - LogoTutar;
    public bool Eslesiyor => Math.Abs(MiktarFarki) < 0.01m && Math.Abs(TutarFarki) < 0.01m;
    public string Durum => Eslesiyor ? "EŞLEŞME TAMAM" : "FARK VAR";
}

public class KarsilastirmaSonuc
{
    public List<KarsilastirmaSatiri> GunlukRapor { get; set; } = new();
    public List<TopluKarsilastirma> TopluRapor { get; set; } = new();
    public int KantarSatirSayisi { get; set; }
    public int LogoSatirSayisi { get; set; }
    public DateTime BaslangicTarihi { get; set; }
    public DateTime BitisTarihi { get; set; }
}

public class KantarHamSatir
{
    public string FisNo { get; set; } = "";
    public string UrunKodu { get; set; } = "";
    public string StokAdi { get; set; } = "";
    public string PlakaNo { get; set; } = "";
    public string SoforAdiSoyadi { get; set; } = "";
    public DateTime Tarih { get; set; }
    public decimal Net { get; set; }
    public decimal BirimFiyat { get; set; }
    public decimal Tutar { get; set; }
}

public class LogoHamSatir
{
    public DateTime Tarih { get; set; }
    public string FisTuru { get; set; } = "";
    public string FisNumarasi { get; set; } = "";
    public string CariKodu { get; set; } = "";
    public string CariUnvani { get; set; } = "";
    public string MalzemeKodu { get; set; } = "";
    public string MalzemeAdi { get; set; } = "";
    public decimal CikisMiktari { get; set; }
    public decimal CikisFiyati { get; set; }
    public decimal CikisTutari { get; set; }
}

public class FisEslestirmesi
{
    // Kantar tarafı
    public string? KantarFisNo { get; set; }
    public string? KantarPlaka { get; set; }
    public string? KantarSofor { get; set; }
    public decimal? KantarNet { get; set; }
    public decimal? KantarBirimFiyat { get; set; }
    public decimal? KantarTutar { get; set; }

    // Logo tarafı
    public string? LogoFisNumarasi { get; set; }
    public string? LogoCariKodu { get; set; }
    public string? LogoCariUnvani { get; set; }
    public decimal? LogoMiktar { get; set; }
    public decimal? LogoBirimFiyat { get; set; }
    public decimal? LogoTutar { get; set; }

    public string Durum { get; set; } = ""; // "EŞLEŞTİ" | "LOGO'DA YOK" | "KANTAR'DA YOK"
    public bool Eslesiyor => Durum == "EŞLEŞTİ";
}
