namespace RaporlamaPortali.Models;

/// <summary>
/// Banka ödeme dosyaları (gönderilen) ile Logo Finans Raporu (gerçekleşen) karşılaştırma sonucu.
/// </summary>
public class OdemeKarsilastirSonuc
{
    public DateTime LogoBaslangic { get; set; }
    public DateTime LogoBitis     { get; set; }
    public int DosyaSatirSayisi   { get; set; }   // banka dosyalarından toplam okunan satır
    public int LogoSatirSayisi    { get; set; }   // Logo'dan çekilen ödeme satır sayısı
    public decimal DosyaToplam    { get; set; }
    public decimal LogoToplam     { get; set; }

    public List<OdemeKarsilastirSatir> Eslesen     { get; set; } = new();
    public List<OdemeKarsilastirSatir> Odenmedi    { get; set; } = new(); // dosyada/sabnet'te var, Logo'da yok
    public List<OdemeKarsilastirSatir> Mukerrer    { get; set; } = new(); // Logo'da aynı kişiye 2+ ödeme
    public List<OdemeKarsilastirSatir> LogoFazlasi { get; set; } = new(); // Logo'da var ama Sabnet/dosyada yok
    public List<string> Hatalar { get; set; } = new();

    public decimal OdenmediToplam    => Odenmedi.Sum(x => x.DosyaTutar);
    public decimal MukerrerToplam    => Mukerrer.Sum(x => x.LogoTutar);
    public decimal LogoFazlasiToplam => LogoFazlasi.Sum(x => x.LogoTutar);
}

public class OdemeKarsilastirSatir
{
    public string Banka      { get; set; } = "";    // Ziraat/Garanti/İş (dosya kaynağı)
    public string AdSoyad    { get; set; } = "";
    public string IBAN       { get; set; } = "";
    public string TcKimlikNo { get; set; } = "";

    public decimal DosyaTutar { get; set; }
    public decimal LogoTutar  { get; set; }
    public int     LogoAdet   { get; set; }         // Logo'da kaç ödeme bulundu
    public DateTime? LogoTarih { get; set; }        // Logo'daki ilk ödeme tarihi
    public string  LogoCariKod { get; set; } = "";  // Logo CH_KOD (ödeme yapılan cari)
    public string  Aciklama   { get; set; } = "";
}
