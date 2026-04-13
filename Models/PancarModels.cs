namespace RaporlamaPortali.Models;

public class PancarAvansKayit
{
    public string  KaynakEvrak  { get; set; } = "";   // "NAKDİ AVANS" | "AYNİ AVANS"
    public string  AvansGrubu   { get; set; } = "";   // "1. Avans", "Gübre", "Tohum" ...
    public decimal TutarToplami { get; set; }
    public decimal StopajToplami{ get; set; }
    public decimal KdvToplami   { get; set; }
}

public class PancarOzetIstatistik
{
    public int     ToplamCiftci  { get; set; }
    public decimal ToplamTaahhut { get; set; }  // kg
    public decimal ToplamNet     { get; set; }  // kg
    public double  OrtFireOrani  { get; set; }  // %
    public double  OrtPolar      { get; set; }  // %
}

public class PancarFinansOzet
{
    public decimal AvansKdv     { get; set; }
    public decimal AlimStopaji  { get; set; }
    public decimal NakliyePrimi { get; set; }
    public decimal KotaCezasi   { get; set; }
    public decimal BagkurBorcu  { get; set; }
    public decimal BorsaTescil  { get; set; }
    public decimal KotaFazlasi  { get; set; }
    public decimal CPancari     { get; set; }
}

public class PancarIcmalKayit
{
    public string  Tip       { get; set; } = "";
    public string  Aciklama  { get; set; } = "";
    public decimal Net       { get; set; }
    public decimal Tutar     { get; set; }
}

public class PancarCiftciDetay
{
    public string  Bolge         { get; set; } = "";
    public string  Koy           { get; set; } = "";
    public string  HesapKodu     { get; set; } = "";
    public string  TcKimlikNo    { get; set; } = "";
    public string  AdiSoyadi     { get; set; } = "";
    public decimal TaahhutTon    { get; set; }
    public decimal NetMiktar     { get; set; }
    public decimal APancari      { get; set; }
    public decimal APancariYuzde { get; set; }
    public decimal KotaFazlasi   { get; set; }
    public decimal OrtalamaPolar { get; set; }
}

public class PancarIcmalDetay
{
    // KarakanRaporu_Detay_2026'dan genel özetler
    public decimal NetMiktarTon       { get; set; }   // SUM(NetMiktar)/1000
    public decimal APancariTon        { get; set; }   // SUM(APancari)/1000
    public decimal APancariBedeli     { get; set; }
    public decimal CPancariTon        { get; set; }
    public decimal CPancariBedeli     { get; set; }
    public decimal KotaFazlasiTon     { get; set; }
    public decimal KotaFazlasiBedeli  { get; set; }
    public decimal KuspePrimi         { get; set; }   // KuspePrimiA + KuspePrimiC
    public decimal KotaTamamlamaPrimi { get; set; }
    public decimal MustahsilNakliye   { get; set; }   // OdenenNakliyePrimi

    // Hesaplanan
    public decimal PancarBedeliToplam => APancariBedeli + CPancariBedeli + KotaFazlasiBedeli;
    public decimal ABirimFiyati       => APancariTon    > 0 ? APancariBedeli    / APancariTon    : 0;
    public decimal CBirimFiyati       => CPancariTon    > 0 ? CPancariBedeli    / CPancariTon    : 0;
    public decimal KFBirimFiyati      => KotaFazlasiTon > 0 ? KotaFazlasiBedeli / KotaFazlasiTon : 0;
    public decimal HakedisToplamı     => PancarBedeliToplam + MustahsilNakliye + KuspePrimi + KotaTamamlamaPrimi;
}

public class PancarDetayTam
{
    public string  HesapKodu         { get; set; } = "";
    public string  TcKimlikNo        { get; set; } = "";
    public string  AdiSoyadi         { get; set; } = "";
    public string  BabaAdi           { get; set; } = "";
    public decimal TaahhutTon        { get; set; }
    public decimal TaahhutTonA       { get; set; }
    public decimal NetMiktar         { get; set; }
    public decimal APancari          { get; set; }
    public decimal APancariBedeli    { get; set; }
    public decimal KotaFazlasi       { get; set; }
    public decimal OrtalamaPolar     { get; set; }
    public decimal NakdiAvans        { get; set; }
    public decimal AvansToplami      { get; set; }
    public decimal Hakedis           { get; set; }
    public decimal NetHakedis        { get; set; }
}
