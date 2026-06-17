namespace RaporlamaPortali.Models;

/// <summary>
/// Alış Kıyaslama Raporu — bir cariye ait iki dönemin malzeme bazında karşılaştırması.
/// Hem satın alma faturası (INVOICE.TRCODE=1) hem alınan hizmet faturası (TRCODE=14)
/// dahil; özet bilgisinde ikisi ayrı kalem olarak gösterilir.
/// </summary>
public class AlisKiyaslamaSonuc
{
    public CariBilgi Cari    { get; set; } = new();
    public AlisDonemOzeti DonemA { get; set; } = new();
    public AlisDonemOzeti DonemB { get; set; } = new();
    /// <summary>Birleşik satırlar — her satırda hem alım hem hizmet kısmı ayrılır.</summary>
    public List<AlisKiyaslamaSatiri> Satirlar { get; set; } = new();
}

public class AlisDonemOzeti
{
    public DateTime Baslangic { get; set; }
    public DateTime Bitis     { get; set; }

    /// <summary>Mal Alım Faturası (INVOICE.TRCODE=1).</summary>
    public int     SatinAlmaFis        { get; set; }
    public decimal SatinAlmaNet        { get; set; }
    public decimal SatinAlmaKdv        { get; set; }
    public decimal SatinAlmaBrut       { get; set; }

    /// <summary>Sabit Kıymet / Duran Varlık Alım Faturası (INVOICE.TRCODE=4).
    /// Doğuş'ta araç/makine/teçhizat alımı bu kanaldan geçer.</summary>
    public int     SabitKiymetFis      { get; set; }
    public decimal SabitKiymetNet      { get; set; }
    public decimal SabitKiymetKdv      { get; set; }
    public decimal SabitKiymetBrut     { get; set; }

    /// <summary>Alınan Hizmet Faturası (INVOICE.TRCODE=14).</summary>
    public int     HizmetFis           { get; set; }
    public decimal HizmetNet           { get; set; }
    public decimal HizmetKdv           { get; set; }
    public decimal HizmetBrut          { get; set; }

    public decimal ToplamNet  => SatinAlmaNet  + SabitKiymetNet  + HizmetNet;
    public decimal ToplamKdv  => SatinAlmaKdv  + SabitKiymetKdv  + HizmetKdv;
    public decimal ToplamBrut => SatinAlmaBrut + SabitKiymetBrut + HizmetBrut;
    public int     ToplamFis  => SatinAlmaFis  + SabitKiymetFis  + HizmetFis;

    public string Etiket => $"{Baslangic:dd.MM.yyyy} - {Bitis:dd.MM.yyyy}";
}

public class AlisKiyaslamaSatiri
{
    public string MalzemeKodu { get; set; } = "";
    public string MalzemeAdi  { get; set; } = "";
    public string Birim       { get; set; } = "";

    /// <summary>"Mal Alım" / "Hizmet" — INVOICE TRCODE'a göre belirlenir.</summary>
    public string FaturaTipi  { get; set; } = "";

    public decimal A_Miktar    { get; set; }
    public decimal A_NetTutar  { get; set; }
    public decimal A_OrtFiyat  => A_Miktar == 0 ? 0 : A_NetTutar / A_Miktar;

    public decimal B_Miktar    { get; set; }
    public decimal B_NetTutar  { get; set; }
    public decimal B_OrtFiyat  => B_Miktar == 0 ? 0 : B_NetTutar / B_Miktar;

    public decimal FarkMiktar  => A_Miktar - B_Miktar;
    public decimal FarkTutar   => A_NetTutar - B_NetTutar;

    public decimal FarkYuzdeTutar => B_NetTutar == 0
        ? (A_NetTutar == 0 ? 0 : 100m)
        : (A_NetTutar - B_NetTutar) / B_NetTutar * 100m;

    public decimal FarkYuzdeMiktar => B_Miktar == 0
        ? (A_Miktar == 0 ? 0 : 100m)
        : (A_Miktar - B_Miktar) / B_Miktar * 100m;
}
