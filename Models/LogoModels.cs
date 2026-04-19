namespace RaporlamaPortali.Models;

public class KasaHareketi
{
    public int LogicalRef { get; set; }
    public DateTime Tarih { get; set; }
    public string IslemNo { get; set; } = "";
    public string BelgeNo { get; set; } = "";
    public string CariUnvani { get; set; } = "";
    public string CariKodu { get; set; } = "";
    public string SatirAciklamasi { get; set; } = "";
    public string OzelKodu { get; set; } = "";
    public string TicariIslemGrubu { get; set; } = "";
    public int TrCode { get; set; }
    public string FisTuru { get; set; } = "";
    public bool Iptal { get; set; }
    public bool Muhasebelesti { get; set; }
    public decimal Tutar { get; set; }
    public int IsYeri { get; set; }
    public int Bolum { get; set; }
    public int TrCurr { get; set; }
    public string DovizTuru { get; set; } = "";
    public decimal Kur { get; set; }
    public decimal IslemDoviziTutari { get; set; }
    public decimal RaporlamaDoviziTutari { get; set; }
    public decimal RaporlamaDoviziKuru { get; set; }
}

public class KurBilgisi
{
    public DateTime Tarih { get; set; }
    public int CrType { get; set; }
    public string DovizKodu { get; set; } = "";
    public string DovizAdi { get; set; } = "";
    public decimal Rate1 { get; set; }
    public decimal Rate2 { get; set; }
    public decimal Rate3 { get; set; }
    public decimal Rate4 { get; set; }
}

public class DovizSecenek
{
    public int CrType { get; set; }
    public string Kod { get; set; } = "";
    public string Ad { get; set; } = "";
}

public class StokSatiri
{
    public int AmbarNo { get; set; }
    public string Ambar { get; set; } = "";
    public string MalzemeKodu { get; set; } = "";
    public string MalzemeAdi { get; set; } = "";
    public decimal Stok { get; set; }
}

public class AmbarSecenek
{
    public int AmbarNo { get; set; }
    public string Ad { get; set; } = "";
}
