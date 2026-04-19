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

public class CariBakiye
{
    public string CariHesapKodu { get; set; } = "";
    public string CariHesapUnvani { get; set; } = "";
    public string TcKimlikNo { get; set; } = "";
    public string VergiNo { get; set; } = "";
    public decimal Borc { get; set; }
    public decimal Alacak { get; set; }
    public decimal Bakiye { get; set; }
    public string OzelKod { get; set; } = "";
    public string OzelKod2 { get; set; } = "";
    public string OzelKod3 { get; set; } = "";
    public string OzelKod4 { get; set; } = "";
    public string OzelKod5 { get; set; } = "";
}

public class CariSecenek
{
    public string Kod { get; set; } = "";
    public string Unvan { get; set; } = "";
    public string TcKimlikNo { get; set; } = "";
    public string VergiNo { get; set; } = "";
    public string Etiket => string.IsNullOrEmpty(Kod) ? Unvan : $"{Kod} — {Unvan}";
}

public class KurFarkiSatiri
{
    public int TrCurr { get; set; }
    public string DovizKodu { get; set; } = "";
    public int IslemSayisi { get; set; }
    public decimal FcBorc { get; set; }
    public decimal FcAlacak { get; set; }
    public decimal FcBakiye { get; set; }
    public decimal TlBorc { get; set; }
    public decimal TlAlacak { get; set; }
    public decimal TlBakiye { get; set; }
    public decimal OrtalamaKur { get; set; }
    public decimal GuncelKur { get; set; }
    public DateTime? GuncelKurTarihi { get; set; }
    public decimal GuncelTlDeger { get; set; }
    public decimal KurFarki { get; set; }
    public string KesenTaraf { get; set; } = "";
    public string Yon { get; set; } = "";
    public string Aciklama { get; set; } = "";
    public int SentetikSayisi { get; set; }
    public decimal SentetikFcBorc { get; set; }
    public decimal SentetikFcAlacak { get; set; }
}

public class CariHareket
{
    public DateTime Tarih { get; set; }
    public DateTime? VadeTarihi { get; set; }
    public int ModuleNr { get; set; }
    public string OdemePlani { get; set; } = "";
    public int TrCode { get; set; }
    public string FisTuru { get; set; } = "";
    public string FisNo { get; set; } = "";
    public string Aciklama { get; set; } = "";
    public string CariKodu { get; set; } = "";
    public string CariUnvan { get; set; } = "";
    public decimal Borc { get; set; }
    public decimal Alacak { get; set; }
    public decimal Bakiye { get; set; }
    public int TrCurr { get; set; }
    public string DovizKodu { get; set; } = "";
    public decimal DovizKur { get; set; }
    public decimal DovizBorc { get; set; }
    public decimal DovizAlacak { get; set; }
    public bool SentetikDoviz { get; set; }
}

public class FifoEslesme
{
    public DateTime FaturaTarihi { get; set; }
    public DateTime FaturaVade { get; set; }
    public string FaturaFisTuru { get; set; } = "";
    public string FaturaFisNo { get; set; } = "";
    public int FaturaTrCurr { get; set; }
    public string FaturaDovizKodu { get; set; } = "";
    public decimal FaturaKuru { get; set; }
    public DateTime TahsilatTarihi { get; set; }
    public string TahsilatFisTuru { get; set; } = "";
    public string TahsilatFisNo { get; set; } = "";
    public int TahsilatTrCurr { get; set; }
    public decimal MahsupTl { get; set; }
    public decimal MahsupFc { get; set; }
    public decimal KullanilanKur { get; set; }
    public decimal GerceklesmisKurFarki { get; set; }
    public string Tip { get; set; } = "";
    public string Aciklama { get; set; } = "";
}

public class FifoAcikSatir
{
    public DateTime Tarih { get; set; }
    public DateTime Vade { get; set; }
    public string FisTuru { get; set; } = "";
    public string FisNo { get; set; } = "";
    public int TrCurr { get; set; }
    public string DovizKodu { get; set; } = "";
    public decimal KalanTl { get; set; }
    public decimal KalanFc { get; set; }
    public decimal FaturaKuru { get; set; }
    public decimal GuncelKur { get; set; }
    public decimal GuncelTlDeger { get; set; }
    public decimal GerceklesmemisKurFarki { get; set; }
    public string Tip { get; set; } = "";
}

public class KurFarkiFifoSonuc
{
    public List<FifoEslesme> Eslesmeler { get; set; } = new();
    public List<FifoAcikSatir> AcikFaturalar { get; set; } = new();
    public List<FifoAcikSatir> FazlaTahsilatlar { get; set; } = new();
    public decimal ToplamGerceklesmis { get; set; }
    public decimal ToplamGerceklesmemis { get; set; }
    public decimal ToplamKurFarki { get; set; }
    public string KesenTaraf { get; set; } = "";
    public string Yon { get; set; } = "";
}
