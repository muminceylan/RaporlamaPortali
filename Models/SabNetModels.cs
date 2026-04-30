namespace RaporlamaPortali.Models;

/// <summary>
/// Kaynak: SabNetKANTAR.dbo.PMHS_KantarHareketleri (raw, tüm kolonlar)
/// SQLite tabloya 1:1 aktarılıyor; tarih/saat alanları Delphi serial (int) olarak saklanır.
/// </summary>
public class SabNetKantarHareketi
{
    public long     Id              { get; set; }
    public int?     Tarih           { get; set; }
    public string?  FisNo           { get; set; }
    public string?  RandevuNo       { get; set; }
    public string?  IslemTipi       { get; set; }
    public string?  UrunKodu        { get; set; }
    public decimal? BirimFiyat      { get; set; }
    public string?  HesapTipi       { get; set; }
    public string?  TcKimlikNo      { get; set; }
    public string?  HesapKodu       { get; set; }
    public string?  SozlesmeYili    { get; set; }
    public string?  Adres           { get; set; }
    public string?  PlakaNo         { get; set; }
    public string?  SoforAdiSoyadi  { get; set; }
    public string?  SoforGsmNo      { get; set; }
    public string?  AracTipi        { get; set; }
    public string?  MuteahhitKodu   { get; set; }
    public string?  MouseKodu       { get; set; }
    public string?  Aciklama        { get; set; }
    public decimal? Sevk            { get; set; }
    public decimal? Brut            { get; set; }
    public decimal? Dara            { get; set; }
    public decimal? Fireli          { get; set; }
    public decimal? Net             { get; set; }
    public decimal? FireOrani       { get; set; }
    public decimal? PolarOrani      { get; set; }
    public decimal? FireMiktari     { get; set; }
    public decimal? Fark            { get; set; }
    public string?  BosaltmaYeri    { get; set; }
    public string?  BosaltmaSekli   { get; set; }
    public int?     CikisTarihi     { get; set; }
    public int?     CikisSaati      { get; set; }
    public string?  CikisKullaniciKodu { get; set; }
    public int?     GirisTarihi     { get; set; }
    public int?     GirisSaati      { get; set; }
    public string?  GirisKullaniciKodu { get; set; }
    public string?  IrsaliyeNo      { get; set; }
    public string?  FaturaNo        { get; set; }
    public string?  SiparisNo       { get; set; }
    public string?  KantarKodu      { get; set; }
    public string?  Branda          { get; set; }
    public string?  KapiKagidiNo    { get; set; }
    public string?  Kod1            { get; set; }
    public string?  Kod2            { get; set; }
    public string?  Kod3            { get; set; }
    public string?  Kod4            { get; set; }
    public string?  Kod5            { get; set; }
    public decimal? Kod6            { get; set; }
    public decimal? Kod7            { get; set; }
    public decimal? Nakit           { get; set; }
    public decimal? KrediKarti      { get; set; }
    public decimal? Cari            { get; set; }
    public decimal? Havale          { get; set; }
    public string?  Kaydeden        { get; set; }
    public int?     KayitTarihi     { get; set; }
    public int?     KayitSaati      { get; set; }
    public string?  OnKayitHostName { get; set; }
    public string?  OnKayitIpAdres  { get; set; }
    public string?  GirisHostName   { get; set; }
    public string?  GirisIpAdres    { get; set; }
    public string?  CikisHostName   { get; set; }
    public string?  CikisIpAdres    { get; set; }
    public int?     RowId           { get; set; }
    public bool?    Kontrol         { get; set; }
    public DateTime ImportedAt      { get; set; } = DateTime.Now;
}

/// <summary>
/// Kaynak: SabNetKANTAR.dbo.PMHS_KantarHareketleri_Log (raw, tüm kolonlar + log alanları)
/// </summary>
public class SabNetKantarHareketiLog
{
    public long     Id              { get; set; }
    public int?     Tarih           { get; set; }
    public string?  FisNo           { get; set; }
    public string?  RandevuNo       { get; set; }
    public string?  IslemTipi       { get; set; }
    public string?  UrunKodu        { get; set; }
    public decimal? BirimFiyat      { get; set; }
    public string?  HesapTipi       { get; set; }
    public string?  TcKimlikNo      { get; set; }
    public string?  HesapKodu       { get; set; }
    public string?  SozlesmeYili    { get; set; }
    public string?  Adres           { get; set; }
    public string?  PlakaNo         { get; set; }
    public string?  SoforAdiSoyadi  { get; set; }
    public string?  SoforGsmNo      { get; set; }
    public string?  AracTipi        { get; set; }
    public string?  MuteahhitKodu   { get; set; }
    public string?  MouseKodu       { get; set; }
    public string?  Aciklama        { get; set; }
    public decimal? Sevk            { get; set; }
    public decimal? Brut            { get; set; }
    public decimal? Dara            { get; set; }
    public decimal? Fireli          { get; set; }
    public decimal? Net             { get; set; }
    public decimal? FireOrani       { get; set; }
    public decimal? PolarOrani      { get; set; }
    public decimal? FireMiktari     { get; set; }
    public decimal? Fark            { get; set; }
    public string?  BosaltmaYeri    { get; set; }
    public string?  BosaltmaSekli   { get; set; }
    public int?     CikisTarihi     { get; set; }
    public int?     CikisSaati      { get; set; }
    public string?  CikisKullaniciKodu { get; set; }
    public int?     GirisTarihi     { get; set; }
    public int?     GirisSaati      { get; set; }
    public string?  GirisKullaniciKodu { get; set; }
    public string?  IrsaliyeNo      { get; set; }
    public string?  FaturaNo        { get; set; }
    public string?  SiparisNo       { get; set; }
    public string?  KantarKodu      { get; set; }
    public string?  Branda          { get; set; }
    public string?  KapiKagidiNo    { get; set; }
    public string?  Kod1            { get; set; }
    public string?  Kod2            { get; set; }
    public string?  Kod3            { get; set; }
    public string?  Kod4            { get; set; }
    public string?  Kod5            { get; set; }
    public decimal? Kod6            { get; set; }
    public decimal? Kod7            { get; set; }
    public decimal? Nakit           { get; set; }
    public decimal? KrediKarti      { get; set; }
    public decimal? Cari            { get; set; }
    public decimal? Havale          { get; set; }
    public string?  Kaydeden        { get; set; }
    public int?     KayitTarihi     { get; set; }
    public int?     KayitSaati      { get; set; }
    public string?  OnKayitHostName { get; set; }
    public string?  OnKayitIpAdres  { get; set; }
    public string?  GirisHostName   { get; set; }
    public string?  GirisIpAdres    { get; set; }
    public string?  CikisHostName   { get; set; }
    public string?  CikisIpAdres    { get; set; }
    public int?     RowId           { get; set; }
    public bool?    Kontrol         { get; set; }
    public string?  LogIslemTipi    { get; set; }
    public string?  LogKaydeden     { get; set; }
    public int?     LogKayitTarihi  { get; set; }
    public int?     LogKayitSaati   { get; set; }
    public string?  LogAciklama     { get; set; }
    public string?  LogHostName     { get; set; }
    public string?  LogIpAdres      { get; set; }
    public DateTime ImportedAt      { get; set; } = DateTime.Now;
}

/// <summary>Kaynak: SabNetKANTAR.dbo.PMHS_CariHesapKarti</summary>
public class SabNetCariHesap
{
    public long    Id           { get; set; }
    public string? HesapKodu    { get; set; }
    public string? Unvan1       { get; set; }
    public string? Unvan2       { get; set; }
    public string? GrupKodu     { get; set; }
    public string? VergiDairesi { get; set; }
    public string? VergiNo      { get; set; }
    public string? FaturaAdres1 { get; set; }
    public string? FaturaAdres2 { get; set; }
    public string? FaturaAdres3 { get; set; }
    public string? ErpHesapKodu { get; set; }
    public string? Gsm          { get; set; }
    public string? Telefon1     { get; set; }
    public string? Aciklama     { get; set; }
    public string? Kaydeden     { get; set; }
    public int?    KayitTarihi  { get; set; }
    public int?    KayitSaati   { get; set; }
    public DateTime ImportedAt  { get; set; } = DateTime.Now;
}

/// <summary>Kaynak: SabNetKANTAR.dbo.PMHS_StokKarti</summary>
public class SabNetStokKarti
{
    public long     Id          { get; set; }
    public int?     StokKodu    { get; set; }
    public string?  StokAdi     { get; set; }
    public string?  Birim       { get; set; }
    public decimal? BirimFiyat  { get; set; }
    public int?     KdvOrani    { get; set; }
    public string?  DoluBos     { get; set; }
    public string?  Fis         { get; set; }
    public string?  Irsaliye    { get; set; }
    public string?  Fatura      { get; set; }
    public string?  AmbarOnayi  { get; set; }
    public string?  KapiKagidi  { get; set; }
    public string?  Aciklama    { get; set; }
    public string?  ERPKodu     { get; set; }
    public string?  GrupKodu    { get; set; }
    public string?  StokTakibi  { get; set; }
    public string?  Kaydeden    { get; set; }
    public int?     KayitTarihi { get; set; }
    public int?     KayitSaati  { get; set; }
    public DateTime ImportedAt  { get; set; } = DateTime.Now;
}

/// <summary>Kullanıcının düzenleyebildiği SabNetKANTAR SQL Server bağlantı ayarları (JSON).</summary>
public class SabNetBaglantiAyari
{
    public string Server   { get; set; } = "192.168.77.7";
    public string Database { get; set; } = "SabNetKANTAR";
    public string Username { get; set; } = "reportuser";
    public string Password { get; set; } = "reportuser";
}
