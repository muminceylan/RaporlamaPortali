using Dapper;
using Microsoft.Data.Sqlite;

namespace RaporlamaPortali.Services;

/// <summary>
/// SabNetKANTAR'dan aktarılan ham verilerin saklandığı SQLite veritabanını yönetir.
/// DB konumu: <see cref="AppDataPaths.SabNetDb"/>.
/// </summary>
public class SabNetDbService
{
    private readonly string _connectionString;

    public SabNetDbService()
    {
        var dbPath = AppDataPaths.SabNetDb;
        Directory.CreateDirectory(Path.GetDirectoryName(dbPath)!);
        _connectionString = $"Data Source={dbPath}";
        InitDb();
    }

    public string ConnectionString => _connectionString;

    public SqliteConnection CreateConnection()
    {
        var conn = new SqliteConnection(_connectionString);
        conn.Open();
        using (var pragma = conn.CreateCommand())
        {
            pragma.CommandText = "PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL; PRAGMA temp_store=MEMORY;";
            pragma.ExecuteNonQuery();
        }
        return conn;
    }

    private void InitDb()
    {
        using var conn = new SqliteConnection(_connectionString);
        conn.Open();

        using (var pragma = conn.CreateCommand())
        {
            pragma.CommandText = "PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL; PRAGMA temp_store=MEMORY;";
            pragma.ExecuteNonQuery();
        }

        conn.Execute(@"
CREATE TABLE IF NOT EXISTS SabNetKantarHareketleri (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Tarih INTEGER,
    FisNo TEXT,
    RandevuNo TEXT,
    IslemTipi TEXT,
    UrunKodu TEXT,
    BirimFiyat REAL,
    HesapTipi TEXT,
    TcKimlikNo TEXT,
    HesapKodu TEXT,
    SozlesmeYili TEXT,
    Adres TEXT,
    PlakaNo TEXT,
    SoforAdiSoyadi TEXT,
    SoforGsmNo TEXT,
    AracTipi TEXT,
    MuteahhitKodu TEXT,
    MouseKodu TEXT,
    Aciklama TEXT,
    Sevk REAL,
    Brut REAL,
    Dara REAL,
    Fireli REAL,
    Net REAL,
    FireOrani REAL,
    PolarOrani REAL,
    FireMiktari REAL,
    Fark REAL,
    BosaltmaYeri TEXT,
    BosaltmaSekli TEXT,
    Cikis_Tarihi INTEGER,
    Cikis_Saati INTEGER,
    Cikis_KullaniciKodu TEXT,
    Giris_Tarihi INTEGER,
    Giris_Saati INTEGER,
    Giris_KullaniciKodu TEXT,
    IrsaliyeNo TEXT,
    FaturaNo TEXT,
    SiparisNo TEXT,
    KantarKodu TEXT,
    Branda TEXT,
    KapiKagidiNo TEXT,
    Kod1 TEXT,
    Kod2 TEXT,
    Kod3 TEXT,
    Kod4 TEXT,
    Kod5 TEXT,
    Kod6 REAL,
    Kod7 REAL,
    Nakit REAL,
    KrediKarti REAL,
    Cari REAL,
    Havale REAL,
    Kaydeden TEXT,
    KayitTarihi INTEGER,
    KayitSaati INTEGER,
    OnKayit_HostName TEXT,
    OnKayit_ipAdres TEXT,
    Giris_HostName TEXT,
    Giris_ipAdres TEXT,
    Cikis_HostName TEXT,
    Cikis_ipAdres TEXT,
    Row_ID INTEGER,
    Kontrol INTEGER,
    ImportedAt TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS IX_SabNetKantarHareketleri_SozlesmeYili
    ON SabNetKantarHareketleri (SozlesmeYili);
CREATE INDEX IF NOT EXISTS IX_SabNetKantarHareketleri_KantarKodu
    ON SabNetKantarHareketleri (KantarKodu);
CREATE INDEX IF NOT EXISTS IX_SabNetKantarHareketleri_Tarih
    ON SabNetKantarHareketleri (Tarih);
CREATE INDEX IF NOT EXISTS IX_SabNetKantarHareketleri_HesapKodu
    ON SabNetKantarHareketleri (HesapKodu);

CREATE TABLE IF NOT EXISTS SabNetKantarHareketleriLog (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Tarih INTEGER,
    FisNo TEXT,
    RandevuNo TEXT,
    IslemTipi TEXT,
    UrunKodu TEXT,
    BirimFiyat REAL,
    HesapTipi TEXT,
    TcKimlikNo TEXT,
    HesapKodu TEXT,
    SozlesmeYili TEXT,
    Adres TEXT,
    PlakaNo TEXT,
    SoforAdiSoyadi TEXT,
    SoforGsmNo TEXT,
    AracTipi TEXT,
    MuteahhitKodu TEXT,
    MouseKodu TEXT,
    Aciklama TEXT,
    Sevk REAL,
    Brut REAL,
    Dara REAL,
    Fireli REAL,
    Net REAL,
    FireOrani REAL,
    PolarOrani REAL,
    FireMiktari REAL,
    Fark REAL,
    BosaltmaYeri TEXT,
    BosaltmaSekli TEXT,
    Cikis_Tarihi INTEGER,
    Cikis_Saati INTEGER,
    Cikis_KullaniciKodu TEXT,
    Giris_Tarihi INTEGER,
    Giris_Saati INTEGER,
    Giris_KullaniciKodu TEXT,
    IrsaliyeNo TEXT,
    FaturaNo TEXT,
    SiparisNo TEXT,
    KantarKodu TEXT,
    Branda TEXT,
    KapiKagidiNo TEXT,
    Kod1 TEXT,
    Kod2 TEXT,
    Kod3 TEXT,
    Kod4 TEXT,
    Kod5 TEXT,
    Kod6 REAL,
    Kod7 REAL,
    Nakit REAL,
    KrediKarti REAL,
    Cari REAL,
    Havale REAL,
    Kaydeden TEXT,
    KayitTarihi INTEGER,
    KayitSaati INTEGER,
    OnKayit_HostName TEXT,
    OnKayit_ipAdres TEXT,
    Giris_HostName TEXT,
    Giris_ipAdres TEXT,
    Cikis_HostName TEXT,
    Cikis_ipAdres TEXT,
    Row_ID INTEGER,
    Kontrol INTEGER,
    LogIslemTipi TEXT,
    LogKaydeden TEXT,
    LogKayitTarihi INTEGER,
    LogKayitSaati INTEGER,
    LogAciklama TEXT,
    LogHostName TEXT,
    LogipAdres TEXT,
    ImportedAt TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS IX_SabNetKantarHareketleriLog_SozlesmeYili
    ON SabNetKantarHareketleriLog (SozlesmeYili);
CREATE INDEX IF NOT EXISTS IX_SabNetKantarHareketleriLog_KantarKodu
    ON SabNetKantarHareketleriLog (KantarKodu);

CREATE TABLE IF NOT EXISTS SabNetCariHesaplar (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    HesapKodu TEXT,
    Unvan1 TEXT,
    Unvan2 TEXT,
    GrupKodu TEXT,
    VergiDairesi TEXT,
    VergiNo TEXT,
    FaturaAdres1 TEXT,
    FaturaAdres2 TEXT,
    FaturaAdres3 TEXT,
    ErpHesapKodu TEXT,
    Gsm TEXT,
    Telefon1 TEXT,
    Aciklama TEXT,
    Kaydeden TEXT,
    KayitTarihi INTEGER,
    KayitSaati INTEGER,
    ImportedAt TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS IX_SabNetCariHesaplar_HesapKodu
    ON SabNetCariHesaplar (HesapKodu);

CREATE TABLE IF NOT EXISTS SabNetStokKartlari (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    StokKodu INTEGER,
    StokAdi TEXT,
    Birim TEXT,
    BirimFiyat REAL,
    KdvOrani INTEGER,
    DoluBos TEXT,
    Fis TEXT,
    Irsaliye TEXT,
    Fatura TEXT,
    AmbarOnayi TEXT,
    KapiKagidi TEXT,
    Aciklama TEXT,
    ERPKodu TEXT,
    GrupKodu TEXT,
    StokTakibi TEXT,
    Kaydeden TEXT,
    KayitTarihi INTEGER,
    KayitSaati INTEGER,
    ImportedAt TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS IX_SabNetStokKartlari_StokKodu
    ON SabNetStokKartlari (StokKodu);
");
    }

    public async Task<(int hareket, int log, int cari, int stok)> KayitSayilariAsync()
    {
        await using var conn = new SqliteConnection(_connectionString);
        await conn.OpenAsync();
        var hareket = await conn.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM SabNetKantarHareketleri");
        var log     = await conn.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM SabNetKantarHareketleriLog");
        var cari    = await conn.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM SabNetCariHesaplar");
        var stok    = await conn.ExecuteScalarAsync<int>("SELECT COUNT(*) FROM SabNetStokKartlari");
        return (hareket, log, cari, stok);
    }

    public async Task<List<(string Yil, int Adet)>> YillaraGoreOzetAsync()
    {
        await using var conn = new SqliteConnection(_connectionString);
        await conn.OpenAsync();
        var rows = await conn.QueryAsync<(string Yil, int Adet)>(@"
            SELECT COALESCE(SozlesmeYili,'?') AS Yil, COUNT(*) AS Adet
            FROM SabNetKantarHareketleri
            GROUP BY COALESCE(SozlesmeYili,'?')
            ORDER BY 1 DESC");
        return rows.ToList();
    }

    public async Task TabloyuTemizleAsync(string tabloAdi)
    {
        var izinli = new[] { "SabNetKantarHareketleri", "SabNetKantarHareketleriLog", "SabNetCariHesaplar", "SabNetStokKartlari" };
        if (!izinli.Contains(tabloAdi))
            throw new ArgumentException($"Geçersiz tablo: {tabloAdi}");

        await using var conn = new SqliteConnection(_connectionString);
        await conn.OpenAsync();
        await conn.ExecuteAsync($"DELETE FROM {tabloAdi}");
        await conn.ExecuteAsync($"DELETE FROM sqlite_sequence WHERE name='{tabloAdi}'");
    }
}
