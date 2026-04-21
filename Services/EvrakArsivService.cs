using Dapper;
using Microsoft.Data.Sqlite;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

public class EvrakArsivService
{
    private readonly string _dbPath;
    private readonly string _connectionString;
    private readonly string _dosyaKokDizini;
    private readonly object _lock = new();

    public EvrakArsivService()
    {
        _dbPath           = AppDataPaths.EvrakArsivDb;
        _connectionString = $"Data Source={_dbPath}";
        _dosyaKokDizini   = AppDataPaths.EvrakArsivDizini;
        Directory.CreateDirectory(_dosyaKokDizini);
        InitDb();
    }

    public string DosyaKokDizini => _dosyaKokDizini;

    private SqliteConnection Open()
    {
        var conn = new SqliteConnection(_connectionString);
        conn.Open();
        return conn;
    }

    private void InitDb()
    {
        using var conn = Open();

        // Migration: CiftciKartiId -> TcKimlikNo (önceki versiyondan) — ana CREATE'lerden ÖNCE
        var tabloVar = conn.ExecuteScalar<long>(
            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='MustahsilEvraki'");
        if (tabloVar > 0)
        {
            var eskiKolon = conn.ExecuteScalar<long>(
                "SELECT COUNT(*) FROM pragma_table_info('MustahsilEvraki') WHERE name='CiftciKartiId'");
            if (eskiKolon > 0)
            {
                conn.Execute("DROP TABLE MustahsilEvraki");
            }
        }

        conn.Execute(@"
CREATE TABLE IF NOT EXISTS Tesis (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Ad TEXT NOT NULL,
    Aktif INTEGER NOT NULL DEFAULT 1,
    OlusturmaTarihi TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS BelgeKategori (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Ad TEXT NOT NULL,
    Tur INTEGER NOT NULL,
    Aktif INTEGER NOT NULL DEFAULT 1,
    OlusturmaTarihi TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS BelgeTipi (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Ad TEXT NOT NULL,
    Kategori INTEGER NOT NULL,
    Aktif INTEGER NOT NULL DEFAULT 1,
    OlusturmaTarihi TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS TesisEvraki (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    DefterYili INTEGER NOT NULL,
    DefterSiraNo INTEGER NOT NULL,
    TesisId INTEGER NOT NULL,
    BelgeTipiId INTEGER NOT NULL,
    DosyaAdi TEXT NOT NULL,
    DosyaYolu TEXT NOT NULL,
    MimeType TEXT,
    DosyaBoyutu INTEGER NOT NULL,
    EvrakTarihi TEXT,
    EvrakNo TEXT,
    TebellugTarihi TEXT NOT NULL,
    GecerlilikBaslangic TEXT,
    GecerlilikBitis TEXT,
    Aciklama TEXT,
    YuklemeTarihi TEXT NOT NULL,
    UNIQUE(DefterYili, DefterSiraNo)
);

CREATE TABLE IF NOT EXISTS MustahsilEvraki (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    DefterYili INTEGER NOT NULL,
    DefterSiraNo INTEGER NOT NULL,
    KampanyaYili INTEGER NOT NULL,
    TcKimlikNo TEXT,
    MustahsilAdSoyadi TEXT NOT NULL,
    MustahsilNo TEXT,
    HesapNo TEXT,
    BelgeTipiId INTEGER NOT NULL,
    DosyaAdi TEXT NOT NULL,
    DosyaYolu TEXT NOT NULL,
    MimeType TEXT,
    DosyaBoyutu INTEGER NOT NULL,
    EvrakTarihi TEXT,
    EvrakNo TEXT,
    TebellugTarihi TEXT NOT NULL,
    Aciklama TEXT,
    YuklemeTarihi TEXT NOT NULL,
    UNIQUE(DefterYili, DefterSiraNo)
);

CREATE TABLE IF NOT EXISTS GenelEvraki (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    DefterYili INTEGER NOT NULL,
    DefterSiraNo INTEGER NOT NULL,
    KategoriId INTEGER NOT NULL,
    BelgeTipiId INTEGER NOT NULL,
    CariUnvani TEXT NOT NULL,
    FaturaNo TEXT,
    EvrakTarihi TEXT,
    Tutar REAL,
    DosyaAdi TEXT NOT NULL,
    DosyaYolu TEXT NOT NULL,
    MimeType TEXT,
    DosyaBoyutu INTEGER NOT NULL,
    TebellugTarihi TEXT NOT NULL,
    Aciklama TEXT,
    YuklemeTarihi TEXT NOT NULL,
    UNIQUE(DefterYili, DefterSiraNo)
);

CREATE INDEX IF NOT EXISTS IX_TesisEvraki_TesisId    ON TesisEvraki(TesisId);
CREATE INDEX IF NOT EXISTS IX_TesisEvraki_BelgeTipi  ON TesisEvraki(BelgeTipiId);
CREATE INDEX IF NOT EXISTS IX_MustEvraki_Kampanya    ON MustahsilEvraki(KampanyaYili);
CREATE INDEX IF NOT EXISTS IX_MustEvraki_Tc          ON MustahsilEvraki(TcKimlikNo);
CREATE INDEX IF NOT EXISTS IX_MustEvraki_BelgeTipi   ON MustahsilEvraki(BelgeTipiId);
CREATE INDEX IF NOT EXISTS IX_GenelEvraki_Kategori   ON GenelEvraki(KategoriId);
CREATE INDEX IF NOT EXISTS IX_GenelEvraki_BelgeTipi  ON GenelEvraki(BelgeTipiId);
");

        // Kategori seed: yalnızca tablo boşsa iki varsayılan satır
        var kategoriSayisi = conn.ExecuteScalar<long>("SELECT COUNT(*) FROM BelgeKategori");
        if (kategoriSayisi == 0)
        {
            var now = DateTime.Now.ToString("O");
            conn.Execute(@"
INSERT INTO BelgeKategori(Id, Ad, Tur, Aktif, OlusturmaTarihi) VALUES (1, 'Tesis', 1, 1, @t);
INSERT INTO BelgeKategori(Id, Ad, Tur, Aktif, OlusturmaTarihi) VALUES (2, 'Müstahsil', 2, 1, @t);", new { t = now });
        }
    }

    // ========== TESIS ==========

    public List<Tesis> TesisListesi(bool? aktif = null)
    {
        using var conn = Open();
        var sql = "SELECT * FROM Tesis";
        if (aktif.HasValue) sql += $" WHERE Aktif = {(aktif.Value ? 1 : 0)}";
        sql += " ORDER BY Ad";
        return conn.Query<Tesis>(sql).ToList();
    }

    public Tesis? TesisGetir(int id)
        => Open().QueryFirstOrDefault<Tesis>("SELECT * FROM Tesis WHERE Id = @id", new { id });

    public int TesisEkle(string ad)
    {
        using var conn = Open();
        return (int)conn.ExecuteScalar<long>(
            @"INSERT INTO Tesis(Ad, Aktif, OlusturmaTarihi) VALUES(@ad, 1, @t);
              SELECT last_insert_rowid();",
            new { ad, t = DateTime.Now.ToString("O") });
    }

    public void TesisGuncelle(int id, string ad, bool aktif)
    {
        using var conn = Open();
        conn.Execute("UPDATE Tesis SET Ad=@ad, Aktif=@aktif WHERE Id=@id",
            new { id, ad, aktif = aktif ? 1 : 0 });
    }

    public void TesisSil(int id)
    {
        using var conn = Open();
        conn.Execute("DELETE FROM Tesis WHERE Id=@id", new { id });
    }

    // ========== BELGE KATEGORI ==========

    public List<BelgeKategori> BelgeKategoriListesi(BelgeKategorisi? tur = null, bool? aktif = null)
    {
        using var conn = Open();
        var sql = "SELECT * FROM BelgeKategori WHERE 1=1";
        if (tur.HasValue)   sql += $" AND Tur = {(int)tur.Value}";
        if (aktif.HasValue) sql += $" AND Aktif = {(aktif.Value ? 1 : 0)}";
        sql += " ORDER BY Tur, Ad";
        return conn.Query<BelgeKategori>(sql).ToList();
    }

    public BelgeKategori? BelgeKategoriGetir(int id)
        => Open().QueryFirstOrDefault<BelgeKategori>("SELECT * FROM BelgeKategori WHERE Id=@id", new { id });

    public int BelgeKategoriEkle(string ad, BelgeKategorisi tur)
    {
        using var conn = Open();
        return (int)conn.ExecuteScalar<long>(
            @"INSERT INTO BelgeKategori(Ad, Tur, Aktif, OlusturmaTarihi)
              VALUES(@ad, @tur, 1, @t);
              SELECT last_insert_rowid();",
            new { ad, tur = (int)tur, t = DateTime.Now.ToString("O") });
    }

    public void BelgeKategoriGuncelle(int id, string ad, BelgeKategorisi tur, bool aktif)
    {
        using var conn = Open();
        conn.Execute("UPDATE BelgeKategori SET Ad=@ad, Tur=@tur, Aktif=@aktif WHERE Id=@id",
            new { id, ad, tur = (int)tur, aktif = aktif ? 1 : 0 });
    }

    public void BelgeKategoriSil(int id)
    {
        using var conn = Open();
        var kullanim = conn.ExecuteScalar<long>("SELECT COUNT(*) FROM BelgeTipi WHERE Kategori=@id", new { id });
        if (kullanim > 0) throw new InvalidOperationException("Bu kategoriye bağlı belge tipi bulunuyor, önce onları silin veya taşıyın.");
        conn.Execute("DELETE FROM BelgeKategori WHERE Id=@id", new { id });
    }

    // ========== BELGE TIPI ==========

    public List<BelgeTipi> BelgeTipiListesi(BelgeKategorisi? tur = null, int? kategoriId = null, bool? aktif = null)
    {
        using var conn = Open();
        var sql = @"
SELECT bt.Id, bt.Ad, bt.Kategori AS KategoriId, bt.Aktif, bt.OlusturmaTarihi,
       k.Ad AS KategoriAdi, k.Tur AS KategoriTuru
FROM BelgeTipi bt
LEFT JOIN BelgeKategori k ON k.Id = bt.Kategori
WHERE 1=1";
        var p = new DynamicParameters();
        if (tur.HasValue)        { sql += " AND k.Tur = @tur";      p.Add("tur", (int)tur.Value); }
        if (kategoriId.HasValue) { sql += " AND bt.Kategori = @kid"; p.Add("kid", kategoriId.Value); }
        if (aktif.HasValue)      { sql += " AND bt.Aktif = @akt";    p.Add("akt", aktif.Value ? 1 : 0); }
        sql += " ORDER BY k.Tur, k.Ad, bt.Ad";
        return conn.Query<BelgeTipi>(sql, p).ToList();
    }

    public BelgeTipi? BelgeTipiGetir(int id)
    {
        using var conn = Open();
        return conn.QueryFirstOrDefault<BelgeTipi>(@"
SELECT bt.Id, bt.Ad, bt.Kategori AS KategoriId, bt.Aktif, bt.OlusturmaTarihi,
       k.Ad AS KategoriAdi, k.Tur AS KategoriTuru
FROM BelgeTipi bt
LEFT JOIN BelgeKategori k ON k.Id = bt.Kategori
WHERE bt.Id=@id", new { id });
    }

    public int BelgeTipiEkle(string ad, int kategoriId)
    {
        using var conn = Open();
        return (int)conn.ExecuteScalar<long>(
            @"INSERT INTO BelgeTipi(Ad, Kategori, Aktif, OlusturmaTarihi)
              VALUES(@ad, @kat, 1, @t);
              SELECT last_insert_rowid();",
            new { ad, kat = kategoriId, t = DateTime.Now.ToString("O") });
    }

    public void BelgeTipiGuncelle(int id, string ad, int kategoriId, bool aktif)
    {
        using var conn = Open();
        conn.Execute("UPDATE BelgeTipi SET Ad=@ad, Kategori=@kat, Aktif=@aktif WHERE Id=@id",
            new { id, ad, kat = kategoriId, aktif = aktif ? 1 : 0 });
    }

    public void BelgeTipiSil(int id)
    {
        using var conn = Open();
        conn.Execute("DELETE FROM BelgeTipi WHERE Id=@id", new { id });
    }

    // ========== NUMARATÖR ==========

    private int SonrakiDefterSiraNo(BelgeKategorisi tur, int yil)
    {
        using var conn = Open();
        var tablo = tur switch
        {
            BelgeKategorisi.Tesis     => "TesisEvraki",
            BelgeKategorisi.Mustahsil => "MustahsilEvraki",
            BelgeKategorisi.Genel     => "GenelEvraki",
            _ => throw new ArgumentOutOfRangeException(nameof(tur))
        };
        var sql = $"SELECT IFNULL(MAX(DefterSiraNo), 0) FROM {tablo} WHERE DefterYili = @yil";
        var son = conn.ExecuteScalar<int>(sql, new { yil });
        return son + 1;
    }

    // ========== TESIS EVRAKI ==========

    public List<TesisEvraki> TesisEvrakListesi(int? tesisId = null, int? belgeTipiId = null, string? arama = null)
    {
        using var conn = Open();
        var sql = @"
SELECT e.*, t.Ad AS TesisAdi, bt.Ad AS BelgeTipiAdi
FROM TesisEvraki e
LEFT JOIN Tesis t       ON t.Id  = e.TesisId
LEFT JOIN BelgeTipi bt  ON bt.Id = e.BelgeTipiId
WHERE 1=1 ";
        var p = new DynamicParameters();
        if (tesisId.HasValue)     { sql += " AND e.TesisId = @tid";      p.Add("tid", tesisId.Value); }
        if (belgeTipiId.HasValue) { sql += " AND e.BelgeTipiId = @btid"; p.Add("btid", belgeTipiId.Value); }
        if (!string.IsNullOrWhiteSpace(arama))
        {
            sql += " AND (e.Aciklama LIKE @q OR e.EvrakNo LIKE @q OR e.DosyaAdi LIKE @q)";
            p.Add("q", "%" + arama.Trim() + "%");
        }
        sql += " ORDER BY e.DefterYili DESC, e.DefterSiraNo DESC";
        return conn.Query<TesisEvraki>(sql, p).ToList();
    }

    public TesisEvraki? TesisEvrakGetir(int id)
    {
        using var conn = Open();
        return conn.QueryFirstOrDefault<TesisEvraki>(@"
SELECT e.*, t.Ad AS TesisAdi, bt.Ad AS BelgeTipiAdi
FROM TesisEvraki e
LEFT JOIN Tesis t      ON t.Id  = e.TesisId
LEFT JOIN BelgeTipi bt ON bt.Id = e.BelgeTipiId
WHERE e.Id = @id", new { id });
    }

    public int TesisEvrakEkle(TesisEvraki e)
    {
        lock (_lock)
        {
            e.DefterYili   = e.TebellugTarihi.Year;
            e.DefterSiraNo = SonrakiDefterSiraNo(BelgeKategorisi.Tesis, e.DefterYili);

            using var conn = Open();
            return (int)conn.ExecuteScalar<long>(@"
INSERT INTO TesisEvraki(
    DefterYili, DefterSiraNo, TesisId, BelgeTipiId,
    DosyaAdi, DosyaYolu, MimeType, DosyaBoyutu,
    EvrakTarihi, EvrakNo, TebellugTarihi, GecerlilikBaslangic, GecerlilikBitis,
    Aciklama, YuklemeTarihi)
VALUES(
    @DefterYili, @DefterSiraNo, @TesisId, @BelgeTipiId,
    @DosyaAdi, @DosyaYolu, @MimeType, @DosyaBoyutu,
    @EvrakTarihi, @EvrakNo, @TebellugTarihi, @GecerlilikBaslangic, @GecerlilikBitis,
    @Aciklama, @YuklemeTarihi);
SELECT last_insert_rowid();", e);
        }
    }

    public void TesisEvrakSil(int id)
    {
        var e = TesisEvrakGetir(id);
        if (e == null) return;
        using var conn = Open();
        conn.Execute("DELETE FROM TesisEvraki WHERE Id=@id", new { id });
        TryDelete(e.DosyaYolu);
    }

    // ========== MUSTAHSIL EVRAKI ==========

    public List<MustahsilEvraki> MustahsilEvrakListesi(int? kampanyaYili = null, int? belgeTipiId = null,
        string? tcKimlikNo = null, string? arama = null)
    {
        using var conn = Open();
        var sql = @"
SELECT e.*, bt.Ad AS BelgeTipiAdi
FROM MustahsilEvraki e
LEFT JOIN BelgeTipi bt ON bt.Id = e.BelgeTipiId
WHERE 1=1 ";
        var p = new DynamicParameters();
        if (kampanyaYili.HasValue)  { sql += " AND e.KampanyaYili = @yil";  p.Add("yil", kampanyaYili.Value); }
        if (belgeTipiId.HasValue)   { sql += " AND e.BelgeTipiId = @btid";  p.Add("btid", belgeTipiId.Value); }
        if (!string.IsNullOrWhiteSpace(tcKimlikNo)) { sql += " AND e.TcKimlikNo = @tc"; p.Add("tc", tcKimlikNo); }
        if (!string.IsNullOrWhiteSpace(arama))
        {
            sql += " AND (e.Aciklama LIKE @q OR e.EvrakNo LIKE @q OR e.MustahsilAdSoyadi LIKE @q OR e.DosyaAdi LIKE @q OR e.TcKimlikNo LIKE @q)";
            p.Add("q", "%" + arama.Trim() + "%");
        }
        sql += " ORDER BY e.DefterYili DESC, e.DefterSiraNo DESC";
        return conn.Query<MustahsilEvraki>(sql, p).ToList();
    }

    public MustahsilEvraki? MustahsilEvrakGetir(int id)
    {
        using var conn = Open();
        return conn.QueryFirstOrDefault<MustahsilEvraki>(@"
SELECT e.*, bt.Ad AS BelgeTipiAdi
FROM MustahsilEvraki e
LEFT JOIN BelgeTipi bt ON bt.Id = e.BelgeTipiId
WHERE e.Id = @id", new { id });
    }

    public int MustahsilEvrakEkle(MustahsilEvraki e)
    {
        lock (_lock)
        {
            e.DefterYili   = e.TebellugTarihi.Year;
            e.DefterSiraNo = SonrakiDefterSiraNo(BelgeKategorisi.Mustahsil, e.DefterYili);

            using var conn = Open();
            return (int)conn.ExecuteScalar<long>(@"
INSERT INTO MustahsilEvraki(
    DefterYili, DefterSiraNo, KampanyaYili, TcKimlikNo, MustahsilAdSoyadi, MustahsilNo, HesapNo,
    BelgeTipiId, DosyaAdi, DosyaYolu, MimeType, DosyaBoyutu,
    EvrakTarihi, EvrakNo, TebellugTarihi, Aciklama, YuklemeTarihi)
VALUES(
    @DefterYili, @DefterSiraNo, @KampanyaYili, @TcKimlikNo, @MustahsilAdSoyadi, @MustahsilNo, @HesapNo,
    @BelgeTipiId, @DosyaAdi, @DosyaYolu, @MimeType, @DosyaBoyutu,
    @EvrakTarihi, @EvrakNo, @TebellugTarihi, @Aciklama, @YuklemeTarihi);
SELECT last_insert_rowid();", e);
        }
    }

    public void MustahsilEvrakSil(int id)
    {
        var e = MustahsilEvrakGetir(id);
        if (e == null) return;
        using var conn = Open();
        conn.Execute("DELETE FROM MustahsilEvraki WHERE Id=@id", new { id });
        TryDelete(e.DosyaYolu);
    }

    // ========== GENEL EVRAKI ==========

    public List<GenelEvraki> GenelEvrakListesi(int? kategoriId = null, int? belgeTipiId = null, string? arama = null)
    {
        using var conn = Open();
        var sql = @"
SELECT e.*, k.Ad AS KategoriAdi, bt.Ad AS BelgeTipiAdi
FROM GenelEvraki e
LEFT JOIN BelgeKategori k ON k.Id = e.KategoriId
LEFT JOIN BelgeTipi bt    ON bt.Id = e.BelgeTipiId
WHERE 1=1 ";
        var p = new DynamicParameters();
        if (kategoriId.HasValue)  { sql += " AND e.KategoriId = @kid";   p.Add("kid", kategoriId.Value); }
        if (belgeTipiId.HasValue) { sql += " AND e.BelgeTipiId = @btid"; p.Add("btid", belgeTipiId.Value); }
        if (!string.IsNullOrWhiteSpace(arama))
        {
            sql += " AND (e.Aciklama LIKE @q OR e.FaturaNo LIKE @q OR e.CariUnvani LIKE @q OR e.DosyaAdi LIKE @q)";
            p.Add("q", "%" + arama.Trim() + "%");
        }
        sql += " ORDER BY e.DefterYili DESC, e.DefterSiraNo DESC";
        return conn.Query<GenelEvraki>(sql, p).ToList();
    }

    public GenelEvraki? GenelEvrakGetir(int id)
    {
        using var conn = Open();
        return conn.QueryFirstOrDefault<GenelEvraki>(@"
SELECT e.*, k.Ad AS KategoriAdi, bt.Ad AS BelgeTipiAdi
FROM GenelEvraki e
LEFT JOIN BelgeKategori k ON k.Id = e.KategoriId
LEFT JOIN BelgeTipi bt    ON bt.Id = e.BelgeTipiId
WHERE e.Id=@id", new { id });
    }

    public int GenelEvrakEkle(GenelEvraki e)
    {
        lock (_lock)
        {
            e.DefterYili   = e.TebellugTarihi.Year;
            e.DefterSiraNo = SonrakiDefterSiraNo(BelgeKategorisi.Genel, e.DefterYili);

            using var conn = Open();
            return (int)conn.ExecuteScalar<long>(@"
INSERT INTO GenelEvraki(
    DefterYili, DefterSiraNo, KategoriId, BelgeTipiId,
    CariUnvani, FaturaNo, EvrakTarihi, Tutar,
    DosyaAdi, DosyaYolu, MimeType, DosyaBoyutu,
    TebellugTarihi, Aciklama, YuklemeTarihi)
VALUES(
    @DefterYili, @DefterSiraNo, @KategoriId, @BelgeTipiId,
    @CariUnvani, @FaturaNo, @EvrakTarihi, @Tutar,
    @DosyaAdi, @DosyaYolu, @MimeType, @DosyaBoyutu,
    @TebellugTarihi, @Aciklama, @YuklemeTarihi);
SELECT last_insert_rowid();", e);
        }
    }

    public void GenelEvrakSil(int id)
    {
        var e = GenelEvrakGetir(id);
        if (e == null) return;
        using var conn = Open();
        conn.Execute("DELETE FROM GenelEvraki WHERE Id=@id", new { id });
        TryDelete(e.DosyaYolu);
    }

    // ========== DOSYA ==========

    public async Task<(string fullPath, string relativePath)> DosyaKaydetTesisAsync(
        string tesisAdi, string belgeTipiAdi, string orjinalAd, Stream icerik)
    {
        var safeTesis = Sanitize(tesisAdi);
        var safeTip   = Sanitize(belgeTipiAdi);
        var klasor    = Path.Combine(_dosyaKokDizini, "Tesisler", safeTesis, safeTip);
        Directory.CreateDirectory(klasor);
        return await DosyaKaydetAsync(klasor, orjinalAd, icerik);
    }

    public async Task<(string fullPath, string relativePath)> DosyaKaydetMustahsilAsync(
        int yil, string mustahsilAdSoyad, string belgeTipiAdi, string orjinalAd, Stream icerik)
    {
        var safeMust = Sanitize(mustahsilAdSoyad);
        var safeTip  = Sanitize(belgeTipiAdi);
        var klasor   = Path.Combine(_dosyaKokDizini, "Mustahsiller", yil.ToString(), safeMust, safeTip);
        Directory.CreateDirectory(klasor);
        return await DosyaKaydetAsync(klasor, orjinalAd, icerik);
    }

    public async Task<(string fullPath, string relativePath)> DosyaKaydetGenelAsync(
        string kategoriAdi, string belgeTipiAdi, string orjinalAd, Stream icerik)
    {
        var safeKat = Sanitize(kategoriAdi);
        var safeTip = Sanitize(belgeTipiAdi);
        var klasor  = Path.Combine(_dosyaKokDizini, "Genel", safeKat, safeTip);
        Directory.CreateDirectory(klasor);
        return await DosyaKaydetAsync(klasor, orjinalAd, icerik);
    }

    private async Task<(string fullPath, string relativePath)> DosyaKaydetAsync(string klasor, string orjinalAd, Stream icerik)
    {
        var uzanti    = Path.GetExtension(orjinalAd);
        var tabanAd   = Sanitize(Path.GetFileNameWithoutExtension(orjinalAd));
        if (string.IsNullOrWhiteSpace(tabanAd)) tabanAd = "dosya";
        var yeniAd    = $"{Guid.NewGuid():N}_{tabanAd}{uzanti}";
        var tamYol    = Path.Combine(klasor, yeniAd);
        using (var fs = new FileStream(tamYol, FileMode.Create, FileAccess.Write))
            await icerik.CopyToAsync(fs);
        var rel = Path.GetRelativePath(_dosyaKokDizini, tamYol).Replace('\\', '/');
        return (tamYol, rel);
    }

    public string TamYolaCevir(string relativePath)
        => Path.Combine(_dosyaKokDizini, relativePath.Replace('/', Path.DirectorySeparatorChar));

    private static void TryDelete(string relativePath)
    {
        try
        {
            var path = Path.Combine(AppDataPaths.EvrakArsivDizini,
                                    relativePath.Replace('/', Path.DirectorySeparatorChar));
            if (File.Exists(path)) File.Delete(path);
        }
        catch { }
    }

    private static string Sanitize(string s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "_";
        var invalid = Path.GetInvalidFileNameChars().Concat(new[] { '/', '\\' }).ToArray();
        var cleaned = new string(s.Select(c => invalid.Contains(c) ? '_' : c).ToArray());
        return cleaned.Trim().Trim('.').Length == 0 ? "_" : cleaned.Trim();
    }
}
