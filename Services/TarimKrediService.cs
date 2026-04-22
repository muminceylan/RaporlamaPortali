using Dapper;
using Microsoft.Data.Sqlite;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Tarım Kredi Raporu servisi.
/// - Kooperatif → Bölge eşleşmesi SQLite'ta saklanır (kullanıcı ekler/düzenler).
/// - Yan ürün hareketleri Logo view'ından (INF_UT_Kısıtlı_Malzeme_Raporu_Afyon_2025) çekilir.
/// - FATURA_NO kolonu view'da varsa kullanılır; yoksa FIS_NUMARASI geçerli sayılır.
/// </summary>
public class TarimKrediService
{
    private readonly DatabaseService _db;
    private readonly string _sqliteConn;

    public TarimKrediService(DatabaseService db)
    {
        _db = db;
        _sqliteConn = $"Data Source={AppDataPaths.TarimKrediDb}";
        InitDb();
    }

    private SqliteConnection OpenSqlite()
    {
        var c = new SqliteConnection(_sqliteConn);
        c.Open();
        return c;
    }

    private void InitDb()
    {
        using var conn = OpenSqlite();
        conn.Execute(@"
CREATE TABLE IF NOT EXISTS TarimKrediBolge (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    FirmaAdi TEXT NOT NULL,
    Bolge TEXT NOT NULL,
    CariKodu TEXT,
    OlusturmaTarihi TEXT NOT NULL,
    UNIQUE(FirmaAdi COLLATE NOCASE)
);
CREATE INDEX IF NOT EXISTS IX_TarimKrediBolge_Bolge ON TarimKrediBolge(Bolge);

CREATE TABLE IF NOT EXISTS TarimKrediBolgeTanimi (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Ad TEXT NOT NULL,
    OlusturmaTarihi TEXT NOT NULL,
    UNIQUE(Ad COLLATE NOCASE)
);
");

        var mevcut = conn.ExecuteScalar<long>("SELECT COUNT(*) FROM TarimKrediBolge");
        if (mevcut == 0) SeedYukle(conn);

        // Var olan TarimKrediBolgeTanimi tablosuna Email/İlgiliKisi/Telefon kolonlarını ekle
        var kolonlar = conn.Query<string>("SELECT name FROM pragma_table_info('TarimKrediBolgeTanimi')").ToList();
        if (!kolonlar.Contains("Email"))
            conn.Execute("ALTER TABLE TarimKrediBolgeTanimi ADD COLUMN Email TEXT");
        if (!kolonlar.Contains("IlgiliKisi"))
            conn.Execute("ALTER TABLE TarimKrediBolgeTanimi ADD COLUMN IlgiliKisi TEXT");
        if (!kolonlar.Contains("Telefon"))
            conn.Execute("ALTER TABLE TarimKrediBolgeTanimi ADD COLUMN Telefon TEXT");

        // Bölge tanımı tablosu boşsa, eşleşmelerden mevcut bölgeleri seed et
        var tanimSayisi = conn.ExecuteScalar<long>("SELECT COUNT(*) FROM TarimKrediBolgeTanimi");
        if (tanimSayisi == 0)
        {
            var now = DateTime.Now.ToString("s");
            conn.Execute(@"INSERT OR IGNORE INTO TarimKrediBolgeTanimi(Ad, OlusturmaTarihi)
                           SELECT DISTINCT Bolge, @N FROM TarimKrediBolge", new { N = now });
        }
    }

    private static void SeedYukle(SqliteConnection conn)
    {
        var now = DateTime.Now.ToString("s");
        using var tx = conn.BeginTransaction();
        foreach (var (firma, bolge) in Seed)
        {
            conn.Execute(@"INSERT OR IGNORE INTO TarimKrediBolge(FirmaAdi, Bolge, OlusturmaTarihi)
                           VALUES(@FirmaAdi, @Bolge, @OlusturmaTarihi)",
                new { FirmaAdi = firma, Bolge = bolge, OlusturmaTarihi = now }, tx);
        }
        tx.Commit();
    }

    // ============ CRUD Bölge Eşleşmesi ============

    public List<TarimKrediBolgeEslesme> BolgeListesi()
    {
        using var conn = OpenSqlite();
        return conn.Query<TarimKrediBolgeEslesme>(
            "SELECT Id, FirmaAdi, Bolge, CariKodu, OlusturmaTarihi FROM TarimKrediBolge ORDER BY Bolge, FirmaAdi"
        ).ToList();
    }

    public List<string> BolgeAdlari()
    {
        using var conn = OpenSqlite();
        return conn.Query<string>(@"
SELECT DISTINCT Ad AS Bolge FROM TarimKrediBolgeTanimi
UNION
SELECT DISTINCT Bolge FROM TarimKrediBolge
ORDER BY Bolge").ToList();
    }

    public List<TarimKrediBolgeTanim> BolgeTanimListesi()
    {
        using var conn = OpenSqlite();
        // Hem tanım tablosundan hem de sadece eşleşmede geçen bölgelerden birleşik liste
        var tanimlar = conn.Query<TarimKrediBolgeTanim>(@"
SELECT Id, Ad, Email, IlgiliKisi, Telefon, OlusturmaTarihi
FROM TarimKrediBolgeTanimi").ToList();

        var tanimliAdlar = new HashSet<string>(tanimlar.Select(t => t.Ad), StringComparer.OrdinalIgnoreCase);
        var eksikler = conn.Query<string>(
            "SELECT DISTINCT Bolge FROM TarimKrediBolge").ToList()
            .Where(b => !tanimliAdlar.Contains(b))
            .Select(b => new TarimKrediBolgeTanim { Id = 0, Ad = b, OlusturmaTarihi = DateTime.Now });

        return tanimlar.Concat(eksikler).OrderBy(t => t.Ad).ToList();
    }

    public int BolgeTanimiEkle(string ad, string? email = null, string? ilgiliKisi = null, string? telefon = null)
    {
        using var conn = OpenSqlite();
        return conn.Execute(@"INSERT OR IGNORE INTO TarimKrediBolgeTanimi(Ad, Email, IlgiliKisi, Telefon, OlusturmaTarihi)
                              VALUES(@Ad, @Email, @IlgiliKisi, @Telefon, @N)",
            new
            {
                Ad = ad.Trim().ToUpper(System.Globalization.CultureInfo.GetCultureInfo("tr-TR")),
                Email = string.IsNullOrWhiteSpace(email) ? null : email!.Trim(),
                IlgiliKisi = string.IsNullOrWhiteSpace(ilgiliKisi) ? null : ilgiliKisi!.Trim(),
                Telefon = string.IsNullOrWhiteSpace(telefon) ? null : telefon!.Trim(),
                N = DateTime.Now.ToString("s")
            });
    }

    public int BolgeTanimiGuncelle(int id, string ad, string? email, string? ilgiliKisi, string? telefon)
    {
        using var conn = OpenSqlite();
        var yeniAd = ad.Trim().ToUpper(System.Globalization.CultureInfo.GetCultureInfo("tr-TR"));

        // Ad değişmişse, eşleşmelerdeki bölge adını da güncelle
        var eski = conn.QuerySingleOrDefault<string>(
            "SELECT Ad FROM TarimKrediBolgeTanimi WHERE Id = @Id", new { Id = id });

        var sonuc = conn.Execute(@"UPDATE TarimKrediBolgeTanimi
                                   SET Ad = @Ad, Email = @Email, IlgiliKisi = @IlgiliKisi, Telefon = @Telefon
                                   WHERE Id = @Id",
            new
            {
                Id = id,
                Ad = yeniAd,
                Email = string.IsNullOrWhiteSpace(email) ? null : email!.Trim(),
                IlgiliKisi = string.IsNullOrWhiteSpace(ilgiliKisi) ? null : ilgiliKisi!.Trim(),
                Telefon = string.IsNullOrWhiteSpace(telefon) ? null : telefon!.Trim()
            });

        if (!string.IsNullOrEmpty(eski) && !string.Equals(eski, yeniAd, StringComparison.OrdinalIgnoreCase))
        {
            conn.Execute("UPDATE TarimKrediBolge SET Bolge = @Yeni WHERE Bolge = @Eski COLLATE NOCASE",
                new { Yeni = yeniAd, Eski = eski });
        }
        return sonuc;
    }

    public int BolgeTanimiSil(int id)
    {
        using var conn = OpenSqlite();
        var ad = conn.QuerySingleOrDefault<string>(
            "SELECT Ad FROM TarimKrediBolgeTanimi WHERE Id = @Id", new { Id = id });
        if (string.IsNullOrEmpty(ad)) return 0;

        var kullanim = conn.ExecuteScalar<long>(
            "SELECT COUNT(*) FROM TarimKrediBolge WHERE Bolge = @Ad COLLATE NOCASE", new { Ad = ad });
        if (kullanim > 0) throw new InvalidOperationException(
            $"\"{ad}\" bölgesi {kullanim} firma eşleşmesinde kullanılıyor. Önce o firmaları düzenleyin.");
        return conn.Execute("DELETE FROM TarimKrediBolgeTanimi WHERE Id = @Id", new { Id = id });
    }

    public int BolgeTanimiSil(string ad)
    {
        using var conn = OpenSqlite();
        var kullanim = conn.ExecuteScalar<long>(
            "SELECT COUNT(*) FROM TarimKrediBolge WHERE Bolge = @Ad COLLATE NOCASE", new { Ad = ad });
        if (kullanim > 0) throw new InvalidOperationException(
            $"\"{ad}\" bölgesi {kullanim} firma eşleşmesinde kullanılıyor. Önce o firmaları düzenleyin.");
        return conn.Execute("DELETE FROM TarimKrediBolgeTanimi WHERE Ad = @Ad COLLATE NOCASE", new { Ad = ad });
    }

    public int BolgeEkle(string firmaAdi, string bolge, string? cariKodu = null)
    {
        using var conn = OpenSqlite();
        return conn.Execute(@"INSERT INTO TarimKrediBolge(FirmaAdi, Bolge, CariKodu, OlusturmaTarihi)
                              VALUES(@FirmaAdi, @Bolge, @CariKodu, @OlusturmaTarihi)",
            new
            {
                FirmaAdi = firmaAdi.Trim(),
                Bolge = bolge.Trim().ToUpper(System.Globalization.CultureInfo.GetCultureInfo("tr-TR")),
                CariKodu = string.IsNullOrWhiteSpace(cariKodu) ? null : cariKodu!.Trim(),
                OlusturmaTarihi = DateTime.Now.ToString("s")
            });
    }

    public int BolgeGuncelle(int id, string firmaAdi, string bolge, string? cariKodu)
    {
        using var conn = OpenSqlite();
        return conn.Execute(@"UPDATE TarimKrediBolge
                              SET FirmaAdi = @FirmaAdi, Bolge = @Bolge, CariKodu = @CariKodu
                              WHERE Id = @Id",
            new
            {
                Id = id,
                FirmaAdi = firmaAdi.Trim(),
                Bolge = bolge.Trim().ToUpper(System.Globalization.CultureInfo.GetCultureInfo("tr-TR")),
                CariKodu = string.IsNullOrWhiteSpace(cariKodu) ? null : cariKodu!.Trim()
            });
    }

    public int BolgeSil(int id)
    {
        using var conn = OpenSqlite();
        return conn.Execute("DELETE FROM TarimKrediBolge WHERE Id = @Id", new { Id = id });
    }

    // ============ Logo Veri Çekme ============

    /// <summary>
    /// Ünvan içinde "TARIM KRED" veya "PANCAR EK" geçen yan ürün hareketleri.
    /// Fatura numarası: STFICHE.INVOICEREF → INVOICE.FICHENO
    /// </summary>
    public async Task<List<YanUrunHareket>> TarimKrediHareketleriAsync(
        DateTime baslangic, DateTime bitis, bool pancarEkicileriDahil = true)
    {
        bitis = SistemTarihi.Clamp(bitis);
        var unvanFiltre = pancarEkicileriDahil
            ? "(v.CARI_HESAP_UNVANI LIKE N'%TARIM KRED%' OR v.CARI_HESAP_UNVANI LIKE N'%PANCAR EK%')"
            : "v.CARI_HESAP_UNVANI LIKE N'%TARIM KRED%'";

        var view = _db.LogoMalzemeTablo;
        var stfTbl = _db.GetPeriodTableName("STFICHE"); // LG_211_01_STFICHE
        var invTbl = _db.GetPeriodTableName("INVOICE"); // LG_211_01_INVOICE

        var sql = $@"
SELECT
    Tarih              = v.TARIH,
    FisTuru            = v.FIS_TURU,
    FisNumarasi        = ISNULL(v.FIS_NUMARASI, ''),
    FaturaNo           = ISNULL(INV.FICHENO, ''),
    CariHesapKodu      = ISNULL(v.CARI_HESAP_KODU, ''),
    CariHesapUnvani    = ISNULL(v.CARI_HESAP_UNVANI, ''),
    MalzemeKodu        = ISNULL(v.MALZEME_KODU, ''),
    MalzemeAciklamasi  = ISNULL(v.MALZEME_ACIKLAMASI, ''),
    GirisMiktari       = ISNULL(v.GIRIS_MIKTARI, 0),
    GirisFiyati        = ISNULL(v.GIRIS_FIYATI, 0),
    GirisTutari        = ISNULL(v.GIRIS_TUTARI, 0),
    CikisMiktari       = ISNULL(v.CIKIS_MIKTARI, 0),
    CikisFiyati        = ISNULL(v.CIKIS_FIYATI, 0),
    CikisTutari        = ISNULL(v.CIKIS_TUTARI, 0)
FROM {view} v WITH(NOLOCK)
LEFT JOIN {stfTbl} STF WITH(NOLOCK)
       ON STF.FICHENO = v.FIS_NUMARASI
      AND CONVERT(date, STF.DATE_) = CONVERT(date, v.TARIH)
LEFT JOIN {invTbl} INV WITH(NOLOCK)
       ON INV.LOGICALREF = STF.INVOICEREF
WHERE v.TARIH >= @Bas AND v.TARIH <= @Bit
  AND {unvanFiltre}
  AND (v.FIS_TURU LIKE N'%Satış%' OR v.FIS_TURU LIKE N'%Toptan%' OR v.FIS_TURU LIKE N'%İade%')
ORDER BY v.CARI_HESAP_UNVANI, v.TARIH, v.FIS_NUMARASI";

        return await CalistirAsync(sql, baslangic, bitis);
    }

    private async Task<List<YanUrunHareket>> CalistirAsync(string sql, DateTime baslangic, DateTime bitis)
    {
        using var conn = _db.CreateConnection();
        var rows = await conn.QueryAsync<YanUrunHareket>(sql, new { Bas = baslangic.Date, Bit = bitis.Date });
        return rows.ToList();
    }

    /// <summary>
    /// Hareketleri bölge → firma → satır yapısına dönüştürür.
    /// Eşleşmesi tanımlı olmayan firmalar "EŞLEŞMEMİŞ" bölgesine düşer.
    /// </summary>
    public List<TarimKrediBolgeRapor> BolgeyeGore(List<YanUrunHareket> hareketler)
    {
        var map = BolgeListesi()
            .GroupBy(b => b.FirmaAdi, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key.ToUpperInvariant(), g => g.First().Bolge, StringComparer.OrdinalIgnoreCase);

        var bolgeler = new Dictionary<string, TarimKrediBolgeRapor>(StringComparer.OrdinalIgnoreCase);

        foreach (var grp in hareketler.GroupBy(h => new { h.CariHesapKodu, h.CariHesapUnvani }))
        {
            var unvan = grp.Key.CariHesapUnvani.Trim();
            var bolge = map.TryGetValue(unvan, out var b) ? b : "EŞLEŞMEMİŞ";

            if (!bolgeler.TryGetValue(bolge, out var rapor))
                bolgeler[bolge] = rapor = new TarimKrediBolgeRapor { Bolge = bolge };

            rapor.Firmalar.Add(new TarimKrediFirmaOzet
            {
                CariHesapKodu = grp.Key.CariHesapKodu,
                CariHesapUnvani = unvan,
                Hareketler = grp.OrderBy(h => h.Tarih).ThenBy(h => h.FisNumarasi).ToList()
            });
        }

        foreach (var r in bolgeler.Values)
            r.Firmalar = r.Firmalar.OrderBy(f => f.CariHesapUnvani).ToList();

        return bolgeler.Values
            .OrderBy(r => r.Bolge == "EŞLEŞMEMİŞ" ? 1 : 0)
            .ThenBy(r => r.Bolge)
            .ToList();
    }

    // ============ Seed Data (77 kooperatif) ============

    private static readonly (string Firma, string Bolge)[] Seed = new (string, string)[]
    {
        ("0808 SAYILI SUSUZMÜSELLİM TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("1083 SAYILI GÜLLÜCE TARIM KREDİ KOOPERATİFİ", "TRABZON"),
        ("1102 SAYILI SÜRMELİ TARIM KREDİ KOOPERATİFİ", "SAMSUN"),
        ("1117 SAYILI ÖRENCİK TARIM KREDİ KOOPERATİFİ", "SAMSUN"),
        ("1168 SAYILI DAMBASLAR TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("1363 SAYILI ALAÇAM TARIM KREDİ KOOPERATİFİ", "SAMSUN"),
        ("1407 SAYILI BARDAKÇI TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("1428 SAYILI SÜLOĞLU TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("1432 SAYILI DOĞANTEPE TARIM KREDİ KOOPERATİFİ", "SAMSUN"),
        ("1442 SAYILI GÖYNÜCEK TARIM KREDİ KOOPERATİFİ", "SAMSUN"),
        ("1452 SAYILI TURGUTBEY TARIM KREDİ KOOPERATİFİ MÜDÜRLÜĞÜ", "TEKİRDAĞ"),
        ("1462 SAYILI ÇAMLICA TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("1866 SAYILI KIZILİNLER TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("188 SAYILI ŞARKÖY TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("1952 SAYILI BALLICA TARIM KREDİ KOOPERATİFİ", "SAMSUN"),
        ("1957 SAYILI AŞAĞIPİRİBEYLİ TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("2010 SAYILI KANDAMIŞ TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("2031 SAYILI AKKONAK TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("2101 SAYILI ALANYURT TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("2210 SAYILI FERHADANLI TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("2233 SAYILI KOZYÖRÜK TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("2253 SAYILI İSAALAN TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("2339 SAYILI KAŞIKÇI TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("2379 SAYILI ŞALGAMLI TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("2620 SAYILI TEKKEKÖY TARIM KREDİ KOOPERATİFİ", "SAMSUN"),
        ("2654 SAYILI EMİRDAĞ TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("2695 SAYILI HEMİT TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("2704 SAYILI ÖZDEMİRCİ TARIM KREDİ KOOPERATİFİ", "DENİZLİ"),
        ("277 SAYILI EDİRNE TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("2884 SAYILI GERZE TARIM KREDİ KOOPERATİFİ", "SAMSUN"),
        ("326 SAYILI KEŞAN TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("347 SAYILI KURŞUNLU TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("369 SAYILI BABAESKİ TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("429 SAYILI GÜRSU TARIM KREDİ KOOPERTİFİ", "BALIKESİR"),
        ("442 SAYILI GÖRÜKLE TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("475 SAYILI KELES TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("482 SAYILI KIZILCASÖĞÜT TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("526 SAYILI AKHİSAR TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("528 SAYILI ORTAKÖY TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("529 SAYILI ORHANELİ TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("552 NOLU HOCAKÖY TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("559 SAYILI YENİCE TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("568 SAYILI HAMİTABAT TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("678 SAYILI PEHLİVANKÖY TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("679 SAYILI HAVSA TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("680 SAYILI LALAPAŞA TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("681 SAYILI TATARLAR TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("696 SAYILI İNECE TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("697 SAYILI YOĞUNTAŞ TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("725 SAYILI KIZILCIKDERE TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("750 SAYILI HİSARBEY TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("806 SAYILI İNEGÖL TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("807 SAYILI TAHTAKÖPRÜ TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("822 SAYILI YAKAKENT TARIM KREDİ KOOPERATİFİ", "SAMSUN"),
        ("823 SAYILI GÜRPINAR TARIM KREDİ KOOPERATİFİ", "DENİZLİ"),
        ("829 SAYILI MELİK TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("830 SAYILI KARAORMAN TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("846 SAYILI KARAHALİL TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("850 SAYILI DEVECİKONAĞI TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("852 SAYILI ŞAHİNBUCAĞI TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("897 SAYILI ÇATALCA TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("919 SAYILI SABUNCUPINAR TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("939 SAYILI KIRKA TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("993 SAYILI KULELİ TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("914 SAYILI DOMANİÇ TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("1635 SAYILI YAPRAKLI TARIM KREDİ KOOPERATİFİ", "ANKARA"),
        ("2309 SAYILI HAN TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("1207 SAYILI GERMEÇ TARIM KREDİ KOOPERATİFİ", "ANKARA"),
        ("2013 SAYILI MESCİT TARIM KREDİ KOOPERATİFİ", "ANKARA"),
        ("1762 SAYILI KORGUN TARIM KREDİ KOOPERATİFİ", "ANKARA"),
        ("1455 SAYILI YUNUSEMRE TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("842 SAYILI SÖĞÜTALAN TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("2572 SAYILI SİVRİHİSAR TARIM KREDİ KOOPERATİFİ", "KÜTAHYA"),
        ("343 SAYILI KIRKLARELİ TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("398 SAYILI ÇALI TARIM KREDİ KOOPERATİFİ", "BALIKESİR"),
        ("1312 SAYILI HASKÖY TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
        ("1038 SAYILI BÜYÜKEVREN TARIM KREDİ KOOPERATİFİ", "TEKİRDAĞ"),
    };
}
