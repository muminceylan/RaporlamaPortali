using System.Diagnostics;
using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// SabNetKANTAR SQL Server'dan SabNet.db SQLite'a toplu veri aktarımı.
/// </summary>
public class SabNetImportService
{
    private readonly SabNetDbService _db;
    private readonly SabNetBaglantiService _baglanti;

    public SabNetImportService(SabNetDbService db, SabNetBaglantiService baglanti)
    {
        _db = db;
        _baglanti = baglanti;
    }

    public async Task<(bool ok, string mesaj, int hareketSayisi, int logSayisi)> BaglantiTestAsync(SabNetBaglantiAyari? ayar = null)
    {
        try
        {
            var connStr = ayar != null ? BuildConnStr(ayar) : _baglanti.GetConnectionString();
            await using var conn = new SqlConnection(connStr);
            await conn.OpenAsync();
            await using var cmd = new SqlCommand(
                "SELECT (SELECT COUNT(*) FROM dbo.PMHS_KantarHareketleri), (SELECT COUNT(*) FROM dbo.PMHS_KantarHareketleri_Log)",
                conn);
            await using var reader = await cmd.ExecuteReaderAsync();
            int kh = 0, lg = 0;
            if (await reader.ReadAsync())
            {
                kh = reader.GetInt32(0);
                lg = reader.GetInt32(1);
            }
            return (true, $"Bağlantı OK. Hareket: {kh:N0} | Log: {lg:N0}", kh, lg);
        }
        catch (Exception ex)
        {
            return (false, ex.Message, 0, 0);
        }
    }

    private static string BuildConnStr(SabNetBaglantiAyari a) =>
        $"Server={a.Server};Database={a.Database};User Id={a.Username};Password={a.Password};TrustServerCertificate=True;Connection Timeout=30;";

    public async Task<int> HareketleriAktarAsync(Action<int, string>? ilerlemeUpdater = null, CancellationToken ct = default)
    {
        int toplam = await SqlScalarAsync<int>("SELECT COUNT(*) FROM dbo.PMHS_KantarHareketleri", ct);
        ilerlemeUpdater?.Invoke(1, $"{toplam:N0} kayıt aktarılacak");

        var sql = @"SELECT Tarih, FisNo, RandevuNo, IslemTipi, UrunKodu, BirimFiyat,
            HesapTipi, TcKimlikNo, HesapKodu, SozlesmeYili,
            Adres, PlakaNo, SoforAdiSoyadi, SoforGsmNo, AracTipi,
            MuteahhitKodu, MouseKodu, Aciklama,
            Sevk, Brut, Dara, Fireli, Net,
            FireOrani, PolarOrani, FireMiktari, Fark,
            BosaltmaYeri, BosaltmaSekli,
            Cikis_Tarihi, Cikis_Saati, Cikis_KullaniciKodu,
            Giris_Tarihi, Giris_Saati, Giris_KullaniciKodu,
            IrsaliyeNo, FaturaNo, SiparisNo, KantarKodu, Branda, KapiKagidiNo,
            Kod1, Kod2, Kod3, Kod4, Kod5, Kod6, Kod7,
            Nakit, KrediKarti, Cari, Havale,
            Kaydeden, KayitTarihi, KayitSaati,
            OnKayit_HostName, OnKayit_ipAdres,
            Giris_HostName, Giris_ipAdres,
            Cikis_HostName, Cikis_ipAdres,
            Row_ID, Kontrol
        FROM dbo.PMHS_KantarHareketleri";

        return await BulkAktarAsync(sql, "SabNetKantarHareketleri", toplam, isLog: false, ilerlemeUpdater, ct);
    }

    public async Task<int> LoglariAktarAsync(Action<int, string>? ilerlemeUpdater = null, CancellationToken ct = default)
    {
        int toplam = await SqlScalarAsync<int>("SELECT COUNT(*) FROM dbo.PMHS_KantarHareketleri_Log", ct);
        ilerlemeUpdater?.Invoke(1, $"{toplam:N0} kayıt aktarılacak");

        var sql = @"SELECT Tarih, FisNo, RandevuNo, IslemTipi, UrunKodu, BirimFiyat,
            HesapTipi, TcKimlikNo, HesapKodu, SozlesmeYili,
            Adres, PlakaNo, SoforAdiSoyadi, SoforGsmNo, AracTipi,
            MuteahhitKodu, MouseKodu, Aciklama,
            Sevk, Brut, Dara, Fireli, Net,
            FireOrani, PolarOrani, FireMiktari, Fark,
            BosaltmaYeri, BosaltmaSekli,
            Cikis_Tarihi, Cikis_Saati, Cikis_KullaniciKodu,
            Giris_Tarihi, Giris_Saati, Giris_KullaniciKodu,
            IrsaliyeNo, FaturaNo, SiparisNo, KantarKodu, Branda, KapiKagidiNo,
            Kod1, Kod2, Kod3, Kod4, Kod5, Kod6, Kod7,
            Nakit, KrediKarti, Cari, Havale,
            Kaydeden, KayitTarihi, KayitSaati,
            OnKayit_HostName, OnKayit_ipAdres,
            Giris_HostName, Giris_ipAdres,
            Cikis_HostName, Cikis_ipAdres,
            Row_ID, Kontrol,
            LogIslemTipi, LogKaydeden, LogKayitTarihi, LogKayitSaati,
            LogAciklama, LogHostName, LogipAdres
        FROM dbo.PMHS_KantarHareketleri_Log";

        return await BulkAktarAsync(sql, "SabNetKantarHareketleriLog", toplam, isLog: true, ilerlemeUpdater, ct);
    }

    public async Task<(int n, double saniye)> CariAktarAsync(CancellationToken ct = default)
    {
        var sw = Stopwatch.StartNew();
        await using var src = new SqlConnection(_baglanti.GetConnectionString());
        await src.OpenAsync(ct);
        await using var cmd = new SqlCommand(@"
            SELECT HesapKodu, Unvan1, Unvan2, GrupKodu, VergiDairesi, VergiNo,
                   FaturaAdres1, FaturaAdres2, FaturaAdres3, ErpHesapKodu,
                   Gsm, Telefon1, Aciklama, Kaydeden, KayitTarihi, KayitSaati
            FROM dbo.PMHS_CariHesapKarti", src);
        cmd.CommandTimeout = 600;
        await using var reader = await cmd.ExecuteReaderAsync(ct);

        await using var dst = _db.CreateConnection();
        await using var tx = (SqliteTransaction)await dst.BeginTransactionAsync(ct);
        await using var insert = dst.CreateCommand();
        insert.Transaction = tx;
        insert.CommandText = @"INSERT INTO SabNetCariHesaplar
            (HesapKodu, Unvan1, Unvan2, GrupKodu, VergiDairesi, VergiNo,
             FaturaAdres1, FaturaAdres2, FaturaAdres3, ErpHesapKodu,
             Gsm, Telefon1, Aciklama, Kaydeden, KayitTarihi, KayitSaati)
            VALUES (@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11,@p12,@p13,@p14,@p15)";
        var prms = new SqliteParameter[16];
        for (int i = 0; i < 16; i++)
        {
            prms[i] = insert.CreateParameter();
            prms[i].ParameterName = $"@p{i}";
            insert.Parameters.Add(prms[i]);
        }
        insert.Prepare();

        int n = 0;
        while (await reader.ReadAsync(ct))
        {
            for (int i = 0; i < 16; i++)
                prms[i].Value = reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i);
            await insert.ExecuteNonQueryAsync(ct);
            n++;
        }
        await tx.CommitAsync(ct);
        sw.Stop();
        return (n, sw.Elapsed.TotalSeconds);
    }

    public async Task<(int n, double saniye)> StokAktarAsync(CancellationToken ct = default)
    {
        var sw = Stopwatch.StartNew();
        await using var src = new SqlConnection(_baglanti.GetConnectionString());
        await src.OpenAsync(ct);
        await using var cmd = new SqlCommand(@"
            SELECT StokKodu, StokAdi, Birim, BirimFiyat, KdvOrani, DoluBos,
                   Fis, Irsaliye, Fatura, AmbarOnayi, KapiKagidi, Aciklama,
                   ERPKodu, GrupKodu, StokTakibi, Kaydeden, KayitTarihi, KayitSaati
            FROM dbo.PMHS_StokKarti", src);
        cmd.CommandTimeout = 600;
        await using var reader = await cmd.ExecuteReaderAsync(ct);

        await using var dst = _db.CreateConnection();
        await using var tx = (SqliteTransaction)await dst.BeginTransactionAsync(ct);
        await using var insert = dst.CreateCommand();
        insert.Transaction = tx;
        insert.CommandText = @"INSERT INTO SabNetStokKartlari
            (StokKodu, StokAdi, Birim, BirimFiyat, KdvOrani, DoluBos,
             Fis, Irsaliye, Fatura, AmbarOnayi, KapiKagidi, Aciklama,
             ERPKodu, GrupKodu, StokTakibi, Kaydeden, KayitTarihi, KayitSaati)
            VALUES (@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11,@p12,@p13,@p14,@p15,@p16,@p17)";
        var prms = new SqliteParameter[18];
        for (int i = 0; i < 18; i++)
        {
            prms[i] = insert.CreateParameter();
            prms[i].ParameterName = $"@p{i}";
            insert.Parameters.Add(prms[i]);
        }
        insert.Prepare();

        int n = 0;
        while (await reader.ReadAsync(ct))
        {
            for (int i = 0; i < 18; i++)
                prms[i].Value = reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i);
            await insert.ExecuteNonQueryAsync(ct);
            n++;
        }
        await tx.CommitAsync(ct);
        sw.Stop();
        return (n, sw.Elapsed.TotalSeconds);
    }

    /// <summary>
    /// Hareket tablosu için inkremental güncelleme:
    /// - Bugün (Tarih = bugünDelphi) ve dünden bugüne açık (Cikis_Tarihi = 0) kayıtları source'tan çeker
    /// - Local'de aynı kriterdeki satırları silip yeni veriyi yazar (UPSERT yerine delete+insert)
    /// Eski kayıtlara dokunmaz.
    /// </summary>
    public async Task<(int silinen, int eklenen, double saniye)> HareketleriGuncelleAsync(int gunGeriye = 1, CancellationToken ct = default)
    {
        var sw = Stopwatch.StartNew();
        int cutoff = DelphiSerialBugun() - gunGeriye;

        var sourceSql = @"SELECT Tarih, FisNo, RandevuNo, IslemTipi, UrunKodu, BirimFiyat,
            HesapTipi, TcKimlikNo, HesapKodu, SozlesmeYili,
            Adres, PlakaNo, SoforAdiSoyadi, SoforGsmNo, AracTipi,
            MuteahhitKodu, MouseKodu, Aciklama,
            Sevk, Brut, Dara, Fireli, Net,
            FireOrani, PolarOrani, FireMiktari, Fark,
            BosaltmaYeri, BosaltmaSekli,
            Cikis_Tarihi, Cikis_Saati, Cikis_KullaniciKodu,
            Giris_Tarihi, Giris_Saati, Giris_KullaniciKodu,
            IrsaliyeNo, FaturaNo, SiparisNo, KantarKodu, Branda, KapiKagidiNo,
            Kod1, Kod2, Kod3, Kod4, Kod5, Kod6, Kod7,
            Nakit, KrediKarti, Cari, Havale,
            Kaydeden, KayitTarihi, KayitSaati,
            OnKayit_HostName, OnKayit_ipAdres,
            Giris_HostName, Giris_ipAdres,
            Cikis_HostName, Cikis_ipAdres,
            Row_ID, Kontrol
        FROM dbo.PMHS_KantarHareketleri
        WHERE Tarih >= @cutoff OR Cikis_Tarihi IS NULL OR Cikis_Tarihi = 0";

        await using var src = new SqlConnection(_baglanti.GetConnectionString());
        await src.OpenAsync(ct);
        await using var srcCmd = new SqlCommand(sourceSql, src);
        srcCmd.Parameters.AddWithValue("@cutoff", cutoff);
        srcCmd.CommandTimeout = 300;
        await using var reader = await srcCmd.ExecuteReaderAsync(ct);

        await using var dst = _db.CreateConnection();
        await using var tx = (SqliteTransaction)await dst.BeginTransactionAsync(ct);

        int silinen;
        await using (var del = dst.CreateCommand())
        {
            del.Transaction = tx;
            del.CommandText = "DELETE FROM SabNetKantarHareketleri WHERE Tarih >= $cutoff OR Cikis_Tarihi IS NULL OR Cikis_Tarihi = 0";
            del.Parameters.AddWithValue("$cutoff", cutoff);
            silinen = await del.ExecuteNonQueryAsync(ct);
        }

        var sutunListe = HareketSutunlari();
        int sutunSayisi = sutunListe.Split(',').Length;
        var paramListesi = string.Join(",", Enumerable.Range(0, sutunSayisi).Select(i => $"@p{i}"));
        var insertSql = $"INSERT INTO SabNetKantarHareketleri ({sutunListe}) VALUES ({paramListesi})";

        await using var ins = dst.CreateCommand();
        ins.Transaction = tx;
        ins.CommandText = insertSql;
        var prms = new SqliteParameter[sutunSayisi];
        for (int i = 0; i < sutunSayisi; i++)
        {
            prms[i] = ins.CreateParameter();
            prms[i].ParameterName = $"@p{i}";
            ins.Parameters.Add(prms[i]);
        }
        ins.Prepare();

        int eklenen = 0;
        while (await reader.ReadAsync(ct))
        {
            for (int i = 0; i < sutunSayisi; i++)
                prms[i].Value = reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i);
            await ins.ExecuteNonQueryAsync(ct);
            eklenen++;
        }
        await tx.CommitAsync(ct);
        sw.Stop();
        return (silinen, eklenen, sw.Elapsed.TotalSeconds);
    }

    /// <summary>
    /// Log tablosu inkremental: Row_ID > local maxRowId olan satırları append eder. Log immutable.
    /// </summary>
    public async Task<(int eklenen, double saniye)> LoglariGuncelleAsync(CancellationToken ct = default)
    {
        var sw = Stopwatch.StartNew();

        long maxLocalRowId;
        await using (var dstRead = _db.CreateConnection())
        await using (var cmdRead = dstRead.CreateCommand())
        {
            cmdRead.CommandText = "SELECT COALESCE(MAX(Row_ID), 0) FROM SabNetKantarHareketleriLog";
            var o = await cmdRead.ExecuteScalarAsync(ct);
            maxLocalRowId = o == null || o is DBNull ? 0L : Convert.ToInt64(o);
        }

        var sourceSql = @"SELECT Tarih, FisNo, RandevuNo, IslemTipi, UrunKodu, BirimFiyat,
            HesapTipi, TcKimlikNo, HesapKodu, SozlesmeYili,
            Adres, PlakaNo, SoforAdiSoyadi, SoforGsmNo, AracTipi,
            MuteahhitKodu, MouseKodu, Aciklama,
            Sevk, Brut, Dara, Fireli, Net,
            FireOrani, PolarOrani, FireMiktari, Fark,
            BosaltmaYeri, BosaltmaSekli,
            Cikis_Tarihi, Cikis_Saati, Cikis_KullaniciKodu,
            Giris_Tarihi, Giris_Saati, Giris_KullaniciKodu,
            IrsaliyeNo, FaturaNo, SiparisNo, KantarKodu, Branda, KapiKagidiNo,
            Kod1, Kod2, Kod3, Kod4, Kod5, Kod6, Kod7,
            Nakit, KrediKarti, Cari, Havale,
            Kaydeden, KayitTarihi, KayitSaati,
            OnKayit_HostName, OnKayit_ipAdres,
            Giris_HostName, Giris_ipAdres,
            Cikis_HostName, Cikis_ipAdres,
            Row_ID, Kontrol,
            LogIslemTipi, LogKaydeden, LogKayitTarihi, LogKayitSaati,
            LogAciklama, LogHostName, LogipAdres
        FROM dbo.PMHS_KantarHareketleri_Log
        WHERE Row_ID > @maxRow";

        await using var src = new SqlConnection(_baglanti.GetConnectionString());
        await src.OpenAsync(ct);
        await using var srcCmd = new SqlCommand(sourceSql, src);
        srcCmd.Parameters.AddWithValue("@maxRow", maxLocalRowId);
        srcCmd.CommandTimeout = 300;
        await using var reader = await srcCmd.ExecuteReaderAsync(ct);

        await using var dst = _db.CreateConnection();
        await using var tx = (SqliteTransaction)await dst.BeginTransactionAsync(ct);

        var sutunListe = HareketSutunlari() + LogEkSutunlari();
        int sutunSayisi = sutunListe.Split(',').Length;
        var paramListesi = string.Join(",", Enumerable.Range(0, sutunSayisi).Select(i => $"@p{i}"));
        var insertSql = $"INSERT INTO SabNetKantarHareketleriLog ({sutunListe}) VALUES ({paramListesi})";

        await using var ins = dst.CreateCommand();
        ins.Transaction = tx;
        ins.CommandText = insertSql;
        var prms = new SqliteParameter[sutunSayisi];
        for (int i = 0; i < sutunSayisi; i++)
        {
            prms[i] = ins.CreateParameter();
            prms[i].ParameterName = $"@p{i}";
            ins.Parameters.Add(prms[i]);
        }
        ins.Prepare();

        int eklenen = 0;
        while (await reader.ReadAsync(ct))
        {
            for (int i = 0; i < sutunSayisi; i++)
                prms[i].Value = reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i);
            await ins.ExecuteNonQueryAsync(ct);
            eklenen++;
        }
        await tx.CommitAsync(ct);
        sw.Stop();
        return (eklenen, sw.Elapsed.TotalSeconds);
    }

    private static int DelphiSerialBugun()
    {
        var baz = new DateTime(1899, 12, 30);
        return (int)(DateTime.Today - baz).TotalDays;
    }

    private static string HareketSutunlari() =>
        "Tarih, FisNo, RandevuNo, IslemTipi, UrunKodu, BirimFiyat, " +
        "HesapTipi, TcKimlikNo, HesapKodu, SozlesmeYili, " +
        "Adres, PlakaNo, SoforAdiSoyadi, SoforGsmNo, AracTipi, " +
        "MuteahhitKodu, MouseKodu, Aciklama, " +
        "Sevk, Brut, Dara, Fireli, Net, " +
        "FireOrani, PolarOrani, FireMiktari, Fark, " +
        "BosaltmaYeri, BosaltmaSekli, " +
        "Cikis_Tarihi, Cikis_Saati, Cikis_KullaniciKodu, " +
        "Giris_Tarihi, Giris_Saati, Giris_KullaniciKodu, " +
        "IrsaliyeNo, FaturaNo, SiparisNo, KantarKodu, Branda, KapiKagidiNo, " +
        "Kod1, Kod2, Kod3, Kod4, Kod5, Kod6, Kod7, " +
        "Nakit, KrediKarti, Cari, Havale, " +
        "Kaydeden, KayitTarihi, KayitSaati, " +
        "OnKayit_HostName, OnKayit_ipAdres, " +
        "Giris_HostName, Giris_ipAdres, " +
        "Cikis_HostName, Cikis_ipAdres, " +
        "Row_ID, Kontrol";

    private static string LogEkSutunlari() =>
        ", LogIslemTipi, LogKaydeden, LogKayitTarihi, LogKayitSaati, LogAciklama, LogHostName, LogipAdres";

    private async Task<T> SqlScalarAsync<T>(string sql, CancellationToken ct)
    {
        await using var conn = new SqlConnection(_baglanti.GetConnectionString());
        await conn.OpenAsync(ct);
        await using var cmd = new SqlCommand(sql, conn);
        cmd.CommandTimeout = 600;
        var o = await cmd.ExecuteScalarAsync(ct);
        if (o == null || o is DBNull) return default!;
        return (T)Convert.ChangeType(o, typeof(T))!;
    }

    private async Task<int> BulkAktarAsync(string sqlServerQuery, string hedefTablo, int toplam, bool isLog,
        Action<int, string>? ilerlemeUpdater, CancellationToken ct)
    {
        var hareketCols = "Tarih, FisNo, RandevuNo, IslemTipi, UrunKodu, BirimFiyat, " +
                          "HesapTipi, TcKimlikNo, HesapKodu, SozlesmeYili, " +
                          "Adres, PlakaNo, SoforAdiSoyadi, SoforGsmNo, AracTipi, " +
                          "MuteahhitKodu, MouseKodu, Aciklama, " +
                          "Sevk, Brut, Dara, Fireli, Net, " +
                          "FireOrani, PolarOrani, FireMiktari, Fark, " +
                          "BosaltmaYeri, BosaltmaSekli, " +
                          "Cikis_Tarihi, Cikis_Saati, Cikis_KullaniciKodu, " +
                          "Giris_Tarihi, Giris_Saati, Giris_KullaniciKodu, " +
                          "IrsaliyeNo, FaturaNo, SiparisNo, KantarKodu, Branda, KapiKagidiNo, " +
                          "Kod1, Kod2, Kod3, Kod4, Kod5, Kod6, Kod7, " +
                          "Nakit, KrediKarti, Cari, Havale, " +
                          "Kaydeden, KayitTarihi, KayitSaati, " +
                          "OnKayit_HostName, OnKayit_ipAdres, " +
                          "Giris_HostName, Giris_ipAdres, " +
                          "Cikis_HostName, Cikis_ipAdres, " +
                          "Row_ID, Kontrol";

        var logEkCols = ", LogIslemTipi, LogKaydeden, LogKayitTarihi, LogKayitSaati, LogAciklama, LogHostName, LogipAdres";

        var sutunListe = isLog ? hareketCols + logEkCols : hareketCols;
        int sutunSayisi = sutunListe.Split(',').Length;
        var paramListesi = string.Join(",", Enumerable.Range(0, sutunSayisi).Select(i => $"@p{i}"));
        var insertSql = $"INSERT INTO {hedefTablo} ({sutunListe}) VALUES ({paramListesi})";

        await using var sqlConn = new SqlConnection(_baglanti.GetConnectionString());
        await sqlConn.OpenAsync(ct);
        await using var sqlCmd = new SqlCommand(sqlServerQuery, sqlConn);
        sqlCmd.CommandTimeout = 600;
        await using var reader = await sqlCmd.ExecuteReaderAsync(ct);

        await using var sqliteConn = _db.CreateConnection();

        int aktarilan = 0;
        const int batchBoyut = 5000;
        var transaction = (SqliteTransaction?)await sqliteConn.BeginTransactionAsync(ct);
        try
        {
            await using var insertCmd = sqliteConn.CreateCommand();
            insertCmd.Transaction = transaction;
            insertCmd.CommandText = insertSql;

            var parameters = new SqliteParameter[sutunSayisi];
            for (int i = 0; i < sutunSayisi; i++)
            {
                parameters[i] = insertCmd.CreateParameter();
                parameters[i].ParameterName = $"@p{i}";
                insertCmd.Parameters.Add(parameters[i]);
            }
            insertCmd.Prepare();

            while (await reader.ReadAsync(ct))
            {
                for (int i = 0; i < sutunSayisi; i++)
                    parameters[i].Value = reader.IsDBNull(i) ? DBNull.Value : reader.GetValue(i);
                await insertCmd.ExecuteNonQueryAsync(ct);
                aktarilan++;

                if (aktarilan % batchBoyut == 0)
                {
                    await transaction!.CommitAsync(ct);
                    await transaction.DisposeAsync();
                    transaction = (SqliteTransaction?)await sqliteConn.BeginTransactionAsync(ct);
                    insertCmd.Transaction = transaction;

                    int yuzde = toplam > 0 ? (int)((long)aktarilan * 100 / toplam) : 0;
                    ilerlemeUpdater?.Invoke(yuzde, $"{aktarilan:N0} / {toplam:N0} kayıt aktarıldı");
                    await Task.Yield();
                }
            }
            await transaction!.CommitAsync(ct);
        }
        catch
        {
            try { await transaction!.RollbackAsync(ct); } catch { }
            throw;
        }
        finally
        {
            if (transaction != null) await transaction.DisposeAsync();
        }

        return aktarilan;
    }
}
