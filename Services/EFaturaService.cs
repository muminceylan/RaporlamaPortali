using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

// eFaturaDogusCayDB.POSTBOX + ELEMENTS üzerinden gelen/giden e-fatura listesi
// ve UBL XML (POSTBOX.DATA içindeki ZIP) ayrıştırma.
//
// ENVELOPETYPE yapısı (bu DB'de gözlenen):
//   SENDERENVELOPE  = gerçek e-faturalar (hem bizim gönderdiğimiz hem bize gelen
//                     burada yer alıyor) — yön SENDER_VKNTCKN / RECEIVER_VKNTCKN
//                     karşılaştırması ile belirlenir
//   POSTBOXENVELOPE = uygulama yanıtları (KABUL/RED) — varsayılan listede yok
//   SYSTEMENVELOPE  = GİB sistem yanıtları — varsayılan listede yok
public class EFaturaService
{
    private readonly string _connStr;
    private readonly DatabaseService _db;
    private readonly ILogger<EFaturaService> _log;
    private readonly string[] _bizimVknler;

    public EFaturaService(IConfiguration cfg, DatabaseService db, ILogger<EFaturaService> log)
    {
        _connStr = cfg.GetConnectionString("EFaturaDB") ?? "";
        _db = db;
        _log = log;
        _bizimVknler = cfg.GetSection("EFatura:BizimVknler").Get<string[]>() ?? Array.Empty<string>();
    }

    public bool BaglantiTanimliMi => !string.IsNullOrWhiteSpace(_connStr);

    // VKN/TCKN → (CariKod, CariUnvan) — LG_xxx_CLCARD üzerinden lookup
    public async Task<Dictionary<string, (string Kod, string Unvan)>> CariEslestirAsync(IEnumerable<string> vknler)
    {
        var sonuc = new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase);
        var set = vknler.Where(x => !string.IsNullOrWhiteSpace(x))
                        .Select(x => x.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToArray();
        if (set.Length == 0) return sonuc;

        try
        {
            var tbl = _db.GetTableName("CLCARD");  // LG_211_CLCARD
            await using var con = _db.CreateConnection();
            await con.OpenAsync();
            var rows = await con.QueryAsync($@"
                SELECT TAXNR, TCKNO, CODE, DEFINITION_
                FROM {tbl} WITH (NOLOCK)
                WHERE (TAXNR IN @set OR TCKNO IN @set)", new { set });

            foreach (var r in rows)
            {
                string kod = ((string?)r.CODE ?? "").Trim();
                string unvan = ((string?)r.DEFINITION_ ?? "").Trim();
                string vn = ((string?)r.TAXNR ?? "").Trim();
                string tc = ((string?)r.TCKNO ?? "").Trim();
                if (!string.IsNullOrEmpty(vn) && !sonuc.ContainsKey(vn)) sonuc[vn] = (kod, unvan);
                if (!string.IsNullOrEmpty(tc) && !sonuc.ContainsKey(tc)) sonuc[tc] = (kod, unvan);
            }
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Cari eşleştirme başarısız");
        }
        return sonuc;
    }

    // Logo cari kart kodu LIKE filtresine uyan tüm VKN/TCKN'leri döndürür.
    // Örn. kodPrefix="320.06" → 320.06* ile başlayan tüm cariler.
    public async Task<List<string>> VknlerByCariKodAsync(string kodPrefix)
    {
        var liste = new List<string>();
        if (string.IsNullOrWhiteSpace(kodPrefix)) return liste;
        try
        {
            var tbl = _db.GetTableName("CLCARD");
            await using var con = _db.CreateConnection();
            await con.OpenAsync();
            var rows = await con.QueryAsync<(string? Tax, string? Tc)>($@"
                SELECT TAXNR, TCKNO
                FROM {tbl} WITH (NOLOCK)
                WHERE CODE LIKE @pat",
                new { pat = kodPrefix.Trim() + "%" });
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in rows)
            {
                if (!string.IsNullOrWhiteSpace(r.Tax)) set.Add(r.Tax.Trim());
                if (!string.IsNullOrWhiteSpace(r.Tc))  set.Add(r.Tc.Trim());
            }
            liste.AddRange(set);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Logo cari kod prefix araması başarısız: {p}", kodPrefix);
        }
        return liste;
    }

    public async Task<List<string>> InvoiceTiplerAsync()
    {
        if (!BaglantiTanimliMi) return new();
        await using var con = new SqlConnection(_connStr);
        await con.OpenAsync();
        var rows = await con.QueryAsync<string>(@"
            SELECT DISTINCT INVOICETYPE
            FROM POSTBOX WITH (NOLOCK)
            WHERE INVOICETYPE IS NOT NULL AND INVOICETYPE <> ''
            ORDER BY INVOICETYPE");
        return rows.ToList();
    }

    public async Task<List<EFaturaListItem>> ListeleAsync(EFaturaFiltre f)
    {
        if (!BaglantiTanimliMi)
            throw new InvalidOperationException("EFaturaDB connection string tanımlı değil.");

        // Bu DB'de hem Gelen hem Giden SENDERENVELOPE içinde — yön SENDER/RECEIVER
        // VKN'ye göre BizimVknler filtresi ile ayrıştırılıyor.
        const string envType = "SENDERENVELOPE";

        // Logo cari kodu prefix filtresi varsa, önce ilgili VKN/TCKN'leri toparla.
        // Eğer hiçbir cari eşleşmezse erken çık.
        string[]? cariVknleri = null;
        if (!string.IsNullOrWhiteSpace(f.LogoCariKodu))
        {
            cariVknleri = (await VknlerByCariKodAsync(f.LogoCariKodu)).ToArray();
            if (cariVknleri.Length == 0) return new List<EFaturaListItem>();
        }

        // Performans: Fatura No filtresi yoksa POSTBOX ile ELEMENTS'i ayrı sorgu yap.
        // (OUTER APPLY her satır için ELEMENTS'a ayrı seek yapıyor, yavaş.)
        bool faturaNoFilter = !string.IsNullOrWhiteSpace(f.FaturaNo);

        var sql = new StringBuilder();
        sql.Append(@"
            SELECT TOP (@Top)
                p.ID                AS Id,
                p.DATETIME          AS Tarih,
                p.INVOICETYPE       AS Tip,
                p.PROFILEID         AS ProfileId,
                p.ENVELOPETYPE      AS EnvelopeType,
                p.SENDER_VKNTCKN    AS SenderVkn,
                p.SENDER_UNVAN      AS SenderUnvan,
                p.RECEIVER_VKNTCKN  AS ReceiverVkn,
                p.RECEIVER_UNVAN    AS ReceiverUnvan,
                p.DESCRIPTION       AS Aciklama,
                p.FILENAME          AS DosyaAdi,
                p.FILESIZE          AS DataSize,
                p.STATUS            AS Status");

        if (faturaNoFilter)
        {
            sql.Append(@",
                e.ELEMENTID         AS FaturaNo,
                e.UUID              AS Uuid
            FROM POSTBOX p WITH (NOLOCK)
            INNER JOIN ELEMENTS e WITH (NOLOCK) ON e.POSTBOXREF = p.ID
            WHERE p.ENVELOPETYPE = @EnvType
              AND p.DATETIME >= @Bas AND p.DATETIME < @Bit
              AND e.ELEMENTID LIKE @FaturaNo");
        }
        else
        {
            sql.Append(@"
            FROM POSTBOX p WITH (NOLOCK)
            WHERE p.ENVELOPETYPE = @EnvType
              AND p.DATETIME >= @Bas AND p.DATETIME < @Bit");
        }

        // Eğer kullanıcı zaten bir fatura tipi seçtiyse, INVOICETYPE = @Tip
        // koşulu zaten GİB yanıt zarflarını (INVOICETYPE NULL) eler.
        // Aksi halde NULL/boş olmamasını ayrıca şart koş.
        if (!string.IsNullOrWhiteSpace(f.Tip))
            sql.Append(" AND p.INVOICETYPE = @Tip");
        else
        {
            if (!f.SistemYanitlariniGoster)
                sql.Append(" AND p.INVOICETYPE IS NOT NULL AND p.INVOICETYPE <> ''");
            // SEVK (irsaliye) faturaları varsayılan listede istenmiyor — finansal
            // toplamı yok ve listeyi şişiriyor. Kullanıcı Tip=SEVK seçerse görür.
            sql.Append(" AND p.INVOICETYPE <> 'SEVK'");
        }

        // Bizim firmaya kesilen/bizden çıkan faturaları filtrele —
        // eLogo POSTBOX'ta bazen farklı şirketlere ait faturalar da bulunabiliyor.
        // Gelen → alıcı bizim olmalı; Giden → satıcı bizim olmalı.
        if (_bizimVknler.Length > 0)
        {
            sql.Append(f.Yon == EFaturaYon.Giden
                ? " AND p.SENDER_VKNTCKN IN @BizimVkn"
                : " AND p.RECEIVER_VKNTCKN IN @BizimVkn");
        }

        // Logo cari kodu prefix filtresi — karşı tarafın VKN'si bu listede olmalı.
        // SQL Server parametre limiti 2100 olduğundan listeyi tek string'e
        // birleştirip STRING_SPLIT ile parse ediyoruz.
        if (cariVknleri != null)
        {
            sql.Append(f.Yon == EFaturaYon.Giden
                ? " AND p.RECEIVER_VKNTCKN IN (SELECT value FROM STRING_SPLIT(@CariVknCsv, ','))"
                : " AND p.SENDER_VKNTCKN   IN (SELECT value FROM STRING_SPLIT(@CariVknCsv, ','))");
        }
        if (!string.IsNullOrWhiteSpace(f.ProfileId))
            sql.Append(" AND p.PROFILEID = @ProfileId");

        if (!string.IsNullOrWhiteSpace(f.KarsiUnvan))
            sql.Append(f.Yon == EFaturaYon.Giden
                ? " AND p.RECEIVER_UNVAN LIKE @Karsi"
                : " AND p.SENDER_UNVAN LIKE @Karsi");

        if (!string.IsNullOrWhiteSpace(f.KarsiVKN))
            sql.Append(f.Yon == EFaturaYon.Giden
                ? " AND p.RECEIVER_VKNTCKN = @VKN"
                : " AND p.SENDER_VKNTCKN = @VKN");

        sql.Append(" ORDER BY p.DATETIME DESC");

        await using var con = new SqlConnection(_connStr);
        await con.OpenAsync();
        var rows = await con.QueryAsync(sql.ToString(), new
        {
            Top       = Math.Max(1, Math.Min(f.MaxKayit, 100000)),
            EnvType   = envType,
            Bas       = f.Baslangic.Date,
            Bit       = f.Bitis.Date.AddDays(1),
            Tip       = f.Tip,
            ProfileId = f.ProfileId,
            Karsi     = "%" + (f.KarsiUnvan ?? "") + "%",
            VKN       = f.KarsiVKN,
            FaturaNo  = "%" + (f.FaturaNo ?? "") + "%",
            BizimVkn   = _bizimVknler,
            CariVknCsv = cariVknleri == null ? "" : string.Join(",", cariVknleri),
        }, commandTimeout: 180);

        var liste = new List<EFaturaListItem>();
        bool gidenFilter = f.Yon == EFaturaYon.Giden;
        foreach (var r in rows)
        {
            string envT = (string?)r.EnvelopeType ?? "";
            liste.Add(new EFaturaListItem
            {
                Id           = Convert.ToInt64(r.Id),
                Tarih        = (DateTime)r.Tarih,
                FaturaNo     = faturaNoFilter ? ((string?)r.FaturaNo ?? "") : "",
                Uuid         = faturaNoFilter ? ((string?)r.Uuid ?? "")     : "",
                Tip          = (string?)r.Tip ?? "",
                ProfileId    = (string?)r.ProfileId ?? "",
                EnvelopeType = envT,
                Yon          = gidenFilter ? "GİDEN" : "GELEN",
                KarsiVKN     = gidenFilter ? ((string?)r.ReceiverVkn ?? "") : ((string?)r.SenderVkn ?? ""),
                KarsiUnvan   = gidenFilter ? ((string?)r.ReceiverUnvan ?? "") : ((string?)r.SenderUnvan ?? ""),
                Aciklama     = (string?)r.Aciklama ?? "",
                DosyaAdi     = (string?)r.DosyaAdi ?? "",
                DataSize     = r.DataSize is null ? 0L : Convert.ToInt64(r.DataSize),
                Status       = r.Status?.ToString() ?? "",
            });
        }

        // FaturaNo/UUID toplu yükle (FaturaNo filtresi yoksa)
        if (!faturaNoFilter && liste.Count > 0)
        {
            var idsCsv = string.Join(",", liste.Select(x => x.Id));
            try
            {
                var elemRows = await con.QueryAsync(@"
                    SELECT POSTBOXREF, ELEMENTID, UUID
                    FROM ELEMENTS WITH (NOLOCK)
                    WHERE POSTBOXREF IN (SELECT CAST(value AS BIGINT) FROM STRING_SPLIT(@idsCsv, ','))",
                    new { idsCsv }, commandTimeout: 60);

                var dict = new Dictionary<long, (string FaturaNo, string Uuid)>();
                foreach (var er in elemRows)
                {
                    long pid = Convert.ToInt64(er.POSTBOXREF);
                    if (!dict.ContainsKey(pid))
                        dict[pid] = ((string?)er.ELEMENTID ?? "", (string?)er.UUID ?? "");
                }
                foreach (var item in liste)
                    if (dict.TryGetValue(item.Id, out var t))
                    {
                        item.FaturaNo = t.FaturaNo;
                        item.Uuid     = t.Uuid;
                    }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "ELEMENTS toplu yüklenemedi");
            }
        }

        // Karşı taraf VKN'lerini Logo cari kart koduyla eşleştir
        var map = await CariEslestirAsync(liste.Select(x => x.KarsiVKN));
        foreach (var item in liste)
        {
            if (!string.IsNullOrEmpty(item.KarsiVKN) && map.TryGetValue(item.KarsiVKN, out var c))
            {
                item.LogoCariKod   = c.Kod;
                item.LogoCariUnvan = c.Unvan;
            }
            // POSTBOX.SENDER_UNVAN / RECEIVER_UNVAN bazı kayıtlarda boş geliyor (eLogo
            // bu kolonu denormalize etmemiş). Boşsa Logo cari ünvanı ile doldur ki
            // listede "Karşı Taraf" sütunu boş görünmesin.
            if (string.IsNullOrWhiteSpace(item.KarsiUnvan) && !string.IsNullOrEmpty(item.LogoCariUnvan))
                item.KarsiUnvan = item.LogoCariUnvan;
        }

        // Fatura toplamlarını (Matrah/KDV/Tutar/Tevkifat) XML'den paralel parse et.
        // POSTBOX.DATA içinde ZIP → UBL XML. Sadece toplam alanları okunur (kalem yok).
        await ToplamlariYukleAsync(liste);
        return liste;
    }

    /// <summary>
    /// Her satır için POSTBOX.DATA'yı çek, içindeki UBL XML'inden sadece toplamları
    /// (Doviz, Matrah, KDV, Fatura Tutarı, Tevkifat) parse et. Sistem yanıtı zarflarında
    /// invoice element olmayabilir — bu durumda 0 kalır.
    /// Performans: önce yerel SQLite cache'inden bak (POSTBOX kayıtları immutable
    /// olduğu için ID bazlı kalıcı cache). Cache'de olmayan ID'ler için BLOB'lar
    /// 125'erlik chunk'lar halinde batch SQL ile çekilir, parse işi paralel yapılır,
    /// sonuçlar cache'e yazılır.
    /// </summary>
    private async Task ToplamlariYukleAsync(List<EFaturaListItem> liste)
    {
        if (liste.Count == 0) return;

        var dict = liste.ToDictionary(x => x.Id);

        // 1) Cache'den mevcut toplamları yükle ve list itemlarına uygula.
        HashSet<long> cachedIds;
        try
        {
            cachedIds = await ApplyCachedTotalsAsync(liste);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "EFatura totals cache okuma hatası — cache atlandı");
            cachedIds = new HashSet<long>();
        }

        // 2) Cache'de olmayan ID'leri belirle.
        var eksik = liste.Where(x => !cachedIds.Contains(x.Id)).Select(x => x.Id).ToArray();
        if (eksik.Length == 0) return;

        // 3) Eksikleri SQL'den paralel olarak çek, parse et.
        const int chunkSize = 125;
        var chunks = new List<long[]>();
        for (int o = 0; o < eksik.Length; o += chunkSize)
            chunks.Add(eksik.Skip(o).Take(chunkSize).ToArray());

        var yeniParse = new System.Collections.Concurrent.ConcurrentBag<EFaturaListItem>();

        using var sem = new SemaphoreSlim(4);
        var tasks = chunks.Select(async chunk =>
        {
            await sem.WaitAsync();
            try
            {
                var idsCsv = string.Join(",", chunk);
                await using var con = new SqlConnection(_connStr);
                await con.OpenAsync();
                var rows = await con.QueryAsync<(long Id, byte[]? Data)>(@"
                    SELECT ID AS Id, DATA AS Data
                    FROM POSTBOX WITH (NOLOCK)
                    WHERE ID IN (SELECT CAST(value AS BIGINT) FROM STRING_SPLIT(@idsCsv, ','))",
                    new { idsCsv }, commandTimeout: 120);

                Parallel.ForEach(rows, new ParallelOptions { MaxDegreeOfParallelism = 4 }, r =>
                {
                    if (r.Data == null || r.Data.Length == 0) return;
                    if (!dict.TryGetValue(r.Id, out var item)) return;
                    try
                    {
                        ParseTotalsFromBlob(r.Data, item);
                        yeniParse.Add(item);
                    }
                    catch (Exception ex)
                    {
                        _log.LogDebug(ex, "Toplam parse hatasi ID={Id}", r.Id);
                    }
                });
            }
            finally { sem.Release(); }
        }).ToArray();
        await Task.WhenAll(tasks);

        // 4) Yeni parse edilenleri cache'e yaz.
        if (yeniParse.Count > 0)
        {
            try { await WriteCachedTotalsAsync(yeniParse); }
            catch (Exception ex) { _log.LogWarning(ex, "EFatura totals cache yazma hatası"); }
        }
    }

    // ---- SQLite cache (parse edilmiş fatura toplamları için kalıcı yerel cache) ----
    //
    // POSTBOX kayıtları immutable olduğundan (bir e-fatura POSTBOX'a düştükten sonra
    // içeriği değişmez), parse edilmiş toplamları kalıcı olarak cache'liyoruz.
    // Bu sayede tarih aralığı tekrar sorgulandığında BLOB çekme + XML parse maliyeti
    // tekrar ödenmez.
    private static readonly SemaphoreSlim _cacheInitLock = new(1, 1);
    private static bool _cacheReady;

    private static async Task EnsureCacheSchemaAsync()
    {
        if (_cacheReady) return;
        await _cacheInitLock.WaitAsync();
        try
        {
            if (_cacheReady) return;
            Directory.CreateDirectory(AppDataPaths.DataRoot);
            await using var con = new SqliteConnection($"Data Source={AppDataPaths.EFaturaTotalsCacheDb}");
            await con.OpenAsync();
            await con.ExecuteAsync(@"
                CREATE TABLE IF NOT EXISTS totals (
                    id        INTEGER PRIMARY KEY,
                    doviz     TEXT NOT NULL DEFAULT '',
                    matrah    REAL NOT NULL DEFAULT 0,
                    kdv       REAL NOT NULL DEFAULT 0,
                    toplam    REAL NOT NULL DEFAULT 0,
                    tevkifat  REAL NOT NULL DEFAULT 0,
                    ilk_kalem TEXT NOT NULL DEFAULT ''
                );");
            _cacheReady = true;
        }
        finally { _cacheInitLock.Release(); }
    }

    /// <summary>
    /// Listedeki ID'ler için cache'den toplamları yükler ve item'lara uygular.
    /// Geri dönüş: cache'de bulunan ID'ler.
    /// </summary>
    private async Task<HashSet<long>> ApplyCachedTotalsAsync(List<EFaturaListItem> liste)
    {
        await EnsureCacheSchemaAsync();
        var bulundu = new HashSet<long>();
        if (liste.Count == 0) return bulundu;

        await using var con = new SqliteConnection($"Data Source={AppDataPaths.EFaturaTotalsCacheDb}");
        await con.OpenAsync();

        // SQLite'da parametre limiti ~999; güvenli olsun diye 500'lük chunk'lar.
        const int chunk = 500;
        var dict = liste.ToDictionary(x => x.Id);
        for (int o = 0; o < liste.Count; o += chunk)
        {
            var ids = liste.Skip(o).Take(chunk).Select(x => x.Id).ToArray();
            var rows = await con.QueryAsync<(long id, string doviz, double matrah, double kdv,
                                              double toplam, double tevkifat, string ilk_kalem)>(
                "SELECT id, doviz, matrah, kdv, toplam, tevkifat, ilk_kalem FROM totals WHERE id IN @ids",
                new { ids });

            foreach (var r in rows)
            {
                if (!dict.TryGetValue(r.id, out var item)) continue;
                item.Doviz         = r.doviz;
                item.Matrah        = (decimal)r.matrah;
                item.KdvTutar      = (decimal)r.kdv;
                item.ToplamTutar   = (decimal)r.toplam;
                item.TevkifatTutar = (decimal)r.tevkifat;
                item.IlkKalem      = r.ilk_kalem;
                bulundu.Add(r.id);
            }
        }
        return bulundu;
    }

    /// <summary>
    /// Yeni parse edilmiş toplamları cache'e yazar (INSERT OR REPLACE, tek transaction).
    /// </summary>
    private async Task WriteCachedTotalsAsync(IEnumerable<EFaturaListItem> items)
    {
        await EnsureCacheSchemaAsync();
        await using var con = new SqliteConnection($"Data Source={AppDataPaths.EFaturaTotalsCacheDb}");
        await con.OpenAsync();
        await using var tx = (SqliteTransaction)await con.BeginTransactionAsync();
        const string sql = @"
            INSERT INTO totals (id, doviz, matrah, kdv, toplam, tevkifat, ilk_kalem)
            VALUES (@id, @doviz, @matrah, @kdv, @toplam, @tevkifat, @ilk_kalem)
            ON CONFLICT(id) DO UPDATE SET
                doviz=excluded.doviz, matrah=excluded.matrah, kdv=excluded.kdv,
                toplam=excluded.toplam, tevkifat=excluded.tevkifat,
                ilk_kalem=excluded.ilk_kalem;";

        foreach (var it in items)
        {
            await con.ExecuteAsync(sql, new
            {
                id        = it.Id,
                doviz     = it.Doviz ?? "",
                matrah    = (double)it.Matrah,
                kdv       = (double)it.KdvTutar,
                toplam    = (double)it.ToplamTutar,
                tevkifat  = (double)it.TevkifatTutar,
                ilk_kalem = it.IlkKalem ?? "",
            }, transaction: tx);
        }
        await tx.CommitAsync();
    }

    /// <summary>
    /// POSTBOX.DATA blob'unu (genelde ZIP'lenmiş UBL XML) doğrudan XmlReader ile akıt.
    /// XmlDocument'e yüklemeden, sadece ihtiyacımız olan birkaç alanı (Doviz, Matrah,
    /// PayableAmount, KDV toplamı, Tevkifat toplamı) sıralı okur. ~5-10x daha hızlı.
    /// </summary>
    private static void ParseTotalsFromBlob(byte[] blob, EFaturaListItem item)
    {
        Stream? src = null;
        ZipArchive? arc = null;
        try
        {
            if (blob.Length >= 2 && blob[0] == 0x50 && blob[1] == 0x4B)
            {
                arc = new ZipArchive(new MemoryStream(blob), ZipArchiveMode.Read);
                var entry = arc.Entries.FirstOrDefault(x =>
                    x.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
                if (entry == null) return;
                src = entry.Open();
            }
            else
            {
                src = new MemoryStream(blob);
            }

            var settings = new XmlReaderSettings
            {
                IgnoreWhitespace = true,
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                DtdProcessing = DtdProcessing.Ignore,
                CheckCharacters = false,
            };
            using var rdr = XmlReader.Create(src, settings);
            ParseTotalsStreaming(rdr, item);
        }
        finally
        {
            src?.Dispose();
            arc?.Dispose();
        }
    }

    private static void ParseTotalsStreaming(XmlReader rdr, EFaturaListItem item)
    {
        bool inInvoice = false;
        int legalTotalDepth = -1;
        int taxSubDepth = -1;
        int withholdDepth = -1;
        string subCode = "";
        decimal subAmt = 0m;
        decimal kdv = 0m, tev = 0m;

        // İlk InvoiceLine içinden ilk kalem adını çıkarmak için state
        bool gotFirstLine = false;
        int  lineDepth    = -1;   // >=0 iken ilk InvoiceLine içindeyiz
        int  itemDepth    = -1;   // >=0 iken ilk satırın cac:Item bloğundayız
        string lineNote   = "";   // cbc:Note (Name boşsa fallback)

        while (rdr.Read())
        {
            // EndElement → state reset
            if (rdr.NodeType == XmlNodeType.EndElement)
            {
                int d = rdr.Depth;
                if (legalTotalDepth >= 0 && d == legalTotalDepth && rdr.LocalName == "LegalMonetaryTotal")
                    legalTotalDepth = -1;
                else if (withholdDepth >= 0 && d == withholdDepth && rdr.LocalName == "WithholdingTaxTotal")
                    withholdDepth = -1;
                else if (taxSubDepth >= 0 && d == taxSubDepth && rdr.LocalName == "TaxSubtotal")
                {
                    if (subCode == "0015") kdv += subAmt;
                    taxSubDepth = -1;
                }
                else if (itemDepth >= 0 && d == itemDepth && rdr.LocalName == "Item")
                    itemDepth = -1;
                else if (lineDepth >= 0 && d == lineDepth && rdr.LocalName == "InvoiceLine")
                {
                    lineDepth = -1;
                    gotFirstLine = true;
                    if (string.IsNullOrEmpty(item.IlkKalem) && !string.IsNullOrEmpty(lineNote))
                        item.IlkKalem = lineNote;
                }
                continue;
            }

            if (rdr.NodeType != XmlNodeType.Element) continue;

            var name = rdr.LocalName;
            if (!inInvoice)
            {
                if (name == "Invoice") inInvoice = true;
                continue;
            }

            // İlk kalemin İÇİNDE iken sadece Item/Name + Note ile ilgileniriz —
            // diğer Tax* handler'ları satır içinde çalıştırmamalı (yoksa toplam bozulur).
            if (lineDepth >= 0)
            {
                switch (name)
                {
                    case "Item":
                        if (!rdr.IsEmptyElement) itemDepth = rdr.Depth;
                        break;
                    case "Name":
                        if (itemDepth >= 0 && rdr.Depth == itemDepth + 1 && string.IsNullOrEmpty(item.IlkKalem))
                            item.IlkKalem = ReadTextNoAdvance(rdr);
                        break;
                    case "Note":
                        if (rdr.Depth == lineDepth + 1 && string.IsNullOrEmpty(lineNote))
                            lineNote = ReadTextNoAdvance(rdr);
                        break;
                }
                continue;
            }

            switch (name)
            {
                case "DocumentCurrencyCode":
                    if (string.IsNullOrEmpty(item.Doviz))
                        item.Doviz = ReadTextNoAdvance(rdr);
                    break;

                case "LegalMonetaryTotal":
                    if (!rdr.IsEmptyElement) legalTotalDepth = rdr.Depth;
                    break;

                case "TaxExclusiveAmount":
                    if (legalTotalDepth >= 0 && rdr.Depth == legalTotalDepth + 1)
                        item.Matrah = ReadDecimalNoAdvance(rdr);
                    break;

                case "PayableAmount":
                    if (legalTotalDepth >= 0 && rdr.Depth == legalTotalDepth + 1)
                        item.ToplamTutar = ReadDecimalNoAdvance(rdr);
                    break;

                case "WithholdingTaxTotal":
                    if (!rdr.IsEmptyElement) withholdDepth = rdr.Depth;
                    break;

                case "TaxSubtotal":
                    if (!rdr.IsEmptyElement && withholdDepth < 0)
                    {
                        taxSubDepth = rdr.Depth;
                        subCode = "";
                        subAmt = 0m;
                    }
                    break;

                case "TaxAmount":
                    if (taxSubDepth >= 0 && rdr.Depth == taxSubDepth + 1)
                        subAmt = ReadDecimalNoAdvance(rdr);
                    else if (withholdDepth >= 0 && rdr.Depth == withholdDepth + 1)
                        tev += ReadDecimalNoAdvance(rdr);
                    break;

                case "TaxTypeCode":
                    if (taxSubDepth >= 0)
                        subCode = ReadTextNoAdvance(rdr);
                    break;

                case "InvoiceLine":
                    // İlk satırın içine in (sadece Item/Name + Note için).
                    // Sonraki satırları doğrudan Skip ile geç → performans korunur.
                    if (!gotFirstLine && !rdr.IsEmptyElement)
                    {
                        lineDepth = rdr.Depth;
                        lineNote = "";
                    }
                    else
                    {
                        rdr.Skip();
                    }
                    break;
            }
        }

        if (!inInvoice) return;
        item.KdvTutar      = kdv;
        item.TevkifatTutar = tev;
    }

    // Element içeriğini text olarak okur; reader'ı bu element'in EndElement'i üzerinde
    // bırakır (yani ana döngünün bir sonraki Read() çağrısı parent veya sibling'e geçer).
    // Bu sayede parent EndElement event'leri kaçırılmaz.
    private static string ReadTextNoAdvance(XmlReader rdr)
    {
        if (rdr.IsEmptyElement) return "";
        int startDepth = rdr.Depth;
        var sb = new StringBuilder();
        while (rdr.Read())
        {
            if (rdr.NodeType == XmlNodeType.Text || rdr.NodeType == XmlNodeType.CDATA)
                sb.Append(rdr.Value);
            else if (rdr.NodeType == XmlNodeType.EndElement && rdr.Depth == startDepth)
                break;
            else if (rdr.NodeType == XmlNodeType.EndElement && rdr.Depth < startDepth)
                break;
        }
        return sb.ToString().Trim();
    }

    private static decimal ReadDecimalNoAdvance(XmlReader rdr)
    {
        var s = ReadTextNoAdvance(rdr);
        return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : 0m;
    }

    /// <summary>
    /// Gelen e-fatura listesini Logo INVOICE tablosu ile cross-check eder.
    /// Üç kademeli eşleştirme:
    ///   1) FICHENO tam eşit (sadeleştirilmiş — alfanümerik karşılaştırma) → Kesin
    ///   2) Aynı VKN + FICHENO son 4 hane eşit + tutar eşit + tarih ±tolerans → Modifiye
    ///   3) Aynı VKN + tutar eşit + tarih ±tolerans (numara hiç tutmadı) → Şüpheli
    ///   4) Hiçbiri → Yok
    /// Tutar karşılaştırması: e-faturanın (Matrah + KDV) ile Logo GROSSTOTAL,
    /// 1 TL tolerans (yuvarlama farkları için).
    /// </summary>
    /// <summary>
    /// Listede UUID'si bilinen e-faturalar için POSTBOX'ta INVOICETYPE='RED' yanıtı var mı kontrol eder.
    /// Bağlantı: POSTBOXENVELOPE.REFERENCE_ENVELOPEID = orijinal SENDERENVELOPE.UUID (case-insensitive).
    /// Eşleşenler için RedEdildi=true, RedTarihi ve RedAciklama doldurulur.
    /// </summary>
    public async Task IptalRedKontroluAsync(List<EFaturaListItem> liste)
    {
        var uuidlist = liste.Where(x => !string.IsNullOrWhiteSpace(x.Uuid))
                            .Select(x => x.Uuid.Trim())
                            .Distinct(StringComparer.OrdinalIgnoreCase)
                            .ToList();
        if (uuidlist.Count == 0) return;

        var uuidCsv = string.Join(",", uuidlist);
        await using var con = new SqlConnection(_connStr);
        await con.OpenAsync();

        // RED yanıtları iki şekilde bağlanabiliyor:
        //   (a) Tek faturalık zarf: POSTBOX.REFERENCE_ENVELOPEID = orijinal fatura UUID
        //   (b) Çok faturalık RED zarfı: ELEMENTS.RESPDOCUMENTREFERENCEID = orijinal fatura UUID
        // İkisini de UNION ile kontrol et.
        var redler = await con.QueryAsync<RedRow>(@"
            SELECT UPPER(REFERENCE_ENVELOPEID) AS RefId,
                   DATETIME                    AS Tarih,
                   DESCRIPTION                 AS Aciklama
            FROM POSTBOX WITH (NOLOCK)
            WHERE INVOICETYPE = 'RED'
              AND REFERENCE_ENVELOPEID IS NOT NULL
              AND UPPER(REFERENCE_ENVELOPEID) IN
                  (SELECT UPPER(value) FROM STRING_SPLIT(@uuidCsv, ','))
            UNION ALL
            SELECT UPPER(e.RESPDOCUMENTREFERENCEID) AS RefId,
                   p.DATETIME                      AS Tarih,
                   ISNULL(p.DESCRIPTION, 'RED')    AS Aciklama
            FROM ELEMENTS e WITH (NOLOCK)
            INNER JOIN POSTBOX p WITH (NOLOCK) ON e.POSTBOXREF = p.ID
            WHERE p.INVOICETYPE = 'RED'
              AND e.RESPDOCUMENTREFERENCEID IS NOT NULL
              AND UPPER(e.RESPDOCUMENTREFERENCEID) IN
                  (SELECT UPPER(value) FROM STRING_SPLIT(@uuidCsv, ','))",
            new { uuidCsv }, commandTimeout: 60);

        var redMap = new Dictionary<string, (DateTime Tarih, string Desc)>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in redler)
        {
            if (string.IsNullOrWhiteSpace(r.RefId)) continue;
            var k = r.RefId.Trim();
            // Aynı UUID'ye birden fazla RED gelirse en yenisini tut
            if (!redMap.TryGetValue(k, out var mevcut) || r.Tarih > mevcut.Tarih)
                redMap[k] = (r.Tarih, r.Aciklama ?? "");
        }
        if (redMap.Count == 0) return;

        foreach (var item in liste)
        {
            if (string.IsNullOrWhiteSpace(item.Uuid)) continue;
            if (redMap.TryGetValue(item.Uuid.Trim(), out var red))
            {
                item.RedEdildi   = true;
                item.RedTarihi   = red.Tarih;
                item.RedAciklama = red.Desc;
            }
        }
    }

    public async Task LogoIslemKontrolAsync(List<EFaturaListItem> liste, int tarihToleransGun = 30)
    {
        if (liste.Count == 0) return;

        // 0) Önce RED yanıtlarını tara — iptal/red edilmiş faturaların Logo'ya işlenmemiş
        //    olması zaten beklenir, "İşlenmedi" listesinde uyarı olarak görünmemeleri lazım.
        try { await IptalRedKontroluAsync(liste); }
        catch (Exception ex) { _log.LogWarning(ex, "İptal/Red kontrolü başarısız"); }

        // Sadece VKN'si olan ve Gelen olan satırları kontrol et
        var hedefler = liste.Where(x => !string.IsNullOrWhiteSpace(x.KarsiVKN)).ToList();
        if (hedefler.Count == 0) return;

        var vknSet = hedefler.Select(x => x.KarsiVKN.Trim())
                             .Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        var minTarih = hedefler.Min(x => x.Tarih).Date.AddDays(-tarihToleransGun);
        var maxTarih = hedefler.Max(x => x.Tarih).Date.AddDays(tarihToleransGun + 1);

        // Logo'daki bu VKN'lere ait tarih aralığındaki tüm faturaları tek sorguda çek
        var invTbl  = _db.GetPeriodTableName("INVOICE");
        var clcTbl  = _db.GetTableName("CLCARD");
        List<LogoInvRow> logoRows;
        try
        {
            await using var con = _db.CreateConnection();
            await con.OpenAsync();
            var rows = await con.QueryAsync<LogoInvRow>($@"
                SELECT i.FICHENO    AS Ficheno,
                       i.DATE_      AS Tarih,
                       i.GROSSTOTAL AS Tutar,
                       i.TRCODE     AS Trcode,
                       c.TAXNR      AS Taxnr,
                       c.TCKNO      AS Tckno
                FROM {invTbl} i WITH (NOLOCK)
                INNER JOIN {clcTbl} c WITH (NOLOCK) ON c.LOGICALREF = i.CLIENTREF
                WHERE (c.TAXNR IN @vknSet OR c.TCKNO IN @vknSet)
                  AND i.DATE_ >= @bas AND i.DATE_ < @bit",
                new { vknSet, bas = minTarih, bit = maxTarih }, commandTimeout: 120);
            logoRows = rows.ToList();
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Logo INVOICE eşleştirme sorgusu başarısız");
            foreach (var it in hedefler) it.LogoDurum = LogoIslemDurumu.Bilinmiyor;
            return;
        }

        // VKN → Logo kayıtları (TAXNR veya TCKNO eşleşmesi)
        var logoByVkn = new Dictionary<string, List<LogoInvRow>>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in logoRows)
        {
            void Ekle(string? vkn)
            {
                if (string.IsNullOrWhiteSpace(vkn)) return;
                var k = vkn.Trim();
                if (!logoByVkn.TryGetValue(k, out var list))
                    logoByVkn[k] = list = new List<LogoInvRow>();
                list.Add(r);
            }
            Ekle(r.Taxnr);
            Ekle(r.Tckno);
        }

        foreach (var item in hedefler)
        {
            if (!logoByVkn.TryGetValue(item.KarsiVKN.Trim(), out var adaylar) || adaylar.Count == 0)
            {
                item.LogoDurum = LogoIslemDurumu.Yok;
                continue;
            }

            var eFNorm = NormalizeFichenoFull(item.FaturaNo);
            var eFSon4 = SonHaneler(item.FaturaNo, 4);
            var eFTarih = item.Tarih.Date;

            LogoInvRow? bestModif = null;
            LogoInvRow? bestSupheli = null;

            // 1) Kesin eşleşme: FICHENO tam eşit (alfanümerik karşılaştırma)
            var kesin = adaylar.FirstOrDefault(r =>
                !string.IsNullOrEmpty(eFNorm) &&
                NormalizeFichenoFull(r.Ficheno) == eFNorm);
            if (kesin != null)
            {
                item.LogoDurum   = LogoIslemDurumu.Kesin;
                item.LogoFicheno = kesin.Ficheno;
                item.LogoTarih   = kesin.Tarih;
                item.LogoTutar   = kesin.Tutar;
                continue;
            }

            // 2/3) Tarih ve tutar filtresinden geçen adaylar
            foreach (var r in adaylar)
            {
                int gunFark = Math.Abs((r.Tarih.Date - eFTarih).Days);
                if (gunFark > tarihToleransGun) continue;
                if (!TutarUyusuyor(item, r.Tutar)) continue;  // tevkifatlı/tevkifatsız 3 değer dener

                // Modifiye yakalama kuralları (herhangi biri tutarsa Modifiye):
                //   a) Son 4 hane aynı (en yaygın — ortada nokta/üçnokta eklenmiş)
                //   b) Karakter benzerliği ≥ %75 (AEF↔AAF gibi 1-2 harf değişimi
                //      veya son hane değişimi de yakalanır)
                var lgNorm = NormalizeFichenoFull(r.Ficheno);
                bool son4Eslesti = !string.IsNullOrEmpty(eFSon4) &&
                                    SonHaneler(r.Ficheno, 4) == eFSon4;
                bool numaraBenzer = !string.IsNullOrEmpty(eFNorm) &&
                                    !string.IsNullOrEmpty(lgNorm) &&
                                    PozisyonelBenzerlik(eFNorm, lgNorm) >= 0.75;

                if (son4Eslesti || numaraBenzer)
                {
                    bestModif = r;
                    break;  // modifiye en güçlü ikinci kademe — döngüyü kes
                }
                bestSupheli ??= r;
            }

            if (bestModif != null)
            {
                item.LogoDurum   = LogoIslemDurumu.Modifiye;
                item.LogoFicheno = bestModif.Ficheno;
                item.LogoTarih   = bestModif.Tarih;
                item.LogoTutar   = bestModif.Tutar;
            }
            else if (bestSupheli != null)
            {
                item.LogoDurum   = LogoIslemDurumu.Supheli;
                item.LogoFicheno = bestSupheli.Ficheno;
                item.LogoTarih   = bestSupheli.Tarih;
                item.LogoTutar   = bestSupheli.Tutar;
            }
            else
            {
                item.LogoDurum = LogoIslemDurumu.Yok;
            }
        }

        // Yok çıkanları ünvan bazlı yedek aramayla tekrar tara —
        // (CLCARD'da VKN'si eksik/yanlış yazılmış mükerrer cari kartları yakalar)
        var yokKalanlar = hedefler
            .Where(x => x.LogoDurum == LogoIslemDurumu.Yok &&
                        !string.IsNullOrWhiteSpace(x.KarsiUnvan))
            .ToList();
        if (yokKalanlar.Count > 0)
        {
            try { await UnvanBazliEslestirmeAsync(yokKalanlar, tarihToleransGun); }
            catch (Exception ex) { _log.LogWarning(ex, "Ünvan bazlı yedek eşleştirme başarısız"); }
        }
    }

    /// <summary>
    /// "Yok" çıkan e-faturaları gönderen ünvanı üzerinden Logo CLCARD'la eşleştirir.
    /// Aynı firmanın Logo'da birden fazla cari kartı olabiliyor (VKN'si eksik olanlar da var).
    /// Ünvandan anlamlı 2 token çıkar, CLCARD.DEFINITION_ içinde her ikisi de geçen
    /// cari kartların INVOICE kayıtlarını aday olarak değerlendirir.
    /// Eşleşirse Logo cari kodunu/ünvanını da günceller (hangi yanlış cariye işlendi görünür).
    /// </summary>
    private async Task UnvanBazliEslestirmeAsync(List<EFaturaListItem> hedefler, int tarihToleransGun)
    {
        if (hedefler.Count == 0) return;

        // E-fatura ünvanlarından anlamlı token setleri çıkar
        var hedefTokens = new Dictionary<long, List<string>>();
        foreach (var it in hedefler)
        {
            var tk = UnvandanAnlmaliKelimeler(it.KarsiUnvan, max: 2);
            if (tk.Count >= 2) hedefTokens[it.Id] = tk;
        }
        if (hedefTokens.Count == 0) return;

        var clcTbl = _db.GetTableName("CLCARD");
        var invTbl = _db.GetPeriodTableName("INVOICE");
        await using var con = _db.CreateConnection();
        await con.OpenAsync();

        // CLCARD'ı tek seferde çek (boyut küçük, full table scan yapılır C# tarafında)
        var tumCariler = (await con.QueryAsync<ClcRow>($@"
            SELECT LOGICALREF AS LogicalRef, CODE AS Code,
                   DEFINITION_ AS Def, TAXNR AS Tax, TCKNO AS Tc
            FROM {clcTbl} WITH (NOLOCK)
            WHERE DEFINITION_ IS NOT NULL AND DEFINITION_ <> ''")).ToList();
        if (tumCariler.Count == 0) return;

        // Tüm CLCARD definition'larını bir kez ASCII normalize et (Ü→U, Ş→S, vb.)
        // — tokenler ASCII olduğu için karşılaştırma da ASCII bazında yapılmalı.
        var defAsciiByRef = new Dictionary<int, string>(tumCariler.Count);
        foreach (var c in tumCariler)
            defAsciiByRef[c.LogicalRef] = TrToAscii(c.Def).ToUpperInvariant();

        // Her hedef için aday CLIENTREF setlerini topla
        var itemAdayRefs = new Dictionary<long, HashSet<int>>();
        var tumAdayRefs  = new HashSet<int>();
        foreach (var (itemId, tokens) in hedefTokens)
        {
            var refs = new HashSet<int>();
            foreach (var c in tumCariler)
            {
                if (!defAsciiByRef.TryGetValue(c.LogicalRef, out var defAscii) ||
                    string.IsNullOrEmpty(defAscii)) continue;
                bool hepsiVar = true;
                foreach (var t in tokens)
                {
                    // tokens zaten ASCII upper, defAscii de ASCII upper → Ordinal yeterli
                    if (defAscii.IndexOf(t, StringComparison.Ordinal) < 0)
                    {
                        hepsiVar = false;
                        break;
                    }
                }
                if (hepsiVar) refs.Add(c.LogicalRef);
            }
            if (refs.Count > 0)
            {
                itemAdayRefs[itemId] = refs;
                foreach (var r in refs) tumAdayRefs.Add(r);
            }
        }
        if (tumAdayRefs.Count == 0) return;

        // Bu adayların INVOICE kayıtlarını çek (tarih aralığında)
        var minTarih = hedefler.Min(x => x.Tarih).Date.AddDays(-tarihToleransGun);
        var maxTarih = hedefler.Max(x => x.Tarih).Date.AddDays(tarihToleransGun + 1);
        var refsCsv  = string.Join(",", tumAdayRefs);
        var invRows  = (await con.QueryAsync<UnvanInvRow>($@"
            SELECT i.FICHENO    AS Ficheno,
                   i.DATE_      AS Tarih,
                   i.GROSSTOTAL AS Tutar,
                   i.CLIENTREF  AS ClientRef
            FROM {invTbl} i WITH (NOLOCK)
            WHERE i.CLIENTREF IN (SELECT CAST(value AS INT) FROM STRING_SPLIT(@refsCsv, ','))
              AND i.DATE_ >= @bas AND i.DATE_ < @bit",
            new { refsCsv, bas = minTarih, bit = maxTarih }, commandTimeout: 120)).ToList();
        if (invRows.Count == 0) return;

        // Ref → cari bilgileri
        var refMap = tumCariler.ToDictionary(c => c.LogicalRef);

        // Her hedef için eşleştirme uygula (aynı kademe mantığı)
        var itemMap = hedefler.ToDictionary(x => x.Id);
        foreach (var (itemId, refSet) in itemAdayRefs)
        {
            if (!itemMap.TryGetValue(itemId, out var item)) continue;

            var adayInv = invRows.Where(r => refSet.Contains(r.ClientRef)).ToList();
            if (adayInv.Count == 0) continue;

            var eFNorm  = NormalizeFichenoFull(item.FaturaNo);
            var eFSon4  = SonHaneler(item.FaturaNo, 4);
            var eFTarih = item.Tarih.Date;

            UnvanInvRow? secilen = null;
            LogoIslemDurumu secilenDurum = LogoIslemDurumu.Yok;

            // 1) Kesin (FICHENO tam eşit)
            secilen = adayInv.FirstOrDefault(r =>
                !string.IsNullOrEmpty(eFNorm) &&
                NormalizeFichenoFull(r.Ficheno) == eFNorm);
            if (secilen != null) secilenDurum = LogoIslemDurumu.Kesin;

            // 2/3) Modifiye veya Şüpheli
            if (secilen == null)
            {
                UnvanInvRow? bestModif = null, bestSup = null;
                foreach (var r in adayInv)
                {
                    int gunFark = Math.Abs((r.Tarih.Date - eFTarih).Days);
                    if (gunFark > tarihToleransGun) continue;
                    if (!TutarUyusuyor(item, r.Tutar)) continue;  // tevkifatlı/tevkifatsız 3 değer dener

                    var lgNorm = NormalizeFichenoFull(r.Ficheno);
                    bool son4   = !string.IsNullOrEmpty(eFSon4) &&
                                  SonHaneler(r.Ficheno, 4) == eFSon4;
                    bool benzer = !string.IsNullOrEmpty(eFNorm) &&
                                  !string.IsNullOrEmpty(lgNorm) &&
                                  PozisyonelBenzerlik(eFNorm, lgNorm) >= 0.75;
                    if (son4 || benzer) { bestModif = r; break; }
                    bestSup ??= r;
                }
                if (bestModif != null)      { secilen = bestModif; secilenDurum = LogoIslemDurumu.Modifiye; }
                else if (bestSup != null)   { secilen = bestSup;   secilenDurum = LogoIslemDurumu.Supheli; }
            }

            if (secilen != null)
            {
                item.LogoDurum   = secilenDurum;
                item.LogoFicheno = secilen.Ficheno;
                item.LogoTarih   = secilen.Tarih;
                item.LogoTutar   = secilen.Tutar;
                // Cari bilgisini güncelle: hangi yanlış cariye işlendiği görünsün
                if (refMap.TryGetValue(secilen.ClientRef, out var c))
                {
                    item.LogoCariKod   = c.Code ?? "";
                    item.LogoCariUnvan = c.Def  ?? "";
                }
            }
        }
    }

    private class ClcRow
    {
        public int     LogicalRef { get; set; }
        public string? Code       { get; set; }
        public string? Def        { get; set; }
        public string? Tax        { get; set; }
        public string? Tc         { get; set; }
    }

    private class UnvanInvRow
    {
        public string   Ficheno   { get; set; } = "";
        public DateTime Tarih     { get; set; }
        public decimal  Tutar     { get; set; }
        public int      ClientRef { get; set; }
    }

    // Türkçe firma ünvanından anlamlı 2-3 token çıkarır.
    // Stop word'leri (LTD, ŞTİ, TİCARET, SANAYİ vb.) eler.
    // "HASAN BALCI BUL. PERAKENDE TİC. LTD. ŞTİ." → ["HASAN", "BALCI"]
    private static readonly HashSet<string> _unvanStopWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "LTD","ŞTİ","STI","LİMİTED","LIMITED","ŞİRKETİ","SIRKETI","ANONİM","ANONIM","AŞ","A.Ş",
        "TİCARET","TICARET","TİC","TIC","SANAYİ","SANAYI","SAN","SAN.","SAN,",
        "İNŞAAT","INSAAT","TEKSTİL","TEKSTIL","GIDA","GIDAVE","ÜRETİM","URETIM",
        "PAZARLAMA","PAZ","PETROL","PETROLCÜLÜK","DAĞITIM","DAGITIM",
        "İTHALAT","ITHALAT","İHRACAT","IHRACAT","NAKLİYAT","NAKLIYAT",
        "PERAKENDE","TOPTAN","MAĞAZA","MAGAZA","MAĞAZACILIK","MARKETÇİLİK","MARKETCILIK",
        "OTOMOTİV","OTOMOTIV","ELEKTRİK","ELEKTRIK","HİZMET","HIZMET","SERVİS","SERVIS",
        "OTOPARK","OTO","MAKİNA","MAKINA","MAKİNE","MAKINE","KİMYA","KIMYA",
        "BUL","BLV","CAD","CAD.","SK","SOK","MH","MAH","MAHALLESİ","MAHALLESI","CADDESİ","CADDESI",
        "İŞ","IS","MERKEZİ","MERKEZI","AVM","PLAZA","BLOK","KAT","NO",
        "VE","İLE","ILE","SAYIN","BAY","BAYAN"
    };

    private static List<string> UnvandanAnlmaliKelimeler(string? unvan, int max)
    {
        var sonuc = new List<string>();
        if (string.IsNullOrWhiteSpace(unvan)) return sonuc;
        // Noktalama temizle
        var temiz = unvan.Replace('.', ' ').Replace(',', ' ').Replace('-', ' ').Replace('/', ' ');
        foreach (var raw in temiz.Split(' ', StringSplitOptions.RemoveEmptyEntries))
        {
            var k = raw.Trim();
            if (k.Length < 4) continue;
            // Türkçe karakterleri ASCII'ye çevir: "BÜYÜK" → "BUYUK"
            // Logo CLCARD genelde Latince yazılıyor (BUYUK), e-fatura UBL ise Türkçe (BÜYÜK).
            var ascii = TrToAscii(k).ToUpperInvariant();
            if (_unvanStopWords.Contains(k) || _unvanStopWords.Contains(ascii)) continue;
            sonuc.Add(ascii);
            if (sonuc.Count >= max) break;
        }
        return sonuc;
    }

    /// <summary>Türkçe karakterleri ASCII karşılıklarına çevirir (Ç→C, Ş→S, Ü→U, vb.).</summary>
    private static string TrToAscii(string? s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        var sb = new StringBuilder(s.Length);
        foreach (var ch in s)
        {
            var c = ch switch
            {
                'Ç' => 'C', 'ç' => 'c',
                'Ğ' => 'G', 'ğ' => 'g',
                'İ' => 'I', 'ı' => 'i',
                'Ö' => 'O', 'ö' => 'o',
                'Ş' => 'S', 'ş' => 's',
                'Ü' => 'U', 'ü' => 'u',
                _ => ch
            };
            sb.Append(c);
        }
        return sb.ToString();
    }

    private class LogoInvRow
    {
        public string    Ficheno { get; set; } = "";
        public DateTime  Tarih   { get; set; }
        public decimal   Tutar   { get; set; }
        public short     Trcode  { get; set; }
        public string?   Taxnr   { get; set; }
        public string?   Tckno   { get; set; }
    }

    private class RedRow
    {
        public string?   RefId    { get; set; }
        public DateTime  Tarih    { get; set; }
        public string?   Aciklama { get; set; }
    }

    // FICHENO'yu sadeleştir: sadece harf ve rakam, büyük harfe çevir
    // ABC20260000000.1 → ABC202600000001
    // ABC202600000...1 → ABC2026000001  (rakam silinmişse haneler kaybolur)
    private static string NormalizeFichenoFull(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return "";
        var sb = new StringBuilder(s.Length);
        foreach (var c in s)
            if (char.IsLetterOrDigit(c)) sb.Append(char.ToUpperInvariant(c));
        return sb.ToString();
    }

    // FICHENO'nun son N rakamını al (modifiye edilen numaralarda sonun korunduğu kuralı için)
    // "ABC202600000...1" → son 4 rakam: "0001" (orijinalin son 4 hanesi)
    private static string SonHaneler(string? s, int n)
    {
        if (string.IsNullOrWhiteSpace(s) || n <= 0) return "";
        var sb = new StringBuilder(n);
        for (int i = s.Length - 1; i >= 0 && sb.Length < n; i--)
            if (char.IsDigit(s[i])) sb.Insert(0, s[i]);
        return sb.ToString();
    }

    // Logo GROSSTOTAL'ın hangi tutara denk geleceği fatura tipine göre değişiyor:
    //  - Tevkifatlı  → Matrah + KDV - Tevkifat (= ödenecek)
    //  - Tevkifatsız → Matrah + KDV (= brüt)
    //  - İade/bazı senaryolar → Matrah (sadece TaxExclusive)
    //  - UBL PayableAmount → her zaman güvenilir (ödenecek)
    // Hepsini ±1 TL toleransla dener; biri tutarsa eşleşmiş sayılır.
    private static bool TutarUyusuyor(EFaturaListItem item, decimal logoTutar)
    {
        const decimal tol = 1m;
        decimal brut         = item.Matrah + item.KdvTutar;
        decimal odenecek     = item.ToplamTutar > 0 ? item.ToplamTutar : brut - item.TevkifatTutar;
        decimal tevkifatsiz  = brut - item.TevkifatTutar;
        decimal matrahOnly   = item.Matrah;
        return Math.Abs(brut         - logoTutar) <= tol
            || Math.Abs(odenecek     - logoTutar) <= tol
            || Math.Abs(tevkifatsiz  - logoTutar) <= tol
            || Math.Abs(matrahOnly   - logoTutar) <= tol;
    }

    /// <summary>
    /// Listede gösterilecek "KDV Dahil Tutar" — UBL PayableAmount mevcutsa o,
    /// yoksa Matrah + KdvTutar fallback. (Line-level iskonto/charge'ların
    /// PayableAmount'a yansıdığı ama bizim toplamımıza yansımadığı senaryoları çözer.)
    /// </summary>
    public static decimal GosterilecekTutar(EFaturaListItem it)
        => it.ToplamTutar > 0 ? it.ToplamTutar : it.Matrah + it.KdvTutar;

    // İki normalize edilmiş FICHENO arasında pozisyonel benzerlik (0..1).
    // Uzunluk farkı dikkate alınır: max(uzunluk) paydaya konur.
    // Örn. "AEF2026000000021" ↔ "AAF2026000000021" → 15/16 = 0.9375
    //      "ABC2026000000001" ↔ "ABC2026000000002" → 15/16 = 0.9375
    private static double PozisyonelBenzerlik(string a, string b)
    {
        if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b)) return 0;
        int maxLen = Math.Max(a.Length, b.Length);
        int minLen = Math.Min(a.Length, b.Length);
        int eslesen = 0;
        for (int i = 0; i < minLen; i++)
            if (a[i] == b[i]) eslesen++;
        return (double)eslesen / maxLen;
    }

    public async Task<byte[]?> ZipGetirAsync(long id)
    {
        if (!BaglantiTanimliMi) return null;
        await using var con = new SqlConnection(_connStr);
        await con.OpenAsync();
        var data = await con.ExecuteScalarAsync<byte[]?>(
            "SELECT TOP 1 DATA FROM POSTBOX WITH (NOLOCK) WHERE ID=@id", new { id });
        return data;
    }

    public async Task<(string Xml, string DosyaAdi)?> XmlGetirAsync(long id)
    {
        var zip = await ZipGetirAsync(id);
        if (zip == null || zip.Length == 0) return null;

        // ZIP mi yoksa düz XML mi? İlk baytlardan tespit et.
        if (zip.Length >= 2 && zip[0] == 0x50 && zip[1] == 0x4B) // "PK"
        {
            using var ms = new MemoryStream(zip);
            using var arc = new ZipArchive(ms, ZipArchiveMode.Read);
            var entry = arc.Entries.FirstOrDefault(x =>
                x.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase));
            if (entry == null) return null;
            using var s = entry.Open();
            using var sr = new StreamReader(s, Encoding.UTF8);
            return (await sr.ReadToEndAsync(), entry.Name);
        }
        else
        {
            // Doğrudan XML
            return (Encoding.UTF8.GetString(zip), $"{id}.xml");
        }
    }

    public async Task<EFaturaDetay?> DetayGetirAsync(long id, string? karsiVknHint = null)
    {
        var pair = await XmlGetirAsync(id);
        if (pair == null) return null;
        var detay = Parse(pair.Value.Xml);
        detay.Id = id;
        detay.XmlIcerik = pair.Value.Xml;
        detay.XmlDosyaAdi = pair.Value.DosyaAdi;

        // Hangi taraf bize "karşı"? POSTBOX ENVELOPETYPE bilgisini detayda doğrudan tutmuyoruz —
        // bunun yerine list satırından gelen ipucu (karsiVknHint) kullanılır.
        string karsi = karsiVknHint ?? "";
        if (string.IsNullOrEmpty(karsi))
        {
            // Hint yoksa: alıcı/satıcıdan birini seç (çoğunlukla biz alıcıyız)
            karsi = string.IsNullOrEmpty(detay.SaticiVkn) ? detay.AliciVkn : detay.SaticiVkn;
        }
        if (!string.IsNullOrEmpty(karsi))
        {
            var map = await CariEslestirAsync(new[] { karsi });
            if (map.TryGetValue(karsi, out var c))
            {
                detay.LogoCariKod   = c.Kod;
                detay.LogoCariUnvan = c.Unvan;
            }
        }
        return detay;
    }

    public static EFaturaDetay Parse(string xmlContent)
    {
        var doc = new XmlDocument { PreserveWhitespace = false };
        doc.LoadXml(xmlContent);

        var ns = new XmlNamespaceManager(doc.NameTable);
        ns.AddNamespace("sbdh", "http://www.unece.org/cefact/namespaces/StandardBusinessDocumentHeader");
        ns.AddNamespace("inv",  "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2");
        ns.AddNamespace("cbc",  "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2");
        ns.AddNamespace("cac",  "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2");

        var inv = doc.SelectSingleNode("//inv:Invoice", ns) ?? doc.DocumentElement!;

        static string G(XmlNode? n, string xp, XmlNamespaceManager ns)
            => n?.SelectSingleNode(xp, ns)?.InnerText.Trim() ?? "";
        static string A(XmlNode? n, string xp, string attr, XmlNamespaceManager ns)
        {
            if (n?.SelectSingleNode(xp, ns) is XmlElement el) return el.GetAttribute(attr);
            return "";
        }

        var d = new EFaturaDetay
        {
            FaturaNo  = G(inv, "cbc:ID", ns),
            Uuid      = G(inv, "cbc:UUID", ns),
            Tarih     = TryDate(G(inv, "cbc:IssueDate", ns)),
            Saat      = G(inv, "cbc:IssueTime", ns),
            Tip       = G(inv, "cbc:InvoiceTypeCode", ns),
            ProfileId = G(inv, "cbc:ProfileID", ns),
            Doviz     = G(inv, "cbc:DocumentCurrencyCode", ns),
            DovizKuru = TryDec(G(inv, "cac:PricingExchangeRate/cbc:CalculationRate", ns)),

            SaticiUnvan       = G(inv, "cac:AccountingSupplierParty/cac:Party/cac:PartyName/cbc:Name", ns),
            SaticiVkn         = G(inv, "cac:AccountingSupplierParty/cac:Party/cac:PartyIdentification/cbc:ID", ns),
            SaticiVergiDairesi= G(inv, "cac:AccountingSupplierParty/cac:Party/cac:PartyTaxScheme/cac:TaxScheme/cbc:Name", ns),
            SaticiTelefon     = G(inv, "cac:AccountingSupplierParty/cac:Party/cac:Contact/cbc:Telephone", ns),
            SaticiEposta      = G(inv, "cac:AccountingSupplierParty/cac:Party/cac:Contact/cbc:ElectronicMail", ns),
            SaticiWeb         = G(inv, "cac:AccountingSupplierParty/cac:Party/cbc:WebsiteURI", ns),

            AliciUnvan        = G(inv, "cac:AccountingCustomerParty/cac:Party/cac:PartyName/cbc:Name", ns),
            AliciVkn          = G(inv, "cac:AccountingCustomerParty/cac:Party/cac:PartyIdentification/cbc:ID", ns),
            AliciVergiDairesi = G(inv, "cac:AccountingCustomerParty/cac:Party/cac:PartyTaxScheme/cac:TaxScheme/cbc:Name", ns),
        };

        // Adres (satıcı)
        var saticiAdr = inv.SelectSingleNode("cac:AccountingSupplierParty/cac:Party/cac:PostalAddress", ns);
        if (saticiAdr != null)
            d.SaticiAdres = BirlestirAdres(saticiAdr, ns);

        var aliciAdr = inv.SelectSingleNode("cac:AccountingCustomerParty/cac:Party/cac:PostalAddress", ns);
        if (aliciAdr != null)
            d.AliciAdres = BirlestirAdres(aliciAdr, ns);

        // Fatura seviyesi notlar (açıklamalar)
        var notes = inv.SelectNodes("cbc:Note", ns);
        if (notes != null)
            foreach (XmlNode n in notes)
            {
                var t = n.InnerText.Trim();
                if (!string.IsNullOrEmpty(t)) d.FaturaNotlari.Add(t);
            }

        // Kalemler
        var lines = inv.SelectNodes("cac:InvoiceLine", ns);
        if (lines != null)
        {
            int i = 0;
            foreach (XmlNode l in lines)
            {
                i++;
                d.Kalemler.Add(new EFaturaKalem
                {
                    Sira       = i,
                    Aciklama   = G(l, "cbc:Note", ns),
                    StokAdi    = G(l, "cac:Item/cbc:Name", ns),
                    StokKodu   = G(l, "cac:Item/cac:SellersItemIdentification/cbc:ID", ns),
                    Miktar     = TryDec(G(l, "cbc:InvoicedQuantity", ns)),
                    Birim      = A(l, "cbc:InvoicedQuantity", "unitCode", ns),
                    BirimFiyat = TryDec(G(l, "cac:Price/cbc:PriceAmount", ns)),
                    Doviz      = A(l, "cbc:LineExtensionAmount", "currencyID", ns),
                    NetTutar   = TryDec(G(l, "cbc:LineExtensionAmount", ns)),
                    KdvOran    = TryDec(G(l, "cac:TaxTotal/cac:TaxSubtotal/cbc:Percent", ns)),
                    KdvTutar   = TryDec(G(l, "cac:TaxTotal/cac:TaxSubtotal/cbc:TaxAmount", ns)),
                    KdvAdi     = G(l, "cac:TaxTotal/cac:TaxSubtotal/cac:TaxCategory/cac:TaxScheme/cbc:Name", ns),
                    IstisnaKodu= G(l, "cac:TaxTotal/cac:TaxSubtotal/cac:TaxCategory/cbc:TaxExemptionReasonCode", ns),
                    IstisnaSebep=G(l, "cac:TaxTotal/cac:TaxSubtotal/cac:TaxCategory/cbc:TaxExemptionReason", ns),
                    TevkifatKodu = G(l, "cac:WithholdingTaxTotal/cac:TaxSubtotal/cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode", ns),
                    TevkifatAdi  = G(l, "cac:WithholdingTaxTotal/cac:TaxSubtotal/cac:TaxCategory/cac:TaxScheme/cbc:Name", ns),
                    TevkifatOran = TryDec(G(l, "cac:WithholdingTaxTotal/cac:TaxSubtotal/cbc:Percent", ns)),
                    TevkifatTutar= TryDec(G(l, "cac:WithholdingTaxTotal/cac:TaxSubtotal/cbc:TaxAmount", ns)),
                });
            }
        }

        // LegalMonetaryTotal
        var mt = inv.SelectSingleNode("cac:LegalMonetaryTotal", ns);
        if (mt != null)
        {
            d.LineExtensionAmount  = TryDec(G(mt, "cbc:LineExtensionAmount", ns));
            d.TaxExclusiveAmount   = TryDec(G(mt, "cbc:TaxExclusiveAmount", ns));
            d.TaxInclusiveAmount   = TryDec(G(mt, "cbc:TaxInclusiveAmount", ns));
            d.AllowanceTotalAmount = TryDec(G(mt, "cbc:AllowanceTotalAmount", ns));
            d.ChargeTotalAmount    = TryDec(G(mt, "cbc:ChargeTotalAmount", ns));
            d.PayableAmount        = TryDec(G(mt, "cbc:PayableAmount", ns));
        }

        // Fatura seviyesi vergi alt-toplamları (KDV vb.)
        var tts = inv.SelectNodes("cac:TaxTotal/cac:TaxSubtotal", ns);
        if (tts != null)
            foreach (XmlNode t in tts)
                d.VergiOzetleri.Add(new EFaturaVergi
                {
                    Adi          = G(t, "cac:TaxCategory/cac:TaxScheme/cbc:Name", ns),
                    Kodu         = G(t, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode", ns),
                    Oran         = TryDec(G(t, "cbc:Percent", ns)),
                    Matrah       = TryDec(G(t, "cbc:TaxableAmount", ns)),
                    Tutar        = TryDec(G(t, "cbc:TaxAmount", ns)),
                    IstisnaKodu  = G(t, "cac:TaxCategory/cbc:TaxExemptionReasonCode", ns),
                    IstisnaSebep = G(t, "cac:TaxCategory/cbc:TaxExemptionReason", ns),
                });

        // Tevkifat alt-toplamları
        var wht = inv.SelectNodes("cac:WithholdingTaxTotal/cac:TaxSubtotal", ns);
        if (wht != null)
            foreach (XmlNode t in wht)
                d.TevkifatOzetleri.Add(new EFaturaVergi
                {
                    Adi    = G(t, "cac:TaxCategory/cac:TaxScheme/cbc:Name", ns),
                    Kodu   = G(t, "cac:TaxCategory/cac:TaxScheme/cbc:TaxTypeCode", ns),
                    Oran   = TryDec(G(t, "cbc:Percent", ns)),
                    Matrah = TryDec(G(t, "cbc:TaxableAmount", ns)),
                    Tutar  = TryDec(G(t, "cbc:TaxAmount", ns)),
                });

        return d;
    }

    private static string BirlestirAdres(XmlNode adr, XmlNamespaceManager ns)
    {
        var p = new List<string>();
        void Add(string xp) { var v = adr.SelectSingleNode(xp, ns)?.InnerText.Trim(); if (!string.IsNullOrWhiteSpace(v)) p.Add(v!); }
        Add("cbc:StreetName");
        Add("cbc:BuildingNumber");
        Add("cbc:Room");
        Add("cbc:CitySubdivisionName");
        Add("cbc:CityName");
        Add("cbc:PostalZone");
        Add("cac:Country/cbc:Name");
        return string.Join(" / ", p);
    }

    private static DateTime? TryDate(string s)
        => DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? d : null;

    private static decimal TryDec(string s)
        => decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ? d : 0m;
}
