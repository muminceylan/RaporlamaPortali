using Dapper;

namespace RaporlamaPortali.Services.Logo;

// E-faturanın karşı firmasının VKN/TCKN'sinden Logo cari kart(lar)ı ve her cari için
// alış faturasında kullanılacak Muhasebe Hesap Kodu'nu (GL_CODE) bulur.
//
// GL_CODE şu join ile alınır (örnek olarak doğrulandı: 320.06.3410 → 320.99.06.3410):
//   LG_211_CLCARD cl
//   INNER JOIN LG_211_CRDACREF rf ON rf.CARDREF = cl.LOGICALREF AND rf.TRCODE = 5
//   INNER JOIN LG_211_EMUHACC acc ON acc.LOGICALREF = rf.ACCOUNTREF
//   => acc.CODE
// TRCODE = 5  satın alma faturası işlemi için cari hesap (alış)
public class LogoCariLookupService
{
    private readonly DatabaseService _db;
    private readonly ILogger<LogoCariLookupService> _log;

    public LogoCariLookupService(DatabaseService db, ILogger<LogoCariLookupService> log)
    {
        _db  = db;
        _log = log;
    }

    public class CariAdayi
    {
        public int     LogicalRef { get; set; }
        public string  Kod        { get; set; } = "";
        public string  Unvan      { get; set; } = "";
        public string  Vkn        { get; set; } = "";
        public string  Tckn       { get; set; } = "";
        public string  GlKod      { get; set; } = "";   // muhasebe hesap kodu (320.99.xxx)
        public bool    GlKodVar   => !string.IsNullOrWhiteSpace(GlKod);
    }

    public class CariBakiye
    {
        public decimal  ToplamBorc      { get; set; }
        public decimal  ToplamAlacak    { get; set; }
        public decimal  Bakiye          => ToplamBorc - ToplamAlacak;   // + → biz alacaklıyız, − → biz borçluyuz
        public string   BakiyeYon       => Bakiye >= 0 ? "BORÇ" : "ALACAK";
        public decimal  BakiyeMutlak    => Math.Abs(Bakiye);
    }

    public enum FaturaMevcutDurumu
    {
        Yok            = 0,
        AyniVknMevcut  = 1,   // ⛔ MÜKERRER — aynı firma için zaten işlenmiş, ödeme riski
        BaskaVknMevcut = 2,   // ⚠️ Başka firmada — duplicate key riski, K-prefix gerekir
    }

    public class FaturaMevcutSonuc
    {
        public FaturaMevcutDurumu Durum   { get; set; }
        public string             FicheNo { get; set; } = "";   // Logo'da bulduğumuz Fiş No
        public string             DocOde  { get; set; } = "";   // Logo'da bulduğumuz Belge No
        public DateTime?          Tarih   { get; set; }
        public string             CariKod   { get; set; } = "";
        public string             CariUnvan { get; set; } = "";
        public string             CariVkn   { get; set; } = "";
        public decimal            Tutar     { get; set; }
    }

    // Logo'da bu fatura no zaten var mı kontrolü.
    //
    // SQL: LG_{firma}_01_INVOICE üzerinde FICHENO veya DOCODE eşleşmesi.
    // VKN match → AyniVknMevcut (mükerrer ödeme riski!), VKN farklı → BaskaVknMevcut (K-prefix gerekir).
    // Hem orijinal hem K-prefix varyantlarını arar (kullanıcı önceden K eklemişse veya silmişse).
    public async Task<FaturaMevcutSonuc> FaturaMevcutMuAsync(string? faturaNo, string? vkn, CancellationToken ct = default)
    {
        var sonuc = new FaturaMevcutSonuc();
        if (string.IsNullOrWhiteSpace(faturaNo)) return sonuc;

        try
        {
            var invTbl = _db.GetPeriodTableName("INVOICE");
            var clcTbl = _db.GetTableName("CLCARD");

            var varyantlar = FaturaNoYardimcisi.Varyantlar(faturaNo).ToArray();

            var sql = $@"
                SELECT TOP 1 i.FICHENO, i.DOCODE, i.DATE_, i.GROSSTOTAL,
                       c.CODE AS CariKod, c.DEFINITION_ AS CariUnvan, c.TAXNR AS CariVkn,
                       CASE WHEN c.TAXNR = @vkn OR c.TCKNO = @vkn THEN 0 ELSE 1 END AS VknOnce
                FROM {invTbl} i WITH (NOLOCK)
                INNER JOIN {clcTbl} c WITH (NOLOCK) ON c.LOGICALREF = i.CLIENTREF
                WHERE (i.FICHENO IN @nos OR i.DOCODE IN @nos) AND i.CANCELLED = 0
                ORDER BY VknOnce, i.DATE_ DESC";

            await using var con = _db.CreateConnection();
            await con.OpenAsync(ct);
            var row = await con.QueryFirstOrDefaultAsync(new CommandDefinition(sql,
                new { nos = varyantlar, vkn = (vkn ?? "").Trim() }, cancellationToken: ct));

            if (row == null) return sonuc;

            string bulunanVkn = ((string?)row.CariVkn ?? "").Trim();
            bool ayniVkn = !string.IsNullOrWhiteSpace(vkn) &&
                           string.Equals(bulunanVkn, vkn.Trim(), StringComparison.OrdinalIgnoreCase);

            sonuc.Durum     = ayniVkn ? FaturaMevcutDurumu.AyniVknMevcut : FaturaMevcutDurumu.BaskaVknMevcut;
            sonuc.FicheNo   = ((string?)row.FICHENO ?? "").Trim();
            sonuc.DocOde    = ((string?)row.DOCODE  ?? "").Trim();
            sonuc.Tarih     = row.DATE_ as DateTime?;
            sonuc.CariKod   = ((string?)row.CariKod   ?? "").Trim();
            sonuc.CariUnvan = ((string?)row.CariUnvan ?? "").Trim();
            sonuc.CariVkn   = bulunanVkn;
            sonuc.Tutar     = row.GROSSTOTAL is null ? 0m : Convert.ToDecimal(row.GROSSTOTAL);
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "FaturaMevcutMu kontrol başarısız (no={N}, vkn={V})", faturaNo, vkn);
        }
        return sonuc;
    }

    // Cari hesabın güncel bakiyesi (LG_{firma}_{donem}_CLFLINE üzerinden).
    // SIGN=0 BORÇ, SIGN=1 ALACAK. Bakiye = BORÇ − ALACAK.
    public async Task<CariBakiye> BakiyeAsync(string cariKod, CancellationToken ct = default)
    {
        var sonuc = new CariBakiye();
        if (string.IsNullOrWhiteSpace(cariKod)) return sonuc;

        try
        {
            var clfTbl = _db.GetPeriodTableName("CLFLINE");
            var clcTbl = _db.GetTableName("CLCARD");
            var sql = $@"
                SELECT
                    SUM(CASE WHEN f.SIGN = 0 THEN f.AMOUNT ELSE 0 END) AS ToplamBorc,
                    SUM(CASE WHEN f.SIGN = 1 THEN f.AMOUNT ELSE 0 END) AS ToplamAlacak
                FROM {clfTbl} f WITH (NOLOCK)
                INNER JOIN {clcTbl} c WITH (NOLOCK) ON c.LOGICALREF = f.CLIENTREF
                WHERE c.CODE = @kod AND f.CANCELLED = 0";

            await using var con = _db.CreateConnection();
            await con.OpenAsync(ct);
            var row = await con.QueryFirstOrDefaultAsync<(decimal? Borc, decimal? Alacak)>(
                new CommandDefinition(sql, new { kod = cariKod }, cancellationToken: ct));
            sonuc.ToplamBorc   = row.Borc   ?? 0m;
            sonuc.ToplamAlacak = row.Alacak ?? 0m;
        }
        catch (Exception ex)
        {
            _log.LogWarning(ex, "Cari bakiye sorgu hata: kod={K}", cariKod);
        }
        return sonuc;
    }

    // VKN/TCKN ile eşleşen tüm carileri (genellikle 1, bazen >1) ve her birinin GL_CODE'unu döner.
    public async Task<List<CariAdayi>> BulAsync(string vknTckn)
    {
        var sonuc = new List<CariAdayi>();
        if (string.IsNullOrWhiteSpace(vknTckn)) return sonuc;

        try
        {
            var clcTbl = _db.GetTableName("CLCARD");
            var rfTbl  = _db.GetTableName("CRDACREF");
            var accTbl = _db.GetTableName("EMUHACC");

            await using var con = _db.CreateConnection();
            await con.OpenAsync();

            // Önce VKN/TCKN ile eşleşen carileri al
            var cariler = (await con.QueryAsync<(int LogicalRef, string? Code, string? Definition, string? Taxnr, string? Tckno)>($@"
                SELECT LOGICALREF AS LogicalRef, CODE AS Code, DEFINITION_ AS Definition_, TAXNR AS Taxnr, TCKNO AS Tckno
                FROM {clcTbl} WITH (NOLOCK)
                WHERE TAXNR = @v OR TCKNO = @v", new { v = vknTckn.Trim() })).ToList();

            if (cariler.Count == 0) return sonuc;

            // Her cari için GL_CODE'u tek query'de al (TRCODE = 5 → satın alma faturası)
            var refs = cariler.Select(c => c.LogicalRef).ToArray();
            var glRows = (await con.QueryAsync<(int CardRef, string? Code)>($@"
                SELECT rf.CARDREF AS CardRef, acc.CODE AS Code
                FROM {rfTbl} rf WITH (NOLOCK)
                INNER JOIN {accTbl} acc WITH (NOLOCK) ON acc.LOGICALREF = rf.ACCOUNTREF
                WHERE rf.CARDREF IN @refs AND rf.TRCODE = 5", new { refs })).ToList();

            var glMap = glRows
                .GroupBy(x => x.CardRef)
                .ToDictionary(g => g.Key, g => (g.First().Code ?? "").Trim());

            foreach (var c in cariler)
            {
                sonuc.Add(new CariAdayi
                {
                    LogicalRef = c.LogicalRef,
                    Kod        = (c.Code ?? "").Trim(),
                    Unvan      = (c.Definition ?? "").Trim(),
                    Vkn        = (c.Taxnr ?? "").Trim(),
                    Tckn       = (c.Tckno ?? "").Trim(),
                    GlKod      = glMap.TryGetValue(c.LogicalRef, out var g) ? g : "",
                });
            }
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "LogoCariLookup hata vknTckn={V}", vknTckn);
        }

        return sonuc;
    }
}
