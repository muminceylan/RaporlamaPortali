using Dapper;

namespace RaporlamaPortali.Services.Logo;

// LogoAktarimDialog'da kullanıcının hizmet/malzeme kartını seçmesi için autocomplete kaynağı.
//   - Alınan Hizmet faturası (TYPE=4)  → LG_211_SRVCARD (hizmet kartları)
//   - Satın Alma faturası   (TYPE=1)  → LG_211_ITEMS   (malzeme kartları)
public class LogoMasterKartLookupService
{
    private readonly DatabaseService _db;
    private readonly ILogger<LogoMasterKartLookupService> _log;

    public LogoMasterKartLookupService(DatabaseService db, ILogger<LogoMasterKartLookupService> log)
    {
        _db = db;
        _log = log;
    }

    public class MasterKart
    {
        public int    LogicalRef { get; set; }
        public string Kod        { get; set; } = "";
        public string Aciklama   { get; set; } = "";
        public bool   Aktif      { get; set; } = true;
    }

    // Hizmet kartı arama (LG_211_SRVCARD)
    public async Task<List<MasterKart>> HizmetAraAsync(string? arama, int limit = 50)
    {
        return await AraInternalAsync(_db.GetTableName("SRVCARD"), arama, limit, codeCol: "CODE", defCol: "DEFINITION_", activeCol: "ACTIVE");
    }

    // Malzeme kartı arama (LG_211_ITEMS, sadece CARDTYPE in (1,2,3,4,10,11) = malzeme + sabit kıymet vb.)
    public async Task<List<MasterKart>> MalzemeAraAsync(string? arama, int limit = 50)
    {
        var tbl = _db.GetTableName("ITEMS");
        var arr = (arama ?? "").Trim();
        var pat = arr.Length == 0 ? "%" : $"%{arr}%";

        try
        {
            await using var con = _db.CreateConnection();
            await con.OpenAsync();
            var rows = await con.QueryAsync<(int Lref, string? Code, string? Defn, short? Active)>($@"
                SELECT TOP (@limit) LOGICALREF AS Lref, CODE AS Code, NAME AS Defn, ACTIVE AS Active
                FROM {tbl} WITH (NOLOCK)
                WHERE (CODE LIKE @pat OR NAME LIKE @pat)
                  AND CARDTYPE IN (1,2,3,4,10,11,12,13,20)
                ORDER BY CODE", new { limit, pat });

            return rows.Select(r => new MasterKart
            {
                LogicalRef = r.Lref,
                Kod        = (r.Code ?? "").Trim(),
                Aciklama   = (r.Defn ?? "").Trim(),
                Aktif      = (r.Active ?? 0) == 0,
            }).ToList();
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "LogoMasterKart Malzeme arama hata");
            return new List<MasterKart>();
        }
    }

    private async Task<List<MasterKart>> AraInternalAsync(string tbl, string? arama, int limit, string codeCol, string defCol, string activeCol)
    {
        var arr = (arama ?? "").Trim();
        var pat = arr.Length == 0 ? "%" : $"%{arr}%";

        try
        {
            await using var con = _db.CreateConnection();
            await con.OpenAsync();
            var rows = await con.QueryAsync<(int Lref, string? Code, string? Defn, short? Active)>($@"
                SELECT TOP (@limit) LOGICALREF AS Lref, {codeCol} AS Code, {defCol} AS Defn, {activeCol} AS Active
                FROM {tbl} WITH (NOLOCK)
                WHERE ({codeCol} LIKE @pat OR {defCol} LIKE @pat)
                ORDER BY {codeCol}", new { limit, pat });

            return rows.Select(r => new MasterKart
            {
                LogicalRef = r.Lref,
                Kod        = (r.Code ?? "").Trim(),
                Aciklama   = (r.Defn ?? "").Trim(),
                Aktif      = (r.Active ?? 0) == 0,
            }).ToList();
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "LogoMasterKart arama hata tbl={Tbl}", tbl);
            return new List<MasterKart>();
        }
    }
}
