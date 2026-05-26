using Dapper;

namespace RaporlamaPortali.Services.Logo;

// Bir malzeme/hizmet kartının Logo'daki birim setini (LG_211_UNITSETL satırlarını) döner.
// Logo'da otomatik UBL→yerel birim eşlemesi GLOBALCODE alanı üzerinden yapılır; ama setteki
// GLOBALCODE'lar eksik/yanlış girilmiş olabiliyor (örn. TON için "26" — TNE olmalı).
// Bu yüzden Logo'ya Aktar dialog'unda kullanıcı set'ten yerel birimi (KG/TON/GR/...) bizzat
// seçer; seçilen yerel kod doğrudan UNIT_CODE alanına yazılır.
public class LogoBirimSetiLookupService
{
    private readonly DatabaseService _db;
    private readonly ILogger<LogoBirimSetiLookupService> _log;

    public LogoBirimSetiLookupService(DatabaseService db, ILogger<LogoBirimSetiLookupService> log)
    {
        _db = db;
        _log = log;
    }

    public class BirimSatir
    {
        public short  LineNr     { get; set; }
        public string Code       { get; set; } = "";   // yerel kod (KG, TON, GR, TORBA…)
        public string Name       { get; set; } = "";
        public string GlobalCode { get; set; } = "";   // UBL evrensel kod (KGM, TNE, EA…)
        public bool   IsMain     { get; set; }
        public double ConvFact1  { get; set; } = 1;
        public double ConvFact2  { get; set; } = 1;

        public string Etiket => string.IsNullOrWhiteSpace(GlobalCode)
            ? Code
            : $"{Code} ({GlobalCode})";
    }

    // isHizmet=true: LG_211_SRVCARD, false: LG_211_ITEMS — UNITSETREF üzerinden UNITSETL satırları.
    public async Task<List<BirimSatir>> KartBirimleriAsync(string kartKodu, bool isHizmet)
    {
        if (string.IsNullOrWhiteSpace(kartKodu))
            return new List<BirimSatir>();

        try
        {
            var kartTbl = _db.GetTableName(isHizmet ? "SRVCARD" : "ITEMS");
            var setlTbl = _db.GetTableName("UNITSETL");

            await using var con = _db.CreateConnection();
            await con.OpenAsync();

            var rows = await con.QueryAsync<(
                short LineNr, string? Code, string? Name, string? GlobalCode,
                short? MainUnit, double? Conv1, double? Conv2)>($@"
                    SELECT UNI.LINENR, UNI.CODE, UNI.NAME, UNI.GLOBALCODE,
                           UNI.MAINUNIT, UNI.CONVFACT1, UNI.CONVFACT2
                    FROM {kartTbl} K WITH (NOLOCK)
                    INNER JOIN {setlTbl} UNI WITH (NOLOCK) ON UNI.UNITSETREF = K.UNITSETREF
                    WHERE K.CODE = @kod
                    ORDER BY UNI.LINENR", new { kod = kartKodu });

            return rows.Select(r => new BirimSatir
            {
                LineNr     = r.LineNr,
                Code       = (r.Code       ?? "").Trim(),
                Name       = (r.Name       ?? "").Trim(),
                GlobalCode = (r.GlobalCode ?? "").Trim(),
                IsMain     = (r.MainUnit   ?? 0) == 1,
                ConvFact1  = r.Conv1       ?? 1,
                ConvFact2  = r.Conv2       ?? 1,
            }).ToList();
        }
        catch (Exception ex)
        {
            _log.LogError(ex, "Birim seti lookup hata kart={Kod} hizmet={H}", kartKodu, isHizmet);
            return new List<BirimSatir>();
        }
    }

    // UBL'deki birim (KGM/TNE/EA…) için set içinde en uygun yerel kodu seçer:
    //   1) GLOBALCODE eşleşmesi  2) yerel CODE eşleşmesi  3) ana birim (MAINUNIT=1)  4) ilk satır
    public BirimSatir? OnerilenBirim(IReadOnlyList<BirimSatir> liste, string? ublBirim)
    {
        if (liste.Count == 0) return null;
        var ubl = (ublBirim ?? "").Trim();
        if (!string.IsNullOrEmpty(ubl))
        {
            var byGlobal = liste.FirstOrDefault(b => b.GlobalCode.Equals(ubl, StringComparison.OrdinalIgnoreCase));
            if (byGlobal != null) return byGlobal;
            var byLocal  = liste.FirstOrDefault(b => b.Code.Equals(ubl, StringComparison.OrdinalIgnoreCase));
            if (byLocal != null) return byLocal;
        }
        return liste.FirstOrDefault(b => b.IsMain) ?? liste[0];
    }
}
