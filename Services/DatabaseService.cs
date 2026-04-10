using Microsoft.Data.SqlClient;
using System.Data;

namespace RaporlamaPortali.Services;

/// <summary>
/// SQL Server veritabanı bağlantı yönetimi
/// </summary>
public class DatabaseService
{
    private readonly string _connectionString;
    private readonly int _firmaNo;
    private readonly int _donemNo;

    public DatabaseService(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("LogoDB") 
            ?? throw new InvalidOperationException("LogoDB connection string bulunamadı!");
        _firmaNo = configuration.GetValue<int>("FirmaNo", 211);
        _donemNo = configuration.GetValue<int>("DonemNo", 1); // Dönem numarası
    }

    public string ConnectionString => _connectionString;
    public int FirmaNo => _firmaNo;
    public int DonemNo => _donemNo;

    /// <summary>
    /// Yeni SQL bağlantısı oluşturur
    /// </summary>
    public SqlConnection CreateConnection()
    {
        return new SqlConnection(_connectionString);
    }

    /// <summary>
    /// Bağlantı testi yapar
    /// </summary>
    public async Task<(bool Success, string Message)> TestConnectionAsync()
    {
        try
        {
            using var conn = CreateConnection();
            await conn.OpenAsync();
            return (true, $"Bağlantı başarılı! Server: {conn.DataSource}, Database: {conn.Database}");
        }
        catch (Exception ex)
        {
            return (false, $"Bağlantı hatası: {ex.Message}");
        }
    }

    /// <summary>
    /// Logo tablo adını firma numarası ile oluşturur (dönem bağımsız tablolar)
    /// Örnek: LG_211_ITEMS, LG_211_CLCARD
    /// </summary>
    public string GetTableName(string baseTable)
    {
        return $"LG_{_firmaNo}_{baseTable}";
    }

    /// <summary>
    /// Logo tablo adını firma ve dönem numarası ile oluşturur (dönem bağımlı tablolar)
    /// Örnek: LG_211_01_STFICHE, LG_211_01_STLINE
    /// </summary>
    public string GetPeriodTableName(string baseTable)
    {
        return $"LG_{_firmaNo}_{_donemNo:D2}_{baseTable}";
    }

    /// <summary>
    /// Logo view adını firma numarası ile oluşturur
    /// Örnek: LV_211_01_STINVTOT
    /// </summary>
    public string GetViewName(string baseView)
    {
        return $"LV_{_firmaNo}_{_donemNo:D2}_{baseView}";
    }
}
