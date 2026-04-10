using Dapper;
using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Pancar Ödeme Raporları servisi
/// </summary>
public class PancarOdemeService
{
    private readonly DatabaseService _db;

    public PancarOdemeService(DatabaseService db)
    {
        _db = db;
    }

    // TODO: Pancar ödeme raporları için metodlar eklenecek
    // Çiftçi borç/alacak, avans takibi
    public Task<object> GetPancarOdemeOzetAsync(DateTime baslangic, DateTime bitis)
    {
        throw new NotImplementedException("Pancar Ödeme modülü ilerleyen aşamada eklenecek.");
    }
}
