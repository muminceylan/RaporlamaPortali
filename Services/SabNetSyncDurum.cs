namespace RaporlamaPortali.Services;

/// <summary>
/// SabNet otomatik senkronizasyon arka plan servisinin paylaşımlı durumunu tutar.
/// Singleton; UI ile background service arasında köprü.
/// </summary>
public class SabNetSyncDurum
{
    public bool OtomatikSyncAcik { get; set; } = true;

    public DateTime? SonSyncZamani { get; set; }
    public int SonSyncEklenenHareket { get; set; }
    public int SonSyncEklenenLog { get; set; }
    public double SonSyncSureSn { get; set; }

    public string? SonHata { get; set; }
    public DateTime? SonHataZamani { get; set; }

    /// <summary>Manuel aktarım/temizleme ile background sync'in çakışmasını önler.</summary>
    public SemaphoreSlim Kilit { get; } = new(1, 1);

    public event Action? Degisti;
    public void Bildir() => Degisti?.Invoke();
}
