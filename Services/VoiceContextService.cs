using RaporlamaPortali.Models;

namespace RaporlamaPortali.Services;

/// <summary>
/// Sayfalar bu servise tool listelerini ve handler'larını kayıt eder.
/// MainLayout'taki global SesliKomutPanel bunları okuyup tek bir mikrofon
/// arayüzü üzerinden çalıştırır.
///
/// Pattern (sayfada):
///   protected override void OnInitialized()
///   {
///       VoiceCtx.Register(_voiceTools, "Sayfa Adı",
///                         getEkContext: VoiceContextOlustur,
///                         onayList: _voiceOnayAksiyonlar,
///                         onAksiyon: SesliAksiyonCalistir);
///   }
///
/// Sayfa değişince MainLayout NavigationManager.LocationChanged ile Clear() çağırır
/// — yani Dispose'ta Clear yapmaya gerek yok (race condition'ı önlemek için).
/// </summary>
public sealed class VoiceContextService
{
    public List<SesliKomutTool>      Tools                    { get; private set; } = new();
    public string                    SayfaAdi                 { get; private set; } = "Ana Sayfa";
    public Func<string>?             GetEkContext             { get; private set; }
    public HashSet<string>?          OnayGerektirenAksiyonlar { get; private set; }
    public Func<VoiceIntent, Task>?  OnAksiyon                { get; private set; }

    /// <summary>Tools veya handler değiştiğinde fırlatılır (UI yeniden render için).</summary>
    public event Action? Changed;

    public void Register(
        List<SesliKomutTool> tools,
        string sayfaAdi,
        Func<string>? getEkContext = null,
        HashSet<string>? onayList = null,
        Func<VoiceIntent, Task>? onAksiyon = null)
    {
        Tools                    = tools ?? new();
        SayfaAdi                 = string.IsNullOrWhiteSpace(sayfaAdi) ? "Sayfa" : sayfaAdi;
        GetEkContext             = getEkContext;
        OnayGerektirenAksiyonlar = onayList;
        OnAksiyon                = onAksiyon;
        Changed?.Invoke();
    }

    public void Clear()
    {
        Tools                    = new();
        SayfaAdi                 = "Sayfa";
        GetEkContext             = null;
        OnayGerektirenAksiyonlar = null;
        OnAksiyon                = null;
        Changed?.Invoke();
    }
}
