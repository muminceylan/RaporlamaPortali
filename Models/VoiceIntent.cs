namespace RaporlamaPortali.Models;

// Kullanıcının sesli komutu Claude API tarafından parse edilip aşağıdaki yapıya çevrilir.
// Frontend bu intent'i alır, sayfanın "voice action handler"ına gönderir, sayfa karar verir.

/// <summary>
/// Claude'dan dönen sesli komut sonucu. Tek bir Action + parametreleri.
/// Action sayfa-bazlı isim — örn "kopyala_son_fis", "satir_ekle", "miktar_degistir".
/// </summary>
public class VoiceIntent
{
    /// <summary>Aksiyon adı (snake_case). Sayfa bunu switch'le yorumlar.</summary>
    public string Action { get; set; } = "";

    /// <summary>Aksiyona göre parametreler (key/value).</summary>
    public Dictionary<string, object?> Params { get; set; } = new();

    /// <summary>Kullanıcıya gösterilecek doğrulama metni — onay penceresinde okunacak.</summary>
    public string OnayMetni { get; set; } = "";

    /// <summary>Claude'un komutu anlamadığı durumlar — UI uyarı gösterir.</summary>
    public bool Anlasildi { get; set; } = true;

    /// <summary>Anlaşılmadıysa açıklama (Claude'dan gelen).</summary>
    public string? Hata { get; set; }

    /// <summary>
    /// Komut zinciri — Action="komut_zinciri" olduğunda sırayla çalıştırılacak alt komutlar.
    /// Her eleman bir VoiceIntent (kendi Action + Params + OnayMetni).
    /// </summary>
    public List<VoiceIntent>? Adimlar { get; set; }
}

/// <summary>
/// Sayfanın desteklediği aksiyonların listesi — Claude'a "tool" olarak verilir.
/// Her sayfa kendi tool listesini tanımlar.
/// </summary>
public class SesliKomutTool
{
    public string Ad           { get; set; } = "";   // "kopyala_son_fis"
    public string Aciklama     { get; set; } = "";   // "Son fişlerden birini seçip forma kopyala"
    public Dictionary<string, SesliKomutParam> Parametreler { get; set; } = new();
}

public class SesliKomutParam
{
    public string Tip       { get; set; } = "string"; // "string", "number", "integer", "boolean"
    public string Aciklama  { get; set; } = "";
    public bool   Zorunlu   { get; set; } = false;
    /// <summary>Enum değerleri (varsa) — Claude bunlardan birini seçer.</summary>
    public List<string>? Secenekler { get; set; }
}
