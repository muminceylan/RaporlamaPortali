namespace RaporlamaPortali.Services;

// Sistem genelinde veri çekiminde kullanılan üst sınır (EXCLUSIVE).
// Bu tarihten ÖNCESİ ve bu tarih dahil değildir; yani dahil edilen son gün = SistemBitisTarihi - 1 gün.
//
// Değişiklik için: aşağıdaki tek satırı güncelle.
// Yeni değer için örnekler:
//   new DateTime(2027, 1, 1)  → 31.12.2026 dahil, 01.01.2027 ve sonrası HARİÇ
//   new DateTime(2026, 9, 1)  → 31.08.2026 dahil, 01.09.2026 ve sonrası HARİÇ
internal static class SistemTarihi
{
    public static readonly DateTime SistemUstSinir = new DateTime(2026, 9, 1);

    // Dahil edilen son gün (SistemUstSinir - 1 gün). Gün bazlı Logo DATE_ alanları için.
    public static DateTime SonDahilGun => SistemUstSinir.AddDays(-1);

    // Bir bitiş tarihini sistem üst sınırına indir (eğer aşıyorsa).
    public static DateTime Clamp(DateTime t)
        => t > SonDahilGun ? SonDahilGun : t;

    public static DateTime? Clamp(DateTime? t)
        => t.HasValue ? Clamp(t.Value) : null;

    // Bugün yerine kullan: bugün üst sınırdan büyükse üst sınıra sabitle.
    public static DateTime Bugun()
        => Clamp(DateTime.Today);
}
