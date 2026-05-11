namespace RaporlamaPortali.Services;

/// <summary>
/// Veri kesim tarihi — bu tarihten sonraki kayıtlar uygulamada işlenmez.
/// Kuralın uygulandığı menüler: Logo İşlemleri, Raporlar, Mutabakat ve Kur Farkı.
/// Kural sessiz çalışır — kullanıcıya uyarı verilmez, kesim tarihinden sonraki
/// veriler arka planda atılır.
///
/// ╔════════════════════════════════════════════════════════════════════╗
/// ║   TARİHİ DEĞİŞTİRMEK İÇİN:                                          ║
/// ║   Aşağıdaki <see cref="VeriKesimTarihi"/> alanını düzenle ve         ║
/// ║   uygulamayı yeniden yayınla (publish).                              ║
/// ║                                                                      ║
/// ║   Dosya yolu:                                                        ║
/// ║   C:\Users\muminceylan\Desktop\RaporlamaPortali\                     ║
/// ║       Services\DonemSiniri.cs                                        ║
/// ╚════════════════════════════════════════════════════════════════════╝
/// </summary>
public static class DonemSiniri
{
    // ╔═══════════════════════════════════════════════════════╗
    // ║   VERİ KESİM TARİHİ — BU SATIRI DEĞİŞTİR              ║
    // ╚═══════════════════════════════════════════════════════╝
    public static readonly DateTime VeriKesimTarihi = new DateTime(2026, 12, 31);

    /// <summary>Tarihi kesim tarihiyle sessizce sınırlar. t > kesim ise kesim döner.</summary>
    public static DateTime Kirp(DateTime t) => t > VeriKesimTarihi ? VeriKesimTarihi : t;

    public static DateTime? Kirp(DateTime? t) => t.HasValue ? Kirp(t.Value) : null;

    /// <summary>Ref ile geçilen bitiş tarihini yerinde sessizce kırpar.</summary>
    public static void Kirp(ref DateTime bitis)
    {
        if (bitis > VeriKesimTarihi) bitis = VeriKesimTarihi;
    }

    public static void Kirp(ref DateTime? bitis)
    {
        if (bitis.HasValue && bitis.Value > VeriKesimTarihi) bitis = VeriKesimTarihi;
    }
}
