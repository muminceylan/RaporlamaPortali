namespace RaporlamaPortali.Services;

/// <summary>
/// Uygulamanın kalıcı verilerini publish klasörünün DIŞINDA, sabit bir konumda tutar.
/// Böylece her publish'te veritabanı, evrak arşivi, WhatsApp oturumu ve yetkili
/// numaralar etkilenmez.
/// </summary>
public static class AppDataPaths
{
    /// <summary>Tüm kalıcı verilerin tutulduğu kök dizin.</summary>
    public static string DataRoot { get; } = @"C:\RaporlamaPortaliData";

    public static string EvrakArsivDb      => Path.Combine(DataRoot, "EvrakArsiv.db");
    public static string TarimKrediDb      => Path.Combine(DataRoot, "TarimKredi.db");
    public static string EvrakArsivDizini  => Path.Combine(DataRoot, "EvrakArsiv");
    public static string GirisAyarlariJson => Path.Combine(DataRoot, "giris_ayarlari.json");
    public static string MailAyarlariJson  => Path.Combine(DataRoot, "mail_ayarlari.json");
    public static string LaunchAuthJson      => Path.Combine(DataRoot, "launch_auth.json");
    public static string MalzemeListeleriJson => Path.Combine(DataRoot, "malzeme_listeleri.json");
    public static string SabNetDb            => Path.Combine(DataRoot, "SabNet.db");
    public static string SabNetBaglantiJson  => Path.Combine(DataRoot, "sabnet_baglanti.json");
    public static string WhatsAppDataDir     => Path.Combine(DataRoot, "WhatsApp");

    /// <summary>
    /// Kök dizini oluşturur ve eski publish klasöründe kalmış verileri buraya taşır.
    /// Uygulama açılışında tek seferde çağrılmalı.
    /// </summary>
    public static void EnsureAndMigrate()
    {
        Directory.CreateDirectory(DataRoot);
        Directory.CreateDirectory(EvrakArsivDizini);
        Directory.CreateDirectory(WhatsAppDataDir);

        var exeDir = Path.GetDirectoryName(Environment.ProcessPath) ?? AppContext.BaseDirectory;

        // Tekil dosyalar: DB'ler ve ayarlar
        TasiEgerVarsa(Path.Combine(exeDir, "EvrakArsiv.db"),      EvrakArsivDb);
        TasiEgerVarsa(Path.Combine(exeDir, "TarimKredi.db"),      TarimKrediDb);
        TasiEgerVarsa(Path.Combine(exeDir, "giris_ayarlari.json"),GirisAyarlariJson);
        TasiEgerVarsa(Path.Combine(exeDir, "mail_ayarlari.json"), MailAyarlariJson);

        // Arşiv dosyaları dizini (alt klasörleriyle)
        var eskiArsiv = Path.Combine(exeDir, "EvrakArsiv");
        if (Directory.Exists(eskiArsiv) && Directory.EnumerateFileSystemEntries(eskiArsiv).Any())
            DizinIcerikTasi(eskiArsiv, EvrakArsivDizini);

        // WhatsApp veri dosyaları (config, status, log, .wwebjs_auth, screenshots)
        var eskiWhatsApp = Path.Combine(exeDir, "WhatsApp");
        if (Directory.Exists(eskiWhatsApp))
        {
            TasiEgerVarsa(Path.Combine(eskiWhatsApp, "whatsapp-config.json"),
                          Path.Combine(WhatsAppDataDir, "whatsapp-config.json"));
            TasiEgerVarsa(Path.Combine(eskiWhatsApp, "whatsapp-status.json"),
                          Path.Combine(WhatsAppDataDir, "whatsapp-status.json"));
            TasiEgerVarsa(Path.Combine(eskiWhatsApp, "whatsapp-log.json"),
                          Path.Combine(WhatsAppDataDir, "whatsapp-log.json"));

            DizinTasi(Path.Combine(eskiWhatsApp, ".wwebjs_auth"),
                      Path.Combine(WhatsAppDataDir, ".wwebjs_auth"));
            DizinTasi(Path.Combine(eskiWhatsApp, ".wwebjs_cache"),
                      Path.Combine(WhatsAppDataDir, ".wwebjs_cache"));
            DizinTasi(Path.Combine(eskiWhatsApp, "screenshots"),
                      Path.Combine(WhatsAppDataDir, "screenshots"));
        }
    }

    private static void TasiEgerVarsa(string kaynak, string hedef)
    {
        try
        {
            if (!File.Exists(kaynak)) return;
            if (File.Exists(hedef)) { File.Delete(kaynak); return; }
            Directory.CreateDirectory(Path.GetDirectoryName(hedef)!);
            File.Move(kaynak, hedef);
        }
        catch { /* migration hatası uygulamanın açılmasını engellemesin */ }
    }

    private static void DizinTasi(string kaynak, string hedef)
    {
        try
        {
            if (!Directory.Exists(kaynak)) return;
            if (Directory.Exists(hedef) && Directory.EnumerateFileSystemEntries(hedef).Any())
            {
                // Hedef zaten dolu — eski kaynağı sil (publish ezmesin diye)
                try { Directory.Delete(kaynak, recursive: true); } catch { }
                return;
            }
            if (Directory.Exists(hedef)) Directory.Delete(hedef);
            Directory.Move(kaynak, hedef);
        }
        catch { }
    }

    private static void DizinIcerikTasi(string kaynak, string hedef)
    {
        try
        {
            Directory.CreateDirectory(hedef);
            foreach (var dosya in Directory.GetFiles(kaynak, "*", SearchOption.AllDirectories))
            {
                var goreli = Path.GetRelativePath(kaynak, dosya);
                var yeni   = Path.Combine(hedef, goreli);
                Directory.CreateDirectory(Path.GetDirectoryName(yeni)!);
                if (!File.Exists(yeni)) File.Move(dosya, yeni);
                else                    File.Delete(dosya);
            }
            try { Directory.Delete(kaynak, recursive: true); } catch { }
        }
        catch { }
    }
}
