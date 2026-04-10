namespace RaporlamaPortali.Models;

/// <summary>
/// Mail gönderim ayarları
/// </summary>
public class MailAyarlari
{
    public string SmtpServer { get; set; } = "smtp.office365.com";
    public int SmtpPort { get; set; } = 587;
    public bool UseSsl { get; set; } = true;
    public string GondericiMail { get; set; } = "muminceylan@doguscay.com.tr";
    public string GondericiAdi { get; set; } = "Muhasebe Raporları";
    public string Sifre { get; set; } = ""; // appsettings.json'dan okunacak
    
    /// <summary>
    /// BCC olarak gönderilecek alıcılar (gizli)
    /// </summary>
    public List<string> Alicilar { get; set; } = new();
}

/// <summary>
/// Mail gönderim sonucu
/// </summary>
public class MailSonuc
{
    public bool Basarili { get; set; }
    public string Mesaj { get; set; } = "";
    public DateTime GonderimZamani { get; set; } = DateTime.Now;
    public int AliciSayisi { get; set; }
}

/// <summary>
/// Zamanlanmış görev ayarları
/// </summary>
public class ZamanlamaAyarlari
{
    public TimeSpan GonderimSaati { get; set; } = new TimeSpan(12, 0, 0); // 12:00
    public bool PazartesiCuma { get; set; } = true; // Hafta içi mi?
    public bool HaftaSonu { get; set; } = false;
}
