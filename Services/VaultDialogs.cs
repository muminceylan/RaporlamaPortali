using System.Windows.Forms;

namespace RaporlamaPortali.Services;

// Vault ile ilgili tüm WinForms dialogları aynı dosyada — sadece LaunchAuthService kullanır.

// İlk kurulumda recovery key'i kullanıcıya bir kez gösteren dialog.
// "Kaydettim" onayı verilmeden continue aktif olmaz.
internal sealed class VaultRecoveryKeyForm : Form
{
    private readonly TextBox  _keyBox;
    private readonly CheckBox _onay;
    private readonly Button   _devam;

    public VaultRecoveryKeyForm(string recoveryKey)
    {
        Text             = "Raporlama Portalı — Kurtarma Anahtarı";
        FormBorderStyle  = FormBorderStyle.FixedDialog;
        StartPosition    = FormStartPosition.CenterScreen;
        MaximizeBox      = false;
        MinimizeBox      = false;
        ShowInTaskbar    = true;
        TopMost          = true;
        ClientSize       = new Size(620, 410);
        ControlBox       = false;

        var basliklbl = new Label
        {
            Text      = "KURTARMA ANAHTARI — bir kez gösterilir, kaybolursa geri getirilemez",
            Font      = new Font("Segoe UI", 11f, FontStyle.Bold),
            ForeColor = Color.DarkRed,
            Location  = new Point(15, 15),
            Size      = new Size(590, 28),
        };

        var aciklama = new Label
        {
            Text =
                "Şifrenizi unutursanız VEYA bilgisayarınız bozulup yeni cihaza geçerseniz " +
                "bu anahtar tek kurtarma yolunuzdur. Şimdi:\n" +
                "  1) Aşağıdaki anahtarı GÜVENLİ bir yere kaydedin (parola yöneticisi / kasa / yazıcı).\n" +
                "  2) Bilgisayarda saklamayın — kötü niyetli biri eline geçirirse şifre olmadan da açabilir.\n" +
                "  3) Kaybolursa kimse (ne ben, ne başka biri) bu programı tekrar açamaz.",
            Location = new Point(15, 50),
            Size     = new Size(590, 95),
        };

        _keyBox = new TextBox
        {
            Text         = recoveryKey,
            ReadOnly     = true,
            Font         = new Font("Consolas", 16f, FontStyle.Bold),
            TextAlign    = HorizontalAlignment.Center,
            BorderStyle  = BorderStyle.FixedSingle,
            BackColor    = Color.LightYellow,
            Location     = new Point(15, 155),
            Size         = new Size(590, 50),
        };

        var kopyala = new Button
        {
            Text     = "Panoya Kopyala",
            Location = new Point(15, 220),
            Size     = new Size(140, 32),
        };
        kopyala.Click += (_, _) =>
        {
            Clipboard.SetText(recoveryKey);
            kopyala.Text = "✓ Kopyalandı";
        };

        var dosyaKaydet = new Button
        {
            Text     = "Dosyaya Kaydet (.txt)",
            Location = new Point(165, 220),
            Size     = new Size(160, 32),
        };
        dosyaKaydet.Click += (_, _) =>
        {
            using var dlg = new SaveFileDialog
            {
                FileName = "raporlama-portali-kurtarma-anahtari.txt",
                Filter   = "Metin dosyası (*.txt)|*.txt",
            };
            // TopMost ana form alt dialog'u arkaya saklayabiliyor — geçici kapat, owner ver.
            var prev = TopMost;
            TopMost = false;
            try
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    File.WriteAllText(dlg.FileName,
                        "RAPORLAMA PORTALI — KURTARMA ANAHTARI\r\n" +
                        "=====================================\r\n\r\n" +
                        recoveryKey + "\r\n\r\n" +
                        "Bu dosyayı GÜVENLİ bir yerde tutun. Bilgisayarda kalmasın.\r\n" +
                        $"Üretildi: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\r\n");
                    dosyaKaydet.Text = "✓ Kaydedildi";
                }
            }
            finally
            {
                TopMost = prev;
                Activate();
            }
        };

        _onay = new CheckBox
        {
            Text     = "Anahtarı kaydettim, kaybedersem programı tekrar açamayacağımı anlıyorum.",
            Location = new Point(15, 280),
            Size     = new Size(590, 24),
            Font     = new Font("Segoe UI", 9f, FontStyle.Bold),
        };

        _devam = new Button
        {
            Text         = "Devam Et",
            DialogResult = DialogResult.OK,
            Location     = new Point(490, 360),
            Size         = new Size(115, 35),
            Enabled      = false,
            Font         = new Font("Segoe UI", 9f, FontStyle.Bold),
        };

        _onay.CheckedChanged += (_, _) => _devam.Enabled = _onay.Checked;

        Controls.AddRange(new Control[] {
            basliklbl, aciklama, _keyBox, kopyala, dosyaKaydet, _onay, _devam,
        });
        AcceptButton = _devam;
    }
}

// Şifre 3 kez yanlış girildiğinde devreye giren kurtarma akışı.
// Önce recovery key alır, doğruysa yeni şifre belirletir.
internal sealed class VaultRecoveryUnlockForm : Form
{
    private readonly TextBox _keyBox;
    private readonly Label   _hata;
    public string RecoveryKey => _keyBox.Text.Trim();

    public VaultRecoveryUnlockForm()
    {
        Text             = "Raporlama Portalı — Kurtarma Anahtarı ile Aç";
        FormBorderStyle  = FormBorderStyle.FixedDialog;
        StartPosition    = FormStartPosition.CenterScreen;
        MaximizeBox      = false;
        MinimizeBox      = false;
        TopMost          = true;
        ClientSize       = new Size(560, 230);

        var info = new Label
        {
            Text =
                "Şifre defalarca yanlış girildi. Eğer kurtarma anahtarınız varsa aşağıya yapıştırın.\n" +
                "(8 grup x 4 karakter; tireler önemli değil, küçük/büyük harf farketmez)",
            Location = new Point(15, 15),
            Size     = new Size(530, 50),
        };

        _keyBox = new TextBox
        {
            Font        = new Font("Consolas", 12f),
            Location    = new Point(15, 75),
            Size        = new Size(530, 30),
            CharacterCasing = CharacterCasing.Upper,
        };

        _hata = new Label
        {
            ForeColor = Color.Red,
            Location  = new Point(15, 115),
            Size      = new Size(530, 22),
        };

        var ok = new Button
        {
            Text         = "Doğrula",
            DialogResult = DialogResult.OK,
            Location     = new Point(360, 175),
            Size         = new Size(90, 32),
        };
        var iptal = new Button
        {
            Text         = "İptal",
            DialogResult = DialogResult.Cancel,
            Location     = new Point(460, 175),
            Size         = new Size(85, 32),
        };

        Controls.AddRange(new Control[] { info, _keyBox, _hata, ok, iptal });
        AcceptButton = ok;
        CancelButton = iptal;
        ActiveControl = _keyBox;
    }

    public void HataGoster(string mesaj) => _hata.Text = mesaj;
}

// Recovery sonrası yeni şifre belirleme.
internal sealed class VaultNewPasswordForm : Form
{
    private readonly TextBox _pwd;
    private readonly TextBox _pwd2;
    private readonly Label   _err;
    public string Password => _pwd.Text;

    public VaultNewPasswordForm()
    {
        Text             = "Raporlama Portalı — Yeni Şifre";
        FormBorderStyle  = FormBorderStyle.FixedDialog;
        StartPosition    = FormStartPosition.CenterScreen;
        MaximizeBox      = false;
        MinimizeBox      = false;
        TopMost          = true;
        ClientSize       = new Size(440, 240);

        var info = new Label
        {
            Text     = "Kurtarma başarılı. Şimdi yeni başlatma şifrenizi belirleyin.",
            Location = new Point(15, 15),
            Size     = new Size(410, 40),
        };
        var lbl1 = new Label
        {
            Text     = "Yeni şifre:",
            Location = new Point(15, 70),
            AutoSize = true,
        };
        _pwd = new TextBox
        {
            UseSystemPasswordChar = true,
            Location              = new Point(140, 67),
            Size                  = new Size(285, 25),
        };
        var lbl2 = new Label
        {
            Text     = "Yeni şifre (tekrar):",
            Location = new Point(15, 105),
            AutoSize = true,
        };
        _pwd2 = new TextBox
        {
            UseSystemPasswordChar = true,
            Location              = new Point(140, 102),
            Size                  = new Size(285, 25),
        };
        _err = new Label
        {
            ForeColor = Color.Red,
            Location  = new Point(15, 140),
            Size      = new Size(410, 22),
        };

        var ok = new Button
        {
            Text     = "Kaydet",
            Location = new Point(245, 185),
            Size     = new Size(85, 32),
        };
        var iptal = new Button
        {
            Text         = "İptal",
            DialogResult = DialogResult.Cancel,
            Location     = new Point(340, 185),
            Size         = new Size(85, 32),
        };
        ok.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_pwd.Text) || _pwd.Text.Length < 4)
            {
                _err.Text = "Şifre en az 4 karakter olmalı."; return;
            }
            if (_pwd.Text != _pwd2.Text)
            {
                _err.Text = "Şifreler eşleşmiyor."; return;
            }
            DialogResult = DialogResult.OK;
            Close();
        };

        Controls.AddRange(new Control[] { info, lbl1, _pwd, lbl2, _pwd2, _err, ok, iptal });
        AcceptButton  = ok;
        CancelButton  = iptal;
        ActiveControl = _pwd;
    }
}
