using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows.Forms;

namespace RaporlamaPortali.Services;

// Uygulama exe'si açılırken:
//   1) Şifre sorar (launch_auth.json'da PBKDF2-SHA256 hash).
//   2) Doğru şifre ile vault.json'u açar — yoksa yeni vault üretip recovery key gösterir.
//   3) SecretsService'i ayağa kaldırır → ConnectionStrings, MailPassword, ANTHROPIC_API_KEY
//      hep şifrelenmiş kasadan okunur.
// Yanlış şifre 3 kez denenince recovery key sorulur; doğruysa yeni şifre belirlenir.
public static class LaunchAuthService
{
    private const int MaxAttempts = 3;
    private const int Iterations  = 100_000;
    private const int HashBytes   = 32;
    private const int SaltBytes   = 16;

    public static void Require()
    {
        ApplicationConfiguration.Initialize();
        Directory.CreateDirectory(AppDataPaths.DataRoot);

        var (hash, salt) = ReadStoredHash();

        string password = string.IsNullOrEmpty(hash) || string.IsNullOrEmpty(salt)
            ? RunFirstTimeSetup()
            : RunPasswordOrRecovery(hash!, salt!);

        // Vault yoksa: ya tamamen yeni kurulum, ya da eski kurulum + ilk Seviye3 başlangıcı.
        // Aynı şifreyle vault'u oluştur ve recovery key'i göster.
        if (!VaultService.VaultVarMi())
        {
            var recoveryKey = VaultService.CreateNewVault(password);
            using var rk = new VaultRecoveryKeyForm(recoveryKey);
            rk.ShowDialog();
        }
        else if (!VaultService.Acik)
        {
            // launch_auth doğru ama vault unlock olmadı (edge case: launch ile vault şifreleri ayrışmış).
            if (!VaultService.TryUnlockWithPassword(password))
            {
                if (!RunRecoveryUnlock())
                {
                    MessageBox.Show(
                        "Vault açılamadı. Kurtarma anahtarı olmadan devam edilemez.",
                        "Raporlama Portalı",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    Environment.Exit(1);
                }
                // Recovery başarılı → launch ile vault şifrelerini eşitle
                ResetLaunchAndVaultPassword();
            }
        }

        SecretsService.LoadOrCreate();
    }

    // ---------- ana akışlar ----------

    private static string RunFirstTimeSetup()
    {
        using var form = new SetupForm();
        if (form.ShowDialog() != DialogResult.OK)
            Environment.Exit(1);

        var newSalt = RandomNumberGenerator.GetBytes(SaltBytes);
        var newHash = Hash(form.Password, newSalt);
        WriteStoredHash(Convert.ToBase64String(newHash), Convert.ToBase64String(newSalt));
        return form.Password;
    }

    private static string RunPasswordOrRecovery(string storedHashB64, string storedSaltB64)
    {
        var storedHash = Convert.FromBase64String(storedHashB64);
        var salt       = Convert.FromBase64String(storedSaltB64);

        for (int i = 0; i < MaxAttempts; i++)
        {
            using var form = new PromptForm(i + 1, MaxAttempts);
            if (form.ShowDialog() != DialogResult.OK)
                Environment.Exit(1);

            var candidate = Hash(form.Password, salt);
            if (!CryptographicOperations.FixedTimeEquals(candidate, storedHash))
                continue;

            // launch_auth eşleşti — vault da varsa aynı şifreyle açmayı dene.
            if (!VaultService.VaultVarMi() || VaultService.TryUnlockWithPassword(form.Password))
                return form.Password;

            // launch doğru, vault yanlış: nadir edge case — recovery flow'a düş.
            break;
        }

        if (VaultService.VaultVarMi() && RunRecoveryUnlock())
            return ResetLaunchAndVaultPassword();

        MessageBox.Show(
            "Hatalı şifre girişi limiti aşıldı. Program kapatılıyor.",
            "Raporlama Portalı",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
        Environment.Exit(1);
        return string.Empty;
    }

    private static bool RunRecoveryUnlock()
    {
        using var form = new VaultRecoveryUnlockForm();
        while (form.ShowDialog() == DialogResult.OK)
        {
            if (VaultService.TryUnlockWithRecovery(form.RecoveryKey))
                return true;
            form.HataGoster("Kurtarma anahtarı geçersiz. Tekrar deneyin.");
        }
        return false;
    }

    // Recovery sonrası: yeni şifre alır, hem vault'ı hem launch_auth'u günceller.
    private static string ResetLaunchAndVaultPassword()
    {
        using var form = new VaultNewPasswordForm();
        if (form.ShowDialog() != DialogResult.OK)
            Environment.Exit(1);

        VaultService.RewrapWithNewPassword(form.Password);

        var newSalt = RandomNumberGenerator.GetBytes(SaltBytes);
        var newHash = Hash(form.Password, newSalt);
        WriteStoredHash(Convert.ToBase64String(newHash), Convert.ToBase64String(newSalt));

        return form.Password;
    }

    // ---------- launch_auth.json I/O ----------

    private static byte[] Hash(string password, byte[] salt) =>
        Rfc2898DeriveBytes.Pbkdf2(
            Encoding.UTF8.GetBytes(password), salt, Iterations, HashAlgorithmName.SHA256, HashBytes);

    private static (string? hash, string? salt) ReadStoredHash()
    {
        var path = AppDataPaths.LaunchAuthJson;
        if (!File.Exists(path)) return (null, null);
        try
        {
            var node = JsonNode.Parse(File.ReadAllText(path));
            var h = node?["PasswordHash"]?.GetValue<string>();
            var s = node?["Salt"]?.GetValue<string>();
            return (h, s);
        }
        catch
        {
            return (null, null);
        }
    }

    private static void WriteStoredHash(string hash, string salt)
    {
        var obj = new JsonObject
        {
            ["PasswordHash"] = hash,
            ["Salt"]         = salt,
        };
        File.WriteAllText(AppDataPaths.LaunchAuthJson,
            obj.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
    }

    // ---------- mevcut prompt formları (DEĞİŞMEDİ) ----------

    private sealed class PromptForm : Form
    {
        private readonly TextBox _pwd;
        public string Password => _pwd.Text;

        public PromptForm(int attempt, int max)
        {
            Text             = "Raporlama Portalı — Başlatma Şifresi";
            FormBorderStyle  = FormBorderStyle.FixedDialog;
            StartPosition    = FormStartPosition.CenterScreen;
            MaximizeBox      = false;
            MinimizeBox      = false;
            ShowInTaskbar    = true;
            TopMost          = true;
            ClientSize       = new Size(360, 150);

            var lbl = new Label
            {
                Text = attempt == 1
                    ? "Başlatma şifresini giriniz:"
                    : $"Hatalı şifre. Tekrar deneyiniz ({attempt}/{max}):",
                Location = new Point(15, 15),
                AutoSize = true,
            };
            _pwd = new TextBox
            {
                UseSystemPasswordChar = true,
                Location              = new Point(15, 45),
                Size                  = new Size(330, 25),
            };
            var ok = new Button
            {
                Text         = "Giriş",
                DialogResult = DialogResult.OK,
                Location     = new Point(175, 90),
                Size         = new Size(80, 30),
            };
            var cancel = new Button
            {
                Text         = "İptal",
                DialogResult = DialogResult.Cancel,
                Location     = new Point(265, 90),
                Size         = new Size(80, 30),
            };

            Controls.AddRange(new Control[] { lbl, _pwd, ok, cancel });
            AcceptButton  = ok;
            CancelButton  = cancel;
            ActiveControl = _pwd;
        }
    }

    private sealed class SetupForm : Form
    {
        private readonly TextBox _pwd;
        private readonly TextBox _pwd2;
        private readonly Label   _err;
        public string Password => _pwd.Text;

        public SetupForm()
        {
            Text             = "Raporlama Portalı — İlk Kurulum";
            FormBorderStyle  = FormBorderStyle.FixedDialog;
            StartPosition    = FormStartPosition.CenterScreen;
            MaximizeBox      = false;
            MinimizeBox      = false;
            ShowInTaskbar    = true;
            TopMost          = true;
            ClientSize       = new Size(420, 280);

            var info = new Label
            {
                Text =
                    "İlk kullanım. Uygulamayı başlatmak için bir şifre belirleyin.\n" +
                    "Bir sonraki adımda gösterilecek 'kurtarma anahtarı'nı GÜVENLİ bir yere kaydedin — " +
                    "şifreyi unutursanız veya bilgisayar değişirse tek kurtarma yolu odur.",
                Location = new Point(15, 10),
                Size     = new Size(390, 90),
            };
            var lbl1 = new Label
            {
                Text     = "Şifre:",
                Location = new Point(15, 110),
                AutoSize = true,
            };
            _pwd = new TextBox
            {
                UseSystemPasswordChar = true,
                Location              = new Point(120, 107),
                Size                  = new Size(285, 25),
            };
            var lbl2 = new Label
            {
                Text     = "Şifre (tekrar):",
                Location = new Point(15, 145),
                AutoSize = true,
            };
            _pwd2 = new TextBox
            {
                UseSystemPasswordChar = true,
                Location              = new Point(120, 142),
                Size                  = new Size(285, 25),
            };
            _err = new Label
            {
                ForeColor = Color.Red,
                Location  = new Point(15, 180),
                Size      = new Size(390, 20),
                Text      = "",
            };
            var ok = new Button
            {
                Text     = "Kaydet",
                Location = new Point(235, 220),
                Size     = new Size(80, 30),
            };
            var cancel = new Button
            {
                Text         = "İptal",
                DialogResult = DialogResult.Cancel,
                Location     = new Point(325, 220),
                Size         = new Size(80, 30),
            };

            ok.Click += (_, _) =>
            {
                if (string.IsNullOrWhiteSpace(_pwd.Text) || _pwd.Text.Length < 4)
                {
                    _err.Text = "Şifre en az 4 karakter olmalı.";
                    return;
                }
                if (_pwd.Text != _pwd2.Text)
                {
                    _err.Text = "Şifreler eşleşmiyor.";
                    return;
                }
                DialogResult = DialogResult.OK;
                Close();
            };

            Controls.AddRange(new Control[] { info, lbl1, _pwd, lbl2, _pwd2, _err, ok, cancel });
            AcceptButton  = ok;
            CancelButton  = cancel;
            ActiveControl = _pwd;
        }
    }
}
