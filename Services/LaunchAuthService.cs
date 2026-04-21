using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Windows.Forms;

namespace RaporlamaPortali.Services;

// Uygulama exe'si başlarken şifre sorar. Doğru şifre girilmezse process kapanır.
// İlk çalıştırmada (hash boşsa) şifre belirleme dialogu çıkar.
// Hash: PBKDF2-SHA256, 100 000 iterasyon, per-install rastgele salt. appsettings.json'da saklanır.
public static class LaunchAuthService
{
    private const int MaxAttempts = 3;
    private const int Iterations = 100_000;
    private const int HashBytes = 32;
    private const int SaltBytes = 16;

    public static void Require()
    {
        ApplicationConfiguration.Initialize();

        Directory.CreateDirectory(AppDataPaths.DataRoot);
        var (hash, salt) = ReadStoredHash();

        if (string.IsNullOrEmpty(hash) || string.IsNullOrEmpty(salt))
        {
            RunFirstTimeSetup();
        }
        else
        {
            RunPasswordPrompt(hash!, salt!);
        }
    }

    private static void RunFirstTimeSetup()
    {
        using var form = new SetupForm();
        var result = form.ShowDialog();
        if (result != DialogResult.OK)
        {
            Environment.Exit(1);
        }

        var newSalt = RandomNumberGenerator.GetBytes(SaltBytes);
        var newHash = Hash(form.Password, newSalt);
        WriteStoredHash(Convert.ToBase64String(newHash), Convert.ToBase64String(newSalt));
    }

    private static void RunPasswordPrompt(string storedHashB64, string storedSaltB64)
    {
        var storedHash = Convert.FromBase64String(storedHashB64);
        var salt = Convert.FromBase64String(storedSaltB64);

        for (int i = 0; i < MaxAttempts; i++)
        {
            using var form = new PromptForm(i + 1, MaxAttempts);
            var result = form.ShowDialog();
            if (result != DialogResult.OK)
            {
                Environment.Exit(1);
            }

            var candidate = Hash(form.Password, salt);
            if (CryptographicOperations.FixedTimeEquals(candidate, storedHash))
            {
                return;
            }
        }

        MessageBox.Show(
            "Hatalı şifre girişi limiti aşıldı. Program kapatılıyor.",
            "Raporlama Portalı",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error);
        Environment.Exit(1);
    }

    private static byte[] Hash(string password, byte[] salt)
    {
        using var pbkdf2 = new Rfc2898DeriveBytes(password, salt, Iterations, HashAlgorithmName.SHA256);
        return pbkdf2.GetBytes(HashBytes);
    }

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
            ["Salt"] = salt,
        };
        File.WriteAllText(AppDataPaths.LaunchAuthJson,
            obj.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
    }

    private sealed class PromptForm : Form
    {
        private readonly TextBox _pwd;
        public string Password => _pwd.Text;

        public PromptForm(int attempt, int max)
        {
            Text = "Raporlama Portalı — Başlatma Şifresi";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = true;
            TopMost = true;
            ClientSize = new Size(360, 150);

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
                Location = new Point(15, 45),
                Size = new Size(330, 25),
            };
            var ok = new Button
            {
                Text = "Giriş",
                DialogResult = DialogResult.OK,
                Location = new Point(175, 90),
                Size = new Size(80, 30),
            };
            var cancel = new Button
            {
                Text = "İptal",
                DialogResult = DialogResult.Cancel,
                Location = new Point(265, 90),
                Size = new Size(80, 30),
            };

            Controls.AddRange(new Control[] { lbl, _pwd, ok, cancel });
            AcceptButton = ok;
            CancelButton = cancel;
            ActiveControl = _pwd;
        }
    }

    private sealed class SetupForm : Form
    {
        private readonly TextBox _pwd;
        private readonly TextBox _pwd2;
        private readonly Label _err;
        public string Password => _pwd.Text;

        public SetupForm()
        {
            Text = "Raporlama Portalı — İlk Kurulum";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = true;
            TopMost = true;
            ClientSize = new Size(420, 240);

            var info = new Label
            {
                Text = "İlk kullanım. Uygulamayı başlatmak için kullanılacak bir şifre belirleyin. Bu şifreyi unutmayın — kurtarmak için appsettings.json düzenlenmelidir.",
                Location = new Point(15, 10),
                Size = new Size(390, 55),
            };
            var lbl1 = new Label
            {
                Text = "Şifre:",
                Location = new Point(15, 75),
                AutoSize = true,
            };
            _pwd = new TextBox
            {
                UseSystemPasswordChar = true,
                Location = new Point(120, 72),
                Size = new Size(285, 25),
            };
            var lbl2 = new Label
            {
                Text = "Şifre (tekrar):",
                Location = new Point(15, 110),
                AutoSize = true,
            };
            _pwd2 = new TextBox
            {
                UseSystemPasswordChar = true,
                Location = new Point(120, 107),
                Size = new Size(285, 25),
            };
            _err = new Label
            {
                ForeColor = Color.Red,
                Location = new Point(15, 145),
                Size = new Size(390, 20),
                Text = "",
            };
            var ok = new Button
            {
                Text = "Kaydet",
                Location = new Point(235, 180),
                Size = new Size(80, 30),
            };
            var cancel = new Button
            {
                Text = "İptal",
                DialogResult = DialogResult.Cancel,
                Location = new Point(325, 180),
                Size = new Size(80, 30),
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
            AcceptButton = ok;
            CancelButton = cancel;
            ActiveControl = _pwd;
        }
    }
}

