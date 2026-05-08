using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Win32;

namespace RaporlamaPortali.Services;

// Envelope encryption.
// Random 32-byte MDK (Master Data Key) üretilir; bu MDK iki ayrı şekilde sarmalanır:
//   WrapPassword  = AES-GCM( PBKDF2( password ⊕ machineId, PasswordSalt, 200k ), MDK )
//   WrapRecovery  = AES-GCM( PBKDF2( recoveryKey,           RecoverySalt, 200k ), MDK )
// vault.json sadece bu iki sarmal + saltleri tutar — şifre veya MDK düz haliyle saklanmaz.
// secrets.enc dosyası da AES-GCM(MDK, secrets_json) ile şifrelenir.
public static class VaultService
{
    private const int Iterations    = 200_000;
    private const int SaltBytes     = 16;
    private const int NonceBytes    = 12;
    private const int TagBytes      = 16;
    private const int MdkBytes      = 32;
    private const int RecoveryBytes = 20; // 160 bit → Base32 ile 32 karakter

    private static byte[]? _cachedMdk;

    public static bool VaultVarMi() => File.Exists(AppDataPaths.VaultJson);
    public static bool Acik => _cachedMdk != null;

    public static byte[] GetMdkOrThrow() =>
        _cachedMdk ?? throw new InvalidOperationException("Vault unlocked değil.");

    // İlk kurulum. MDK üretir, password+machine ile + recoveryKey ile iki ayrı sarmalı diske yazar.
    // Geri dönen string: kullanıcıya gösterilecek 32 karakterlik (8x4 grup) recovery key.
    public static string CreateNewVault(string password)
    {
        var mdk          = RandomNumberGenerator.GetBytes(MdkBytes);
        var passwordSalt = RandomNumberGenerator.GetBytes(SaltBytes);
        var recoverySalt = RandomNumberGenerator.GetBytes(SaltBytes);

        var passwordKey = DerivePasswordKey(password, passwordSalt);
        var (passNonce, passCipher, passTag) = AesGcmEncrypt(passwordKey, mdk);

        var recoveryRaw      = RandomNumberGenerator.GetBytes(RecoveryBytes);
        var recoveryDisplay  = FormatRecoveryKey(recoveryRaw);
        var recoveryKey      = DeriveRecoveryKey(recoveryRaw, recoverySalt);
        var (recNonce, recCipher, recTag) = AesGcmEncrypt(recoveryKey, mdk);

        // Recovery key'in doğruluğunu hızlıca kontrol etmek için (kullanıcı yanlış girince
        // 200k PBKDF2 iterasyonuna girmeden önce aday eşleşiyor mu görelim).
        var recoveryVerify = SHA256.HashData(recoveryRaw);

        WriteVault(new VaultData(
            CreatedUtc:    DateTime.UtcNow.ToString("o"),
            PasswordSalt:  Convert.ToBase64String(passwordSalt),
            RecoverySalt:  Convert.ToBase64String(recoverySalt),
            WrapPassword:  new VaultWrap(B64(passNonce), B64(passCipher), B64(passTag)),
            WrapRecovery:  new VaultWrap(B64(recNonce),  B64(recCipher),  B64(recTag)),
            RecoveryVerify: B64(recoveryVerify)));

        _cachedMdk = mdk;
        return recoveryDisplay;
    }

    public static bool TryUnlockWithPassword(string password)
    {
        try
        {
            var v     = ReadVault();
            var key   = DerivePasswordKey(password, FromB64(v.PasswordSalt));
            _cachedMdk = AesGcmDecrypt(
                key,
                FromB64(v.WrapPassword.Nonce),
                FromB64(v.WrapPassword.Cipher),
                FromB64(v.WrapPassword.Tag));
            return true;
        }
        catch
        {
            _cachedMdk = null;
            return false;
        }
    }

    public static bool TryUnlockWithRecovery(string recoveryDisplay)
    {
        try
        {
            var raw = ParseRecoveryKey(recoveryDisplay);
            if (raw == null) return false;

            var v        = ReadVault();
            var verify   = FromB64(v.RecoveryVerify);
            var calc     = SHA256.HashData(raw);
            if (!CryptographicOperations.FixedTimeEquals(verify, calc)) return false;

            var key   = DeriveRecoveryKey(raw, FromB64(v.RecoverySalt));
            _cachedMdk = AesGcmDecrypt(
                key,
                FromB64(v.WrapRecovery.Nonce),
                FromB64(v.WrapRecovery.Cipher),
                FromB64(v.WrapRecovery.Tag));
            return true;
        }
        catch
        {
            _cachedMdk = null;
            return false;
        }
    }

    // Recovery sonrası yeni şifre belirleme — MDK aynı kalır, password wrap baştan üretilir.
    // Recovery wrap'ı bozmaz, kullanıcı eski recovery anahtarını kullanmaya devam edebilir.
    public static void RewrapWithNewPassword(string newPassword)
    {
        if (_cachedMdk == null) throw new InvalidOperationException("Önce unlock olmalı.");

        var v               = ReadVault();
        var newPasswordSalt = RandomNumberGenerator.GetBytes(SaltBytes);
        var newKey          = DerivePasswordKey(newPassword, newPasswordSalt);
        var (n, c, t)       = AesGcmEncrypt(newKey, _cachedMdk);

        WriteVault(v with
        {
            PasswordSalt = B64(newPasswordSalt),
            WrapPassword = new VaultWrap(B64(n), B64(c), B64(t))
        });
    }

    // SecretsService kullansın diye dışa açık AES-GCM yardımcıları
    public static (byte[] Nonce, byte[] Cipher, byte[] Tag) EncryptWithMdk(byte[] plaintext)
    {
        if (_cachedMdk == null) throw new InvalidOperationException("Vault locked.");
        return AesGcmEncrypt(_cachedMdk, plaintext);
    }

    public static byte[] DecryptWithMdk(byte[] nonce, byte[] cipher, byte[] tag)
    {
        if (_cachedMdk == null) throw new InvalidOperationException("Vault locked.");
        return AesGcmDecrypt(_cachedMdk, nonce, cipher, tag);
    }

    // ---------- KDF ----------

    private static byte[] DerivePasswordKey(string password, byte[] salt)
    {
        // password ⊕ machineId — makineye bağlı. Aynı şifre, başka makinede MDK çıkmaz.
        var pwBytes      = Encoding.UTF8.GetBytes(password);
        var machineId    = GetMachineId();
        var combined     = new byte[pwBytes.Length + machineId.Length];
        Buffer.BlockCopy(pwBytes,   0, combined, 0,                pwBytes.Length);
        Buffer.BlockCopy(machineId, 0, combined, pwBytes.Length,   machineId.Length);
        return Rfc2898DeriveBytes.Pbkdf2(combined, salt, Iterations, HashAlgorithmName.SHA256, 32);
    }

    private static byte[] DeriveRecoveryKey(byte[] rawRecovery, byte[] salt)
    {
        // Recovery key makineden bağımsız — donanım değişiminde kullanılabilsin.
        return Rfc2898DeriveBytes.Pbkdf2(rawRecovery, salt, Iterations, HashAlgorithmName.SHA256, 32);
    }

    private static byte[] GetMachineId()
    {
        try
        {
            using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Cryptography");
            var guid = key?.GetValue("MachineGuid") as string;
            if (!string.IsNullOrWhiteSpace(guid))
                return SHA256.HashData(Encoding.UTF8.GetBytes(guid));
        }
        catch { }
        var fallback = Environment.MachineName + "/" + Environment.OSVersion;
        return SHA256.HashData(Encoding.UTF8.GetBytes(fallback));
    }

    // ---------- AES-GCM ----------

    private static (byte[] Nonce, byte[] Cipher, byte[] Tag) AesGcmEncrypt(byte[] key, byte[] plaintext)
    {
        var nonce  = RandomNumberGenerator.GetBytes(NonceBytes);
        var cipher = new byte[plaintext.Length];
        var tag    = new byte[TagBytes];
        using var aes = new AesGcm(key, TagBytes);
        aes.Encrypt(nonce, plaintext, cipher, tag);
        return (nonce, cipher, tag);
    }

    private static byte[] AesGcmDecrypt(byte[] key, byte[] nonce, byte[] cipher, byte[] tag)
    {
        var plain = new byte[cipher.Length];
        using var aes = new AesGcm(key, TagBytes);
        aes.Decrypt(nonce, cipher, tag, plain);
        return plain;
    }

    // ---------- Recovery key formatlama ----------

    // Crockford-benzeri Base32 (1, 0, I, O harfleri yok — okunurluk için)
    private const string Base32Alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";

    private static string FormatRecoveryKey(byte[] raw)
    {
        // 20 byte = 160 bit → 32 base32 karakter → 8 grup x 4 karakter
        var sb       = new StringBuilder(40);
        int buffer   = 0, bits = 0;
        foreach (var b in raw)
        {
            buffer = (buffer << 8) | b;
            bits  += 8;
            while (bits >= 5)
            {
                bits -= 5;
                sb.Append(Base32Alphabet[(buffer >> bits) & 0x1F]);
            }
        }
        if (bits > 0) sb.Append(Base32Alphabet[(buffer << (5 - bits)) & 0x1F]);

        var s        = sb.ToString();
        var groups   = new List<string>(s.Length / 4 + 1);
        for (int i = 0; i < s.Length; i += 4)
            groups.Add(s.Substring(i, Math.Min(4, s.Length - i)));
        return string.Join("-", groups);
    }

    private static byte[]? ParseRecoveryKey(string display)
    {
        if (string.IsNullOrWhiteSpace(display)) return null;
        var clean = new StringBuilder();
        foreach (var ch in display)
        {
            var u = char.ToUpperInvariant(ch);
            if (Base32Alphabet.Contains(u)) clean.Append(u);
        }
        if (clean.Length != 32) return null;

        var raw    = new byte[20];
        int buffer = 0, bits = 0, idx = 0;
        foreach (var c in clean.ToString())
        {
            int v = Base32Alphabet.IndexOf(c);
            if (v < 0) return null;
            buffer = (buffer << 5) | v;
            bits  += 5;
            if (bits >= 8)
            {
                bits -= 8;
                raw[idx++] = (byte)((buffer >> bits) & 0xFF);
            }
        }
        return idx == 20 ? raw : null;
    }

    // ---------- vault.json I/O ----------

    private record VaultData(
        string CreatedUtc,
        string PasswordSalt,
        string RecoverySalt,
        VaultWrap WrapPassword,
        VaultWrap WrapRecovery,
        string RecoveryVerify);

    private record VaultWrap(string Nonce, string Cipher, string Tag);

    private static VaultData ReadVault()
    {
        var j = JsonNode.Parse(File.ReadAllText(AppDataPaths.VaultJson))!;
        return new VaultData(
            j["CreatedUtc"]!.GetValue<string>(),
            j["PasswordSalt"]!.GetValue<string>(),
            j["RecoverySalt"]!.GetValue<string>(),
            ReadWrap(j["WrapPassword"]!),
            ReadWrap(j["WrapRecovery"]!),
            j["RecoveryVerify"]!.GetValue<string>()
        );
    }

    private static VaultWrap ReadWrap(JsonNode n) =>
        new(n["Nonce"]!.GetValue<string>(),
            n["Cipher"]!.GetValue<string>(),
            n["Tag"]!.GetValue<string>());

    private static void WriteVault(VaultData v)
    {
        var obj = new JsonObject
        {
            ["Version"]        = 1,
            ["CreatedUtc"]     = v.CreatedUtc,
            ["PasswordSalt"]   = v.PasswordSalt,
            ["RecoverySalt"]   = v.RecoverySalt,
            ["WrapPassword"]   = WrapToJson(v.WrapPassword),
            ["WrapRecovery"]   = WrapToJson(v.WrapRecovery),
            ["RecoveryVerify"] = v.RecoveryVerify,
        };
        File.WriteAllText(AppDataPaths.VaultJson,
            obj.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
    }

    private static JsonObject WrapToJson(VaultWrap w) => new()
    {
        ["Nonce"]  = w.Nonce,
        ["Cipher"] = w.Cipher,
        ["Tag"]    = w.Tag,
    };

    private static string B64(byte[] b)   => Convert.ToBase64String(b);
    private static byte[] FromB64(string s) => Convert.FromBase64String(s);
}
