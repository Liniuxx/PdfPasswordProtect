using System.Security.Cryptography;
using System.Text;
namespace VipaPDFProtect.Services;
public sealed class PasswordCryptoService
{
    private readonly IConfiguration _configuration;
    public PasswordCryptoService(IConfiguration configuration) => _configuration = configuration;
    public string ResolvePassword(string? plainPassword, string? encryptedPassword)
    {
        if (!string.IsNullOrWhiteSpace(plainPassword)) return plainPassword;
        if (string.IsNullOrWhiteSpace(encryptedPassword)) throw new ArgumentException("PlainPassword or EncryptedPassword is required.");
        var key = _configuration["PDF_PASSWORD_ENCRYPTION_KEY"];
        if (string.IsNullOrEmpty(key)) throw new InvalidOperationException("PDF_PASSWORD_ENCRYPTION_KEY is not configured.");
        var keyBytes = Encoding.UTF8.GetBytes(key);
        if (keyBytes.Length != 16) throw new InvalidOperationException("PDF_PASSWORD_ENCRYPTION_KEY must be exactly 16 UTF-8 bytes for AES-128.");
        byte[] cipher;
        try { cipher = Convert.FromBase64String(encryptedPassword); }
        catch (FormatException ex) { throw new ArgumentException("EncryptedPassword is not valid Base64.", ex); }
        using var aes = Aes.Create();
        aes.Key = keyBytes; aes.Mode = CipherMode.ECB; aes.Padding = PaddingMode.PKCS7;
        using var decryptor = aes.CreateDecryptor();
        var clear = decryptor.TransformFinalBlock(cipher, 0, cipher.Length);
        return Encoding.UTF8.GetString(clear);
    }
}
