using System.Security.Cryptography;
using PdfSharp.Pdf.IO;
namespace VipaPDFProtect.Services;
public sealed class PdfProtectService
{
    public MemoryStream Protect(Stream input, string password)
    {
        if (string.IsNullOrEmpty(password)) throw new ArgumentException("PDF password cannot be empty.");
        input.Position = 0;
        using var document = PdfReader.Open(input, PdfDocumentOpenMode.Modify);
        document.SecuritySettings.UserPassword = password;
        document.SecuritySettings.OwnerPassword = Convert.ToHexString(RandomNumberGenerator.GetBytes(32));
        document.SecurityHandler.SetEncryptionToV4UsingAES();
        var output = new MemoryStream();
        document.Save(output, false);
        output.Position = 0;
        return output;
    }
}
