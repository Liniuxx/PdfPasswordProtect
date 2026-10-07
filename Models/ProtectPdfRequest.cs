namespace VipaPDFProtect.Models;
public sealed class ProtectPdfRequest
{
    public string OneDriveUrl { get; set; } = string.Empty;
    public string SourceUrl { get; set; } = string.Empty;
    public string DestinationUrl { get; set; } = string.Empty;
    public string? PlainPassword { get; set; }
    public string? EncryptedPassword { get; set; }
}
