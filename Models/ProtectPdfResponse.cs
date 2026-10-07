using System.Text.Json.Serialization;
namespace VipaPDFProtect.Models;
public sealed class ProtectPdfResponse { [JsonPropertyName("result")] public ProtectPdfResult Result { get; set; } = new(); }
public sealed class ProtectPdfResult
{
    [JsonPropertyName("Message")] public string Message { get; set; } = string.Empty;
    [JsonPropertyName("ErrorMessage")] public string ErrorMessage { get; set; } = string.Empty;
    [JsonPropertyName("InnerErrorMessage")] public string InnerErrorMessage { get; set; } = string.Empty;
    [JsonPropertyName("IsSuccessful")] public bool IsSuccessful { get; set; }
    [JsonPropertyName("ProtectedPDFFilePath")] public string ProtectedPDFFilePath { get; set; } = string.Empty;
}
