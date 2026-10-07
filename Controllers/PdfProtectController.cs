using Microsoft.AspNetCore.Mvc;
using VipaPDFProtect.Models;
using VipaPDFProtect.Services;

namespace VipaPDFProtect.Controllers;

[ApiController]
[Route("api")]
public sealed class PdfProtectController : ControllerBase
{
    private readonly GraphOneDriveService _oneDrive;
    private readonly PasswordCryptoService _crypto;
    private readonly PdfProtectService _pdfProtect;
    private readonly ILogger<PdfProtectController> _logger;
    private readonly IConfiguration _configuration;

    public PdfProtectController(GraphOneDriveService oneDrive, PasswordCryptoService crypto, PdfProtectService pdfProtect, ILogger<PdfProtectController> logger, IConfiguration configuration)
    { _oneDrive = oneDrive; _crypto = crypto; _pdfProtect = pdfProtect; _logger = logger; _configuration = configuration; }

    [HttpPost("protect-pdf")]
    public async Task<IActionResult> ProtectPdf([FromBody] ProtectPdfRequest input, CancellationToken cancellationToken)
    {
        var response = new ProtectPdfResponse();
        try
        {
            var configuredApiKey = _configuration["ApiKey"];
            if (!string.IsNullOrWhiteSpace(configuredApiKey))
            {
                var supplied = Request.Headers["x-api-key"].FirstOrDefault();
                if (!string.Equals(configuredApiKey, supplied, StringComparison.Ordinal))
                { response.Result.ErrorMessage = "Unauthorized."; return Unauthorized(response); }
            }

            if (string.IsNullOrWhiteSpace(input.OneDriveUrl) || string.IsNullOrWhiteSpace(input.SourceUrl) || string.IsNullOrWhiteSpace(input.DestinationUrl))
            { response.Result.ErrorMessage = "OneDriveUrl, SourceUrl and DestinationUrl are required."; return BadRequest(response); }

            if (!Uri.TryCreate(input.OneDriveUrl, UriKind.Absolute, out var rootUri) ||
                !Uri.TryCreate(input.SourceUrl, UriKind.Absolute, out var sourceUri) ||
                !Uri.TryCreate(input.DestinationUrl, UriKind.Absolute, out var destUri))
            { response.Result.ErrorMessage = "OneDriveUrl, SourceUrl and DestinationUrl must be valid absolute URLs."; return BadRequest(response); }

            if (rootUri.Scheme != Uri.UriSchemeHttps || sourceUri.Scheme != Uri.UriSchemeHttps || destUri.Scheme != Uri.UriSchemeHttps)
                throw new ArgumentException("Only HTTPS URLs are allowed.");
            if (!string.Equals(rootUri.Host, sourceUri.Host, StringComparison.OrdinalIgnoreCase) || !string.Equals(rootUri.Host, destUri.Host, StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("OneDriveUrl, SourceUrl and DestinationUrl must use the same host.");

            var sitePath = Uri.UnescapeDataString(rootUri.AbsolutePath).TrimEnd('/');
            var sourcePath = ToDriveRelative(sitePath, sourceUri.AbsolutePath);
            var destinationPath = ToDriveRelative(sitePath, destUri.AbsolutePath);
            if (destinationPath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                destinationPath = destinationPath.Contains('/') ? destinationPath[..destinationPath.LastIndexOf('/')] : string.Empty;

            var password = _crypto.ResolvePassword(input.PlainPassword, input.EncryptedPassword);

            // Resolve the drive using the source parent folder, exactly like PDFMerge.
            var sourceParent = sourcePath.Contains('/') ? sourcePath[..sourcePath.LastIndexOf('/')] : string.Empty;
            var location = await _oneDrive.ResolveLocationAsync(rootUri.Host, sitePath, sourceParent, cancellationToken);
            var sourceItem = await _oneDrive.GetItemByPathAsync(location.DriveId, sourcePath, cancellationToken);
            if (sourceItem.File is null) throw new InvalidOperationException("SourceUrl does not point to a file.");
            if (!sourceItem.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("SourceUrl must point to a PDF file.");

            await using var sourceStream = await _oneDrive.DownloadFileAsync(location.DriveId, sourceItem.Id, cancellationToken);
            await using var protectedStream = _pdfProtect.Protect(sourceStream, password);

            var protectedFolderPath = string.IsNullOrWhiteSpace(destinationPath) ? "Protected" : $"{destinationPath}/Protected";
            var protectedFolder = await _oneDrive.EnsureFolderPathAsync(location.DriveId, protectedFolderPath, cancellationToken);
            var baseName = Path.GetFileNameWithoutExtension(sourceItem.Name);
            var outputName = $"{baseName}_withpassword.pdf";
            var uploaded = await _oneDrive.UploadFileAsync(location.DriveId, protectedFolder.Id, outputName, protectedStream, cancellationToken);

            response.Result.IsSuccessful = true;
            response.Result.ProtectedPDFFilePath = uploaded.WebUrl ?? BuildWebUrl(input.OneDriveUrl, protectedFolderPath, outputName);
            return Ok(response);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "PDF protection failed.");
            response.Result.IsSuccessful = false;
            response.Result.ErrorMessage = ex.Message;
            response.Result.InnerErrorMessage = ex.InnerException?.Message ?? string.Empty;
            return Ok(response);
        }
    }

    private static string ToDriveRelative(string sitePath, string absolutePath)
    {
        var full = Uri.UnescapeDataString(absolutePath).TrimEnd('/');
        if (!full.Equals(sitePath, StringComparison.OrdinalIgnoreCase) && !full.StartsWith(sitePath + "/", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("URL must be inside OneDriveUrl.");
        var relative = full.Length == sitePath.Length ? string.Empty : full[(sitePath.Length + 1)..].Trim('/');
        if (relative.Equals("Documents", StringComparison.OrdinalIgnoreCase) || relative.Equals("Shared Documents", StringComparison.OrdinalIgnoreCase)) return string.Empty;
        if (relative.StartsWith("Documents/", StringComparison.OrdinalIgnoreCase)) return relative["Documents/".Length..];
        if (relative.StartsWith("Shared Documents/", StringComparison.OrdinalIgnoreCase)) return relative["Shared Documents/".Length..];
        return relative;
    }

    private static string BuildWebUrl(string oneDriveUrl, string folderPath, string fileName)
    {
        var baseUrl = oneDriveUrl.TrimEnd('/');
        var escaped = string.Join("/", (folderPath + "/" + fileName).Split('/', StringSplitOptions.RemoveEmptyEntries).Select(Uri.EscapeDataString));
        return $"{baseUrl}/Documents/{escaped}";
    }
}
