using System.Net;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;
using PdfSharp.Pdf.IO;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddHttpClient("graph", client =>
{
    client.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/");
    client.Timeout = TimeSpan.FromMinutes(10);
});

builder.Services.AddSingleton<TokenCredential>(_ => new DefaultAzureCredential());

var app = builder.Build();

app.MapGet("/", () => Results.Ok(new
{
    service = "VipaPDFProtect",
    status = "OK"
}));

app.MapGet("/health", () => Results.Ok(new { status = "OK" }));

app.MapPost("/api/pdf/protect", async (
    ProtectPdfRequest request,
    IHttpClientFactory httpClientFactory,
    TokenCredential credential,
    CancellationToken cancellationToken) =>
{
    try
    {
        ValidateRequest(request);

        var password = ResolvePassword(request);

        var source = SharePointLocation.Parse(request.OneDriveUrl!, request.SourceUrl!);
        var destination = SharePointLocation.Parse(request.OneDriveUrl!, request.DestinationUrl!);

        if (!string.Equals(source.HostName, destination.HostName, StringComparison.OrdinalIgnoreCase) ||
            !string.Equals(source.SitePath, destination.SitePath, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "SourceUrl and DestinationUrl must belong to the SharePoint site specified in OneDriveUrl.");
        }

        var graph = httpClientFactory.CreateClient("graph");

        var token = await credential.GetTokenAsync(
            new TokenRequestContext(["https://graph.microsoft.com/.default"]),
            cancellationToken);

        graph.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", token.Token);

        var siteId = await GraphApi.GetSiteIdAsync(
            graph, source.HostName, source.SitePath, cancellationToken);

        var driveId = await GraphApi.GetDefaultDriveIdAsync(
            graph, siteId, cancellationToken);

        var pdfBytes = await GraphApi.DownloadFileAsync(
            graph, driveId, source.RelativePath, cancellationToken);

        var protectedPdf = ProtectPdf(pdfBytes, password);

        var destinationFolder = destination.RelativePath.Trim('/');
        var mergedFolder = CombinePath(destinationFolder, "Merged");

        await GraphApi.EnsureFolderPathAsync(
            graph, driveId, mergedFolder, cancellationToken);

        var sourceFileName = Path.GetFileName(Uri.UnescapeDataString(source.RelativePath));
        var outputFileName =
            $"{Path.GetFileNameWithoutExtension(sourceFileName)}_withpassword.pdf";

        var outputRelativePath = CombinePath(mergedFolder, outputFileName);

        await GraphApi.UploadFileAsync(
            graph, driveId, outputRelativePath, protectedPdf, cancellationToken);

        var protectedUrl = BuildSharePointWebUrl(
            request.OneDriveUrl!, outputRelativePath);

        return Results.Ok(new ApiEnvelope(new ApiResult(
            Message: "",
            ErrorMessage: "",
            InnerErrorMessage: "",
            IsSuccessful: true,
            ProtectedPDFFilePath: protectedUrl
        )));
    }
    catch (Exception ex)
    {
        return Results.Ok(new ApiEnvelope(new ApiResult(
            Message: "",
            ErrorMessage: ex.Message,
            InnerErrorMessage: ex.InnerException?.Message ?? "",
            IsSuccessful: false,
            ProtectedPDFFilePath: ""
        )));
    }
});

app.Run();

static void ValidateRequest(ProtectPdfRequest request)
{
    if (string.IsNullOrWhiteSpace(request.OneDriveUrl))
        throw new ArgumentException("OneDriveUrl is required.");

    if (string.IsNullOrWhiteSpace(request.SourceUrl))
        throw new ArgumentException("SourceUrl is required.");

    if (string.IsNullOrWhiteSpace(request.DestinationUrl))
        throw new ArgumentException("DestinationUrl is required.");

    if (string.IsNullOrWhiteSpace(request.PlainPassword) &&
        string.IsNullOrWhiteSpace(request.EncryptedPassword))
        throw new ArgumentException(
            "PlainPassword or EncryptedPassword is required.");

    if (!string.IsNullOrWhiteSpace(request.PlainPassword) &&
        !string.IsNullOrWhiteSpace(request.EncryptedPassword))
        throw new ArgumentException(
            "Send only one password field: PlainPassword or EncryptedPassword.");
}

static string ResolvePassword(ProtectPdfRequest request)
{
    if (!string.IsNullOrWhiteSpace(request.PlainPassword))
        return request.PlainPassword;

    var key = Environment.GetEnvironmentVariable(
        "PDF_PASSWORD_ENCRYPTION_KEY");

    if (string.IsNullOrWhiteSpace(key))
        throw new InvalidOperationException(
            "Azure setting PDF_PASSWORD_ENCRYPTION_KEY is missing.");

    var keyBytes = Encoding.UTF8.GetBytes(key);

    if (keyBytes.Length != 16)
        throw new InvalidOperationException(
            "PDF_PASSWORD_ENCRYPTION_KEY must be exactly 16 bytes for AES-128.");

    byte[] encryptedBytes;

    try
    {
        encryptedBytes = Convert.FromBase64String(request.EncryptedPassword!);
    }
    catch (FormatException ex)
    {
        throw new InvalidOperationException(
            "EncryptedPassword is not valid Base64.", ex);
    }

    using var aes = Aes.Create();
    aes.Key = keyBytes;
    aes.Mode = CipherMode.ECB;
    aes.Padding = PaddingMode.PKCS7;

    try
    {
        using var decryptor = aes.CreateDecryptor();
        var decryptedBytes = decryptor.TransformFinalBlock(
            encryptedBytes, 0, encryptedBytes.Length);

        return Encoding.UTF8.GetString(decryptedBytes);
    }
    catch (CryptographicException ex)
    {
        throw new InvalidOperationException(
            "EncryptedPassword could not be decrypted. Check the AES key, ECB mode and PKCS7 padding.",
            ex);
    }
}

static byte[] ProtectPdf(byte[] inputPdf, string password)
{
    using var input = new MemoryStream(inputPdf);
    using var document = PdfReader.Open(input, PdfDocumentOpenMode.Modify);

    document.SecuritySettings.UserPassword = password;

    // Separate random owner password prevents the user password from being used
    // as the administrative/owner password.
    document.SecuritySettings.OwnerPassword =
        Convert.ToHexString(RandomNumberGenerator.GetBytes(32));

    // PDF AES-128 encryption (PDF V4).
    document.SecurityHandler.SetEncryptionToV4UsingAES();

    using var output = new MemoryStream();
    document.Save(output, false);

    return output.ToArray();
}

static string BuildSharePointWebUrl(string siteUrl, string relativePath)
{
    var baseUrl = siteUrl.TrimEnd('/');
    var escaped = string.Join("/",
        relativePath
            .Split('/', StringSplitOptions.RemoveEmptyEntries)
            .Select(Uri.EscapeDataString));

    return $"{baseUrl}/{escaped}";
}

static string CombinePath(params string[] parts) =>
    string.Join("/",
        parts
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim('/'))
            .Where(x => x.Length > 0));

public sealed record ProtectPdfRequest(
    string? OneDriveUrl,
    string? SourceUrl,
    string? DestinationUrl,
    string? PlainPassword,
    string? EncryptedPassword);

public sealed record ApiEnvelope(ApiResult result);

public sealed record ApiResult(
    string Message,
    string ErrorMessage,
    string InnerErrorMessage,
    bool IsSuccessful,
    string ProtectedPDFFilePath);

public sealed record SharePointLocation(
    string HostName,
    string SitePath,
    string RelativePath)
{
    public static SharePointLocation Parse(string siteUrl, string targetUrl)
    {
        var site = new Uri(siteUrl);
        var target = new Uri(targetUrl);

        if (!string.Equals(
                site.Host,
                target.Host,
                StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"URL host '{target.Host}' does not match OneDriveUrl host '{site.Host}'.");
        }

        var sitePath = Uri.UnescapeDataString(site.AbsolutePath).TrimEnd('/');
        var targetPath = Uri.UnescapeDataString(target.AbsolutePath);

        if (!targetPath.Equals(sitePath, StringComparison.OrdinalIgnoreCase) &&
            !targetPath.StartsWith(sitePath + "/", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException(
                $"URL '{targetUrl}' is outside SharePoint site '{siteUrl}'.");
        }

        var relative = targetPath.Length == sitePath.Length
            ? ""
            : targetPath[(sitePath.Length + 1)..];

        return new SharePointLocation(
            site.Host,
            sitePath,
            relative.Trim('/'));
    }
}

static class GraphApi
{
    public static async Task<string> GetSiteIdAsync(
        HttpClient graph,
        string hostName,
        string sitePath,
        CancellationToken ct)
    {
        var encodedPath = EncodeGraphPath(sitePath);
        var url = $"sites/{hostName}:/{encodedPath}";

        using var response = await graph.GetAsync(url, ct);
        var json = await ReadSuccessJsonAsync(response, "Get SharePoint site", ct);

        return GetRequiredString(json.RootElement, "id");
    }

    public static async Task<string> GetDefaultDriveIdAsync(
        HttpClient graph,
        string siteId,
        CancellationToken ct)
    {
        using var response = await graph.GetAsync(
            $"sites/{Uri.EscapeDataString(siteId)}/drive?$select=id",
            ct);

        var json = await ReadSuccessJsonAsync(
            response, "Get SharePoint default drive", ct);

        return GetRequiredString(json.RootElement, "id");
    }

    public static async Task<byte[]> DownloadFileAsync(
        HttpClient graph,
        string driveId,
        string relativePath,
        CancellationToken ct)
    {
        var path = EncodeGraphPath(relativePath);
        var url =
            $"drives/{Uri.EscapeDataString(driveId)}/root:/{path}:/content";

        using var response = await graph.GetAsync(
            url,
            HttpCompletionOption.ResponseHeadersRead,
            ct);

        if (!response.IsSuccessStatusCode)
            await ThrowGraphErrorAsync(response, "Download source PDF", ct);

        return await response.Content.ReadAsByteArrayAsync(ct);
    }

    public static async Task EnsureFolderPathAsync(
        HttpClient graph,
        string driveId,
        string folderPath,
        CancellationToken ct)
    {
        var current = "";

        foreach (var segment in folderPath
                     .Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            var next = string.IsNullOrEmpty(current)
                ? segment
                : $"{current}/{segment}";

            if (!await ItemExistsAsync(graph, driveId, next, ct))
            {
                await CreateFolderAsync(
                    graph, driveId, current, segment, ct);
            }

            current = next;
        }
    }

    public static async Task UploadFileAsync(
        HttpClient graph,
        string driveId,
        string relativePath,
        byte[] bytes,
        CancellationToken ct)
    {
        // Simple upload supports files up to 250 MB.
        var path = EncodeGraphPath(relativePath);
        var url =
            $"drives/{Uri.EscapeDataString(driveId)}/root:/{path}:/content";

        using var content = new ByteArrayContent(bytes);
        content.Headers.ContentType =
            new MediaTypeHeaderValue("application/pdf");

        using var response = await graph.PutAsync(url, content, ct);

        if (!response.IsSuccessStatusCode)
            await ThrowGraphErrorAsync(response, "Upload protected PDF", ct);
    }

    private static async Task<bool> ItemExistsAsync(
        HttpClient graph,
        string driveId,
        string relativePath,
        CancellationToken ct)
    {
        var path = EncodeGraphPath(relativePath);
        var url =
            $"drives/{Uri.EscapeDataString(driveId)}/root:/{path}?$select=id";

        using var response = await graph.GetAsync(url, ct);

        if (response.StatusCode == HttpStatusCode.NotFound)
            return false;

        if (!response.IsSuccessStatusCode)
            await ThrowGraphErrorAsync(response, "Check destination folder", ct);

        return true;
    }

    private static async Task CreateFolderAsync(
        HttpClient graph,
        string driveId,
        string parentPath,
        string folderName,
        CancellationToken ct)
    {
        string url;

        if (string.IsNullOrEmpty(parentPath))
        {
            url =
                $"drives/{Uri.EscapeDataString(driveId)}/root/children";
        }
        else
        {
            var path = EncodeGraphPath(parentPath);
            url =
                $"drives/{Uri.EscapeDataString(driveId)}/root:/{path}:/children";
        }

        var payload = JsonSerializer.Serialize(new Dictionary<string, object?>
        {
            ["name"] = folderName,
            ["folder"] = new Dictionary<string, object>(),
            ["@microsoft.graph.conflictBehavior"] = "fail"
        });

        using var content = new StringContent(
            payload, Encoding.UTF8, "application/json");

        using var response = await graph.PostAsync(url, content, ct);

        if (response.StatusCode == HttpStatusCode.Conflict)
            return;

        if (!response.IsSuccessStatusCode)
            await ThrowGraphErrorAsync(response, "Create destination folder", ct);
    }

    private static async Task<JsonDocument> ReadSuccessJsonAsync(
        HttpResponseMessage response,
        string operation,
        CancellationToken ct)
    {
        if (!response.IsSuccessStatusCode)
            await ThrowGraphErrorAsync(response, operation, ct);

        var stream = await response.Content.ReadAsStreamAsync(ct);
        return await JsonDocument.ParseAsync(stream, cancellationToken: ct);
    }

    private static string GetRequiredString(
        JsonElement element,
        string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var value) ||
            string.IsNullOrWhiteSpace(value.GetString()))
        {
            throw new InvalidOperationException(
                $"Microsoft Graph response does not contain '{propertyName}'.");
        }

        return value.GetString()!;
    }

    private static async Task ThrowGraphErrorAsync(
        HttpResponseMessage response,
        string operation,
        CancellationToken ct)
    {
        var body = await response.Content.ReadAsStringAsync(ct);

        if (body.Length > 2000)
            body = body[..2000];

        throw new HttpRequestException(
            $"{operation} failed: {(int)response.StatusCode} {response.ReasonPhrase}. {body}");
    }

    private static string EncodeGraphPath(string path) =>
        string.Join("/",
            path
                .Split('/', StringSplitOptions.RemoveEmptyEntries)
                .Select(Uri.EscapeDataString));
}
