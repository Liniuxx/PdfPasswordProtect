using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace VipaPDFProtect.Services;

public sealed class GraphOneDriveService
{
    private const string GraphBase = "https://graph.microsoft.com/v1.0";

    private readonly IHttpClientFactory _httpClientFactory;
    private readonly GraphTokenService _tokenService;

    public GraphOneDriveService(
        IHttpClientFactory httpClientFactory,
        GraphTokenService tokenService)
    {
        _httpClientFactory = httpClientFactory;
        _tokenService = tokenService;
    }

    public async Task<DriveLocation> ResolveLocationAsync(
        string hostname,
        string sitePath,
        string folderPath,
        CancellationToken cancellationToken)
    {
        var site = await GetJsonAsync<GraphSite>(
            $"/sites/{hostname}:{EncodeGraphPath(sitePath)}?$select=id,webUrl",
            cancellationToken);

        var drive = await GetJsonAsync<GraphDrive>(
            $"/sites/{Uri.EscapeDataString(site.Id)}/drive?$select=id,webUrl",
            cancellationToken);

        var folder = await GetJsonAsync<GraphDriveItem>(
            $"/drives/{Uri.EscapeDataString(drive.Id)}/root:{EncodeGraphPath(folderPath)}?$select=id,name,webUrl,folder",
            cancellationToken);

        if (folder.Folder is null)
            throw new InvalidOperationException("FolderPath does not point to a folder.");

        return new DriveLocation(site.Id, drive.Id, folder.Id);
    }

    public async Task<GraphDriveItem> GetItemByPathAsync(
        string driveId,
        string itemPath,
        CancellationToken cancellationToken)
    {
        return await GetJsonAsync<GraphDriveItem>(
            $"/drives/{Uri.EscapeDataString(driveId)}/root:{EncodeGraphPath(itemPath)}?$select=id,name,webUrl,file,folder",
            cancellationToken);
    }

    public async Task<List<GraphDriveItem>> GetPdfFilesAsync(
        string driveId,
        string folderItemId,
        CancellationToken cancellationToken)
    {
        var result = new List<GraphDriveItem>();

        string? url =
            $"{GraphBase}/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(folderItemId)}/children" +
            "?$select=id,name,file,folder,webUrl&$top=200";

        while (!string.IsNullOrWhiteSpace(url))
        {
            var page =
                await GetJsonAbsoluteAsync<GraphCollection<GraphDriveItem>>(
                    url,
                    cancellationToken);

            result.AddRange(page.Value.Where(x =>
                x.File is not null &&
                x.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)));

            url = page.NextLink;
        }

        return result;
    }

    public async Task<Stream> DownloadFileAsync(
        string driveId,
        string itemId,
        CancellationToken cancellationToken)
    {
        var client = await CreateClientAsync(cancellationToken);

        using var response = await client.GetAsync(
            $"{GraphBase}/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(itemId)}/content",
            HttpCompletionOption.ResponseHeadersRead,
            cancellationToken);

        await EnsureSuccessAsync(response, cancellationToken);

        var memory = new MemoryStream();
        await response.Content.CopyToAsync(memory, cancellationToken);
        memory.Position = 0;

        return memory;
    }

    public async Task<GraphDriveItem> EnsureFolderPathAsync(
        string driveId,
        string folderPath,
        CancellationToken cancellationToken)
    {
        var parts = folderPath.Split(
            '/',
            StringSplitOptions.RemoveEmptyEntries |
            StringSplitOptions.TrimEntries);

        var current = await GetJsonAsync<GraphDriveItem>(
            $"/drives/{Uri.EscapeDataString(driveId)}/root?$select=id,name,webUrl,folder",
            cancellationToken);

        var accumulated = new List<string>();

        foreach (var part in parts)
        {
            accumulated.Add(part);
            var path = string.Join("/", accumulated);

            try
            {
                current = await GetJsonAsync<GraphDriveItem>(
                    $"/drives/{Uri.EscapeDataString(driveId)}/root:{EncodeGraphPath(path)}?$select=id,name,webUrl,folder",
                    cancellationToken);
            }
            catch (GraphNotFoundException)
            {
                current = await CreateFolderAsync(
                    driveId,
                    current.Id,
                    part,
                    cancellationToken);
            }
        }

        return current;
    }

    public async Task<GraphDriveItem> UploadFileAsync(
        string driveId,
        string parentFolderId,
        string fileName,
        Stream content,
        CancellationToken cancellationToken)
    {
        content.Position = 0;

        var client = await CreateClientAsync(cancellationToken);

        using var streamContent = new StreamContent(content);
        streamContent.Headers.ContentType =
            new MediaTypeHeaderValue("application/pdf");

        var url =
            $"{GraphBase}/drives/{Uri.EscapeDataString(driveId)}/items/" +
            $"{Uri.EscapeDataString(parentFolderId)}:/{Uri.EscapeDataString(fileName)}:/content";

        using var response =
            await client.PutAsync(url, streamContent, cancellationToken);

        await EnsureSuccessAsync(response, cancellationToken);

        var json =
            await response.Content.ReadAsStringAsync(cancellationToken);

        return JsonSerializer.Deserialize<GraphDriveItem>(json, JsonOptions)
            ?? throw new InvalidOperationException(
                "Graph returned an empty upload response.");
    }

    private async Task<GraphDriveItem> CreateFolderAsync(
        string driveId,
        string parentId,
        string name,
        CancellationToken cancellationToken)
    {
        var payload = JsonSerializer.Serialize(new Dictionary<string, object?>
        {
            ["name"] = name,
            ["folder"] = new Dictionary<string, object>(),
            ["@microsoft.graph.conflictBehavior"] = "fail"
        });

        var client = await CreateClientAsync(cancellationToken);

        using var response = await client.PostAsync(
            $"{GraphBase}/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(parentId)}/children",
            new StringContent(payload, Encoding.UTF8, "application/json"),
            cancellationToken);

        if (response.StatusCode == HttpStatusCode.Conflict)
        {
            // Folder was likely created by another request.
            // Resolve the child by name from the parent.
            var children = await GetJsonAsync<GraphCollection<GraphDriveItem>>(
                $"/drives/{Uri.EscapeDataString(driveId)}/items/{Uri.EscapeDataString(parentId)}/children?$select=id,name,webUrl,folder",
                cancellationToken);

            var existing = children.Value.FirstOrDefault(x =>
                x.Folder is not null &&
                string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase));

            if (existing is not null)
                return existing;
        }

        await EnsureSuccessAsync(response, cancellationToken);

        var json =
            await response.Content.ReadAsStringAsync(cancellationToken);

        return JsonSerializer.Deserialize<GraphDriveItem>(json, JsonOptions)
            ?? throw new InvalidOperationException(
                "Graph returned an empty create-folder response.");
    }

    private async Task<T> GetJsonAsync<T>(
        string relativeUrl,
        CancellationToken cancellationToken)
    {
        var url = relativeUrl.StartsWith(
            "http",
            StringComparison.OrdinalIgnoreCase)
            ? relativeUrl
            : GraphBase + relativeUrl;

        return await GetJsonAbsoluteAsync<T>(url, cancellationToken);
    }

    private async Task<T> GetJsonAbsoluteAsync<T>(
        string url,
        CancellationToken cancellationToken)
    {
        var client = await CreateClientAsync(cancellationToken);

        using var response =
            await client.GetAsync(url, cancellationToken);

        if (response.StatusCode == HttpStatusCode.NotFound)
            throw new GraphNotFoundException(url);

        await EnsureSuccessAsync(response, cancellationToken);

        var json =
            await response.Content.ReadAsStringAsync(cancellationToken);

        return JsonSerializer.Deserialize<T>(json, JsonOptions)
            ?? throw new InvalidOperationException(
                "Graph returned an empty response.");
    }

    private async Task<HttpClient> CreateClientAsync(
        CancellationToken cancellationToken)
    {
        var client = _httpClientFactory.CreateClient();
        var token =
            await _tokenService.GetTokenAsync(cancellationToken);

        client.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", token);

        return client;
    }

    private static async Task EnsureSuccessAsync(
        HttpResponseMessage response,
        CancellationToken cancellationToken)
    {
        if (response.IsSuccessStatusCode)
            return;

        var body =
            await response.Content.ReadAsStringAsync(cancellationToken);

        throw new HttpRequestException(
            $"Microsoft Graph returned {(int)response.StatusCode} " +
            $"{response.ReasonPhrase}. {body}");
    }

    private static string EncodeGraphPath(string path)
    {
        var clean =
            Uri.UnescapeDataString(path).Trim('/');

        if (string.IsNullOrEmpty(clean))
            return "/";

        return "/" + string.Join("/",
            clean.Split('/', StringSplitOptions.RemoveEmptyEntries)
                 .Select(Uri.EscapeDataString));
    }

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };
}

public sealed record DriveLocation(
    string SiteId,
    string DriveId,
    string FolderItemId);

public sealed class GraphSite
{
    public string Id { get; set; } = string.Empty;
    public string? WebUrl { get; set; }
}

public sealed class GraphDrive
{
    public string Id { get; set; } = string.Empty;
    public string? WebUrl { get; set; }
}

public sealed class GraphDriveItem
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string? WebUrl { get; set; }
    public JsonElement? File { get; set; }
    public JsonElement? Folder { get; set; }
}

public sealed class GraphCollection<T>
{
    public List<T> Value { get; set; } = new();

    [JsonPropertyName("@odata.nextLink")]
    public string? NextLink { get; set; }
}

public sealed class GraphNotFoundException : Exception
{
    public GraphNotFoundException(string url)
        : base($"Graph resource not found: {url}") { }
}
