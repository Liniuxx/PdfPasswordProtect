using Azure.Core;
using Azure.Identity;

namespace VipaPDFProtect.Services;

public sealed class GraphTokenService
{
    private static readonly string[] Scopes =
        ["https://graph.microsoft.com/.default"];

    private readonly TokenCredential _credential;

    public GraphTokenService()
    {
        // Azure App Service:
        //   uses System Assigned Managed Identity automatically.
        //
        // Local development:
        //   can use Azure CLI / Visual Studio / VS Code credentials.
        _credential = new DefaultAzureCredential();
    }

    public async Task<string> GetTokenAsync(CancellationToken cancellationToken)
    {
        var token = await _credential.GetTokenAsync(
            new TokenRequestContext(Scopes),
            cancellationToken);

        return token.Token;
    }
}
