param(
    [Parameter(Mandatory=$true)]
    [string]$Hostname,

    [Parameter(Mandatory=$true)]
    [string]$SitePath,

    [Parameter(Mandatory=$true)]
    [string]$ManagedIdentityClientId,

    [string]$DisplayName = "VipaPDFProtect Managed Identity"
)

$ErrorActionPreference = "Stop"

$site = az rest `
    --method GET `
    --uri "https://graph.microsoft.com/v1.0/sites/$Hostname`:$SitePath?`$select=id,webUrl" |
    ConvertFrom-Json

if (-not $site.id) {
    throw "Target site could not be resolved."
}

$body = @{
    roles = @("write")
    grantedToIdentities = @(
        @{
            application = @{
                id = $ManagedIdentityClientId
                displayName = $DisplayName
            }
        }
    )
} | ConvertTo-Json -Depth 10 -Compress

az rest `
    --method POST `
    --uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/permissions" `
    --headers "Content-Type=application/json" `
    --body $body

Write-Host "WRITE permission granted to:" $site.webUrl
