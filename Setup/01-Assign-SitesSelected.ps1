param(
    [Parameter(Mandatory=$true)]
    [string]$ManagedIdentityObjectId
)

$ErrorActionPreference = "Stop"

$graphSpId = az ad sp show `
    --id "00000003-0000-0000-c000-000000000000" `
    --query id -o tsv

if (-not $graphSpId) {
    throw "Microsoft Graph service principal not found."
}

$sitesSelectedRoleId =
    "883ea226-0bf2-4a8f-9f9d-92c9162a727d"

az rest `
    --method POST `
    --uri "https://graph.microsoft.com/v1.0/servicePrincipals/$ManagedIdentityObjectId/appRoleAssignments" `
    --headers "Content-Type=application/json" `
    --body "{`"principalId`":`"$ManagedIdentityObjectId`",`"resourceId`":`"$graphSpId`",`"appRoleId`":`"$sitesSelectedRoleId`"}"

Write-Host "Sites.Selected assigned."
