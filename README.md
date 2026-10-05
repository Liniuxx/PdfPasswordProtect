# VipaPDFProtect

ASP.NET Core .NET 10 API that:
- downloads a PDF from a SharePoint site's default document library using Microsoft Graph;
- accepts `PlainPassword` or decrypts `EncryptedPassword` with AES-128 / ECB / PKCS7;
- protects the PDF with a user password using PDFsharp AES-128 PDF encryption;
- creates `<DestinationUrl>/Merged` if needed;
- uploads `<original>_withpassword.pdf`;
- returns the protected SharePoint URL.

## Required Azure setting

Set this only in Azure App Service -> Settings -> Environment variables:

`PDF_PASSWORD_ENCRYPTION_KEY`

The value must be exactly 16 UTF-8 bytes. Do not commit the real key to GitHub.

## Authentication

Uses `DefaultAzureCredential`. In Azure App Service enable System Assigned Managed Identity and grant that identity Microsoft Graph / SharePoint access to the target site.

## Endpoint

`POST /api/pdf/protect`

Encrypted password example:

```json
{
  "OneDriveUrl": "https://xxx.sharepoint.com/sites/correspondence/",
  "SourceUrl": "https://xxx.sharepoint.com/sites/correspondence/folder/file.pdf",
  "DestinationUrl": "https://xxx.sharepoint.com/sites/correspondence/folder",
  "EncryptedPassword": "BASE64_ENCRYPTED_PASSWORD"
}
```

Plain password example:

```json
{
  "OneDriveUrl": "https://xxx.sharepoint.com/sites/correspondence/",
  "SourceUrl": "https://xxx.sharepoint.com/sites/correspondence/folder/file.pdf",
  "DestinationUrl": "https://xxx.sharepoint.com/sites/correspondence/folder",
  "PlainPassword": "123456"
}
```

Success response:

```json
{
  "result": {
    "message": "",
    "errorMessage": "",
    "innerErrorMessage": "",
    "isSuccessful": true,
    "protectedPDFFilePath": "https://xxx.sharepoint.com/sites/correspondence/folder/Merged/file_withpassword.pdf"
  }
}
```

Health check:

`GET /health`

## Notes

- This implementation uses the SharePoint site's **default document library**.
- Microsoft Graph simple upload is intended for files up to 250 MB.
- `DestinationUrl` is treated as a folder.
