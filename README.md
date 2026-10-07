# VipaPDFProtect

.NET 10 API based on the same Microsoft Graph / Managed Identity pattern as VipaPDFMerge.

## Endpoint
`POST /api/protect-pdf`

```json
{
  "OneDriveUrl": "https://vipagentura.sharepoint.com/sites/correspondence/",
  "SourceUrl": "https://vipagentura.sharepoint.com/sites/correspondence/Masiniai_2026_DNMF1/.../Pranesimas.pdf",
  "DestinationUrl": "https://vipagentura.sharepoint.com/sites/correspondence/Masiniai_2026_DNMF1/...",
  "EncryptedPassword": "BASE64..."
}
```

`PlainPassword` may be supplied instead of `EncryptedPassword`.

Output is uploaded to `DestinationUrl/Protected/<source>_withpassword.pdf`.

## Azure setting
Set `PDF_PASSWORD_ENCRYPTION_KEY` to the 16-byte AES-128 key. It is intentionally not stored in this repository.

EncryptedPassword decryption: AES-128 / ECB / PKCS7, Base64 input.

## Health
`GET /health`
