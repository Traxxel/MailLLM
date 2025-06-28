# mailhandler_cs

## Zweck

Dieses Projekt ist ein reiner Mail-Downloader für Microsoft 365/Exchange Online (Graph API). Es lädt E-Mails (inkl. PDF-Anhänge) aus Posteingang, Unterordnern und Archiv und speichert sie als Textdateien/PDFs lokal ab.

## Voraussetzungen
- .NET 8 SDK
- Microsoft 365/Exchange Online-Konto mit App-Registrierung (ClientId, TenantId, ClientSecret)
- Die App muss die Berechtigung `Mail.Read` (Application) im Azure-Portal besitzen

## Konfiguration

Lege eine Datei `appsettings.json` im Projektverzeichnis an (wird **nicht** ins Git eingecheckt, siehe `.gitignore`).
Eine Beispielkonfiguration findest du in [`appsettings.json.example`](./appsettings.json.example):

```json
{
  "MailSettings": {
    "EmailAddress": "deine-adresse@beispiel.de",
    "ClientId": "<CLIENT_ID>",
    "TenantId": "<TENANT_ID>",
    "ClientSecret": "<CLIENT_SECRET>",
    "MailDir": "mails",
    "IncludeFolders": "true",
    "IncludeArchive": "true",
    "ChunkSize": "50",
    "LoadAllEmails": "true",
    "MaxEmailsPerFolder": "0",
    "DaysBack": "30",
    "MaxEmails": "100"
  }
}
```

**Wichtiger Hinweis:**
- Trage deine echten Zugangsdaten nur in `appsettings.json` ein, niemals in die Beispiel- oder Git-Dateien!

## Nutzung

1. Restore & Build:
   ```
   dotnet restore
   dotnet build
   ```
2. Ausführen:
   ```
   dotnet run --project mailhandler_cs
   ```
3. Die E-Mails werden im Verzeichnis `MailDir` (Standard: `mails/`) als Textdateien gespeichert. PDF-Anhänge landen in `mails/pdf/`.

## Hinweise
- Die Datei `.gitignore` schützt sensible Daten und Build-Artefakte.
- Nur die Datei `appsettings.json.example` ist im Repository enthalten.
- Bei Problemen prüfe die Azure-App-Berechtigungen und die Konfiguration. 