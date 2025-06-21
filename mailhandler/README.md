# MailLLM

**Achtung:** Dieses Projekt unterstützt ausschließlich Microsoft Graph API für den E-Mail-Download. POP3 und EWS werden nicht mehr unterstützt.

## Nutzung

1. Konfiguriere deine Zugangsdaten und Parameter in der `.env`-Datei (siehe `env.example`).
2. Starte den Download:

```bash
python mail_downloader_graph.py
```

- Es werden alle E-Mails aus Posteingang, Unterordnern und Archiv geladen (je nach Konfiguration).
- Die E-Mails werden im Verzeichnis `mails/` als Textdateien gespeichert.

## Wichtige Hinweise
- Nur Microsoft 365/Exchange Online-Konten mit OAuth2/Graph API werden unterstützt.
- Für Legacy-Protokolle (IMAP, POP3, EWS) gibt es keine Unterstützung mehr.

## Konfiguration
Alle Einstellungen erfolgen über die `.env`-Datei. Siehe `env.example` für Beispiele.

## Programme im Überblick
- **mail_downloader_graph.py**: Hauptskript für den Download via Microsoft Graph API
- **mail_search.py**: Suche in den gespeicherten E-Mails
- **llm_integration_example.py**: Beispiel für LLM-Integration 