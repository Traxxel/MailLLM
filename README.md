# MailLLM - E-Mail Downloader für LLM-Recherche

Dieses Projekt lädt E-Mails aus Microsoft 365 herunter und extrahiert den Text für die Verwendung in LLM-basierten Recherchen.

## Features

- Download von E-Mails über Exchange Web Services (EWS) oder POP3
- Automatische Text-Extraktion aus HTML- und Plain-Text E-Mails
- Speicherung als strukturierte TXT-Dateien mit Zeitstempel
- Unterstützung für verschiedene E-Mail-Formate

## Installation

1. Python 3.8+ installieren
2. Abhängigkeiten installieren:
```bash
pip install -r requirements.txt
```

## Konfiguration

1. `.env` Datei erstellen:
```bash
cp .env.example .env
```

2. Azure App-Registrierung durchführen:
   - Registriere eine Anwendung im [Azure-Portal](https://portal.azure.com) unter Azure Active Directory > App-Registrierungen.
   - Notiere dir die folgenden Werte:
     - `AZURE_CLIENT_ID` (Anwendungs-ID)
     - `AZURE_TENANT_ID` (Verzeichnis-ID)
     - `AZURE_CLIENT_SECRET` (Client-Geheimnis)
   - Weise die nötigen Berechtigungen für Microsoft Graph (z.B. `Mail.Read`) zu.

3. Trage die Zugangsdaten in `.env` ein:
```
EMAIL_ADDRESS=deine.email@domain.com
AZURE_CLIENT_ID=deine_client_id
AZURE_TENANT_ID=dein_tenant_id
AZURE_CLIENT_SECRET=dein_client_secret
EMAIL_SERVER=outlook.office365.com
USE_EWS=true

# Optionale Einstellungen für Ordner und Archive
INCLUDE_FOLDERS=true
INCLUDE_ARCHIVE=true
FOLDER_NAMES=Wichtig,Projekte,Newsletter

# Chunk-basiertes Laden (für große E-Mail-Mengen)
CHUNK_SIZE=50
LOAD_ALL_EMAILS=true
MAX_EMAILS_PER_FOLDER=0
```

> **Hinweis:** Das Feld `EMAIL_PASSWORD` wird nicht mehr benötigt, wenn OAuth2/Azure verwendet wird.

## Verwendung

```bash
# E-Mails aus Posteingang, Ordnern und Archiv herunterladen
python mail_downloader_graph.py

# Nur aus dem Posteingang (klassisch)
python mail_downloader.py
```

Die E-Mails werden im Verzeichnis `mails/` gespeichert mit dem Format:
`yyyy-mm-dd-hh-mm-ss--[Ordner]--Betreff.txt`

**Unterstützte Quellen:**
- Posteingang (Inbox)
- Benutzerdefinierte Ordner (konfigurierbar)
- Archiv (falls verfügbar)

**Chunk-basiertes Laden:**
- Lädt E-Mails in kleinen Paketen (standardmäßig 50 pro Chunk)
- Verhindert Timeouts bei großen E-Mail-Mengen
- `LOAD_ALL_EMAILS=true` lädt alle verfügbaren E-Mails
- `CHUNK_SIZE=50` bestimmt die Größe jedes Pakets
- `MAX_EMAILS_PER_FOLDER=0` = unbegrenzt (oder setze einen Wert für Limit)

## Programme im Überblick

| Programm                   | Beschreibung                                                      | Beispielaufruf                                  |
|----------------------------|-------------------------------------------------------------------|-------------------------------------------------|
| mail_downloader.py         | Lädt E-Mails per EWS (klassisch, Passwort oder App-Passwort)      | `python mail_downloader.py`                     |
| mail_downloader_graph.py   | Lädt E-Mails per Microsoft Graph API (OAuth2, Azure Client Secret)| `python mail_downloader_graph.py`               |
| mail_search.py             | Sucht, fasst zusammen und exportiert E-Mails für LLMs             | `python mail_search.py --search Meeting`         |
|                            |                                                                   | `python mail_search.py --summary`                |
|                            |                                                                   | `python mail_search.py --export`                 |
| llm_integration_example.py | Beispiel für LLM-Integration mit geladenen E-Mails                | `python llm_integration_example.py`              |

## Sicherheit

- Speichere niemals echte Passwörter oder Secrets in öffentlichen Repositories
- Nutze ausschließlich Umgebungsvariablen oder eine `.env` Datei (diese ist in `.gitignore` aufgenommen)
- Erstelle und verwalte das Azure Client Secret sicher im Azure-Portal
- Vergib nur die minimal notwendigen Berechtigungen für die App-Registrierung 