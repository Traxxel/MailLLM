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

2. E-Mail-Zugangsdaten in `.env` eintragen:
```
EMAIL_ADDRESS=deine.email@domain.com
EMAIL_PASSWORD=dein_passwort
EMAIL_SERVER=outlook.office365.com
EMAIL_PORT=587
USE_EWS=true
```

## Verwendung

```bash
python mail_downloader.py
```

Die E-Mails werden im Verzeichnis `mails/` gespeichert mit dem Format:
`yyyy-mm-dd-hh-mm-ss--Betreff.txt`

## Sicherheit

- Verwende niemals echte Passwörter in der `.env` Datei
- Nutze App-Passwörter oder OAuth2-Tokens
- Die `.env` Datei ist in `.gitignore` aufgenommen 