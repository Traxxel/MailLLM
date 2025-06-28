#!/usr/bin/env python3
"""
MailLLM - E-Mail Downloader mit Microsoft Graph API
Verwendet OAuth2 und Microsoft Graph API für moderne M365-Konten
"""

import os
import re
import sys
import json
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Dict, Any, Optional
import logging

from dotenv import load_dotenv
from tqdm import tqdm
import html2text
import requests

# OAuth2-Bibliotheken
try:
    import msal
    OAUTH2_AVAILABLE = True
except ImportError:
    OAUTH2_AVAILABLE = False
    print("Warnung: msal nicht verfügbar. Installiere mit: pip install msal")

# Logging konfigurieren
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('mail_downloader.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


class MailDownloaderGraph:
    """Hauptklasse für den E-Mail-Download mit Microsoft Graph API"""
    
    def __init__(self):
        load_dotenv()
        self.email_address = os.getenv('EMAIL_ADDRESS')
        self.client_id = os.getenv('CLIENT_ID')
        self.client_secret = os.getenv('CLIENT_SECRET')
        self.tenant_id = os.getenv('TENANT_ID')
        self.max_emails = int(os.getenv('MAX_EMAILS', '100'))
        self.days_back = int(os.getenv('DAYS_BACK', '30'))
        self.mail_dir = Path(os.getenv('MAIL_DIR', 'mails'))
        
        # PDF-Verzeichnis für Attachments
        self.pdf_dir = self.mail_dir / 'pdf'
        
        # Neue Optionen für Unterverzeichnisse und Archive
        self.include_folders = os.getenv('INCLUDE_FOLDERS', 'true').lower() == 'true'
        self.include_archive = os.getenv('INCLUDE_ARCHIVE', 'true').lower() == 'true'
        
        # Chunk-basiertes Laden
        self.chunk_size = int(os.getenv('CHUNK_SIZE', '50'))
        self.load_all_emails = os.getenv('LOAD_ALL_EMAILS', 'true').lower() == 'true'
        self.max_emails_per_folder = int(os.getenv('MAX_EMAILS_PER_FOLDER', '0'))  # 0 = unbegrenzt
        
        # Graph API Endpoints
        self.graph_endpoint = "https://graph.microsoft.com/v1.0"
        
        # HTML zu Text Konverter
        self.html_converter = html2text.HTML2Text()
        self.html_converter.ignore_links = False
        self.html_converter.ignore_images = False
        self.html_converter.body_width = 0
        
        # Deduplizierung für PDF-Attachments
        self.seen_pdf_attachment_ids = set()
        
        # Verzeichnisse erstellen
        self.mail_dir.mkdir(exist_ok=True)
        self.pdf_dir.mkdir(exist_ok=True)
        
        self._validate_config()
    
    def _validate_config(self):
        """Validiert die Konfiguration"""
        if not self.email_address:
            raise ValueError("EMAIL_ADDRESS muss in .env gesetzt werden")
        
        if not OAUTH2_AVAILABLE:
            raise ValueError("msal ist nicht verfügbar")
        
        if not self.client_id or not self.client_secret or not self.tenant_id:
            raise ValueError("CLIENT_ID, CLIENT_SECRET und TENANT_ID müssen in .env gesetzt werden")
    
    def get_access_token(self) -> str:
        """Holt ein OAuth2-Token für Microsoft Graph API"""
        logger.info("Hole OAuth2-Token für Microsoft Graph...")
        
        # MSAL-Anwendung erstellen
        app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}"
        )
        
        # Scopes für E-Mail-Zugriff
        scopes = ['https://graph.microsoft.com/.default']
        
        # Token anfordern
        result = app.acquire_token_silent(scopes, account=None)
        if not result:
            result = app.acquire_token_for_client(scopes=scopes)
        
        if result and "access_token" in result:
            logger.info("OAuth2-Token erfolgreich erhalten")
            return result['access_token']
        else:
            error_msg = f"Fehler beim OAuth2-Token: {result.get('error_description', result.get('error', 'Unbekannter Fehler')) if result else 'Kein Token erhalten'}"
            logger.error(error_msg)
            raise ValueError(error_msg)
    
    def sanitize_filename(self, filename: str) -> str:
        """Bereinigt Dateinamen von ungültigen Zeichen und kürzt auf 50 Zeichen"""
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        filename = re.sub(r'_+', '_', filename)
        filename = filename.strip()
        if len(filename) > 50:
            filename = filename[:50]
        return filename
    
    def extract_text_from_email(self, email_content: str, content_type: str = 'text/plain') -> str:
        """Extrahiert reinen Text aus E-Mail-Inhalt"""
        if content_type.startswith('text/html'):
            # HTML zu Text konvertieren
            text = self.html_converter.handle(email_content)
        else:
            # Plain text
            text = email_content
        
        # Text bereinigen
        text = text.strip()
        
        # HTML-Tags entfernen (falls noch welche übrig sind)
        import re
        text = re.sub(r'<[^>]+>', '', text)
        
        # HTML-Entities dekodieren
        import html
        text = html.unescape(text)
        
        # Mehrfache Leerzeichen entfernen
        text = re.sub(r'\s+', ' ', text)
        
        # Mehrfache Zeilenumbrüche entfernen
        text = re.sub(r'\n\s*\n', '\n\n', text)
        
        # Leerzeichen am Anfang und Ende von Zeilen entfernen
        text = '\n'.join(line.strip() for line in text.split('\n'))
        
        return text
    
    def get_emails_from_graph(self, access_token: str) -> List[Dict[str, Any]]:
        """Holt E-Mails von Microsoft Graph API"""
        logger.info("Hole E-Mails von Microsoft Graph API...")
        
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Zeitraum definieren
        since_date = (datetime.now() - timedelta(days=self.days_back)).strftime('%Y-%m-%dT%H:%M:%SZ')
        
        all_emails = []
        
        # 1. Posteingang (Inbox)
        logger.info("Lade E-Mails aus dem Posteingang...")
        inbox_emails = self._get_emails_from_folder(access_token, headers, since_date, "Inbox")
        all_emails.extend(inbox_emails)
        
        logger.info(f"E-Mails aus Inbox geladen: {len(inbox_emails)}")
        
        # 2. Unterverzeichnisse (Ordner)
        if self.include_folders:
            logger.info("Lade E-Mails aus Unterverzeichnissen...")
            folder_emails = self._get_emails_from_folders(access_token, headers, since_date)
            all_emails.extend(folder_emails)
            
            logger.info(f"E-Mails aus Ordnern geladen: {len(folder_emails)}")
        
        # 3. Archiv
        if self.include_archive:
            logger.info("Lade E-Mails aus dem Archiv...")
            try:
                archive_emails = self._get_emails_from_folder(access_token, headers, since_date, "Archive")
                all_emails.extend(archive_emails)
                
                logger.info(f"E-Mails aus Archiv geladen: {len(archive_emails)}")
            except Exception as e:
                logger.warning(f"Archiv nicht verfügbar: {e}")
        
        logger.info(f"Gesamt E-Mails von Graph API erhalten: {len(all_emails)}")
        return all_emails
    
    def _get_emails_from_folder(self, access_token: str, headers: Dict[str, str], since_date: str, folder_name: str) -> List[Dict[str, Any]]:
        """Holt E-Mails aus einem spezifischen Ordner mit Chunk-basiertem Laden"""
        try:
            # Graph API Anfrage für spezifischen Ordner
            if folder_name.lower() == "inbox":
                url = f"{self.graph_endpoint}/users/{self.email_address}/mailFolders/inbox/messages"
            elif folder_name.lower() == "archive":
                url = f"{self.graph_endpoint}/users/{self.email_address}/mailFolders/archive/messages"
            else:
                # Für andere Ordner müssen wir zuerst die Ordner-ID finden
                folder_id = self._get_folder_id(access_token, headers, folder_name)
                if not folder_id:
                    logger.warning(f"Ordner '{folder_name}' nicht gefunden")
                    return []
                url = f"{self.graph_endpoint}/users/{self.email_address}/mailFolders/{folder_id}/messages"
            
            all_emails = []
            skip_count = 0
            
            while True:
                params = {
                    '$top': self.chunk_size,
                    '$skip': skip_count,
                    '$orderby': 'receivedDateTime desc',
                    '$filter': f"receivedDateTime ge {since_date}",
                    '$select': 'id,subject,from,toRecipients,receivedDateTime,body,bodyPreview',
                    '$expand': 'attachments'
                }
                
                logger.info(f"Lade Chunk {skip_count//self.chunk_size + 1} für {folder_name} (Skip: {skip_count})")
                
                response = requests.get(url, headers=headers, params=params)
                response.raise_for_status()
                
                data = response.json()
                emails = data.get('value', [])
                
                if not emails:
                    logger.info(f"Keine weiteren E-Mails in {folder_name}")
                    break
                
                # Ordner-Information zu jeder E-Mail hinzufügen
                for email in emails:
                    email['folder_name'] = folder_name
                
                all_emails.extend(emails)
                logger.info(f"Chunk geladen: {len(emails)} E-Mails aus {folder_name}")
                
                # Prüfe ob wir alle E-Mails laden sollen oder nur bis max_emails
                if not self.load_all_emails and len(all_emails) >= self.max_emails:
                    logger.info(f"Maximale Anzahl E-Mails ({self.max_emails}) erreicht für {folder_name}")
                    all_emails = all_emails[:self.max_emails]
                    break
                
                # Prüfe ob es weitere E-Mails gibt
                if len(emails) < self.chunk_size:
                    logger.info(f"Alle E-Mails aus {folder_name} geladen")
                    break
                
                skip_count += self.chunk_size
                
                # Optionaler Sicherheitscheck: Maximal E-Mails pro Ordner
                if self.max_emails_per_folder > 0 and skip_count >= self.max_emails_per_folder:
                    logger.warning(f"Maximale Anzahl E-Mails ({self.max_emails_per_folder}) für {folder_name} erreicht")
                    break
            
            logger.info(f"Gesamt E-Mails aus {folder_name}: {len(all_emails)}")
            return all_emails
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Fehler bei Graph API Anfrage für {folder_name}: {e}")
            return []
    
    def _get_folder_id(self, access_token: str, headers: Dict[str, str], folder_name: str) -> Optional[str]:
        """Findet die ID eines Ordners anhand des Namens"""
        try:
            url = f"{self.graph_endpoint}/users/{self.email_address}/mailFolders"
            params = {
                '$filter': f"displayName eq '{folder_name}'"
            }
            
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()
            
            data = response.json()
            folders = data.get('value', [])
            
            if folders:
                return folders[0]['id']
            else:
                return None
                
        except requests.exceptions.RequestException as e:
            logger.error(f"Fehler beim Suchen des Ordners {folder_name}: {e}")
            return None
    
    def _get_emails_from_folders(self, access_token: str, headers: Dict[str, str], since_date: str) -> List[Dict[str, Any]]:
        """Holt E-Mails aus allen verfügbaren Ordnern"""
        all_emails = []
        
        try:
            # Alle Ordner auflisten
            url = f"{self.graph_endpoint}/users/{self.email_address}/mailFolders"
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            data = response.json()
            folders = data.get('value', [])
            
            for folder in folders:
                folder_name = folder.get('displayName', '')
                
                # Überspringe spezielle Ordner
                if folder_name.lower() in ['inbox', 'archive', 'sent items', 'deleted items', 'drafts']:
                    continue
                
                logger.info(f"Prüfe Ordner: {folder_name}")
                try:
                    folder_emails = self._get_emails_from_folder(access_token, headers, since_date, folder_name)
                    all_emails.extend(folder_emails)
                except Exception as e:
                    logger.warning(f"Fehler beim Zugriff auf Ordner {folder_name}: {e}")
                    continue
        
        except requests.exceptions.RequestException as e:
            logger.error(f"Fehler beim Auflisten der Ordner: {e}")
        
        return all_emails
    
    def download_via_graph_api(self) -> List[str]:
        """Lädt E-Mails über Microsoft Graph API herunter"""
        logger.info("Starte E-Mail-Download über Microsoft Graph API...")
        
        # OAuth2-Token holen
        access_token = self.get_access_token()
        
        # E-Mails direkt laden und speichern
        downloaded_files = []
        seen_email_ids = set()
        
        # 1. Posteingang (Inbox)
        logger.info("Lade und speichere E-Mails aus dem Posteingang...")
        inbox_files = self._download_and_save_emails_from_folder(access_token, "Inbox", seen_email_ids)
        downloaded_files.extend(inbox_files)
        
        # 2. Unterverzeichnisse (Ordner)
        if self.include_folders:
            logger.info("Lade und speichere E-Mails aus Unterverzeichnissen...")
            folder_files = self._download_and_save_emails_from_folders(access_token, seen_email_ids)
            downloaded_files.extend(folder_files)
        
        # 3. Archiv
        if self.include_archive:
            logger.info("Lade und speichere E-Mails aus dem Archiv...")
            try:
                archive_files = self._download_and_save_emails_from_folder(access_token, "Archive", seen_email_ids)
                downloaded_files.extend(archive_files)
            except Exception as e:
                logger.warning(f"Archiv nicht verfügbar: {e}")
        
        logger.info(f"Gesamt eindeutige E-Mails heruntergeladen: {len(downloaded_files)}")
        return downloaded_files
    
    def download_emails(self) -> List[str]:
        """Hauptmethode für den E-Mail-Download"""
        logger.info(f"Starte E-Mail-Download für {self.email_address}")
        logger.info(f"Verwende Microsoft Graph API")
        logger.info(f"Maximale Anzahl E-Mails: {self.max_emails}")
        logger.info(f"Zeitraum: {self.days_back} Tage zurück")
        
        try:
            return self.download_via_graph_api()
        except Exception as e:
            logger.error(f"Fehler beim E-Mail-Download: {e}")
            return []
    
    def _download_and_save_emails_from_folder(self, access_token: str, folder_name: str, seen_email_ids: set) -> List[str]:
        """Lädt E-Mails aus einem Ordner und speichert sie direkt"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # Zeitraum definieren
        since_date = (datetime.now() - timedelta(days=self.days_back)).strftime('%Y-%m-%dT%H:%M:%SZ')
        
        downloaded_files = []
        skip_count = 0
        
        try:
            # Graph API Anfrage für spezifischen Ordner
            if folder_name.lower() == "inbox":
                url = f"{self.graph_endpoint}/users/{self.email_address}/mailFolders/inbox/messages"
            elif folder_name.lower() == "archive":
                url = f"{self.graph_endpoint}/users/{self.email_address}/mailFolders/archive/messages"
            else:
                # Für andere Ordner müssen wir zuerst die Ordner-ID finden
                folder_id = self._get_folder_id(access_token, headers, folder_name)
                if not folder_id:
                    logger.warning(f"Ordner '{folder_name}' nicht gefunden")
                    return []
                url = f"{self.graph_endpoint}/users/{self.email_address}/mailFolders/{folder_id}/messages"
            
            while True:
                params = {
                    '$top': self.chunk_size,
                    '$skip': skip_count,
                    '$orderby': 'receivedDateTime desc',
                    '$filter': f"receivedDateTime ge {since_date}",
                    '$select': 'id,subject,from,toRecipients,receivedDateTime,body,bodyPreview',
                    '$expand': 'attachments'
                }
                
                logger.info(f"Lade Chunk {skip_count//self.chunk_size + 1} für {folder_name} (Skip: {skip_count})")
                
                response = requests.get(url, headers=headers, params=params)
                response.raise_for_status()
                
                data = response.json()
                emails = data.get('value', [])
                
                if not emails:
                    logger.info(f"Keine weiteren E-Mails in {folder_name}")
                    break
                
                # E-Mails verarbeiten und speichern
                chunk_saved = 0
                for email_data in emails:
                    email_id = email_data.get('id')
                    
                    # Deduplizierung
                    if email_id and email_id in seen_email_ids:
                        continue
                    
                    seen_email_ids.add(email_id)
                    
                    try:
                        # E-Mail speichern
                        filepath = self._save_email_data(email_data, folder_name)
                        if filepath:
                            downloaded_files.append(filepath)
                            chunk_saved += 1
                    except Exception as e:
                        logger.error(f"Fehler beim Speichern der E-Mail: {e}")
                        continue
                
                logger.info(f"Chunk geladen: {len(emails)} E-Mails aus {folder_name}, {chunk_saved} neue gespeichert")
                
                # Prüfe ob wir alle E-Mails laden sollen oder nur bis max_emails
                if not self.load_all_emails and len(downloaded_files) >= self.max_emails:
                    logger.info(f"Maximale Anzahl E-Mails ({self.max_emails}) erreicht für {folder_name}")
                    break
                
                # Prüfe ob es weitere E-Mails gibt
                if len(emails) < self.chunk_size:
                    logger.info(f"Alle E-Mails aus {folder_name} geladen")
                    break
                
                skip_count += self.chunk_size
                
                # Optionaler Sicherheitscheck: Maximal E-Mails pro Ordner
                if self.max_emails_per_folder > 0 and skip_count >= self.max_emails_per_folder:
                    logger.warning(f"Maximale Anzahl E-Mails ({self.max_emails_per_folder}) für {folder_name} erreicht")
                    break
            
            logger.info(f"Gesamt E-Mails aus {folder_name} gespeichert: {len(downloaded_files)}")
            return downloaded_files
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Fehler bei Graph API Anfrage für {folder_name}: {e}")
            return []
    
    def _download_and_save_emails_from_folders(self, access_token: str, seen_email_ids: set) -> List[str]:
        """Lädt E-Mails aus allen verfügbaren Ordnern und speichert sie direkt"""
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        downloaded_files = []
        
        try:
            # Alle Ordner auflisten
            url = f"{self.graph_endpoint}/users/{self.email_address}/mailFolders"
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            data = response.json()
            folders = data.get('value', [])
            
            for folder in folders:
                folder_name = folder.get('displayName', '')
                
                # Überspringe spezielle Ordner
                if folder_name.lower() in ['inbox', 'archive', 'sent items', 'deleted items', 'drafts']:
                    continue
                
                logger.info(f"Prüfe Ordner: {folder_name}")
                try:
                    folder_files = self._download_and_save_emails_from_folder(access_token, folder_name, seen_email_ids)
                    downloaded_files.extend(folder_files)
                except Exception as e:
                    logger.warning(f"Fehler beim Zugriff auf Ordner {folder_name}: {e}")
                    continue
        
        except requests.exceptions.RequestException as e:
            logger.error(f"Fehler beim Auflisten der Ordner: {e}")
        
        return downloaded_files
    
    def _save_email_data(self, email_data: Dict[str, Any], folder_name: str) -> Optional[str]:
        """Speichert eine einzelne E-Mail-Nachricht und lädt PDF-Attachments herunter"""
        try:
            # Empfangsdatum
            received_date_str = email_data.get('receivedDateTime', '')
            if received_date_str:
                try:
                    received_date = datetime.fromisoformat(received_date_str.replace('Z', '+00:00'))
                    date_str = received_date.strftime('%Y-%m-%d-%H-%M-%S')
                except:
                    date_str = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
            else:
                date_str = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
            
            # Betreff
            subject = email_data.get('subject', 'Kein_Betreff')
            subject = self.sanitize_filename(subject)
            
            # Dateiname mit Ordner-Präfix
            filename = f"{date_str}--[{folder_name}]--{subject}.txt"
            filepath = self.mail_dir / filename
            
            # E-Mail-Inhalt extrahieren
            body = email_data.get('body', {})
            content_type = body.get('contentType', 'text/plain')
            text_content = body.get('content', '')
            
            if not text_content:
                text_content = email_data.get('bodyPreview', 'Kein Inhalt verfügbar')
            
            # Text extrahieren
            extracted_text = self.extract_text_from_email(text_content, content_type)
            
            # Absender und Empfänger
            from_info = email_data.get('from', {})
            from_email = from_info.get('emailAddress', {}).get('address', 'Unbekannt')
            from_name = from_info.get('emailAddress', {}).get('name', '')
            
            to_recipients = email_data.get('toRecipients', [])
            to_email = to_recipients[0].get('emailAddress', {}).get('address', 'Unbekannt') if to_recipients else 'Unbekannt'
            
            # Metadaten hinzufügen
            email_text = f"""Von: {from_name} <{from_email}>
An: {to_email}
Datum: {received_date_str}
Betreff: {email_data.get('subject', 'Kein Betreff')}
Ordner: {folder_name}

{extracted_text}
"""
            
            # Datei speichern
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(email_text)
            
            logger.info(f"E-Mail gespeichert: {filename}")
            
            # PDF-Attachments herunterladen
            pdf_files = self._download_pdf_attachments(email_data, folder_name, date_str)
            if pdf_files:
                logger.info(f"{len(pdf_files)} PDF-Attachments für E-Mail heruntergeladen")
            
            return str(filepath)
            
        except Exception as e:
            logger.error(f"Fehler beim Speichern der E-Mail: {e}")
            return None
    
    def _download_pdf_attachments(self, email_data: Dict[str, Any], folder_name: str, date_str: str) -> List[str]:
        """Lädt PDF-Attachments aus einer E-Mail herunter"""
        downloaded_pdfs = []
        
        try:
            attachments = email_data.get('attachments', [])
            if not attachments:
                return downloaded_pdfs
            
            # Nur PDF-Attachments verarbeiten
            pdf_attachments = [att for att in attachments if att.get('contentType') and att.get('contentType', '').lower() == 'application/pdf']
            
            if not pdf_attachments:
                return downloaded_pdfs
            
            logger.info(f"Lade {len(pdf_attachments)} PDF-Attachments...")
            
            for attachment in pdf_attachments:
                try:
                    attachment_name = attachment.get('name', 'Unbekannt.pdf')
                    attachment_id = attachment.get('id')
                    email_id = email_data.get('id')
                    
                    if not attachment_id or not email_id:
                        logger.warning(f"Keine ID für Attachment oder E-Mail: {attachment_name}")
                        continue
                    
                    # Deduplizierung: Überspringe bereits gesehene PDF-Attachments
                    if attachment_id in self.seen_pdf_attachment_ids:
                        logger.info(f"PDF-Attachment bereits heruntergeladen, überspringe: {attachment_name}")
                        continue
                    
                    # PDF-Dateiname mit Timestamp erstellen
                    safe_name = self.sanitize_filename(attachment_name)
                    if not safe_name.lower().endswith('.pdf'):
                        safe_name += '.pdf'
                    
                    pdf_filename = f"{date_str}--[{folder_name}]--{safe_name}"
                    pdf_filepath = self.pdf_dir / pdf_filename
                    
                    # PDF-Daten herunterladen
                    pdf_content = self._download_attachment_content(email_id, attachment_id)
                    if pdf_content:
                        with open(pdf_filepath, 'wb') as f:
                            f.write(pdf_content)
                        
                        # Attachment-ID als gesehen markieren
                        self.seen_pdf_attachment_ids.add(attachment_id)
                        
                        downloaded_pdfs.append(str(pdf_filepath))
                        logger.info(f"PDF gespeichert: {pdf_filename}")
                    
                except Exception as e:
                    logger.error(f"Fehler beim Herunterladen von PDF-Attachment {attachment.get('name', 'Unbekannt')}: {e}")
                    continue
            
        except Exception as e:
            logger.error(f"Fehler beim Verarbeiten der PDF-Attachments: {e}")
        
        return downloaded_pdfs
    
    def _download_attachment_content(self, email_id: str, attachment_id: str) -> Optional[bytes]:
        """Lädt den Inhalt eines Attachments herunter"""
        try:
            access_token = self.get_access_token()
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            
            # Attachment-Inhalt herunterladen - korrekte URL über E-Mail-ID
            url = f"{self.graph_endpoint}/users/{self.email_address}/messages/{email_id}/attachments/{attachment_id}/$value"
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            
            return response.content
            
        except Exception as e:
            logger.error(f"Fehler beim Herunterladen des Attachment-Inhalts: {e}")
            return None


def main():
    """Hauptfunktion"""
    try:
        downloader = MailDownloaderGraph()
        downloaded_files = downloader.download_emails()
        
        logger.info(f"Download abgeschlossen. {len(downloaded_files)} E-Mails heruntergeladen.")
        logger.info(f"E-Mails gespeichert in: {downloader.mail_dir}")
        
        if downloaded_files:
            print(f"\n✅ Erfolgreich {len(downloaded_files)} E-Mails heruntergeladen!")
            print(f"📁 Speicherort: {downloader.mail_dir}")
        else:
            print("❌ Keine E-Mails heruntergeladen.")
            
    except Exception as e:
        logger.error(f"Kritischer Fehler: {e}")
        print(f"❌ Fehler: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main() 