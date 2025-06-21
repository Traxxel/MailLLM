#!/usr/bin/env python3
"""
MailLLM - E-Mail Downloader für LLM-Recherche
Lädt E-Mails aus M365 herunter und speichert sie als TXT-Dateien
"""

import os
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import List, Optional
import logging

from dotenv import load_dotenv
from tqdm import tqdm
import html2text

# E-Mail-Bibliotheken
try:
    from exchangelib import Credentials, Account, DELEGATE, Configuration
    from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
    EWS_AVAILABLE = True
except ImportError:
    EWS_AVAILABLE = False
    print("Warnung: exchangelib nicht verfügbar. EWS-Funktionalität deaktiviert.")

try:
    import poplib
    import email
    from email.header import decode_header
    from email import utils as email_utils
    POP3_AVAILABLE = True
except ImportError:
    POP3_AVAILABLE = False
    print("Warnung: poplib nicht verfügbar. POP3-Funktionalität deaktiviert.")

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


class MailDownloader:
    """Hauptklasse für den E-Mail-Download"""
    
    def __init__(self):
        load_dotenv()
        self.email_address = os.getenv('EMAIL_ADDRESS')
        self.email_password = os.getenv('EMAIL_PASSWORD')
        self.email_server = os.getenv('EMAIL_SERVER', 'outlook.office365.com')
        self.email_port = int(os.getenv('EMAIL_PORT', '587'))
        self.use_ews = os.getenv('USE_EWS', 'true').lower() == 'true'
        self.max_emails = int(os.getenv('MAX_EMAILS', '100'))
        self.days_back = int(os.getenv('DAYS_BACK', '30'))
        self.mail_dir = Path(os.getenv('MAIL_DIR', 'mails'))
        
        # Neue Optionen für Unterverzeichnisse und Archive
        self.include_folders = os.getenv('INCLUDE_FOLDERS', 'true').lower() == 'true'
        self.include_archive = os.getenv('INCLUDE_ARCHIVE', 'true').lower() == 'true'
        self.folder_names = os.getenv('FOLDER_NAMES', '').split(',') if os.getenv('FOLDER_NAMES') else []
        
        # Chunk-basiertes Laden
        self.chunk_size = int(os.getenv('CHUNK_SIZE', '50'))
        self.load_all_emails = os.getenv('LOAD_ALL_EMAILS', 'true').lower() == 'true'
        
        # HTML zu Text Konverter
        self.html_converter = html2text.HTML2Text()
        self.html_converter.ignore_links = False
        self.html_converter.ignore_images = False
        self.html_converter.body_width = 0  # Keine Zeilenumbrüche
        
        # Verzeichnis erstellen
        self.mail_dir.mkdir(exist_ok=True)
        
        self._validate_config()
    
    def _validate_config(self):
        """Validiert die Konfiguration"""
        if not self.email_address or not self.email_password:
            raise ValueError("EMAIL_ADDRESS und EMAIL_PASSWORD müssen in .env gesetzt werden")
        
        if self.use_ews and not EWS_AVAILABLE:
            raise ValueError("EWS ist aktiviert aber exchangelib ist nicht verfügbar")
        
        if not self.use_ews and not POP3_AVAILABLE:
            raise ValueError("POP3 ist aktiviert aber poplib ist nicht verfügbar")
    
    def sanitize_filename(self, filename: str) -> str:
        """Bereinigt Dateinamen von ungültigen Zeichen"""
        # Ungültige Zeichen entfernen/ersetzen
        filename = re.sub(r'[<>:"/\\|?*]', '_', filename)
        # Mehrfache Unterstriche zusammenfassen
        filename = re.sub(r'_+', '_', filename)
        # Leerzeichen am Anfang/Ende entfernen
        filename = filename.strip()
        # Maximale Länge begrenzen
        if len(filename) > 100:
            filename = filename[:100]
        return filename
    
    def extract_text_from_email(self, email_content: str, content_type: str = 'text/plain') -> str:
        """Extrahiert Text aus E-Mail-Inhalt"""
        if content_type.startswith('text/html'):
            # HTML zu Text konvertieren
            text = self.html_converter.handle(email_content)
        else:
            # Plain text
            text = email_content
        
        # Text bereinigen
        text = text.strip()
        # Mehrfache Leerzeichen entfernen
        text = re.sub(r'\s+', ' ', text)
        # Mehrfache Zeilenumbrüche entfernen
        text = re.sub(r'\n\s*\n', '\n\n', text)
        
        return text
    
    def download_via_ews(self) -> List[str]:
        """Lädt E-Mails über Exchange Web Services herunter"""
        logger.info("Starte E-Mail-Download über EWS...")
        
        # SSL-Verifizierung deaktivieren (falls nötig)
        BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
        
        # Verbindung aufbauen
        credentials = Credentials(self.email_address, self.email_password)
        config = Configuration(service_endpoint=f'https://{self.email_server}/EWS/Exchange.asmx', credentials=credentials)
        account = Account(primary_smtp_address=self.email_address, config=config, autodiscover=False, access_type=DELEGATE)
        
        # Zeitraum definieren
        start_date = datetime.now() - timedelta(days=self.days_back)
        
        downloaded_files = []
        
        # 1. Posteingang (Inbox)
        logger.info("Lade E-Mails aus dem Posteingang...")
        inbox_files = self._download_from_folder(account.inbox, start_date, "Inbox")
        downloaded_files.extend(inbox_files)
        
        # 2. Unterverzeichnisse (Ordner)
        if self.include_folders:
            logger.info("Lade E-Mails aus Unterverzeichnissen...")
            folder_files = self._download_from_folders(account, start_date)
            downloaded_files.extend(folder_files)
        
        # 3. Archiv
        if self.include_archive:
            logger.info("Lade E-Mails aus dem Archiv...")
            try:
                # Prüfe ob Archive-Attribut existiert
                if hasattr(account, 'archive'):
                    archive_files = self._download_from_folder(account.archive, start_date, "Archive")
                    downloaded_files.extend(archive_files)
                else:
                    logger.info("Archiv nicht verfügbar für dieses Konto")
            except Exception as e:
                logger.warning(f"Archiv nicht verfügbar: {e}")
        
        return downloaded_files
    
    def _download_from_folder(self, folder, start_date: datetime, folder_name: str) -> List[str]:
        """Lädt E-Mails aus einem spezifischen Ordner mit Chunk-basiertem Laden"""
        try:
            downloaded_files = []
            offset = 0
            
            while True:
                # E-Mails in Chunks abrufen
                messages = folder.filter(received__gte=start_date).order_by('-datetime_received')[offset:offset + self.chunk_size]
                messages_list = list(messages)
                
                if not messages_list:
                    logger.info(f"Keine weiteren E-Mails in {folder_name}")
                    break
                
                logger.info(f"Lade Chunk {offset//self.chunk_size + 1} für {folder_name} (Offset: {offset})")
                
                for message in tqdm(messages_list, desc=f"E-Mails aus {folder_name} (Chunk {offset//self.chunk_size + 1})"):
                    try:
                        filepath = self._save_email_message(message, folder_name)
                        if filepath:
                            downloaded_files.append(filepath)
                    except Exception as e:
                        logger.error(f"Fehler beim Verarbeiten der E-Mail aus {folder_name}: {e}")
                        continue
                
                logger.info(f"Chunk geladen: {len(messages_list)} E-Mails aus {folder_name}")
                
                # Prüfe ob wir alle E-Mails laden sollen oder nur bis max_emails
                if not self.load_all_emails and len(downloaded_files) >= self.max_emails:
                    logger.info(f"Maximale Anzahl E-Mails ({self.max_emails}) erreicht für {folder_name}")
                    downloaded_files = downloaded_files[:self.max_emails]
                    break
                
                # Prüfe ob es weitere E-Mails gibt
                if len(messages_list) < self.chunk_size:
                    logger.info(f"Alle E-Mails aus {folder_name} geladen")
                    break
                
                offset += self.chunk_size
                
                # Sicherheitscheck: Maximal 1000 E-Mails pro Ordner
                if offset >= 1000:
                    logger.warning(f"Maximale Anzahl E-Mails (1000) für {folder_name} erreicht")
                    break
            
            logger.info(f"Gesamt E-Mails aus {folder_name}: {len(downloaded_files)}")
            return downloaded_files
        except Exception as e:
            logger.error(f"Fehler beim Zugriff auf {folder_name}: {e}")
            return []
    
    def _download_from_folders(self, account, start_date: datetime) -> List[str]:
        """Lädt E-Mails aus allen verfügbaren Ordnern"""
        downloaded_files = []
        
        try:
            # Alle Ordner auflisten
            folders = account.root.walk()
            
            for folder in folders:
                # Überspringe spezielle Ordner
                if folder.folder_class in ['IPF.Note', 'IPF'] and folder.name not in ['Inbox', 'Archive', 'Sent Items', 'Deleted Items']:
                    # Prüfe ob spezifische Ordner gefiltert werden sollen
                    if self.folder_names and folder.name not in self.folder_names:
                        continue
                    
                    logger.info(f"Prüfe Ordner: {folder.name}")
                    try:
                        folder_files = self._download_from_folder(folder, start_date, folder.name)
                        downloaded_files.extend(folder_files)
                    except Exception as e:
                        logger.warning(f"Fehler beim Zugriff auf Ordner {folder.name}: {e}")
                        continue
        
        except Exception as e:
            logger.error(f"Fehler beim Auflisten der Ordner: {e}")
        
        return downloaded_files
    
    def _save_email_message(self, message, folder_name: str = "Inbox") -> Optional[str]:
        """Speichert eine einzelne E-Mail-Nachricht"""
        try:
            # Empfangsdatum
            received_date = message.datetime_received
            date_str = received_date.strftime('%Y-%m-%d-%H-%M-%S')
            
            # Betreff
            subject = message.subject or "Kein_Betreff"
            subject = self.sanitize_filename(subject)
            
            # Dateiname mit Ordner-Präfix
            filename = f"{date_str}--[{folder_name}]--{subject}.txt"
            filepath = self.mail_dir / filename
            
            # E-Mail-Inhalt extrahieren
            if message.body:
                content_type = 'text/plain'
                if hasattr(message, 'body_type') and message.body_type == 'HTML':
                    content_type = 'text/html'
                
                text_content = self.extract_text_from_email(message.body, content_type)
            else:
                text_content = "Kein Inhalt verfügbar"
            
            # Metadaten hinzufügen
            email_text = f"""Von: {message.sender.email_address if message.sender else 'Unbekannt'}
An: {message.to_recipients[0].email_address if message.to_recipients else 'Unbekannt'}
Datum: {received_date.strftime('%Y-%m-%d %H:%M:%S')}
Betreff: {message.subject or 'Kein Betreff'}
Ordner: {folder_name}

{text_content}
"""
            
            # Datei speichern
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(email_text)
            
            logger.info(f"E-Mail gespeichert: {filename}")
            return str(filepath)
            
        except Exception as e:
            logger.error(f"Fehler beim Speichern der E-Mail: {e}")
            return None
    
    def download_via_pop3(self) -> List[str]:
        """Lädt E-Mails über POP3 herunter"""
        logger.info("Starte E-Mail-Download über POP3...")
        
        # Type-Check für Anmeldedaten
        if not self.email_address or not self.email_password:
            logger.error("EMAIL_ADDRESS oder EMAIL_PASSWORD nicht gesetzt")
            return []
        
        # POP3-Verbindung aufbauen
        pop3_server = poplib.POP3_SSL(self.email_server, 995)
        pop3_server.user(self.email_address)
        pop3_server.pass_(self.email_password)
        
        # Anzahl E-Mails abrufen
        num_messages = len(pop3_server.list()[1])
        messages_to_download = min(self.max_emails, num_messages)
        
        downloaded_files = []
        
        for i in tqdm(range(messages_to_download), desc="E-Mails herunterladen"):
            try:
                # E-Mail abrufen
                response, lines, octets = pop3_server.retr(num_messages - i)
                email_content = b'\n'.join(lines).decode('utf-8', errors='ignore')
                
                # E-Mail parsen
                msg = email.message_from_string(email_content)
                
                # Empfangsdatum
                date_str = msg['date']
                if date_str:
                    try:
                        parsed_date = email_utils.parsedate_to_datetime(date_str)
                        date_str = parsed_date.strftime('%Y-%m-%d-%H-%M-%S')
                    except:
                        date_str = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
                else:
                    date_str = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
                
                # Betreff
                subject = msg['subject'] or "Kein_Betreff"
                if subject:
                    subject = decode_header(subject)[0][0]
                    if isinstance(subject, bytes):
                        subject = subject.decode('utf-8', errors='ignore')
                subject = self.sanitize_filename(subject)
                
                # Dateiname
                filename = f"{date_str}--[POP3]--{subject}.txt"
                filepath = self.mail_dir / filename
                
                # E-Mail-Inhalt extrahieren
                text_content = ""
                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_maintype() == 'multipart':
                            continue
                        if part.get_content_maintype() == 'text':
                            payload = part.get_payload(decode=True)
                            if payload:
                                text_content = payload.decode('utf-8', errors='ignore')
                            content_type = part.get_content_type()
                            break
                else:
                    payload = msg.get_payload(decode=True)
                    if payload:
                        text_content = payload.decode('utf-8', errors='ignore')
                    content_type = msg.get_content_type()
                
                # Text extrahieren
                extracted_text = self.extract_text_from_email(text_content, content_type)
                
                # Metadaten hinzufügen
                email_text = f"""Von: {msg['from'] or 'Unbekannt'}
An: {msg['to'] or 'Unbekannt'}
Datum: {msg['date'] or 'Unbekannt'}
Betreff: {msg['subject'] or 'Kein Betreff'}
Ordner: POP3

{extracted_text}
"""
                
                # Datei speichern
                with open(filepath, 'w', encoding='utf-8') as f:
                    f.write(email_text)
                
                downloaded_files.append(str(filepath))
                logger.info(f"E-Mail gespeichert: {filename}")
                
            except Exception as e:
                logger.error(f"Fehler beim Verarbeiten der E-Mail: {e}")
                continue
        
        pop3_server.quit()
        return downloaded_files
    
    def download_emails(self) -> List[str]:
        """Hauptmethode für den E-Mail-Download"""
        logger.info(f"Starte E-Mail-Download für {self.email_address}")
        logger.info(f"Verwende {'EWS' if self.use_ews else 'POP3'}")
        logger.info(f"Maximale Anzahl E-Mails: {self.max_emails}")
        logger.info(f"Zeitraum: {self.days_back} Tage zurück")
        
        try:
            if self.use_ews:
                return self.download_via_ews()
            else:
                return self.download_via_pop3()
        except Exception as e:
            logger.error(f"Fehler beim E-Mail-Download: {e}")
            return []


def main():
    """Hauptfunktion"""
    try:
        downloader = MailDownloader()
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