#!/usr/bin/env python3
"""
MailLLM - E-Mail Such- und Vorbereitungstool
Bereitet E-Mails für LLM-Recherchen vor
"""

import os
import re
import json
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any
import logging

from dotenv import load_dotenv

# Logging konfigurieren
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class MailSearch:
    """Klasse für E-Mail-Suche und -Vorbereitung"""
    
    def __init__(self, mail_dir: str = "mails"):
        self.mail_dir = Path(mail_dir)
        self.mail_index = []
        
        if not self.mail_dir.exists():
            raise ValueError(f"E-Mail-Verzeichnis {mail_dir} existiert nicht")
    
    def load_mail_index(self) -> List[Dict[str, Any]]:
        """Lädt alle E-Mails und erstellt einen Index"""
        logger.info(f"Lade E-Mail-Index aus {self.mail_dir}")
        
        mail_files = list(self.mail_dir.glob("*.txt"))
        index = []
        
        for mail_file in mail_files:
            try:
                mail_data = self._parse_mail_file(mail_file)
                if mail_data:
                    index.append(mail_data)
            except Exception as e:
                logger.error(f"Fehler beim Parsen von {mail_file}: {e}")
                continue
        
        # Nach Datum sortieren (neueste zuerst)
        index.sort(key=lambda x: x['received_date'], reverse=True)
        
        self.mail_index = index
        logger.info(f"Index erstellt: {len(index)} E-Mails")
        return index
    
    def _parse_mail_file(self, filepath: Path) -> Dict[str, Any]:
        """Parst eine E-Mail-Datei und extrahiert Metadaten"""
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Metadaten extrahieren
        metadata = {}
        lines = content.split('\n')
        
        for line in lines[:10]:  # Nur die ersten 10 Zeilen für Metadaten
            if line.startswith('Von: '):
                metadata['from'] = line[5:].strip()
            elif line.startswith('An: '):
                metadata['to'] = line[4:].strip()
            elif line.startswith('Datum: '):
                date_str = line[7:].strip()
                try:
                    metadata['received_date'] = datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
                except:
                    metadata['received_date'] = datetime.now()
            elif line.startswith('Betreff: '):
                metadata['subject'] = line[9:].strip()
                break
        
        # Inhalt extrahieren (alles nach den Metadaten)
        content_start = content.find('\n\n')
        if content_start != -1:
            body = content[content_start + 2:].strip()
        else:
            body = content
        
        return {
            'filename': filepath.name,
            'filepath': str(filepath),
            'from': metadata.get('from', 'Unbekannt'),
            'to': metadata.get('to', 'Unbekannt'),
            'received_date': metadata.get('received_date', datetime.now()),
            'subject': metadata.get('subject', 'Kein Betreff'),
            'body': body,
            'body_length': len(body),
            'word_count': len(body.split())
        }
    
    def search_emails(self, query: str, max_results: int = 10) -> List[Dict[str, Any]]:
        """Sucht E-Mails nach Text"""
        if not self.mail_index:
            self.load_mail_index()
        
        query_lower = query.lower()
        results = []
        
        for mail in self.mail_index:
            score = 0
            
            # Suche in Betreff
            if query_lower in mail['subject'].lower():
                score += 10
            
            # Suche im Absender
            if query_lower in mail['from'].lower():
                score += 5
            
            # Suche im Inhalt
            if query_lower in mail['body'].lower():
                score += 1
            
            if score > 0:
                mail_copy = mail.copy()
                mail_copy['search_score'] = score
                results.append(mail_copy)
        
        # Nach Relevanz sortieren
        results.sort(key=lambda x: x['search_score'], reverse=True)
        
        return results[:max_results]
    
    def get_emails_by_date_range(self, start_date: datetime, end_date: datetime) -> List[Dict[str, Any]]:
        """Holt E-Mails aus einem bestimmten Zeitraum"""
        if not self.mail_index:
            self.load_mail_index()
        
        results = []
        for mail in self.mail_index:
            if start_date <= mail['received_date'] <= end_date:
                results.append(mail)
        
        return results
    
    def export_for_llm(self, output_file: str = "emails_for_llm.json", 
                      max_emails: int = 100) -> str:
        """Exportiert E-Mails in einem LLM-freundlichen Format"""
        if not self.mail_index:
            self.load_mail_index()
        
        # E-Mails für LLM vorbereiten
        llm_data = []
        
        for mail in self.mail_index[:max_emails]:
            llm_entry = {
                'id': mail['filename'],
                'date': mail['received_date'].isoformat(),
                'from': mail['from'],
                'subject': mail['subject'],
                'content': mail['body'],
                'word_count': mail['word_count']
            }
            llm_data.append(llm_entry)
        
        # Als JSON speichern
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(llm_data, f, ensure_ascii=False, indent=2)
        
        logger.info(f"LLM-Export erstellt: {output_file} ({len(llm_data)} E-Mails)")
        return output_file
    
    def create_summary(self) -> Dict[str, Any]:
        """Erstellt eine Zusammenfassung der E-Mails"""
        if not self.mail_index:
            self.load_mail_index()
        
        if not self.mail_index:
            return {}
        
        total_emails = len(self.mail_index)
        total_words = sum(mail['word_count'] for mail in self.mail_index)
        
        # Datumsbereich
        dates = [mail['received_date'] for mail in self.mail_index]
        earliest_date = min(dates)
        latest_date = max(dates)
        
        # Häufigste Absender
        senders = {}
        for mail in self.mail_index:
            sender = mail['from']
            senders[sender] = senders.get(sender, 0) + 1
        
        top_senders = sorted(senders.items(), key=lambda x: x[1], reverse=True)[:5]
        
        return {
            'total_emails': total_emails,
            'total_words': total_words,
            'average_words_per_email': total_words / total_emails if total_emails > 0 else 0,
            'date_range': {
                'earliest': earliest_date.isoformat(),
                'latest': latest_date.isoformat()
            },
            'top_senders': top_senders
        }


def main():
    """Hauptfunktion für Kommandozeilen-Interface"""
    import argparse
    
    parser = argparse.ArgumentParser(description='MailLLM - E-Mail Such- und Vorbereitungstool')
    parser.add_argument('--mail-dir', default='mails', help='E-Mail-Verzeichnis')
    parser.add_argument('--search', help='Suchbegriff')
    parser.add_argument('--export', action='store_true', help='Export für LLM erstellen')
    parser.add_argument('--summary', action='store_true', help='Zusammenfassung anzeigen')
    parser.add_argument('--max-results', type=int, default=10, help='Maximale Anzahl Ergebnisse')
    
    args = parser.parse_args()
    
    try:
        mail_search = MailSearch(args.mail_dir)
        
        if args.search:
            print(f"🔍 Suche nach: {args.search}")
            results = mail_search.search_emails(args.search, args.max_results)
            
            if results:
                print(f"\n📧 Gefunden: {len(results)} E-Mails")
                for i, mail in enumerate(results, 1):
                    print(f"\n{i}. {mail['subject']}")
                    print(f"   Von: {mail['from']}")
                    print(f"   Datum: {mail['received_date'].strftime('%Y-%m-%d %H:%M')}")
                    print(f"   Wörter: {mail['word_count']}")
                    print(f"   Relevanz: {mail['search_score']}")
            else:
                print("❌ Keine E-Mails gefunden")
        
        elif args.export:
            output_file = mail_search.export_for_llm()
            print(f"✅ LLM-Export erstellt: {output_file}")
        
        elif args.summary:
            summary = mail_search.create_summary()
            if summary:
                print("📊 E-Mail-Zusammenfassung:")
                print(f"   Gesamt E-Mails: {summary['total_emails']}")
                print(f"   Gesamt Wörter: {summary['total_words']}")
                print(f"   Durchschnitt Wörter/E-Mail: {summary['average_words_per_email']:.1f}")
                print(f"   Zeitraum: {summary['date_range']['earliest']} bis {summary['date_range']['latest']}")
                print("\n   Top Absender:")
                for sender, count in summary['top_senders']:
                    print(f"     {sender}: {count} E-Mails")
            else:
                print("❌ Keine E-Mails gefunden")
        
        else:
            # Standard: Zusammenfassung anzeigen
            summary = mail_search.create_summary()
            if summary:
                print("📊 E-Mail-Zusammenfassung:")
                print(f"   Gesamt E-Mails: {summary['total_emails']}")
                print(f"   Gesamt Wörter: {summary['total_words']}")
                print(f"   Durchschnitt Wörter/E-Mail: {summary['average_words_per_email']:.1f}")
            else:
                print("❌ Keine E-Mails gefunden")
    
    except Exception as e:
        logger.error(f"Fehler: {e}")
        print(f"❌ Fehler: {e}")


if __name__ == "__main__":
    main() 