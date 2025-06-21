#!/usr/bin/env python3
"""
MailLLM - LLM-Integrationsbeispiel
Zeigt, wie heruntergeladene E-Mails mit einem LLM verwendet werden können
"""

import json
import os
from typing import List, Dict, Any
from mail_search import MailSearch

# Beispiel für OpenAI Integration (optional)
try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False
    print("Hinweis: OpenAI-Bibliothek nicht installiert. Installiere mit: pip install openai")


class LLMIntegration:
    """Beispiel-Klasse für LLM-Integration mit E-Mails"""
    
    def __init__(self, mail_dir: str = "mails"):
        self.mail_search = MailSearch(mail_dir)
        self.mail_index = []
        
        # OpenAI Setup (falls verfügbar)
        if OPENAI_AVAILABLE:
            self.openai_client = openai.OpenAI(
                api_key=os.getenv('OPENAI_API_KEY')
            )
    
    def load_emails(self):
        """Lädt alle E-Mails in den Index"""
        self.mail_index = self.mail_search.load_mail_index()
        print(f"📧 {len(self.mail_index)} E-Mails geladen")
    
    def create_context_from_emails(self, query: str, max_emails: int = 5) -> str:
        """Erstellt Kontext aus relevanten E-Mails für LLM-Anfragen"""
        # Relevante E-Mails suchen
        relevant_emails = self.mail_search.search_emails(query, max_emails)
        
        if not relevant_emails:
            return "Keine relevanten E-Mails gefunden."
        
        # Kontext erstellen
        context = f"Basierend auf {len(relevant_emails)} relevanten E-Mails:\n\n"
        
        for i, email in enumerate(relevant_emails, 1):
            context += f"E-Mail {i}:\n"
            context += f"Datum: {email['received_date'].strftime('%Y-%m-%d %H:%M')}\n"
            context += f"Von: {email['from']}\n"
            context += f"Betreff: {email['subject']}\n"
            context += f"Inhalt: {email['body'][:500]}...\n\n"
        
        return context
    
    def ask_llm_about_emails(self, question: str, max_context_emails: int = 5) -> str:
        """Stellt eine Frage an das LLM basierend auf den E-Mails"""
        if not OPENAI_AVAILABLE:
            return "OpenAI-Bibliothek nicht verfügbar. Installiere mit: pip install openai"
        
        # Kontext aus relevanten E-Mails erstellen
        context = self.create_context_from_emails(question, max_context_emails)
        
        # Prompt erstellen
        prompt = f"""Du bist ein hilfreicher Assistent, der E-Mails analysiert. 
        
Kontext aus E-Mails:
{context}

Frage: {question}

Antworte basierend auf den bereitgestellten E-Mail-Informationen. 
Falls die Informationen nicht ausreichen, gib das an."""

        try:
            response = self.openai_client.chat.completions.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "Du bist ein hilfreicher Assistent für E-Mail-Analyse."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=1000,
                temperature=0.7
            )
            
            return response.choices[0].message.content
            
        except Exception as e:
            return f"Fehler bei der LLM-Anfrage: {e}"
    
    def create_email_summary(self, date_range_days: int = 7) -> str:
        """Erstellt eine Zusammenfassung der E-Mails der letzten Tage"""
        from datetime import datetime, timedelta
        
        end_date = datetime.now()
        start_date = end_date - timedelta(days=date_range_days)
        
        recent_emails = self.mail_search.get_emails_by_date_range(start_date, end_date)
        
        if not recent_emails:
            return f"Keine E-Mails in den letzten {date_range_days} Tagen gefunden."
        
        summary = f"E-Mail-Zusammenfassung der letzten {date_range_days} Tage:\n\n"
        summary += f"Anzahl E-Mails: {len(recent_emails)}\n"
        
        # Nach Absendern gruppieren
        senders = {}
        for email in recent_emails:
            sender = email['from']
            senders[sender] = senders.get(sender, 0) + 1
        
        summary += "\nE-Mails nach Absendern:\n"
        for sender, count in sorted(senders.items(), key=lambda x: x[1], reverse=True):
            summary += f"  {sender}: {count} E-Mails\n"
        
        # Wichtigste Themen (basierend auf Betreffzeilen)
        subjects = [email['subject'] for email in recent_emails]
        summary += f"\nAnzahl verschiedene Betreffzeilen: {len(set(subjects))}\n"
        
        return summary
    
    def export_for_vector_database(self, output_file: str = "emails_for_vector_db.json") -> str:
        """Exportiert E-Mails in einem Format für Vector-Datenbanken"""
        if not self.mail_index:
            self.load_emails()
        
        vector_data = []
        
        for email in self.mail_index:
            # Text für Embeddings vorbereiten
            text_for_embedding = f"""
            Betreff: {email['subject']}
            Von: {email['from']}
            Datum: {email['received_date'].strftime('%Y-%m-%d %H:%M')}
            Inhalt: {email['body']}
            """.strip()
            
            vector_entry = {
                'id': email['filename'],
                'text': text_for_embedding,
                'metadata': {
                    'date': email['received_date'].isoformat(),
                    'from': email['from'],
                    'subject': email['subject'],
                    'word_count': email['word_count'],
                    'filepath': email['filepath']
                }
            }
            vector_data.append(vector_entry)
        
        # Als JSON speichern
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(vector_data, f, ensure_ascii=False, indent=2)
        
        print(f"✅ Vector-Datenbank-Export erstellt: {output_file}")
        return output_file


def main():
    """Hauptfunktion mit Beispielen"""
    print("🚀 MailLLM - LLM-Integrationsbeispiel")
    print("=" * 50)
    
    # LLM-Integration initialisieren
    llm_integration = LLMIntegration()
    
    # E-Mails laden
    print("\n📧 Lade E-Mails...")
    llm_integration.load_emails()
    
    # Zusammenfassung erstellen
    print("\n📊 Erstelle E-Mail-Zusammenfassung...")
    summary = llm_integration.create_email_summary(7)
    print(summary)
    
    # Beispiel-Suche
    print("\n🔍 Beispiel-Suche nach 'Meeting'...")
    search_results = llm_integration.mail_search.search_emails("Meeting", 3)
    if search_results:
        print(f"Gefunden: {len(search_results)} E-Mails")
        for i, email in enumerate(search_results, 1):
            print(f"{i}. {email['subject']} (von {email['from']})")
    else:
        print("Keine E-Mails mit 'Meeting' gefunden")
    
    # LLM-Anfrage (falls OpenAI verfügbar)
    if OPENAI_AVAILABLE and os.getenv('OPENAI_API_KEY'):
        print("\n🤖 Stelle LLM-Anfrage...")
        question = "Welche wichtigen Termine oder Meetings wurden in den letzten E-Mails erwähnt?"
        answer = llm_integration.ask_llm_about_emails(question)
        print(f"Frage: {question}")
        print(f"Antwort: {answer}")
    else:
        print("\n🤖 LLM-Anfragen deaktiviert (OpenAI API Key nicht gesetzt)")
    
    # Vector-Datenbank-Export
    print("\n💾 Erstelle Vector-Datenbank-Export...")
    llm_integration.export_for_vector_database()
    
    print("\n✅ Beispiel abgeschlossen!")
    print("\nNächste Schritte:")
    print("1. Setze OPENAI_API_KEY in .env für LLM-Funktionen")
    print("2. Verwende die exportierten JSON-Dateien mit Vector-Datenbanken")
    print("3. Integriere in deine eigene LLM-Anwendung")


if __name__ == "__main__":
    main() 