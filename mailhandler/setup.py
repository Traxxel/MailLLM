#!/usr/bin/env python3
"""
MailLLM - Setup-Skript
Installiert und konfiguriert das MailLLM-Projekt
"""

import os
import sys
import subprocess
from pathlib import Path


def run_command(command: str, description: str) -> bool:
    """Führt einen Befehl aus und gibt Status zurück"""
    print(f"🔄 {description}...")
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f"✅ {description} erfolgreich")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} fehlgeschlagen: {e}")
        print(f"Fehlerausgabe: {e.stderr}")
        return False


def check_python_version() -> bool:
    """Überprüft die Python-Version"""
    version = sys.version_info
    if version.major < 3 or (version.major == 3 and version.minor < 8):
        print(f"❌ Python 3.8+ erforderlich, aber {version.major}.{version.minor} gefunden")
        return False
    print(f"✅ Python {version.major}.{version.minor}.{version.micro} gefunden")
    return True


def install_dependencies() -> bool:
    """Installiert die Python-Abhängigkeiten"""
    return run_command("pip3 install -r requirements.txt", "Installiere Python-Abhängigkeiten")


def create_env_file() -> bool:
    """Erstellt die .env Datei aus der Vorlage"""
    if os.path.exists('.env'):
        print("ℹ️  .env Datei existiert bereits")
        return True
    
    if not os.path.exists('env.example'):
        print("❌ env.example nicht gefunden")
        return False
    
    try:
        with open('env.example', 'r') as src, open('.env', 'w') as dst:
            dst.write(src.read())
        print("✅ .env Datei aus env.example erstellt")
        return True
    except Exception as e:
        print(f"❌ Fehler beim Erstellen der .env Datei: {e}")
        return False


def create_mail_directory() -> bool:
    """Erstellt das E-Mail-Verzeichnis"""
    try:
        mail_dir = Path('mails')
        mail_dir.mkdir(exist_ok=True)
        print("✅ E-Mail-Verzeichnis 'mails' erstellt")
        return True
    except Exception as e:
        print(f"❌ Fehler beim Erstellen des E-Mail-Verzeichnisses: {e}")
        return False


def make_scripts_executable() -> bool:
    """Macht die Python-Skripte ausführbar"""
    scripts = ['mail_downloader.py', 'mail_search.py', 'llm_integration_example.py']
    
    for script in scripts:
        if os.path.exists(script):
            try:
                os.chmod(script, 0o755)
                print(f"✅ {script} ausführbar gemacht")
            except Exception as e:
                print(f"⚠️  Konnte {script} nicht ausführbar machen: {e}")
    
    return True


def test_installation() -> bool:
    """Testet die Installation"""
    print("\n🧪 Teste Installation...")
    
    # Teste Import der Hauptmodule
    try:
        import dotenv
        print("✅ python-dotenv importiert")
    except ImportError:
        print("❌ python-dotenv nicht verfügbar")
        return False
    
    try:
        import tqdm
        print("✅ tqdm importiert")
    except ImportError:
        print("❌ tqdm nicht verfügbar")
        return False
    
    try:
        import html2text
        print("✅ html2text importiert")
    except ImportError:
        print("❌ html2text nicht verfügbar")
        return False
    
    # Teste EWS/POP3-Bibliotheken
    try:
        from exchangelib import Credentials
        print("✅ exchangelib verfügbar")
    except ImportError:
        print("⚠️  exchangelib nicht verfügbar (EWS deaktiviert)")
    
    try:
        import poplib
        print("✅ poplib verfügbar (Standardbibliothek)")
    except ImportError:
        print("⚠️  poplib nicht verfügbar (POP3 deaktiviert)")
    
    # Teste Standardbibliotheken
    try:
        import email
        print("✅ email (Standardbibliothek) verfügbar")
    except ImportError:
        print("⚠️  email nicht verfügbar")
    
    return True


def show_next_steps():
    """Zeigt die nächsten Schritte an"""
    print("\n" + "="*60)
    print("🎉 MailLLM Setup abgeschlossen!")
    print("="*60)
    
    print("\n📋 Nächste Schritte:")
    print("1. Bearbeite die .env Datei mit deinen E-Mail-Zugangsdaten:")
    print("   nano .env")
    print("\n2. Teste den E-Mail-Download:")
    print("   python3 mail_downloader.py")
    print("\n3. Suche in deinen E-Mails:")
    print("   python3 mail_search.py --search 'Meeting'")
    print("\n4. Erstelle eine Zusammenfassung:")
    print("   python3 mail_search.py --summary")
    print("\n5. Exportiere für LLM:")
    print("   python3 mail_search.py --export")
    print("\n6. Teste LLM-Integration (optional):")
    print("   python3 llm_integration_example.py")
    
    print("\n🔧 Konfiguration:")
    print("- EWS (Exchange Web Services): Standard für M365")
    print("- POP3: Alternative für ältere Systeme")
    print("- Verwende App-Passwörter für bessere Sicherheit")
    
    print("\n📚 Dokumentation:")
    print("- README.md: Projektübersicht")
    print("- mail_downloader.py: Hauptprogramm")
    print("- mail_search.py: Such- und Exportfunktionen")
    print("- llm_integration_example.py: LLM-Integration")


def main():
    """Hauptfunktion"""
    print("🚀 MailLLM - Setup-Skript")
    print("="*40)
    
    # Python-Version prüfen
    if not check_python_version():
        sys.exit(1)
    
    # Abhängigkeiten installieren
    if not install_dependencies():
        print("❌ Installation der Abhängigkeiten fehlgeschlagen")
        sys.exit(1)
    
    # .env Datei erstellen
    if not create_env_file():
        print("❌ Erstellung der .env Datei fehlgeschlagen")
        sys.exit(1)
    
    # E-Mail-Verzeichnis erstellen
    if not create_mail_directory():
        print("❌ Erstellung des E-Mail-Verzeichnisses fehlgeschlagen")
        sys.exit(1)
    
    # Skripte ausführbar machen
    make_scripts_executable()
    
    # Installation testen
    if not test_installation():
        print("❌ Installationstest fehlgeschlagen")
        sys.exit(1)
    
    # Nächste Schritte anzeigen
    show_next_steps()


if __name__ == "__main__":
    main() 