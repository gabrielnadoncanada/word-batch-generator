# -*- coding: utf-8 -*-
"""
Script de test pour la configuration SMTP
"""
import logging
from pathlib import Path
from smtp_email_sender import SMTPEmailSender
from config import SMTP_SERVER, SMTP_PORT, SMTP_USERNAME, SMTP_PASSWORD

# Configuration du logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s'
)

def test_smtp_connection():
    """Teste la connexion SMTP."""
    print("=== Test de connexion SMTP ===")
    print(f"SMTP Server: {SMTP_SERVER}")
    print(f"SMTP Port: {SMTP_PORT}")
    print(f"Username: {SMTP_USERNAME}")
    print(f"Password: {'*' * len(SMTP_PASSWORD) if SMTP_PASSWORD else 'NON CONFIGURÉ'}")
    print()
    
    if not SMTP_PASSWORD:
        print("❌ ERREUR: Mot de passe SMTP non configuré dans config.py")
        print("Veuillez ajouter votre mot de passe dans la variable SMTP_PASSWORD")
        return False
    
    sender = SMTPEmailSender(enabled=True)
    success = sender.test_connection()
    
    if success:
        print("[OK] Connexion SMTP reussie!")
        return True
    else:
        print("[ERREUR] Echec de la connexion SMTP")
        return False

def test_send_email():
    """Teste l'envoi d'un email de test."""
    print("\n=== Test d'envoi d'email ===")
    
    # Données de test
    test_data = [{
        "nom": "Test User",
        "email": "test@example.com"  # Changez cette adresse pour tester
    }]
    
    # Créer un fichier PDF de test (vide)
    test_pdf = Path("test.pdf")
    test_pdf.write_bytes(b"Test PDF content")
    
    try:
        sender = SMTPEmailSender(enabled=True)
        sent = sender.send_emails_batch(test_data, [test_pdf])
        
        if sent > 0:
            print("[OK] Email de test envoye avec succes!")
        else:
            print("[ERREUR] Aucun email envoye")
        
        # Nettoyer le fichier de test
        test_pdf.unlink()
        
    except Exception as e:
        print(f"[ERREUR] Erreur lors de l'envoi: {e}")
        # Nettoyer le fichier de test
        if test_pdf.exists():
            test_pdf.unlink()

if __name__ == "__main__":
    print("Test de la configuration SMTP")
    print("=" * 40)
    
    
    # Test de connexion
    if test_smtp_connection():
        # Test d'envoi (décommentez pour tester l'envoi réel)
        # test_send_email()
        print("\n[OK] Configuration SMTP prete a utiliser!")
    else:
        print("\n[ERREUR] Veuillez corriger la configuration SMTP")
