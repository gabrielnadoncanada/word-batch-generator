# -*- coding: utf-8 -*-
"""
Test pour vérifier l'attachement de fichiers PDF
"""
import logging
from pathlib import Path
from smtp_email_sender import SMTPEmailSender

# Configuration du logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s'
)

def test_attachment():
    """Teste l'attachement d'un fichier PDF."""
    print("=== Test d'attachement de fichier PDF ===")
    
    # Créer un fichier PDF de test
    test_pdf = Path("test_attachment.pdf")
    test_content = b"Test PDF content for attachment verification"
    test_pdf.write_bytes(test_content)
    
    print(f"Fichier de test créé: {test_pdf}")
    print(f"Taille du fichier: {test_pdf.stat().st_size} bytes")
    
    # Données de test
    test_data = [{
        "nom": "Test Attachment",
        "email": "gabrielnadoncanada@gmail.com"  # Votre email pour recevoir le test
    }]
    
    try:
        sender = SMTPEmailSender(enabled=True)
        sent = sender.send_emails_batch(test_data, [test_pdf])
        
        if sent > 0:
            print("[OK] Email avec pièce jointe envoyé avec succès!")
            print("Vérifiez votre boîte email (et le dossier spam) pour voir la pièce jointe.")
        else:
            print("[ERREUR] Aucun email envoyé")
        
    except Exception as e:
        print(f"[ERREUR] Erreur lors de l'envoi: {e}")
    
    finally:
        # Nettoyer le fichier de test
        if test_pdf.exists():
            test_pdf.unlink()
            print("Fichier de test nettoyé")

if __name__ == "__main__":
    test_attachment()
