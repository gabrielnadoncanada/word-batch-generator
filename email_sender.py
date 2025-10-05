# -*- coding: utf-8 -*-
"""
Gestionnaire d'envoi d'emails via SMTP
"""
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional

from config import SEND_EMAIL
from smtp_email_sender import SMTPEmailSender


class EmailSender:
    """Classe pour gérer l'envoi d'emails via SMTP."""
    
    def __init__(self, enabled: bool = SEND_EMAIL):
        self.enabled = enabled
        self.sender = SMTPEmailSender(enabled)
        logging.info("Système SMTP activé")
    
    def send_emails_batch(self, rows: List[Dict[str, Any]], pdf_files: List[Path]) -> int:
        """Envoie les emails pour une liste de données."""
        if not self.enabled:
            logging.info("Envoi d'emails désactivé")
            return 0
        
        return self.sender.send_emails_batch(rows, pdf_files)
    
    def test_connection(self) -> bool:
        """Teste la connexion SMTP."""
        return self.sender.test_connection()

