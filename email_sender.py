# -*- coding: utf-8 -*-
"""
Gestionnaire d'envoi d'emails
"""
import logging
from pathlib import Path
from typing import List, Dict, Any, Optional

from config import (
    SEND_EMAIL, FROM_ACCOUNT, CC, BCC, SUBJECT_TEMPLATE, FALLBACK_BODY_HTML_TEMPLATE,
    USE_SYSTEM_SIGNATURE, USE_PROJECT_SIGNATURE, USE_EMAIL_TEMPLATE, SIGNATURE_NAME, STRICT_EMBED_SIGNATURE_IMAGES,
    MAX_RETRIES, DELAY_SECONDS
)
from outlook_utils import send_email_via_outlook, debug_list_outlook_accounts


class EmailSender:
    """Classe pour gérer l'envoi d'emails via Outlook."""
    
    def __init__(self, enabled: bool = SEND_EMAIL):
        self.enabled = enabled
        self.from_account = FROM_ACCOUNT
        self.cc = CC
        self.bcc = BCC
        self.subject_template = SUBJECT_TEMPLATE
        self.body_template = FALLBACK_BODY_HTML_TEMPLATE
        self.use_signature = USE_SYSTEM_SIGNATURE
        self.use_project_signature = USE_PROJECT_SIGNATURE
        self.use_email_template = USE_EMAIL_TEMPLATE
        self.signature_name = SIGNATURE_NAME
        self.strict_embed_images = STRICT_EMBED_SIGNATURE_IMAGES
        self.max_retries = MAX_RETRIES
        self.delay_seconds = DELAY_SECONDS
    
    def send_emails_batch(self, rows: List[Dict[str, Any]], pdf_files: List[Path]) -> int:
        """Envoie les emails pour une liste de données."""
        if not self.enabled:
            logging.info("Envoi d'emails désactivé")
            return 0
        
        sent = 0
        debug_list_outlook_accounts()
        
        for idx, row in enumerate(rows):
            to_email = (row.get("email") or "").strip()
            if not to_email:
                logging.info(f"Pas d'email pour la ligne {idx+1}: {row.get('nom')}")
                continue
            
            try:
                self._send_single_email(row, pdf_files[idx] if idx < len(pdf_files) else None)
                sent += 1
            except Exception as e:
                logging.error(f"Échec envoi à {to_email}: {e}")
        
        logging.info(f"Emails envoyés: {sent}/{len(rows)}")
        return sent
    
    def _send_single_email(self, row: Dict[str, Any], pdf_path: Optional[Path]) -> None:
        """Envoie un email pour une ligne de données."""
        to_email = row.get("email", "").strip()
        name = row.get("nom", "")
        
        subject = self.subject_template.format(nom=name)
        body_html = self.body_template.format(nom=name)
        
        attachments = [pdf_path] if pdf_path and pdf_path.exists() else []
        
        send_email_via_outlook(
            to_email=to_email,
            subject=subject,
            html_body=body_html,
            attachments=attachments,
            from_account=self.from_account,
            cc=self.cc,
            bcc=self.bcc,
            use_signature=self.use_signature,
            signature_name=self.signature_name,
            use_email_template=self.use_email_template,
            strict_embed_images=self.strict_embed_images,
            max_retries=self.max_retries,
            delay_seconds=self.delay_seconds
        )

