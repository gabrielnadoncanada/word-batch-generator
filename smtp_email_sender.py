# -*- coding: utf-8 -*-
"""
Gestionnaire d'envoi d'emails via SMTP (remplace Outlook)
"""
import logging
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from typing import List, Dict, Any, Optional
import traceback

from config import (
    SEND_EMAIL, FROM_ACCOUNT, CC, BCC, SUBJECT_TEMPLATE, FALLBACK_BODY_HTML_TEMPLATE,
    USE_SYSTEM_SIGNATURE, USE_PROJECT_SIGNATURE, USE_EMAIL_TEMPLATE, SIGNATURE_NAME,
    MAX_RETRIES, DELAY_SECONDS, SMTP_SERVER, SMTP_PORT, SMTP_USERNAME, 
    SMTP_PASSWORD, SMTP_USE_TLS, SMTP_USE_SSL
)
from file_utils import read_text_smart


class SMTPEmailSender:
    """Classe pour gérer l'envoi d'emails via SMTP."""
    
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
        self.max_retries = MAX_RETRIES
        self.delay_seconds = DELAY_SECONDS
        
        # Configuration SMTP
        self.smtp_server = SMTP_SERVER
        self.smtp_port = SMTP_PORT
        self.smtp_username = SMTP_USERNAME
        self.smtp_password = SMTP_PASSWORD
        self.smtp_use_tls = SMTP_USE_TLS
        self.smtp_use_ssl = SMTP_USE_SSL
    
    def send_emails_batch(self, rows: List[Dict[str, Any]], pdf_files: List[Path]) -> int:
        """Envoie les emails pour une liste de données."""
        if not self.enabled:
            logging.info("Envoi d'emails désactivé")
            return 0
        
        
        if not self.smtp_password:
            logging.error("Mot de passe SMTP non configuré dans config.py")
            return 0
        
        sent = 0
        
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
                logging.error(traceback.format_exc())
        
        logging.info(f"Emails envoyés: {sent}/{len(rows)}")
        return sent
    
    def _send_single_email(self, row: Dict[str, Any], pdf_path: Optional[Path]) -> None:
        """Envoie un email pour une ligne de données."""
        to_email = row.get("email", "").strip()
        name = row.get("nom", "")
        
        subject = self.subject_template.format(nom=name)
        body_html = self._prepare_email_body(name)
        
        # Créer le message
        msg = MIMEMultipart('alternative')
        msg['From'] = self.from_account
        msg['To'] = to_email
        msg['Subject'] = subject
        
        if self.cc:
            msg['Cc'] = self.cc
        if self.bcc:
            msg['Bcc'] = self.bcc
        
        # Ajouter le contenu HTML
        html_part = MIMEText(body_html, 'html', 'utf-8')
        msg.attach(html_part)
        
        # Ajouter les pièces jointes
        if pdf_path and pdf_path.exists():
            self._attach_file(msg, pdf_path)
        
        # Envoyer l'email
        self._send_via_smtp(msg, to_email)
    
    def _prepare_email_body(self, name: str) -> str:
        """Prépare le corps de l'email avec template et signature."""
        # Charger le template d'email si activé
        if self.use_email_template:
            try:
                from config import EMAIL_TEMPLATE_FILE
                if EMAIL_TEMPLATE_FILE.exists():
                    with open(EMAIL_TEMPLATE_FILE, "r", encoding="utf-8") as f:
                        template_html = f.read()
                    # Remplacer les variables dans le template
                    from datetime import datetime
                    template_html = template_html.replace("{{DATE_SOUMISSION}}", datetime.now().strftime("%d/%m/%Y"))
                    body_html = template_html
                    logging.debug("[SMTP] Template d'email chargé")
                else:
                    body_html = self.body_template.format(nom=name)
            except Exception as e:
                logging.warning(f"[SMTP] Erreur chargement template: {e}")
                body_html = self.body_template.format(nom=name)
        else:
            body_html = self.body_template.format(nom=name)
        
        # Ajouter la signature
        signature_html = self._load_signature()
        if signature_html:
            body_html += signature_html
            logging.debug("[SMTP] Signature ajoutée")
        
        return body_html
    
    def _load_signature(self) -> str:
        """Charge la signature appropriée."""
        try:
            # Priorité: signature du projet
            if self.use_project_signature:
                from config import PROJECT_SIGNATURE_FILE
                if PROJECT_SIGNATURE_FILE.exists():
                    with open(PROJECT_SIGNATURE_FILE, "r", encoding="utf-8") as f:
                        return f.read()
        except Exception as e:
            logging.warning(f"[SMTP] Erreur chargement signature: {e}")
        
        return ""
    
    def _attach_file(self, msg: MIMEMultipart, file_path: Path) -> None:
        """Attache un fichier au message."""
        try:
            with open(file_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {file_path.name}'
            )
            msg.attach(part)
            logging.debug(f"[SMTP] Fichier attaché: {file_path}")
        except Exception as e:
            logging.warning(f"[SMTP] Impossible d'attacher {file_path}: {e}")
    
    def _send_via_smtp(self, msg: MIMEMultipart, to_email: str) -> None:
        """Envoie l'email via SMTP avec retry."""
        for attempt in range(1, self.max_retries + 1):
            try:
                # Créer la connexion SMTP
                if self.smtp_use_ssl:
                    context = ssl.create_default_context()
                    server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port, context=context)
                else:
                    server = smtplib.SMTP(self.smtp_server, self.smtp_port)
                    if self.smtp_use_tls:
                        context = ssl.create_default_context()
                        server.starttls(context=context)
                
                # Authentification
                server.login(self.smtp_username, self.smtp_password)
                
                # Préparer les destinataires
                recipients = [to_email]
                if self.cc:
                    recipients.extend([email.strip() for email in self.cc.split(',') if email.strip()])
                if self.bcc:
                    recipients.extend([email.strip() for email in self.bcc.split(',') if email.strip()])
                
                # Envoyer l'email
                text = msg.as_string()
                server.sendmail(self.from_account, recipients, text)
                server.quit()
                
                logging.info(f"[SMTP] Email envoyé avec succès à {to_email}")
                return
                
            except Exception as e:
                logging.warning(f"[SMTP] Tentative {attempt} échouée: {e}")
                if attempt < self.max_retries:
                    import time
                    time.sleep(self.delay_seconds)
                else:
                    raise
        
        raise RuntimeError(f"[SMTP] Échec envoi à {to_email} après {self.max_retries} tentatives.")
    
    def test_connection(self) -> bool:
        """Teste la connexion SMTP."""
        try:
            if self.smtp_use_ssl:
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port, context=context)
            else:
                server = smtplib.SMTP(self.smtp_server, self.smtp_port)
                if self.smtp_use_tls:
                    context = ssl.create_default_context()
                    server.starttls(context=context)
            
            server.login(self.smtp_username, self.smtp_password)
            server.quit()
            logging.info("[SMTP] Connexion testée avec succès")
            return True
        except Exception as e:
            logging.error(f"[SMTP] Échec test connexion: {e}")
            return False
