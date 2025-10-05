# -*- coding: utf-8 -*-
"""
Configuration du générateur de documents Word et envoi d'emails
"""
import os
from pathlib import Path
from typing import Optional
from dotenv import load_dotenv

# Charger les variables d'environnement depuis .env
load_dotenv()

# Chemins de base
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE = BASE_DIR / "templates" / "modele.docx"
CSV_FILE = BASE_DIR / "data" / "entrepreneurs.csv"
OUT_DOCX_DIR = BASE_DIR / "out" / "docx"
OUT_PDF_DIR = BASE_DIR / "out" / "pdf"
LOG_DIR = OUT_PDF_DIR.parent / "logs"
LOG_FILE = LOG_DIR / "mail.log"

# Configuration des placeholders
PLACEHOLDER = "{{VENDEUR}}"

# Configuration email
SEND_EMAIL = True
FROM_ACCOUNT: Optional[str] = os.getenv("FROM_ACCOUNT", "gabriel@dilamco.com")
CC = os.getenv("CC", "")
BCC = os.getenv("BCC", "")
SUBJECT_TEMPLATE = "Soumission - 25142 - École Arc-en-ciel Pavillon 1 (Laval)"

# Configuration SMTP (depuis variables d'environnement)
SMTP_SERVER = os.getenv("SMTP_SERVER", "secure.emailsrvr.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))
SMTP_USERNAME = os.getenv("SMTP_USERNAME", "gabriel@dilamco.com")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD", "")
SMTP_USE_TLS = os.getenv("SMTP_USE_TLS", "false").lower() == "true"
SMTP_USE_SSL = os.getenv("SMTP_USE_SSL", "true").lower() == "true"

# Configuration templates d'emails
USE_EMAIL_TEMPLATE = True
EMAIL_TEMPLATE_FILE = BASE_DIR / "templates" / "emails" / "soumission_template.html"
FALLBACK_BODY_HTML_TEMPLATE = (
    "Bonjour, <br>Je me permets de vous transmettre notre soumission pour le projet "
    "25142 - École Arc-en-ciel Pavillon 1 (Laval).<br><br>Vous trouverez ci-joint le "
    "document détaillant notre proposition pour la fourniture et l'installation de la vanité."
    "<br>Nous restons disponibles pour toute précision ou information complémentaire."
)

# Configuration signature
USE_SYSTEM_SIGNATURE = False  # Désactiver les signatures Outlook
USE_PROJECT_SIGNATURE = True   # Utiliser la signature du projet
SIGNATURE_NAME = "Dilamco (gabriel@dilamco.com)"
STRICT_EMBED_SIGNATURE_IMAGES = False
PROJECT_SIGNATURE_FILE = BASE_DIR / "signatures" / "dilamco_signature.html"

# Configuration retry
MAX_RETRIES = 5
DELAY_SECONDS = 2.0

# Configuration logging
LOG_LEVEL = "DEBUG"
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"

