# -*- coding: utf-8 -*-
"""
Configuration du système de logging
"""
import logging
from pathlib import Path
from config import LOG_FILE, LOG_LEVEL, LOG_FORMAT


def setup_logging():
    """Configure le système de logging."""
    # Créer le répertoire de logs s'il n'existe pas
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    
    # Configurer le logging
    logging.basicConfig(
        level=getattr(logging, LOG_LEVEL.upper()),
        format=LOG_FORMAT,
        handlers=[
            logging.FileHandler(LOG_FILE, encoding="utf-8"),
            logging.StreamHandler()
        ],
    )
    
    # Logger spécifique pour les emails
    email_logger = logging.getLogger("email")
    email_logger.setLevel(logging.INFO)
    
    return logging.getLogger(__name__)

