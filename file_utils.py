# -*- coding: utf-8 -*-
"""
Utilitaires pour la gestion des fichiers
"""
import csv
import logging
from pathlib import Path
from typing import List, Dict, Any
import platform


def safe_filename(name: str) -> str:
    """Crée un nom de fichier sécurisé à partir d'une chaîne."""
    return "".join(c for c in name if c.isalnum() or c in (" ", "-", "_")).strip().replace(" ", "_")


def safe_email_for_filename(email: str) -> str:
    """Crée un nom de fichier sécurisé à partir d'un email."""
    email = (email or "").replace("@", "_at_")
    return "".join(c for c in email if c.isalnum() or c in ("-", "_", "."))


def read_text_smart(path: Path) -> str:
    """Lit un fichier texte en essayant plusieurs encodages."""
    data = path.read_bytes()
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    # Dernier recours : ne pas perdre les caractères
    return data.decode("utf-8", errors="replace")


def read_csv_rows(csv_file: Path) -> List[Dict[str, Any]]:
    """Lit les lignes d'un fichier CSV et retourne une liste de dictionnaires."""
    rows: List[Dict[str, Any]] = []
    
    if not csv_file.exists():
        logging.error(f"Fichier CSV introuvable: {csv_file}")
        return rows
    
    try:
        with open(csv_file, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for row in reader:
                name = (row.get("nom") or "").strip()
                email = (row.get("email") or "").strip()
                if name:
                    rows.append({"nom": name, "email": email})
    except Exception as e:
        logging.error(f"Erreur lors de la lecture du CSV: {e}")
    
    return rows
