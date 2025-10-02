# -*- coding: utf-8 -*-
"""
Validateurs pour les données d'entrée
"""
import re
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple


class DataValidator:
    """Classe pour valider les données d'entrée."""
    
    @staticmethod
    def validate_email(email: str) -> bool:
        """Valide le format d'un email."""
        if not email:
            return False
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))
    
    @staticmethod
    def validate_name(name: str) -> bool:
        """Valide qu'un nom n'est pas vide."""
        return bool(name and name.strip())
    
    @staticmethod
    def validate_file_exists(file_path: Path) -> bool:
        """Valide qu'un fichier existe."""
        return file_path.exists()
    
    @staticmethod
    def validate_csv_data(rows: List[Dict[str, Any]]) -> Tuple[bool, List[str]]:
        """Valide les données CSV et retourne les erreurs."""
        errors = []
        
        if not rows:
            errors.append("Aucune donnée trouvée dans le CSV")
            return False, errors
        
        for i, row in enumerate(rows, 1):
            name = row.get('nom', '').strip()
            email = row.get('email', '').strip()
            
            if not DataValidator.validate_name(name):
                errors.append(f"Ligne {i}: nom manquant ou invalide")
            
            if email and not DataValidator.validate_email(email):
                errors.append(f"Ligne {i}: email invalide '{email}'")
        
        return len(errors) == 0, errors
    
    @staticmethod
    def validate_template_file(template_path: Path) -> Tuple[bool, str]:
        """Valide qu'un fichier template existe et est valide."""
        if not template_path.exists():
            return False, f"Fichier template introuvable: {template_path}"
        
        if not template_path.suffix.lower() == '.docx':
            return False, f"Le fichier template doit être un .docx: {template_path}"
        
        return True, ""
    
