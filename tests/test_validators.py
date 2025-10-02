# -*- coding: utf-8 -*-
"""
Tests unitaires pour les validateurs
"""
import unittest
from pathlib import Path
from validators import DataValidator


class TestDataValidator(unittest.TestCase):
    """Tests pour la classe DataValidator."""
    
    def test_validate_email(self):
        """Test de validation des emails."""
        # Emails valides
        self.assertTrue(DataValidator.validate_email("test@example.com"))
        self.assertTrue(DataValidator.validate_email("user.name@domain.co.uk"))
        self.assertTrue(DataValidator.validate_email("test+tag@example.org"))
        
        # Emails invalides
        self.assertFalse(DataValidator.validate_email(""))
        self.assertFalse(DataValidator.validate_email("invalid-email"))
        self.assertFalse(DataValidator.validate_email("@example.com"))
        self.assertFalse(DataValidator.validate_email("test@"))
        self.assertFalse(DataValidator.validate_email("test@.com"))
    
    def test_validate_name(self):
        """Test de validation des noms."""
        # Noms valides
        self.assertTrue(DataValidator.validate_name("John Doe"))
        self.assertTrue(DataValidator.validate_name("  Jane Smith  "))
        self.assertTrue(DataValidator.validate_name("A"))
        
        # Noms invalides
        self.assertFalse(DataValidator.validate_name(""))
        self.assertFalse(DataValidator.validate_name("   "))
        self.assertFalse(DataValidator.validate_name(None))
    
    def test_validate_file_exists(self):
        """Test de validation d'existence de fichier."""
        # Fichier existant
        self.assertTrue(DataValidator.validate_file_exists(Path(__file__)))
        
        # Fichier inexistant
        self.assertFalse(DataValidator.validate_file_exists(Path("nonexistent.txt")))
    
    def test_validate_csv_data(self):
        """Test de validation des données CSV."""
        # Données valides
        valid_data = [
            {"nom": "John Doe", "email": "john@example.com"},
            {"nom": "Jane Smith", "email": "jane@example.com"},
            {"nom": "Bob Wilson", "email": ""}  # Email vide est OK
        ]
        is_valid, errors = DataValidator.validate_csv_data(valid_data)
        self.assertTrue(is_valid)
        self.assertEqual(len(errors), 0)
        
        # Données invalides
        invalid_data = [
            {"nom": "", "email": "john@example.com"},  # Nom vide
            {"nom": "Jane Smith", "email": "invalid-email"},  # Email invalide
            {"nom": "Bob Wilson", "email": "bob@example.com"}
        ]
        is_valid, errors = DataValidator.validate_csv_data(invalid_data)
        self.assertFalse(is_valid)
        self.assertGreater(len(errors), 0)
        
        # Données vides
        is_valid, errors = DataValidator.validate_csv_data([])
        self.assertFalse(is_valid)
        self.assertIn("Aucune donnée trouvée", errors[0])


if __name__ == "__main__":
    unittest.main()

