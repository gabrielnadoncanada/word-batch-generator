# -*- coding: utf-8 -*-
"""
Tests unitaires pour le générateur de documents
"""
import unittest
import tempfile
from pathlib import Path
from unittest.mock import Mock, patch, MagicMock
from document_generator import DocumentGenerator


class TestDocumentGenerator(unittest.TestCase):
    """Tests pour la classe DocumentGenerator."""
    
    def setUp(self):
        """Configuration des tests."""
        # Créer un fichier template temporaire
        self.temp_dir = tempfile.mkdtemp()
        self.template_path = Path(self.temp_dir) / "test_template.docx"
        self.template_path.touch()  # Créer un fichier vide
        
        self.generator = DocumentGenerator(self.template_path)
    
    def tearDown(self):
        """Nettoyage après les tests."""
        # Nettoyer le répertoire temporaire
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_init_with_nonexistent_template(self):
        """Test d'initialisation avec un template inexistant."""
        with self.assertRaises(FileNotFoundError):
            DocumentGenerator(Path("nonexistent.docx"))
    
    def test_replace_placeholder_in_paragraph(self):
        """Test de remplacement de placeholder dans un paragraphe."""
        # Mock d'un paragraphe
        paragraph = Mock()
        paragraph.text = "Hello {{VENDEUR}}"
        paragraph.runs = [Mock()]
        paragraph.runs[0].text = "Hello {{VENDEUR}}"
        
        self.generator.replace_placeholder_in_paragraph(paragraph, "{{VENDEUR}}", "John Doe")
        
        self.assertEqual(paragraph.runs[0].text, "Hello John Doe")
    
    def test_force_replace_across_runs(self):
        """Test de remplacement forcé sur plusieurs runs."""
        # Mock d'un paragraphe avec plusieurs runs
        paragraph = Mock()
        paragraph.text = "Hello {{VENDEUR}} world"
        paragraph.runs = [Mock(), Mock()]
        paragraph.runs[0].text = "Hello {{VENDEUR}}"
        paragraph.runs[1].text = " world"
        
        # Mock des méthodes clear et add_run
        paragraph.clear = Mock()
        paragraph.add_run = Mock()
        
        result = self.generator.force_replace_across_runs(paragraph, "{{VENDEUR}}", "John Doe")
        
        self.assertTrue(result)
        paragraph.clear.assert_called_once()
        paragraph.add_run.assert_called_once_with("Hello John Doe world")
    
    @patch('document_generator.Document')
    def test_generate_document(self, mock_document_class):
        """Test de génération de document."""
        # Mock du document
        mock_doc = Mock()
        mock_document_class.return_value = mock_doc
        
        # Mock des paragraphes
        mock_paragraph = Mock()
        mock_paragraph.text = "Hello {{VENDEUR}}"
        mock_paragraph.runs = [Mock()]
        mock_paragraph.runs[0].text = "Hello {{VENDEUR}}"
        mock_doc.paragraphs = [mock_paragraph]
        
        # Mock des tableaux
        mock_doc.tables = []
        
        # Mock de la sauvegarde
        mock_doc.save = Mock()
        
        with patch('document_generator.OUT_DOCX_DIR', Path(self.temp_dir)):
            result = self.generator.generate_document("John Doe", 1)
        
        # Vérifications
        self.assertIsInstance(result, Path)
        mock_doc.save.assert_called_once()
        self.assertEqual(mock_paragraph.runs[0].text, "Hello John Doe")


if __name__ == "__main__":
    unittest.main()

