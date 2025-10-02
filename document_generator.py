# -*- coding: utf-8 -*-
"""
Générateur de documents Word à partir de modèles
"""
import logging
from pathlib import Path
from typing import List, Dict, Any, Tuple
from docx import Document
from docx2pdf import convert

from config import PLACEHOLDER, OUT_DOCX_DIR, OUT_PDF_DIR
from file_utils import safe_filename, safe_email_for_filename


class DocumentGenerator:
    """Classe pour générer des documents Word à partir de modèles."""
    
    def __init__(self, template_path: Path):
        self.template_path = template_path
        if not template_path.exists():
            raise FileNotFoundError(f"Modèle introuvable: {template_path}")
    
    def replace_placeholder_in_paragraph(self, paragraph, placeholder: str, replacement: str) -> None:
        """Remplace un placeholder dans un paragraphe."""
        if placeholder in paragraph.text:
            for run in paragraph.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, replacement)
    
    def force_replace_across_runs(self, paragraph, placeholder: str, replacement: str) -> bool:
        """Force le remplacement d'un placeholder qui s'étend sur plusieurs runs."""
        full = paragraph.text
        if placeholder in full:
            new_text = full.replace(placeholder, replacement)
            # Purger les runs et réécrire
            for _ in range(len(paragraph.runs)-1, -1, -1):
                try:
                    paragraph.runs[_].text = ""
                except Exception:
                    pass
            try:
                paragraph.clear()
            except Exception:
                # Fallback pour les versions plus anciennes
                while paragraph.runs:
                    paragraph.runs[-1].text = ""
            paragraph.add_run(new_text)
            return True
        return False
    
    def replace_in_tables(self, doc, placeholder: str, replacement: str) -> None:
        """Remplace les placeholders dans les tableaux."""
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if placeholder in paragraph.text:
                            self.replace_placeholder_in_paragraph(paragraph, placeholder, replacement) or \
                            self.force_replace_across_runs(paragraph, placeholder, replacement)
    
    def generate_document(self, name: str, index: int) -> Path:
        """Génère un document Word pour un nom donné."""
        doc = Document(str(self.template_path))
        
        # Remplacer dans les paragraphes
        for p in doc.paragraphs:
            if PLACEHOLDER in p.text:
                self.replace_placeholder_in_paragraph(p, PLACEHOLDER, name)
                if PLACEHOLDER in p.text:
                    self.force_replace_across_runs(p, PLACEHOLDER, name)
        
        # Remplacer dans les tableaux
        self.replace_in_tables(doc, PLACEHOLDER, name)
        
        # Sauvegarder
        out_path = OUT_DOCX_DIR / f"{index:02d}_{safe_filename(name)}.docx"
        doc.save(str(out_path))
        return out_path
    
    def convert_to_pdf(self, docx_path: Path, email: str = "") -> Path:
        """Convertit un document Word en PDF."""
        email_suffix = ("_" + safe_email_for_filename(email)) if email else ""
        pdf_path = OUT_PDF_DIR / (docx_path.stem + f"{email_suffix}.pdf")
        convert(str(docx_path), str(pdf_path))
        return pdf_path
    
    def generate_documents_batch(self, rows: List[Dict[str, Any]]) -> Tuple[List[Path], List[Path]]:
        """Génère tous les documents pour une liste de données."""
        docx_files = []
        pdf_files = []
        
        for i, row in enumerate(rows):
            try:
                # Générer le document Word
                docx_path = self.generate_document(row['nom'], i + 1)
                docx_files.append(docx_path)
                
                # Convertir en PDF
                pdf_path = self.convert_to_pdf(docx_path, row.get('email', ''))
                pdf_files.append(pdf_path)
                
                logging.info(f"Document généré: {docx_path.name} -> {pdf_path.name}")
                
            except Exception as e:
                logging.error(f"Erreur lors de la génération du document pour {row.get('nom', 'inconnu')}: {e}")
                raise
        
        return docx_files, pdf_files
