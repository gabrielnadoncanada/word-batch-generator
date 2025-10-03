# -*- coding: utf-8 -*-
"""
Point d'entrée principal du générateur de documents Word
Support du mode CLI et GUI
"""
import sys
from pathlib import Path
from typing import List, Dict, Any, Tuple

from config import TEMPLATE, CSV_FILE, OUT_DOCX_DIR, OUT_PDF_DIR
from logger_config import setup_logging
from file_utils import read_csv_rows
from document_generator import DocumentGenerator
from email_sender import EmailSender
from validators import DataValidator


class WordBatchGenerator:
    """Classe principale pour orchestrer la génération de documents et l'envoi d'emails."""

    def __init__(self):
        self.logger = setup_logging()
        self.document_generator = None
        self.email_sender = EmailSender()

    def validate_environment(self) -> bool:
        """Valide l'environnement et les fichiers requis."""
        self.logger.info("Validation de l'environnement...")

        # Valider le fichier template
        is_valid, error = DataValidator.validate_template_file(TEMPLATE)
        if not is_valid:
            self.logger.error(f"[ERREUR] {error}")
            return False

        # Valider le fichier CSV
        if not CSV_FILE.exists():
            self.logger.error(f"[ERREUR] Fichier CSV introuvable: {CSV_FILE}")
            return False

        # Créer les répertoires de sortie
        try:
            OUT_DOCX_DIR.mkdir(parents=True, exist_ok=True)
            OUT_PDF_DIR.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            self.logger.error(f"[ERREUR] Impossible de créer les répertoires: {e}")
            return False

        return True

    def load_data(self) -> List[Dict[str, Any]]:
        """Charge et valide les données CSV."""
        self.logger.info("Chargement des données...")

        rows = read_csv_rows(CSV_FILE)
        if not rows:
            self.logger.error("[ERREUR] Aucune ligne valide dans le CSV (colonne 'nom' requise).")
            return []

        # Valider les données
        is_valid, errors = DataValidator.validate_csv_data(rows)
        if not is_valid:
            for error in errors:
                self.logger.warning(f"[WARN] {error}")

        self.logger.info(f"Données chargées: {len(rows)} entrée(s)")
        return rows

    def generate_documents(self, rows: List[Dict[str, Any]]) -> Tuple[List[Path], List[Path]]:
        """Génère les documents Word et PDF."""
        self.logger.info("Génération des documents...")

        try:
            self.document_generator = DocumentGenerator(TEMPLATE)
            docx_files, pdf_files = self.document_generator.generate_documents_batch(rows)

            self.logger.info(f"[OK] {len(docx_files)} DOCX générés -> {OUT_DOCX_DIR}")
            self.logger.info(f"[OK] {len(pdf_files)} PDF générés -> {OUT_PDF_DIR}")

            return docx_files, pdf_files

        except Exception as e:
            self.logger.error(f"[ERREUR] Génération des documents: {e}")
            if "docx2pdf" in str(e).lower():
                self.logger.error("docx2pdf nécessite Microsoft Word installé sous Windows.")
                self.logger.error("Alternative: LibreOffice headless -> voir README.md.")
            raise

    def send_emails(self, rows: List[Dict[str, Any]], pdf_files: List[Path]) -> int:
        """Envoie les emails avec les PDF en pièce jointe."""
        self.logger.info("Envoi des emails...")

        try:
            sent_count = self.email_sender.send_emails_batch(rows, pdf_files)
            return sent_count
        except Exception as e:
            self.logger.error(f"[ERREUR] Envoi des emails: {e}")
            raise

    def run(self) -> int:
        """Exécute le processus complet."""
        try:
            # Validation de l'environnement
            if not self.validate_environment():
                return 1

            # Chargement des données
            rows = self.load_data()
            if not rows:
                return 1

            # Génération des documents
            docx_files, pdf_files = self.generate_documents(rows)

            # Envoi des emails
            sent_count = self.send_emails(rows, pdf_files)

            # Résumé final
            self.logger.info("=== RÉSUMÉ ===")
            self.logger.info(f"Documents générés: {len(docx_files)} DOCX, {len(pdf_files)} PDF")
            self.logger.info(f"Emails envoyés: {sent_count}/{len(rows)}")

            return 0

        except Exception as e:
            self.logger.error(f"[ERREUR] Processus interrompu: {e}")
            return 1


def main() -> int:
    """Point d'entrée principal - détecte le mode CLI ou GUI."""
    # Vérifier si l'utilisateur demande le mode GUI
    if len(sys.argv) > 1 and sys.argv[1] == "--gui":
        # Lancer l'interface graphique
        try:
            from gui_controller import main as gui_main
            gui_main()
            return 0
        except ImportError:
            print("ERREUR: CustomTkinter n'est pas installé.")
            print("Installez-le avec: pip install customtkinter")
            return 1
    else:
        # Mode CLI par défaut
        generator = WordBatchGenerator()
        return generator.run()


if __name__ == "__main__":
    sys.exit(main())
