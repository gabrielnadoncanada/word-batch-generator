# -*- coding: utf-8 -*-
"""
Contrôleur pour l'interface graphique - gère la logique métier
"""
import threading
from pathlib import Path
from typing import Optional

from gui import DocumentGeneratorGUI
from file_utils import read_csv_rows
from document_generator import DocumentGenerator
from email_sender import EmailSender
from validators import DataValidator


class GUIController:
    """Contrôleur pour gérer la logique métier de l'interface graphique."""

    def __init__(self, gui: DocumentGeneratorGUI):
        self.gui = gui
        self.gui.on_generate = self.start_generation
        self._stop_requested = False
        self._worker_thread: Optional[threading.Thread] = None

    def start_generation(self):
        """Démarre le processus de génération dans un thread séparé."""
        if self._worker_thread and self._worker_thread.is_alive():
            self.gui.show_warning("Traitement en cours", "Un traitement est déjà en cours.")
            return

        # Réinitialiser l'état
        self._stop_requested = False
        self.gui.app_state.reset_progress()

        # Validation initiale
        if not self._validate_inputs():
            return

        # Démarrer le thread de travail
        self._worker_thread = threading.Thread(target=self._run_generation, daemon=True)
        self._worker_thread.start()

    def _validate_inputs(self) -> bool:
        """Valide les entrées utilisateur."""
        state = self.gui.app_state

        # Valider le fichier template
        is_valid, error = DataValidator.validate_template_file(state.template_path)
        if not is_valid:
            self.gui.show_error("Erreur de validation", error)
            return False

        # Valider le fichier CSV
        if not state.csv_path.exists():
            self.gui.show_error("Erreur de validation", f"Fichier CSV introuvable: {state.csv_path}")
            return False

        # Valider le placeholder
        if not state.placeholder.strip():
            self.gui.show_error("Erreur de validation", "Le placeholder ne peut pas être vide.")
            return False

        # Valider les paramètres email si activé
        if state.send_email:
            if not state.subject.strip():
                self.gui.show_error("Erreur de validation", "Le sujet de l'email est requis.")
                return False

        return True

    def _run_generation(self):
        """Exécute le processus de génération (dans un thread séparé)."""

        try:
            self.gui.set_processing_state(True)
            self.gui.add_log("🚀 Démarrage de la génération...", "INFO")

            # Étape 1: Créer les répertoires de sortie
            self._create_output_directories()

            # Étape 2: Charger les données CSV
            rows = self._load_csv_data()
            if not rows:
                return

            # Étape 3: Générer les documents
            docx_files, pdf_files = self._generate_documents(rows)
            if not docx_files:
                return

            # Étape 4: Envoyer les emails (si activé)
            sent_count = 0
            if self.gui.app_state.send_email:
                sent_count = self._send_emails(rows, pdf_files)

            # Résumé final
            self._show_summary(len(docx_files), len(pdf_files), sent_count, len(rows))

        except Exception as e:
            self.gui.add_log(f"❌ Erreur fatale: {e}", "ERROR")
            self.gui.show_error("Erreur", f"Une erreur est survenue:\n{str(e)}")

        finally:
            self.gui.set_processing_state(False)
            self.gui.update_progress(0, 0, "Terminé")

    def _create_output_directories(self):
        """Crée les répertoires de sortie."""
        try:
            self.gui.app_state.output_docx_dir.mkdir(parents=True, exist_ok=True)
            self.gui.app_state.output_pdf_dir.mkdir(parents=True, exist_ok=True)
            self.gui.add_log("✅ Répertoires de sortie créés", "INFO")
        except Exception as e:
            raise Exception(f"Impossible de créer les répertoires de sortie: {e}")

    def _load_csv_data(self):
        """Charge et valide les données CSV."""
        self.gui.add_log("📂 Chargement des données CSV...", "INFO")

        try:
            rows = read_csv_rows(self.gui.app_state.csv_path)

            if not rows:
                self.gui.add_log("❌ Aucune donnée valide dans le CSV", "ERROR")
                self.gui.show_error("Erreur", "Aucune ligne valide dans le CSV (colonne 'nom' requise).")
                return []

            # Valider les données
            is_valid, errors = DataValidator.validate_csv_data(rows)
            if not is_valid:
                for error in errors:
                    self.gui.add_log(f"⚠️ {error}", "WARNING")

            self.gui.add_log(f"✅ {len(rows)} ligne(s) chargée(s)", "INFO")
            return rows

        except Exception as e:
            self.gui.add_log(f"❌ Erreur lors du chargement du CSV: {e}", "ERROR")
            self.gui.show_error("Erreur", f"Erreur lors du chargement du CSV:\n{str(e)}")
            return []

    def _generate_documents(self, rows):
        """Génère les documents Word et PDF."""
        self.gui.add_log("📝 Génération des documents...", "INFO")

        try:
            generator = DocumentGenerator(self.gui.app_state.template_path)
            docx_files = []
            pdf_files = []

            total = len(rows)
            for i, row in enumerate(rows):
                if self._stop_requested:
                    self.gui.add_log("⏹️ Génération arrêtée par l'utilisateur", "WARNING")
                    break

                name = row.get('nom', 'inconnu')

                # Mettre à jour la progression
                self.gui.update_progress(i, total, f"Génération: {name}")

                try:
                    # Générer le document Word
                    docx_path = generator.generate_document(name, i + 1)
                    docx_files.append(docx_path)

                    # Convertir en PDF
                    pdf_path = generator.convert_to_pdf(docx_path, row.get('email', ''))
                    pdf_files.append(pdf_path)

                    self.gui.add_log(f"✅ Document généré: {name}", "INFO")

                except Exception as e:
                    self.gui.add_log(f"❌ Erreur pour {name}: {e}", "ERROR")
                    continue

            self.gui.update_progress(total, total, "Génération terminée")
            self.gui.add_log(f"✅ {len(docx_files)} DOCX et {len(pdf_files)} PDF générés", "INFO")

            return docx_files, pdf_files

        except Exception as e:
            self.gui.add_log(f"❌ Erreur lors de la génération: {e}", "ERROR")
            self.gui.show_error("Erreur", f"Erreur lors de la génération:\n{str(e)}")
            return [], []

    def _send_emails(self, rows, pdf_files):
        """Envoie les emails avec les PDF en pièce jointe."""
        self.gui.add_log("📧 Envoi des emails...", "INFO")

        try:
            email_sender = EmailSender(enabled=True)
            self.gui.add_log("Envoi des emails...", "INFO")
            sent_count = email_sender.send_emails_batch(rows, pdf_files)
            self.gui.add_log(f"✅ {sent_count}/{len(rows)} emails envoyés", "INFO")
            return sent_count

        except Exception as e:
            self.gui.add_log(f"❌ Erreur lors de l'envoi des emails: {e}", "ERROR")
            return 0

    def _show_summary(self, docx_count: int, pdf_count: int, email_count: int, total_rows: int):
        """Affiche le résumé final."""
        summary = (
            f"🎉 Traitement terminé!\n\n"
            f"Documents générés:\n"
            f"  • {docx_count} fichiers DOCX\n"
            f"  • {pdf_count} fichiers PDF\n\n"
            f"Emails envoyés: {email_count}/{total_rows}"
        )

        self.gui.add_log("=" * 50, "INFO")
        self.gui.add_log("📊 RÉSUMÉ", "INFO")
        self.gui.add_log(f"Documents: {docx_count} DOCX, {pdf_count} PDF", "INFO")
        self.gui.add_log(f"Emails: {email_count}/{total_rows}", "INFO")
        self.gui.add_log("=" * 50, "INFO")

        self.gui.show_success("Traitement terminé", summary)

    def request_stop(self):
        """Demande l'arrêt du traitement."""
        self._stop_requested = True


def main():
    """Point d'entrée de l'application avec contrôleur."""
    app = DocumentGeneratorGUI()
    controller = GUIController(app)
    app.mainloop()


if __name__ == "__main__":
    main()
