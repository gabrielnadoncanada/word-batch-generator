# -*- coding: utf-8 -*-
"""
Interface graphique moderne pour le générateur de documents Word
"""
import customtkinter as ctk
from pathlib import Path
from typing import Optional, Callable, List, Dict, Any
import threading
from datetime import datetime

from config import (
    TEMPLATE, CSV_FILE, OUT_DOCX_DIR, OUT_PDF_DIR,
    PLACEHOLDER, SEND_EMAIL, FROM_ACCOUNT, SUBJECT_TEMPLATE,
    USE_EMAIL_TEMPLATE, USE_SYSTEM_SIGNATURE, USE_PROJECT_SIGNATURE
)


class ApplicationState:
    """Gestionnaire d'état de l'application."""

    def __init__(self):
        self.template_path: Path = TEMPLATE
        self.csv_path: Path = CSV_FILE
        self.output_docx_dir: Path = OUT_DOCX_DIR
        self.output_pdf_dir: Path = OUT_PDF_DIR
        self.placeholder: str = PLACEHOLDER
        self.send_email: bool = SEND_EMAIL
        self.from_account: str = FROM_ACCOUNT or ""
        self.subject: str = SUBJECT_TEMPLATE
        self.is_processing: bool = False
        self.total_rows: int = 0
        self.current_progress: int = 0
        self.logs: List[str] = []

    def reset_progress(self):
        """Réinitialise le progrès."""
        self.current_progress = 0
        self.total_rows = 0

    def add_log(self, message: str, level: str = "INFO"):
        """Ajoute un message aux logs."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] [{level}] {message}"
        self.logs.append(log_entry)
        return log_entry


class DocumentGeneratorGUI(ctk.CTk):
    """Interface graphique principale."""

    def __init__(self):
        super().__init__()

        # Configuration de base
        self.title("Générateur de Documents Word - Dilamco")
        self.geometry("1200x800")

        # Thème moderne
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # État de l'application
        self.app_state = ApplicationState()

        # Initialisation de l'interface
        self._create_widgets()

        # Callback pour le contrôleur (sera défini par gui_controller)
        self.on_generate: Optional[Callable] = None

    def _create_widgets(self):
        """Crée tous les widgets de l'interface."""

        # Grid configuration
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header
        self._create_header()

        # Main content area
        main_frame = ctk.CTkFrame(self, corner_radius=0)
        main_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)

        # Configuration panel
        self._create_config_panel(main_frame)

        # Progress and logs panel
        self._create_progress_panel(main_frame)

        # Footer avec boutons
        self._create_footer()

    def _create_header(self):
        """Crée l'en-tête de l'application."""
        header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color=("#1f6aa5", "#144870"))
        header_frame.grid(row=0, column=0, sticky="ew")

        title_label = ctk.CTkLabel(
            header_frame,
            text="🎯 Générateur de Documents Word",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color="white"
        )
        title_label.pack(pady=20, padx=20)

        subtitle_label = ctk.CTkLabel(
            header_frame,
            text="Génération automatisée de documents personnalisés et envoi d'emails",
            font=ctk.CTkFont(size=12),
            text_color=("#E0E0E0", "#B0B0B0")
        )
        subtitle_label.pack(pady=(0, 15), padx=20)

    def _create_config_panel(self, parent):
        """Crée le panneau de configuration."""
        config_frame = ctk.CTkScrollableFrame(parent, corner_radius=10)
        config_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 5))
        config_frame.grid_columnconfigure(1, weight=1)

        # Section: Fichiers
        self._create_section_label(config_frame, "📁 Fichiers", 0)

        self._create_file_input(config_frame, "Modèle Word:", self.app_state.template_path, 1,
                               lambda p: setattr(self.app_state, 'template_path', Path(p)),
                               [("Word Documents", "*.docx")])

        self._create_file_input(config_frame, "Fichier CSV:", self.app_state.csv_path, 2,
                               lambda p: setattr(self.app_state, 'csv_path', Path(p)),
                               [("CSV Files", "*.csv")])

        self._create_dir_input(config_frame, "Répertoire DOCX:", self.app_state.output_docx_dir, 3,
                              lambda p: setattr(self.app_state, 'output_docx_dir', Path(p)))

        self._create_dir_input(config_frame, "Répertoire PDF:", self.app_state.output_pdf_dir, 4,
                              lambda p: setattr(self.app_state, 'output_pdf_dir', Path(p)))

        # Section: Personnalisation
        self._create_section_label(config_frame, "✏️ Personnalisation", 5)

        self._create_text_input(config_frame, "Placeholder:", self.app_state.placeholder, 6,
                               lambda v: setattr(self.app_state, 'placeholder', v))

        # Section: Email
        self._create_section_label(config_frame, "📧 Configuration Email", 7)

        self.email_enabled_var = ctk.BooleanVar(value=self.app_state.send_email)
        email_checkbox = ctk.CTkCheckBox(
            config_frame,
            text="Envoyer les emails automatiquement",
            variable=self.email_enabled_var,
            command=lambda: setattr(self.app_state, 'send_email', self.email_enabled_var.get())
        )
        email_checkbox.grid(row=8, column=0, columnspan=2, sticky="w", padx=20, pady=5)

        self._create_text_input(config_frame, "Compte Email:", self.app_state.from_account, 9,
                               lambda v: setattr(self.app_state, 'from_account', v))

        self._create_text_input(config_frame, "Sujet:", self.app_state.subject, 10,
                               lambda v: setattr(self.app_state, 'subject', v))

    def _create_progress_panel(self, parent):
        """Crée le panneau de progression et logs."""
        progress_frame = ctk.CTkFrame(parent, corner_radius=10)
        progress_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(5, 10))
        progress_frame.grid_columnconfigure(0, weight=1)
        progress_frame.grid_rowconfigure(2, weight=1)

        # Titre
        title_label = ctk.CTkLabel(
            progress_frame,
            text="📊 Progression et Logs",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        title_label.grid(row=0, column=0, sticky="w", padx=20, pady=(15, 5))

        # Barre de progression
        self.progress_label = ctk.CTkLabel(
            progress_frame,
            text="Prêt à démarrer",
            font=ctk.CTkFont(size=12)
        )
        self.progress_label.grid(row=1, column=0, sticky="w", padx=20, pady=(5, 5))

        self.progress_bar = ctk.CTkProgressBar(progress_frame)
        self.progress_bar.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 10))
        self.progress_bar.set(0)

        # Zone de logs
        self.log_text = ctk.CTkTextbox(progress_frame, height=200, font=ctk.CTkFont(family="Consolas", size=11))
        self.log_text.grid(row=3, column=0, sticky="nsew", padx=20, pady=(5, 15))

    def _create_footer(self):
        """Crée le pied de page avec les boutons d'action."""
        footer_frame = ctk.CTkFrame(self, corner_radius=0)
        footer_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))

        # Conteneur pour centrer les boutons
        button_container = ctk.CTkFrame(footer_frame, fg_color="transparent")
        button_container.pack(expand=True)

        self.generate_button = ctk.CTkButton(
            button_container,
            text="🚀 Générer les Documents",
            command=self._on_generate_clicked,
            width=200,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("#2CC985", "#2FA572"),
            hover_color=("#28B374", "#268B5F")
        )
        self.generate_button.pack(side="left", padx=10)

        self.stop_button = ctk.CTkButton(
            button_container,
            text="⏹️ Arrêter",
            command=self._on_stop_clicked,
            width=120,
            height=40,
            font=ctk.CTkFont(size=14),
            fg_color=("#D32F2F", "#B71C1C"),
            hover_color=("#C62828", "#A11515"),
            state="disabled"
        )
        self.stop_button.pack(side="left", padx=10)

        clear_logs_button = ctk.CTkButton(
            button_container,
            text="🗑️ Effacer Logs",
            command=self._clear_logs,
            width=120,
            height=40,
            font=ctk.CTkFont(size=14)
        )
        clear_logs_button.pack(side="left", padx=10)

    def _create_section_label(self, parent, text: str, row: int):
        """Crée un label de section."""
        label = ctk.CTkLabel(
            parent,
            text=text,
            font=ctk.CTkFont(size=14, weight="bold"),
            anchor="w"
        )
        label.grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=(15, 5))

    def _create_file_input(self, parent, label_text: str, initial_value: Path, row: int,
                          on_change: Callable, file_types: List):
        """Crée un champ de sélection de fichier."""
        label = ctk.CTkLabel(parent, text=label_text, anchor="w", width=150)
        label.grid(row=row, column=0, sticky="w", padx=20, pady=5)

        entry_frame = ctk.CTkFrame(parent, fg_color="transparent")
        entry_frame.grid(row=row, column=1, sticky="ew", padx=(0, 20), pady=5)
        entry_frame.grid_columnconfigure(0, weight=1)

        entry = ctk.CTkEntry(entry_frame)
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        entry.insert(0, str(initial_value))

        def browse():
            from tkinter import filedialog
            path = filedialog.askopenfilename(filetypes=file_types)
            if path:
                entry.delete(0, "end")
                entry.insert(0, path)
                on_change(path)

        browse_btn = ctk.CTkButton(entry_frame, text="Parcourir", width=100, command=browse)
        browse_btn.grid(row=0, column=1)

    def _create_dir_input(self, parent, label_text: str, initial_value: Path, row: int,
                         on_change: Callable):
        """Crée un champ de sélection de répertoire."""
        label = ctk.CTkLabel(parent, text=label_text, anchor="w", width=150)
        label.grid(row=row, column=0, sticky="w", padx=20, pady=5)

        entry_frame = ctk.CTkFrame(parent, fg_color="transparent")
        entry_frame.grid(row=row, column=1, sticky="ew", padx=(0, 20), pady=5)
        entry_frame.grid_columnconfigure(0, weight=1)

        entry = ctk.CTkEntry(entry_frame)
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        entry.insert(0, str(initial_value))

        def browse():
            from tkinter import filedialog
            path = filedialog.askdirectory()
            if path:
                entry.delete(0, "end")
                entry.insert(0, path)
                on_change(path)

        browse_btn = ctk.CTkButton(entry_frame, text="Parcourir", width=100, command=browse)
        browse_btn.grid(row=0, column=1)

    def _create_text_input(self, parent, label_text: str, initial_value: str, row: int,
                          on_change: Callable):
        """Crée un champ de texte."""
        label = ctk.CTkLabel(parent, text=label_text, anchor="w", width=150)
        label.grid(row=row, column=0, sticky="w", padx=20, pady=5)

        entry = ctk.CTkEntry(parent)
        entry.grid(row=row, column=1, sticky="ew", padx=(0, 20), pady=5)
        entry.insert(0, str(initial_value))
        entry.bind("<KeyRelease>", lambda e: on_change(entry.get()))

    def _on_generate_clicked(self):
        """Gère le clic sur le bouton Générer."""
        if self.on_generate:
            self.on_generate()

    def _on_stop_clicked(self):
        """Gère le clic sur le bouton Arrêter."""
        self.add_log("Arrêt demandé par l'utilisateur...", "WARNING")
        # TODO: Implémenter la logique d'arrêt

    def _clear_logs(self):
        """Efface les logs."""
        self.log_text.delete("1.0", "end")
        self.app_state.logs.clear()

    def set_processing_state(self, is_processing: bool):
        """Définit l'état de traitement."""
        self.app_state.is_processing = is_processing

        if is_processing:
            self.generate_button.configure(state="disabled")
            self.stop_button.configure(state="normal")
        else:
            self.generate_button.configure(state="normal")
            self.stop_button.configure(state="disabled")

    def update_progress(self, current: int, total: int, message: str = ""):
        """Met à jour la barre de progression."""
        self.app_state.current_progress = current
        self.app_state.total_rows = total

        if total > 0:
            progress_value = current / total
            self.progress_bar.set(progress_value)

            status_text = f"Progression: {current}/{total}"
            if message:
                status_text += f" - {message}"

            self.progress_label.configure(text=status_text)
        else:
            self.progress_bar.set(0)
            self.progress_label.configure(text=message or "Prêt à démarrer")

    def add_log(self, message: str, level: str = "INFO"):
        """Ajoute un message aux logs avec coloration."""
        log_entry = self.app_state.add_log(message, level)

        # Définir les tags de couleur si ce n'est pas déjà fait
        if not hasattr(self, '_tags_configured'):
            self.log_text.tag_config("ERROR", foreground="#FF5555")
            self.log_text.tag_config("WARNING", foreground="#FFB86C")
            self.log_text.tag_config("INFO", foreground="#50FA7B")
            self.log_text.tag_config("DEBUG", foreground="#8BE9FD")
            self._tags_configured = True

        # Insérer avec la couleur appropriée
        start_index = self.log_text.index("end-1c")
        self.log_text.insert("end", log_entry + "\n")
        end_index = self.log_text.index("end-1c")

        # Appliquer le tag de couleur
        self.log_text.tag_add(level, start_index, end_index)
        self.log_text.see("end")

    def show_error(self, title: str, message: str):
        """Affiche une boîte de dialogue d'erreur."""
        from tkinter import messagebox
        messagebox.showerror(title, message)

    def show_success(self, title: str, message: str):
        """Affiche une boîte de dialogue de succès."""
        from tkinter import messagebox
        messagebox.showinfo(title, message)

    def show_warning(self, title: str, message: str):
        """Affiche une boîte de dialogue d'avertissement."""
        from tkinter import messagebox
        messagebox.showwarning(title, message)


def main():
    """Point d'entrée de l'application GUI."""
    app = DocumentGeneratorGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
