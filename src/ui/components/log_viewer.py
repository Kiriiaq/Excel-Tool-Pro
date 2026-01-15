"""
Visualiseur de logs en temps réel avec filtrage
"""

import customtkinter as ctk
from tkinter import ttk
import tkinter as tk
from typing import Optional, List
from datetime import datetime

from .tooltip import Tooltip
from ...core.constants import COLORS
from ...core.logger import LogEntry, LogLevel


class LogViewer(ctk.CTkFrame):
    """
    Widget d'affichage des logs en temps réel

    Fonctionnalités:
    - Affichage coloré par niveau
    - Filtrage par niveau
    - Recherche textuelle
    - Export des logs
    - Auto-scroll
    """

    def __init__(
        self,
        parent,
        max_entries: int = 1000,
        show_filters: bool = True,
        show_search: bool = True,
        height: int = 200,
        **kwargs
    ):
        super().__init__(parent, **kwargs)

        self.max_entries = max_entries
        self.entries: List[LogEntry] = []
        self.auto_scroll = True
        self.filter_level: Optional[LogLevel] = None
        self.search_term = ""

        self.configure(fg_color=COLORS["bg_card"], corner_radius=10)
        self._create_widgets(show_filters, show_search, height)

    def _create_widgets(self, show_filters: bool, show_search: bool, height: int):
        """Crée les widgets d'affichage"""
        # Header avec titre et contrôles
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=(10, 5))

        ctk.CTkLabel(
            header,
            text="📋 Journal des opérations",
            font=("Segoe UI", 12, "bold")
        ).pack(side="left")

        # Bouton de nettoyage
        clear_btn = ctk.CTkButton(
            header,
            text="🗑️",
            width=30,
            height=24,
            command=self.clear
        )
        clear_btn.pack(side="right", padx=2)
        Tooltip(clear_btn, "Effacer les logs")

        # Toggle auto-scroll
        self.auto_scroll_var = ctk.BooleanVar(value=True)
        auto_scroll_cb = ctk.CTkCheckBox(
            header,
            text="Auto-scroll",
            variable=self.auto_scroll_var,
            width=24,
            checkbox_height=16,
            checkbox_width=16,
            font=("Segoe UI", 10)
        )
        auto_scroll_cb.pack(side="right", padx=10)
        Tooltip(auto_scroll_cb, "Défiler automatiquement vers les nouveaux logs")

        # Barre de filtres (optionnel)
        if show_filters:
            filter_frame = ctk.CTkFrame(self, fg_color="transparent")
            filter_frame.pack(fill="x", padx=10, pady=(0, 5))

            ctk.CTkLabel(filter_frame, text="Niveau:", font=("Segoe UI", 10)).pack(side="left")

            self.level_combo = ctk.CTkComboBox(
                filter_frame,
                values=["Tous", "DEBUG", "INFO", "SUCCESS", "WARNING", "ERROR"],
                width=100,
                height=24,
                command=self._on_filter_change
            )
            self.level_combo.set("Tous")
            self.level_combo.pack(side="left", padx=(5, 10))

            # Recherche (optionnel)
            if show_search:
                ctk.CTkLabel(filter_frame, text="🔍", font=("Segoe UI", 10)).pack(side="left")

                self.search_entry = ctk.CTkEntry(
                    filter_frame,
                    placeholder_text="Rechercher...",
                    width=150,
                    height=24
                )
                self.search_entry.pack(side="left", padx=(5, 0))
                self.search_entry.bind("<KeyRelease>", self._on_search)
        else:
            self.level_combo = None
            self.search_entry = None

        # Zone de texte avec scrollbar
        text_frame = ctk.CTkFrame(self, fg_color="transparent")
        text_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Utiliser un widget Text Tkinter pour plus de contrôle
        self.log_text = tk.Text(
            text_frame,
            wrap="word",
            font=("Consolas", 10),
            bg="#1e1e2e",
            fg="#ffffff",
            insertbackground="#ffffff",
            selectbackground="#3d3d6a",
            height=height // 20,  # Approximation
            state="disabled",
            borderwidth=0,
            highlightthickness=0
        )
        self.log_text.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # Configurer les tags de couleur
        self._setup_tags()

    def _setup_tags(self):
        """Configure les tags de couleur pour les différents niveaux"""
        self.log_text.tag_configure("DEBUG", foreground=LogLevel.DEBUG.color)
        self.log_text.tag_configure("INFO", foreground=LogLevel.INFO.color)
        self.log_text.tag_configure("SUCCESS", foreground=LogLevel.SUCCESS.color)
        self.log_text.tag_configure("WARNING", foreground=LogLevel.WARNING.color)
        self.log_text.tag_configure("ERROR", foreground=LogLevel.ERROR.color)
        self.log_text.tag_configure("CRITICAL", foreground=LogLevel.CRITICAL.color)
        self.log_text.tag_configure("timestamp", foreground="#6c757d")

    def add_entry(self, entry: LogEntry):
        """Ajoute une entrée de log"""
        self.entries.append(entry)

        # Limiter le nombre d'entrées
        if len(self.entries) > self.max_entries:
            self.entries = self.entries[-self.max_entries:]

        # Appliquer les filtres
        if self._should_display(entry):
            self._display_entry(entry)

    def _should_display(self, entry: LogEntry) -> bool:
        """Vérifie si une entrée doit être affichée selon les filtres"""
        # Filtre par niveau
        if self.filter_level and entry.level != self.filter_level:
            return False

        # Filtre par recherche
        if self.search_term:
            if self.search_term.lower() not in entry.message.lower():
                return False

        return True

    def _display_entry(self, entry: LogEntry):
        """Affiche une entrée dans le widget"""
        self.log_text.configure(state="normal")

        # Formater le message
        timestamp = f"[{entry.timestamp.strftime('%H:%M:%S')}] "
        level_str = f"[{entry.level.name_str}] "
        message = f"{entry.message}\n"

        # Insérer avec les tags appropriés
        self.log_text.insert("end", timestamp, "timestamp")
        self.log_text.insert("end", level_str, entry.level.name_str)
        self.log_text.insert("end", message, entry.level.name_str)

        self.log_text.configure(state="disabled")

        # Auto-scroll si activé
        if self.auto_scroll_var.get():
            self.log_text.see("end")

    def _on_filter_change(self, selection: str):
        """Callback lors du changement de filtre"""
        if selection == "Tous":
            self.filter_level = None
        else:
            self.filter_level = LogLevel[selection]

        self._refresh_display()

    def _on_search(self, event=None):
        """Callback lors de la recherche"""
        if self.search_entry:
            self.search_term = self.search_entry.get()
            self._refresh_display()

    def _refresh_display(self):
        """Rafraîchit l'affichage avec les filtres actuels"""
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        for entry in self.entries:
            if self._should_display(entry):
                self._display_entry(entry)

    def clear(self):
        """Efface tous les logs"""
        self.entries.clear()
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def log(self, message: str, level: LogLevel = LogLevel.INFO, source: str = ""):
        """Méthode raccourci pour ajouter un log"""
        entry = LogEntry(
            timestamp=datetime.now(),
            level=level,
            message=message,
            source=source
        )
        self.add_entry(entry)

    def info(self, message: str, source: str = ""):
        self.log(message, LogLevel.INFO, source)

    def success(self, message: str, source: str = ""):
        self.log(message, LogLevel.SUCCESS, source)

    def warning(self, message: str, source: str = ""):
        self.log(message, LogLevel.WARNING, source)

    def error(self, message: str, source: str = ""):
        self.log(message, LogLevel.ERROR, source)

    def get_text(self) -> str:
        """Retourne tout le texte des logs"""
        return self.log_text.get("1.0", "end-1c")

    def export_to_file(self, filepath: str) -> bool:
        """Exporte les logs vers un fichier"""
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                for entry in self.entries:
                    f.write(entry.format() + "\n")
            return True
        except Exception:
            return False
