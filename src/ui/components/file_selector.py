"""
Composant de sélection de fichier avec prévisualisation
"""

import customtkinter as ctk
from tkinter import filedialog
from pathlib import Path
from typing import Optional, List, Callable
import pandas as pd

from .tooltip import Tooltip
from ...core.constants import COLORS


class FileSelector(ctk.CTkFrame):
    """
    Widget de sélection de fichier Excel avec chargement d'onglets

    Fonctionnalités:
    - Sélection de fichier avec dialogue
    - Chargement automatique des onglets
    - Callback lors du chargement
    - Affichage des informations (lignes, colonnes)
    """

    def __init__(
        self,
        parent,
        label: str,
        tooltip: str = "",
        filetypes: Optional[List[tuple]] = None,
        on_file_loaded: Optional[Callable] = None,
        show_sheet_selector: bool = True,
        **kwargs
    ):
        super().__init__(parent, **kwargs)

        self.filepath: Optional[str] = None
        self.df: Optional[pd.DataFrame] = None
        self.sheets: List[str] = []
        self.current_sheet: Optional[str] = None
        self.on_file_loaded = on_file_loaded
        self.show_sheet_selector = show_sheet_selector

        self.filetypes = filetypes or [
            ("Fichiers Excel", "*.xlsx *.xls *.xlsm"),
            ("Tous les fichiers", "*.*")
        ]

        self.configure(fg_color="transparent")
        self._create_widgets(label, tooltip)

    def _create_widgets(self, label: str, tooltip: str):
        """Crée les widgets du sélecteur"""
        # Label principal
        self.label = ctk.CTkLabel(
            self,
            text=label,
            font=("Segoe UI", 12, "bold")
        )
        self.label.pack(anchor="w", pady=(0, 5))
        if tooltip:
            Tooltip(self.label, tooltip)

        # Frame pour chemin + bouton
        path_frame = ctk.CTkFrame(self, fg_color="transparent")
        path_frame.pack(fill="x", pady=(0, 5))

        self.path_entry = ctk.CTkEntry(
            path_frame,
            placeholder_text="Aucun fichier sélectionné...",
            state="disabled",
            width=350
        )
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.browse_btn = ctk.CTkButton(
            path_frame,
            text="📁 Parcourir",
            width=120,
            command=self.browse_file
        )
        self.browse_btn.pack(side="right")
        Tooltip(self.browse_btn, "Sélectionner un fichier")

        # Sélecteur d'onglet (optionnel)
        if self.show_sheet_selector:
            sheet_frame = ctk.CTkFrame(self, fg_color="transparent")
            sheet_frame.pack(fill="x", pady=(0, 5))

            ctk.CTkLabel(sheet_frame, text="Onglet:", width=60).pack(side="left")

            self.sheet_combo = ctk.CTkComboBox(
                sheet_frame,
                values=["(Charger un fichier)"],
                state="disabled",
                width=250,
                command=self._on_sheet_change
            )
            self.sheet_combo.pack(side="left", padx=(10, 0))
            Tooltip(self.sheet_combo, "Sélectionnez l'onglet à utiliser")
        else:
            self.sheet_combo = None

        # Label d'information
        self.info_label = ctk.CTkLabel(
            self,
            text="",
            font=("Segoe UI", 10),
            text_color=COLORS["text_muted"]
        )
        self.info_label.pack(anchor="w")

    def browse_file(self):
        """Ouvre le dialogue de sélection de fichier"""
        filepath = filedialog.askopenfilename(
            title="Sélectionner un fichier",
            filetypes=self.filetypes
        )
        if filepath:
            self.load_file(filepath)

    def load_file(self, filepath: str) -> bool:
        """
        Charge un fichier

        Returns:
            True si le chargement a réussi
        """
        try:
            self.filepath = filepath

            # Charger les noms d'onglets
            xl = pd.ExcelFile(filepath)
            self.sheets = xl.sheet_names

            # Mettre à jour l'interface
            self.path_entry.configure(state="normal")
            self.path_entry.delete(0, "end")
            self.path_entry.insert(0, Path(filepath).name)
            self.path_entry.configure(state="disabled")

            if self.sheet_combo:
                self.sheet_combo.configure(state="normal", values=self.sheets)
                self.sheet_combo.set(self.sheets[0])

            # Charger le premier onglet
            self._load_sheet(self.sheets[0])
            return True

        except Exception as e:
            self._set_error(f"Erreur: {str(e)}")
            return False

    def _load_sheet(self, sheet_name: str):
        """Charge les données d'un onglet"""
        try:
            self.df = pd.read_excel(self.filepath, sheet_name=sheet_name, dtype=str)
            self.df.columns = self.df.columns.str.strip()
            self.current_sheet = sheet_name

            self.info_label.configure(
                text=f"✓ {len(self.df)} lignes, {len(self.df.columns)} colonnes",
                text_color=COLORS["success"]
            )

            if self.on_file_loaded:
                self.on_file_loaded(self.df)

        except Exception as e:
            self._set_error(f"Erreur: {str(e)}")

    def _on_sheet_change(self, sheet_name: str):
        """Callback lors du changement d'onglet"""
        self._load_sheet(sheet_name)

    def _set_error(self, message: str):
        """Affiche un message d'erreur"""
        self.info_label.configure(
            text=f"✗ {message}",
            text_color=COLORS["error"]
        )

    def get_columns(self) -> List[str]:
        """Retourne la liste des colonnes"""
        if self.df is not None:
            return list(self.df.columns)
        return []

    def get_dataframe(self) -> Optional[pd.DataFrame]:
        """Retourne le DataFrame chargé"""
        return self.df

    def get_filepath(self) -> Optional[str]:
        """Retourne le chemin du fichier"""
        return self.filepath

    def get_current_sheet(self) -> Optional[str]:
        """Retourne le nom de l'onglet actuel"""
        return self.current_sheet

    def reset(self):
        """Réinitialise le sélecteur"""
        self.filepath = None
        self.df = None
        self.sheets = []
        self.current_sheet = None

        self.path_entry.configure(state="normal")
        self.path_entry.delete(0, "end")
        self.path_entry.configure(state="disabled")

        if self.sheet_combo:
            self.sheet_combo.configure(state="disabled", values=["(Charger un fichier)"])
            self.sheet_combo.set("(Charger un fichier)")

        self.info_label.configure(text="", text_color=COLORS["text_muted"])

    def is_loaded(self) -> bool:
        """Vérifie si un fichier est chargé"""
        return self.df is not None
