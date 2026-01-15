"""
Tableau de prévisualisation des données avec Treeview
"""

import customtkinter as ctk
from tkinter import ttk
import tkinter as tk
from typing import Optional, List
import pandas as pd

from ...core.constants import COLORS


class PreviewTable(ctk.CTkFrame):
    """
    Widget de prévisualisation des données avec Treeview

    Fonctionnalités:
    - Affichage tabulaire des données
    - Scrollbars horizontale et verticale
    - Tri par colonne
    - Style sombre moderne
    - Limitation du nombre de lignes affichées
    """

    def __init__(
        self,
        parent,
        title: str = "Aperçu des données",
        max_rows: int = 100,
        max_columns: int = 20,
        **kwargs
    ):
        super().__init__(parent, **kwargs)

        self.title = title
        self.max_rows = max_rows
        self.max_columns = max_columns
        self.df: Optional[pd.DataFrame] = None

        self.configure(fg_color=COLORS["bg_card"], corner_radius=10)
        self._create_widgets()

    def _create_widgets(self):
        """Crée les widgets du tableau"""
        # Titre
        title_label = ctk.CTkLabel(
            self,
            text=self.title,
            font=("Segoe UI", 12, "bold")
        )
        title_label.pack(anchor="w", padx=15, pady=(15, 10))

        # Info label
        self.info_label = ctk.CTkLabel(
            self,
            text="",
            font=("Segoe UI", 10),
            text_color=COLORS["text_muted"]
        )
        self.info_label.pack(anchor="w", padx=15, pady=(0, 5))

        # Frame pour le Treeview
        tree_frame = ctk.CTkFrame(self, fg_color="transparent")
        tree_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Configurer le style du Treeview
        self._setup_style()

        # Créer le Treeview
        self.tree = ttk.Treeview(tree_frame, style="Custom.Treeview", show="headings")
        self.tree.pack(side="left", fill="both", expand=True)

        # Scrollbar verticale
        v_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        v_scrollbar.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=v_scrollbar.set)

        # Scrollbar horizontale
        h_scrollbar = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        h_scrollbar.pack(fill="x", padx=10, pady=(0, 10))
        self.tree.configure(xscrollcommand=h_scrollbar.set)

        # Bind pour le tri
        self.tree.bind("<Button-1>", self._on_header_click)

    def _setup_style(self):
        """Configure le style du Treeview"""
        style = ttk.Style()
        style.theme_use("clam")

        # Style du Treeview
        style.configure(
            "Custom.Treeview",
            background="#1e1e2e",
            foreground="#ffffff",
            fieldbackground="#1e1e2e",
            rowheight=25,
            font=("Segoe UI", 10)
        )

        # Style des en-têtes
        style.configure(
            "Custom.Treeview.Heading",
            background="#2d2d4a",
            foreground="#ffffff",
            font=("Segoe UI", 10, "bold"),
            padding=5
        )

        # Style de sélection
        style.map(
            "Custom.Treeview",
            background=[("selected", "#3d3d6a")],
            foreground=[("selected", "#ffffff")]
        )

    def load_data(self, df: pd.DataFrame, max_rows: Optional[int] = None):
        """
        Charge les données dans le tableau

        Args:
            df: DataFrame à afficher
            max_rows: Nombre max de lignes (None = utiliser self.max_rows)
        """
        self.df = df
        max_rows = max_rows or self.max_rows

        # Effacer les données existantes
        self.tree.delete(*self.tree.get_children())

        if df is None or df.empty:
            self.info_label.configure(text="Aucune donnée")
            return

        # Limiter les colonnes
        columns = list(df.columns)[:self.max_columns]
        self.tree["columns"] = columns

        # Configurer les colonnes
        for col in columns:
            self.tree.heading(col, text=col, anchor="w")
            # Calculer la largeur
            max_width = max(len(str(col)), 5) * 10
            self.tree.column(col, width=min(max_width + 20, 200), minwidth=50, anchor="w")

        # Ajouter les lignes
        display_df = df.head(max_rows)
        for idx, row in display_df.iterrows():
            values = []
            for col in columns:
                val = row[col]
                if pd.isna(val):
                    val = ""
                else:
                    val = str(val)[:100]  # Limiter la longueur
                values.append(val)
            self.tree.insert("", "end", values=values)

        # Mettre à jour l'info
        total_rows = len(df)
        displayed_rows = len(display_df)
        total_cols = len(df.columns)

        if total_rows > displayed_rows:
            info_text = f"{displayed_rows}/{total_rows} lignes affichées, {total_cols} colonnes"
        else:
            info_text = f"{total_rows} lignes, {total_cols} colonnes"

        self.info_label.configure(text=info_text)

    def _on_header_click(self, event):
        """Gère le clic sur un en-tête pour trier"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            column = self.tree.identify_column(event.x)
            col_idx = int(column[1:]) - 1  # Convertir #1 en 0

            if self.df is not None and col_idx < len(self.df.columns):
                col_name = self.df.columns[col_idx]
                self._sort_by_column(col_name)

    def _sort_by_column(self, column: str):
        """Trie le tableau par une colonne"""
        if self.df is None or column not in self.df.columns:
            return

        # Toggle ascending/descending
        if not hasattr(self, '_sort_ascending'):
            self._sort_ascending = {}

        ascending = self._sort_ascending.get(column, True)
        self._sort_ascending[column] = not ascending

        # Trier
        sorted_df = self.df.sort_values(by=column, ascending=ascending)
        self.load_data(sorted_df)

    def clear(self):
        """Efface toutes les données"""
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []
        self.df = None
        self.info_label.configure(text="")

    def get_selected_rows(self) -> List[int]:
        """Retourne les indices des lignes sélectionnées"""
        selected = self.tree.selection()
        indices = []
        for item in selected:
            idx = self.tree.index(item)
            indices.append(idx)
        return indices

    def get_dataframe(self) -> Optional[pd.DataFrame]:
        """Retourne le DataFrame chargé"""
        return self.df
