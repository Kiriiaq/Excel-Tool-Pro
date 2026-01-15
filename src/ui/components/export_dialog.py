"""
Dialogue d'export avancé pour ExcelToolsPro
Permet de configurer tous les paramètres d'export avant la sauvegarde
"""

import customtkinter as ctk
from tkinter import filedialog, colorchooser
from typing import Dict, Any, Optional, Callable
from pathlib import Path

from ...core.constants import COLORS
from ...core.config import ConfigManager, ExcelExportConfig
from .tooltip import Tooltip


class ExportDialog(ctk.CTkToplevel):
    """
    Dialogue d'export avancé

    Permet à l'utilisateur de configurer:
    - Le dossier et nom de fichier de destination
    - Le nom de l'onglet
    - Les options de formatage Excel
    - La prévisualisation du formatage
    """

    def __init__(
        self,
        parent,
        config_manager: ConfigManager,
        default_filename: str = "export",
        sheet_name: str = "Données",
        on_export: Optional[Callable[[Dict[str, Any]], None]] = None,
        **kwargs
    ):
        super().__init__(parent, **kwargs)

        self.config_manager = config_manager
        self.on_export = on_export
        self.result: Optional[Dict[str, Any]] = None

        # Configuration de la fenêtre
        self.title("Options d'export")
        self.geometry("550x650")
        self.transient(parent)
        self.grab_set()

        # Centrer
        self.update_idletasks()
        x = (self.winfo_screenwidth() // 2) - 275
        y = (self.winfo_screenheight() // 2) - 325
        self.geometry(f"+{x}+{y}")

        # Variables
        config = config_manager.config.excel_export
        default_dir = config_manager.config.default_output_dir or str(Path.home())

        self.filepath_var = ctk.StringVar(
            value=str(Path(default_dir) / f"{default_filename}.xlsx")
        )
        self.sheet_name_var = ctk.StringVar(value=sheet_name)

        # Options de formatage
        self.freeze_header_var = ctk.BooleanVar(value=config.freeze_header)
        self.auto_fit_var = ctk.BooleanVar(value=config.auto_fit_columns)
        self.alternate_rows_var = ctk.BooleanVar(value=config.alternate_row_colors)
        self.add_borders_var = ctk.BooleanVar(value=config.add_borders)
        self.header_color_var = ctk.StringVar(value=config.header_bg_color)
        self.min_width_var = ctk.IntVar(value=config.min_column_width)
        self.max_width_var = ctk.IntVar(value=config.max_column_width)

        self._create_interface()

    def _create_interface(self):
        """Crée l'interface du dialogue"""
        # Container principal avec scroll
        main_frame = ctk.CTkScrollableFrame(self, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # === SECTION FICHIER ===
        self._create_file_section(main_frame)

        # === SECTION FORMATAGE ===
        self._create_format_section(main_frame)

        # === SECTION COLONNES ===
        self._create_columns_section(main_frame)

        # === PRÉVISUALISATION ===
        self._create_preview_section(main_frame)

        # === BOUTONS ===
        self._create_buttons()

    def _create_file_section(self, parent):
        """Section de sélection du fichier"""
        section = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"], corner_radius=10)
        section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            section,
            text="📁 FICHIER DE DESTINATION",
            font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        # Chemin du fichier
        path_frame = ctk.CTkFrame(section, fg_color="transparent")
        path_frame.pack(fill="x", padx=15, pady=(0, 10))

        self.path_entry = ctk.CTkEntry(
            path_frame,
            textvariable=self.filepath_var,
            width=380
        )
        self.path_entry.pack(side="left")

        ctk.CTkButton(
            path_frame,
            text="...",
            width=40,
            command=self._browse_save_path
        ).pack(side="left", padx=(10, 0))

        # Nom de l'onglet
        sheet_frame = ctk.CTkFrame(section, fg_color="transparent")
        sheet_frame.pack(fill="x", padx=15, pady=(0, 15))

        lbl = ctk.CTkLabel(sheet_frame, text="Nom de l'onglet:", width=120)
        lbl.pack(side="left")
        Tooltip(lbl, "Nom de l'onglet Excel à créer")

        ctk.CTkEntry(
            sheet_frame,
            textvariable=self.sheet_name_var,
            width=200
        ).pack(side="left", padx=(10, 0))

    def _create_format_section(self, parent):
        """Section des options de formatage"""
        section = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"], corner_radius=10)
        section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            section,
            text="🎨 FORMATAGE",
            font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        options_frame = ctk.CTkFrame(section, fg_color="transparent")
        options_frame.pack(fill="x", padx=15, pady=(0, 15))

        # Ligne 1
        row1 = ctk.CTkFrame(options_frame, fg_color="transparent")
        row1.pack(fill="x", pady=3)

        cb1 = ctk.CTkCheckBox(
            row1,
            text="Geler l'en-tête",
            variable=self.freeze_header_var
        )
        cb1.pack(side="left", padx=(0, 30))
        Tooltip(cb1, "Gèle la première ligne pour qu'elle reste visible au défilement")

        cb2 = ctk.CTkCheckBox(
            row1,
            text="Ajuster colonnes auto",
            variable=self.auto_fit_var
        )
        cb2.pack(side="left")
        Tooltip(cb2, "Ajuste automatiquement la largeur des colonnes au contenu")

        # Ligne 2
        row2 = ctk.CTkFrame(options_frame, fg_color="transparent")
        row2.pack(fill="x", pady=3)

        cb3 = ctk.CTkCheckBox(
            row2,
            text="Alternance couleurs",
            variable=self.alternate_rows_var
        )
        cb3.pack(side="left", padx=(0, 30))
        Tooltip(cb3, "Alterne les couleurs de fond des lignes pour une meilleure lisibilité")

        cb4 = ctk.CTkCheckBox(
            row2,
            text="Ajouter bordures",
            variable=self.add_borders_var
        )
        cb4.pack(side="left")
        Tooltip(cb4, "Ajoute des bordures fines autour de chaque cellule")

        # Couleur de l'en-tête
        color_frame = ctk.CTkFrame(options_frame, fg_color="transparent")
        color_frame.pack(fill="x", pady=(10, 0))

        lbl = ctk.CTkLabel(color_frame, text="Couleur en-tête:", width=120)
        lbl.pack(side="left")

        self.color_preview = ctk.CTkLabel(
            color_frame,
            text="",
            width=30,
            height=25,
            fg_color=self.header_color_var.get(),
            corner_radius=5
        )
        self.color_preview.pack(side="left", padx=(10, 5))

        ctk.CTkButton(
            color_frame,
            text="Choisir",
            width=70,
            height=25,
            command=self._choose_header_color
        ).pack(side="left")

    def _create_columns_section(self, parent):
        """Section des paramètres de colonnes"""
        section = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"], corner_radius=10)
        section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            section,
            text="📏 LARGEUR DES COLONNES",
            font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        cols_frame = ctk.CTkFrame(section, fg_color="transparent")
        cols_frame.pack(fill="x", padx=15, pady=(0, 15))

        # Largeur min
        min_frame = ctk.CTkFrame(cols_frame, fg_color="transparent")
        min_frame.pack(fill="x", pady=3)

        lbl = ctk.CTkLabel(min_frame, text="Largeur minimale:", width=140)
        lbl.pack(side="left")
        Tooltip(lbl, "Largeur minimale des colonnes en caractères")

        self.min_width_slider = ctk.CTkSlider(
            min_frame,
            from_=5,
            to=30,
            number_of_steps=25,
            variable=self.min_width_var,
            width=150,
            command=lambda v: self.min_label.configure(text=str(int(v)))
        )
        self.min_width_slider.pack(side="left", padx=(10, 5))

        self.min_label = ctk.CTkLabel(min_frame, text=str(self.min_width_var.get()), width=30)
        self.min_label.pack(side="left")

        # Largeur max
        max_frame = ctk.CTkFrame(cols_frame, fg_color="transparent")
        max_frame.pack(fill="x", pady=3)

        lbl = ctk.CTkLabel(max_frame, text="Largeur maximale:", width=140)
        lbl.pack(side="left")
        Tooltip(lbl, "Largeur maximale des colonnes en caractères")

        self.max_width_slider = ctk.CTkSlider(
            max_frame,
            from_=20,
            to=100,
            number_of_steps=80,
            variable=self.max_width_var,
            width=150,
            command=lambda v: self.max_label.configure(text=str(int(v)))
        )
        self.max_width_slider.pack(side="left", padx=(10, 5))

        self.max_label = ctk.CTkLabel(max_frame, text=str(self.max_width_var.get()), width=30)
        self.max_label.pack(side="left")

    def _create_preview_section(self, parent):
        """Section de prévisualisation"""
        section = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"], corner_radius=10)
        section.pack(fill="x", pady=(0, 15))

        ctk.CTkLabel(
            section,
            text="👁️ APERÇU DU FORMATAGE",
            font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        preview_frame = ctk.CTkFrame(section, fg_color="white", corner_radius=5)
        preview_frame.pack(fill="x", padx=15, pady=(0, 15))

        # Simulation d'un tableau Excel
        self.preview_canvas = ctk.CTkFrame(preview_frame, fg_color="white", height=80)
        self.preview_canvas.pack(fill="x", padx=2, pady=2)

        self._update_preview()

    def _update_preview(self):
        """Met à jour l'aperçu du formatage"""
        # Nettoyer
        for widget in self.preview_canvas.winfo_children():
            widget.destroy()

        header_color = self.header_color_var.get()
        alternate = self.alternate_rows_var.get()

        # En-tête
        header_frame = ctk.CTkFrame(self.preview_canvas, fg_color=header_color, height=25)
        header_frame.pack(fill="x")

        for col in ["Colonne A", "Colonne B", "Colonne C"]:
            ctk.CTkLabel(
                header_frame,
                text=col,
                text_color="white",
                font=("Segoe UI", 10, "bold"),
                width=100
            ).pack(side="left", padx=5, pady=3)

        # Lignes de données
        for i in range(3):
            bg = "#F2F2F2" if alternate and i % 2 == 0 else "white"
            row_frame = ctk.CTkFrame(self.preview_canvas, fg_color=bg, height=22)
            row_frame.pack(fill="x")

            for j, val in enumerate([f"Valeur {i+1}-{j+1}" for j in range(3)]):
                ctk.CTkLabel(
                    row_frame,
                    text=val,
                    text_color="black",
                    font=("Segoe UI", 9),
                    width=100
                ).pack(side="left", padx=5, pady=2)

    def _create_buttons(self):
        """Crée les boutons d'action"""
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(0, 20))

        # Sauvegarder comme défaut
        ctk.CTkButton(
            btn_frame,
            text="💾 Comme défaut",
            width=120,
            fg_color=COLORS["info"],
            command=self._save_as_default
        ).pack(side="left")
        Tooltip(btn_frame.winfo_children()[-1], "Enregistrer ces options comme valeurs par défaut")

        # Annuler
        ctk.CTkButton(
            btn_frame,
            text="Annuler",
            width=100,
            fg_color=COLORS["text_muted"],
            command=self._cancel
        ).pack(side="right", padx=(10, 0))

        # Exporter
        ctk.CTkButton(
            btn_frame,
            text="📤 Exporter",
            width=120,
            fg_color=COLORS["success"],
            command=self._export
        ).pack(side="right")

    def _browse_save_path(self):
        """Ouvre le dialogue de sélection du fichier"""
        filepath = filedialog.asksaveasfilename(
            title="Enregistrer sous",
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx")],
            initialfile=Path(self.filepath_var.get()).name
        )
        if filepath:
            self.filepath_var.set(filepath)

    def _choose_header_color(self):
        """Ouvre le sélecteur de couleur"""
        color = colorchooser.askcolor(initialcolor=self.header_color_var.get())[1]
        if color:
            self.header_color_var.set(color)
            self.color_preview.configure(fg_color=color)
            self._update_preview()

    def _save_as_default(self):
        """Sauvegarde les options comme valeurs par défaut"""
        config = self.config_manager.config.excel_export

        config.freeze_header = self.freeze_header_var.get()
        config.auto_fit_columns = self.auto_fit_var.get()
        config.alternate_row_colors = self.alternate_rows_var.get()
        config.add_borders = self.add_borders_var.get()
        config.header_bg_color = self.header_color_var.get()
        config.min_column_width = self.min_width_var.get()
        config.max_column_width = self.max_width_var.get()

        self.config_manager.save()

        # Feedback visuel
        from tkinter import messagebox
        messagebox.showinfo("Succès", "Options enregistrées comme valeurs par défaut")

    def _cancel(self):
        """Annule et ferme le dialogue"""
        self.result = None
        self.destroy()

    def _export(self):
        """Valide et lance l'export"""
        filepath = self.filepath_var.get().strip()

        if not filepath:
            from tkinter import messagebox
            messagebox.showwarning("Attention", "Veuillez spécifier un fichier de destination")
            return

        sheet_name = self.sheet_name_var.get().strip()
        if not sheet_name:
            sheet_name = "Données"

        # Construire le résultat
        self.result = {
            "filepath": filepath,
            "sheet_name": sheet_name,
            "freeze_header": self.freeze_header_var.get(),
            "auto_fit_columns": self.auto_fit_var.get(),
            "alternate_row_colors": self.alternate_rows_var.get(),
            "add_borders": self.add_borders_var.get(),
            "header_bg_color": self.header_color_var.get(),
            "min_column_width": self.min_width_var.get(),
            "max_column_width": self.max_width_var.get(),
        }

        if self.on_export:
            self.on_export(self.result)

        self.destroy()

    def get_result(self) -> Optional[Dict[str, Any]]:
        """Retourne le résultat de l'export ou None si annulé"""
        return self.result
