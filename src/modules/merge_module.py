"""
Module de fusion de fichiers Excel
Fusion de deux fichiers Excel sur une colonne commune
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
from typing import Dict, Any, Optional, List
import pandas as pd

from .base_module import BaseModule
from ..ui.components.tooltip import Tooltip
from ..ui.components.file_selector import FileSelector
from ..ui.components.preview_table import PreviewTable
from ..ui.components.stat_card import StatCardGroup
from ..ui.components.settings_panel import SettingsPanel, SettingDefinition, SettingType
from ..core.constants import COLORS
from ..utils.excel_utils import ExcelUtils


class MergeModule(BaseModule):
    """
    Module de fusion de références documentaires

    Fonctionnalités:
    - Chargement de deux fichiers Excel
    - Sélection des colonnes clés
    - Fusion par jointure gauche
    - Prévisualisation du résultat
    - Export avec formatage professionnel
    """

    MODULE_ID = "merge"
    MODULE_NAME = "Fusion de documents"
    MODULE_DESCRIPTION = "Fusionne deux fichiers Excel sur une colonne commune"
    MODULE_ICON = "🔗"

    def _create_interface(self):
        """Crée l'interface du module de fusion"""
        # Container principal avec scroll
        main_scroll = ctk.CTkScrollableFrame(self.frame, fg_color="transparent")
        main_scroll.pack(fill="both", expand=True, padx=10, pady=10)

        # Layout en deux colonnes
        left_panel = ctk.CTkFrame(main_scroll, fg_color="transparent", width=450)
        left_panel.pack(side="left", fill="y", padx=(0, 10))

        right_panel = ctk.CTkFrame(main_scroll, fg_color="transparent")
        right_panel.pack(side="right", fill="both", expand=True)

        # === PANNEAU GAUCHE - Configuration ===
        self._create_files_section(left_panel)
        self._create_join_config_section(left_panel)
        self._create_options_section(left_panel)
        self._create_action_buttons(left_panel)

        # === PANNEAU DROIT - Prévisualisation ===
        self._create_preview_section(right_panel)

    def _create_files_section(self, parent):
        """Section de sélection des fichiers"""
        section = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"], corner_radius=10)
        section.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            section,
            text="📁 FICHIERS",
            font=("Segoe UI", 14, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        # Fichier source
        self.source_selector = FileSelector(
            section,
            label="Fichier Source",
            tooltip="Fichier Excel principal. Un nouvel onglet sera créé avec les données fusionnées.",
            on_file_loaded=self._on_source_loaded
        )
        self.source_selector.pack(fill="x", padx=15, pady=(0, 10))

        # Fichier référence
        self.ref_selector = FileSelector(
            section,
            label="Fichier Référence",
            tooltip="Fichier contenant les données à ajouter au fichier source.",
            on_file_loaded=self._on_ref_loaded
        )
        self.ref_selector.pack(fill="x", padx=15, pady=(0, 15))

    def _create_join_config_section(self, parent):
        """Section de configuration de la jointure"""
        section = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"], corner_radius=10)
        section.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            section,
            text="🔗 CONFIGURATION DE LA JOINTURE",
            font=("Segoe UI", 14, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        config_frame = ctk.CTkFrame(section, fg_color="transparent")
        config_frame.pack(fill="x", padx=15, pady=(0, 15))

        # Colonne source
        src_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        src_frame.pack(fill="x", pady=5)

        lbl = ctk.CTkLabel(src_frame, text="Colonne clé (Source):", width=180)
        lbl.pack(side="left")
        Tooltip(lbl, "Colonne de référence dans le fichier source")

        self.col_source_combo = ctk.CTkComboBox(
            src_frame,
            values=["(Charger fichier)"],
            state="disabled",
            width=220
        )
        self.col_source_combo.pack(side="left", padx=(10, 0))

        # Colonne référence
        ref_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        ref_frame.pack(fill="x", pady=5)

        lbl = ctk.CTkLabel(ref_frame, text="Colonne clé (Référence):", width=180)
        lbl.pack(side="left")
        Tooltip(lbl, "Colonne de référence dans le fichier référentiel")

        self.col_ref_combo = ctk.CTkComboBox(
            ref_frame,
            values=["(Charger fichier)"],
            state="disabled",
            width=220
        )
        self.col_ref_combo.pack(side="left", padx=(10, 0))

        # Nom de l'onglet de sortie
        output_frame = ctk.CTkFrame(config_frame, fg_color="transparent")
        output_frame.pack(fill="x", pady=(10, 0))

        lbl = ctk.CTkLabel(output_frame, text="Nom onglet sortie:", width=180)
        lbl.pack(side="left")
        Tooltip(lbl, "Nom du nouvel onglet à créer")

        self.output_sheet_entry = ctk.CTkEntry(output_frame, width=220)
        self.output_sheet_entry.insert(0, "Données_Fusionnées")
        self.output_sheet_entry.pack(side="left", padx=(10, 0))

    def _create_options_section(self, parent):
        """Section des options avancées"""
        section = ctk.CTkFrame(parent, fg_color=COLORS["bg_card"], corner_radius=10)
        section.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            section,
            text="⚙️ OPTIONS",
            font=("Segoe UI", 14, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        options_frame = ctk.CTkFrame(section, fg_color="transparent")
        options_frame.pack(fill="x", padx=15, pady=(0, 15))

        # Option LAST
        self.filter_last_var = ctk.BooleanVar(value=False)
        cb1 = ctk.CTkCheckBox(
            options_frame,
            text="Filtrer uniquement LAST = Y",
            variable=self.filter_last_var
        )
        cb1.pack(anchor="w", pady=3)
        Tooltip(cb1, "Ne garde que les lignes où la colonne LAST vaut 'Y'")

        # Option correspondances uniquement
        self.match_only_var = ctk.BooleanVar(value=False)
        cb2 = ctk.CTkCheckBox(
            options_frame,
            text="Exporter uniquement les correspondances",
            variable=self.match_only_var
        )
        cb2.pack(anchor="w", pady=3)
        Tooltip(cb2, "N'exporte que les lignes ayant une correspondance")

        # Option ajout colonne MATCH
        self.add_match_col_var = ctk.BooleanVar(value=True)
        cb3 = ctk.CTkCheckBox(
            options_frame,
            text="Ajouter colonne MATCH (OUI/NON)",
            variable=self.add_match_col_var
        )
        cb3.pack(anchor="w", pady=3)
        Tooltip(cb3, "Ajoute une colonne indiquant si une correspondance a été trouvée")

    def _create_action_buttons(self, parent):
        """Section des boutons d'action"""
        action_frame = ctk.CTkFrame(parent, fg_color="transparent")
        action_frame.pack(fill="x", pady=(0, 10))

        self.preview_btn = ctk.CTkButton(
            action_frame,
            text="👁️ Prévisualiser",
            font=("Segoe UI", 12),
            height=40,
            fg_color=COLORS["accent_primary"],
            command=self._preview_merge
        )
        self.preview_btn.pack(fill="x", pady=(0, 10))
        Tooltip(self.preview_btn, "Prévisualiser le résultat sans modifier les fichiers")

        self.execute_btn = ctk.CTkButton(
            action_frame,
            text="🚀 FUSIONNER ET EXPORTER",
            font=("Segoe UI", 14, "bold"),
            height=50,
            fg_color=COLORS["success"],
            hover_color="#00a050",
            command=self.start_execution
        )
        self.execute_btn.pack(fill="x")
        Tooltip(self.execute_btn, "Exécuter la fusion et créer le nouvel onglet")

        # Barre de progression
        self.progress_bar = ctk.CTkProgressBar(action_frame, height=8)
        self.progress_bar.pack(fill="x", pady=(10, 0))
        self.progress_bar.set(0)

        # Label de statut
        self.status_label = ctk.CTkLabel(
            action_frame,
            text="Prêt",
            font=("Segoe UI", 10),
            text_color=COLORS["text_muted"]
        )
        self.status_label.pack(anchor="w", pady=(5, 0))

    def _create_preview_section(self, parent):
        """Section de prévisualisation"""
        # Onglets
        self.tabview = ctk.CTkTabview(parent, fg_color=COLORS["bg_card"])
        self.tabview.pack(fill="both", expand=True)

        self.tabview.add("📄 Source")
        self.tabview.add("📋 Référence")
        self.tabview.add("✨ Résultat")
        self.tabview.add("📊 Statistiques")

        # Tables de prévisualisation
        self.source_preview = PreviewTable(
            self.tabview.tab("📄 Source"),
            title="Aperçu du fichier source"
        )
        self.source_preview.pack(fill="both", expand=True, padx=5, pady=5)

        self.ref_preview = PreviewTable(
            self.tabview.tab("📋 Référence"),
            title="Aperçu du fichier référence"
        )
        self.ref_preview.pack(fill="both", expand=True, padx=5, pady=5)

        self.result_preview = PreviewTable(
            self.tabview.tab("✨ Résultat"),
            title="Aperçu du résultat fusionné"
        )
        self.result_preview.pack(fill="both", expand=True, padx=5, pady=5)

        # Statistiques
        stats_frame = ctk.CTkFrame(
            self.tabview.tab("📊 Statistiques"),
            fg_color="transparent"
        )
        stats_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Cartes de statistiques
        self.stats_cards = StatCardGroup(stats_frame)
        self.stats_cards.pack(fill="x", pady=(0, 20))

        self.stats_cards.add_card("source_rows", "Lignes source", "0", "📄", color=COLORS["info"])
        self.stats_cards.add_card("ref_rows", "Lignes référence", "0", "📋", color=COLORS["info"])
        self.stats_cards.add_card("matches", "Correspondances", "0", "✓", color=COLORS["success"])
        self.stats_cards.add_card("no_matches", "Sans correspondance", "0", "✗", color=COLORS["error"])

        # Texte détaillé
        self.stats_text = ctk.CTkTextbox(stats_frame, font=("Consolas", 11))
        self.stats_text.pack(fill="both", expand=True)

    def _on_source_loaded(self, df: pd.DataFrame):
        """Callback quand le fichier source est chargé"""
        self.source_preview.load_data(df)
        columns = list(df.columns)
        self.col_source_combo.configure(state="normal", values=columns)

        # Auto-sélection de colonnes communes
        for col in ["REF", "Référence", "Reference", "ID", "DOC_REF"]:
            if col in columns:
                self.col_source_combo.set(col)
                break
        else:
            self.col_source_combo.set(columns[0])

        self.log_info(f"Fichier source chargé: {len(df)} lignes")

    def _on_ref_loaded(self, df: pd.DataFrame):
        """Callback quand le fichier référence est chargé"""
        self.ref_preview.load_data(df)
        columns = list(df.columns)
        self.col_ref_combo.configure(state="normal", values=columns)

        for col in ["REF", "Référence", "Reference", "Legacy Number", "TC Reference", "ID"]:
            if col in columns:
                self.col_ref_combo.set(col)
                break
        else:
            self.col_ref_combo.set(columns[0])

        self.log_info(f"Fichier référence chargé: {len(df)} lignes")

    def validate_inputs(self) -> tuple[bool, str]:
        """Valide les entrées utilisateur"""
        if not self.source_selector.is_loaded():
            return False, "Veuillez charger un fichier source"

        if not self.ref_selector.is_loaded():
            return False, "Veuillez charger un fichier référence"

        if not self.output_sheet_entry.get().strip():
            return False, "Veuillez spécifier un nom d'onglet de sortie"

        return True, ""

    def _merge_data(self) -> pd.DataFrame:
        """Effectue la fusion des données"""
        df_source = self.source_selector.get_dataframe().copy()
        df_ref = self.ref_selector.get_dataframe().copy()

        col_source = self.col_source_combo.get()
        col_ref = self.col_ref_combo.get()

        # Nettoyage des colonnes clés
        df_source[col_source] = df_source[col_source].astype(str).str.strip()
        df_ref[col_ref] = df_ref[col_ref].astype(str).str.strip()

        # Filtrage LAST si demandé
        if self.filter_last_var.get():
            for last_col in ["LAST", "Last", "last"]:
                if last_col in df_ref.columns:
                    df_ref = df_ref[df_ref[last_col].str.upper() == "Y"]
                    break

        # Renommer colonnes du référentiel pour éviter les conflits
        ref_cols_renamed = {}
        for col in df_ref.columns:
            if col != col_ref and col in df_source.columns:
                ref_cols_renamed[col] = f"{col}_REF"
        df_ref_renamed = df_ref.rename(columns=ref_cols_renamed)

        # Fusion (jointure gauche)
        df_merged = pd.merge(
            df_source,
            df_ref_renamed,
            left_on=col_source,
            right_on=col_ref,
            how='left',
            suffixes=('', '_REF2')
        )

        # Ajouter colonne MATCH si demandé
        if self.add_match_col_var.get():
            df_merged['MATCH'] = df_merged[col_ref].notna().map({True: 'OUI', False: 'NON'})

        # Supprimer colonne dupliquée
        if col_ref in df_merged.columns and col_source != col_ref:
            df_merged = df_merged.drop(columns=[col_ref])

        # Filtrer correspondances uniquement si demandé
        if self.match_only_var.get() and self.add_match_col_var.get():
            df_merged = df_merged[df_merged['MATCH'] == 'OUI']

        return df_merged

    def _preview_merge(self):
        """Prévisualise la fusion"""
        valid, error = self.validate_inputs()
        if not valid:
            messagebox.showwarning("Attention", error)
            return

        try:
            self.update_status("Fusion en cours...")
            self.progress_bar.set(0.3)
            self.frame.update()

            df_merged = self._merge_data()

            self.progress_bar.set(0.7)
            self.frame.update()

            self.result_preview.load_data(df_merged, max_rows=50)
            self._update_statistics(df_merged)

            self.progress_bar.set(1.0)
            self.update_status(f"Prévisualisation: {len(df_merged)} lignes")
            self.tabview.set("✨ Résultat")

            self.log_success(f"Prévisualisation terminée: {len(df_merged)} lignes")

        except Exception as e:
            self.update_status(f"Erreur: {str(e)}")
            self.log_error(str(e))
            messagebox.showerror("Erreur", str(e))

        finally:
            self.frame.after(2000, lambda: self.progress_bar.set(0))

    def _update_statistics(self, df_merged: pd.DataFrame):
        """Met à jour les statistiques avec analyse avancée"""
        df_source = self.source_selector.get_dataframe()
        df_ref = self.ref_selector.get_dataframe()

        matches = (df_merged['MATCH'] == 'OUI').sum() if 'MATCH' in df_merged.columns else 0
        no_matches = len(df_merged) - matches
        rate = matches / len(df_source) * 100 if len(df_source) > 0 else 0

        self.stats_cards.update_card("source_rows", str(len(df_source)))
        self.stats_cards.update_card("ref_rows", str(len(df_ref)))
        self.stats_cards.update_card("matches", str(matches))
        self.stats_cards.update_card("no_matches", str(no_matches))

        # Analyse avancée
        col_source = self.col_source_combo.get()
        col_ref = self.col_ref_combo.get()

        # Valeurs uniques et doublons
        source_unique = df_source[col_source].nunique()
        source_duplicates = len(df_source) - source_unique
        ref_unique = df_ref[col_ref].nunique()
        ref_duplicates = len(df_ref) - ref_unique

        # Valeurs communes
        source_values = set(df_source[col_source].dropna().astype(str))
        ref_values = set(df_ref[col_ref].dropna().astype(str))
        common_values = len(source_values & ref_values)
        only_in_source = len(source_values - ref_values)
        only_in_ref = len(ref_values - source_values)

        # Analyse des valeurs vides
        source_empty = df_source[col_source].isna().sum() + (df_source[col_source].astype(str).str.strip() == '').sum()
        ref_empty = df_ref[col_ref].isna().sum() + (df_ref[col_ref].astype(str).str.strip() == '').sum()

        # Top 5 des non correspondances
        non_match_sample = ""
        if 'MATCH' in df_merged.columns:
            non_matches_df = df_merged[df_merged['MATCH'] == 'NON']
            if len(non_matches_df) > 0:
                sample = non_matches_df[col_source].head(5).tolist()
                non_match_sample = "\n".join([f"    - {v}" for v in sample])
                if len(non_matches_df) > 5:
                    non_match_sample += f"\n    ... et {len(non_matches_df) - 5} autres"

        # Texte détaillé
        stats_text = f"""
╔══════════════════════════════════════════════════════════╗
║           STATISTIQUES DE LA FUSION                       ║
╠══════════════════════════════════════════════════════════╣

📄 FICHIER SOURCE
   • Lignes totales: {len(df_source):,}
   • Colonnes: {len(df_source.columns)}
   • Colonne clé: {col_source}
   • Valeurs uniques: {source_unique:,}
   • Doublons: {source_duplicates:,}
   • Valeurs vides: {source_empty:,}

📋 FICHIER RÉFÉRENCE
   • Lignes totales: {len(df_ref):,}
   • Colonnes: {len(df_ref.columns)}
   • Colonne clé: {col_ref}
   • Valeurs uniques: {ref_unique:,}
   • Doublons: {ref_duplicates:,}
   • Valeurs vides: {ref_empty:,}

╠══════════════════════════════════════════════════════════╣

🔗 ANALYSE DES CORRESPONDANCES
   • Valeurs communes aux deux fichiers: {common_values:,}
   • Valeurs uniquement dans source: {only_in_source:,}
   • Valeurs uniquement dans référence: {only_in_ref:,}

✨ RÉSULTAT DE LA FUSION
   ✓ Correspondances trouvées: {matches:,}
   ✗ Sans correspondance: {no_matches:,}
   📊 Taux de correspondance: {rate:.1f}%
   📁 Colonnes résultat: {len(df_merged.columns)}
   📝 Lignes exportées: {len(df_merged):,}

╠══════════════════════════════════════════════════════════╣

📋 COLONNES DU RÉSULTAT:
{chr(10).join(['   • ' + col for col in df_merged.columns])}
"""

        if non_match_sample:
            stats_text += f"""
╠══════════════════════════════════════════════════════════╣

⚠️ ÉCHANTILLON NON CORRESPONDANCES (clé source):
{non_match_sample}
"""

        stats_text += """
╚══════════════════════════════════════════════════════════╝
"""

        self.stats_text.delete("1.0", "end")
        self.stats_text.insert("1.0", stats_text)

    def _execute_task(self) -> Dict[str, Any]:
        """Exécute la fusion et l'export"""
        self.update_progress(0.1)
        self.update_status("Fusion des données...")

        df_merged = self._merge_data()

        self.update_progress(0.5)
        self.update_status("Création de l'onglet...")

        # Créer l'onglet dans le fichier source
        output_sheet = self.output_sheet_entry.get().strip() or "Données_Fusionnées"
        filepath = self.source_selector.get_filepath()

        success, error = ExcelUtils.add_sheet_to_workbook(
            filepath,
            output_sheet,
            df_merged,
            apply_formatting=True
        )

        if not success:
            raise Exception(error)

        self.update_progress(1.0)

        # Mettre à jour l'interface (dans le thread principal)
        self.frame.after(0, lambda: self._on_task_complete(df_merged, output_sheet, filepath))

        return {
            "rows": len(df_merged),
            "columns": len(df_merged.columns),
            "sheet": output_sheet,
            "filepath": filepath
        }

    def _on_task_complete(self, df_merged: pd.DataFrame, sheet: str, filepath: str):
        """Callback après la fin de la tâche"""
        self.result_preview.load_data(df_merged)
        self._update_statistics(df_merged)
        self.tabview.set("📊 Statistiques")

        messagebox.showinfo(
            "Succès",
            f"Fusion terminée!\n\nOnglet '{sheet}' créé dans:\n{filepath}"
        )

    def update_status(self, message: str, level: str = "info"):
        """Met à jour le statut affiché"""
        color_map = {
            "info": COLORS["text_muted"],
            "success": COLORS["success"],
            "warning": COLORS["warning"],
            "error": COLORS["error"]
        }
        self.status_label.configure(
            text=message,
            text_color=color_map.get(level, COLORS["text_muted"])
        )

    def update_progress(self, progress: float):
        """Met à jour la barre de progression"""
        self.progress_bar.set(progress)

    def reset(self):
        """Réinitialise le module"""
        super().reset()
        self.source_selector.reset()
        self.ref_selector.reset()
        self.source_preview.clear()
        self.ref_preview.clear()
        self.result_preview.clear()
        self.stats_cards.reset_all()
        self.output_sheet_entry.delete(0, "end")
        self.output_sheet_entry.insert(0, "Données_Fusionnées")
