"""
Module de gestion de fichiers
Déplacement, copie et organisation de fichiers basés sur liste Excel
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
from typing import Dict, Any, Optional, List
from pathlib import Path
import pandas as pd
import shutil
import threading
from datetime import datetime
from dataclasses import dataclass
from enum import Enum

from .base_module import BaseModule
from ..ui.components.tooltip import Tooltip
from ..ui.components.file_selector import FileSelector
from ..ui.components.preview_table import PreviewTable
from ..ui.components.stat_card import StatCardGroup
from ..core.constants import COLORS


class OperationType(Enum):
    MOVE = "move"
    COPY = "copy"


@dataclass
class OperationStats:
    total: int = 0
    success: int = 0
    errors: int = 0
    not_found: int = 0
    skipped: int = 0


class FileManagerModule(BaseModule):
    """
    Module de gestion de fichiers en masse

    Fonctionnalités:
    - Déplacement/copie de fichiers basé sur liste Excel
    - Gestion des conflits (renommer, écraser, ignorer)
    - Prévisualisation avant exécution
    - Logs détaillés avec statistiques
    - Organisation automatique en dossiers
    """

    MODULE_ID = "file_manager"
    MODULE_NAME = "Gestionnaire de fichiers"
    MODULE_DESCRIPTION = "Déplace ou copie des fichiers basés sur liste Excel"
    MODULE_ICON = "📂"

    def __init__(self, *args, **kwargs):
        self._stop_event = threading.Event()
        self._stats = OperationStats()
        super().__init__(*args, **kwargs)

    def _create_interface(self):
        """Crée l'interface du module"""
        # Onglets
        self.tabview = ctk.CTkTabview(self.frame, fg_color=COLORS["bg_card"])
        self.tabview.pack(fill="both", expand=True, padx=10, pady=10)

        self.tabview.add("Configuration")
        self.tabview.add("Prévisualisation")
        self.tabview.add("Journaux")

        self._create_config_tab()
        self._create_preview_tab()
        self._create_logs_tab()

    def _create_config_tab(self):
        """Onglet de configuration"""
        tab = self.tabview.tab("Configuration")

        scroll = ctk.CTkScrollableFrame(tab, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=10, pady=10)

        # === Section Fichier Liste ===
        section1 = ctk.CTkFrame(scroll, fg_color=COLORS["bg_card"], corner_radius=10)
        section1.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            section1, text="📋 FICHIER LISTE", font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        self.list_selector = FileSelector(
            section1,
            label="Liste Excel",
            tooltip="Fichier Excel contenant les chemins des fichiers à traiter",
            on_file_loaded=self._on_list_loaded
        )
        self.list_selector.pack(fill="x", padx=15, pady=(0, 10))

        col_frame = ctk.CTkFrame(section1, fg_color="transparent")
        col_frame.pack(fill="x", padx=15, pady=(0, 15))

        ctk.CTkLabel(col_frame, text="Colonne des chemins:", width=150).pack(side="left")
        self.path_col_combo = ctk.CTkComboBox(
            col_frame, values=["(Charger fichier)"], state="disabled", width=200
        )
        self.path_col_combo.pack(side="left", padx=(10, 0))
        Tooltip(self.path_col_combo, "Colonne contenant les chemins complets des fichiers")

        # === Section Dossier Destination ===
        section2 = ctk.CTkFrame(scroll, fg_color=COLORS["bg_card"], corner_radius=10)
        section2.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            section2, text="📂 DESTINATION", font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        dest_frame = ctk.CTkFrame(section2, fg_color="transparent")
        dest_frame.pack(fill="x", padx=15, pady=(0, 15))

        self.dest_path_var = ctk.StringVar()
        ctk.CTkEntry(
            dest_frame, textvariable=self.dest_path_var, width=350, state="disabled"
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            dest_frame, text="📂 Parcourir", width=100,
            command=self._browse_destination
        ).pack(side="left")

        # === Section Options ===
        section3 = ctk.CTkFrame(scroll, fg_color=COLORS["bg_card"], corner_radius=10)
        section3.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            section3, text="⚙️ OPTIONS", font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        options_frame = ctk.CTkFrame(section3, fg_color="transparent")
        options_frame.pack(fill="x", padx=15, pady=(0, 10))

        # Type d'opération
        ctk.CTkLabel(options_frame, text="Opération:", width=100).pack(side="left")

        self.operation_var = ctk.StringVar(value="move")
        ctk.CTkRadioButton(
            options_frame, text="Déplacer", variable=self.operation_var, value="move"
        ).pack(side="left", padx=10)
        ctk.CTkRadioButton(
            options_frame, text="Copier", variable=self.operation_var, value="copy"
        ).pack(side="left", padx=10)

        # Gestion des conflits
        conflict_frame = ctk.CTkFrame(section3, fg_color="transparent")
        conflict_frame.pack(fill="x", padx=15, pady=(0, 10))

        ctk.CTkLabel(conflict_frame, text="Si fichier existe:", width=100).pack(side="left")

        self.conflict_var = ctk.StringVar(value="rename")
        ctk.CTkRadioButton(
            conflict_frame, text="Renommer", variable=self.conflict_var, value="rename"
        ).pack(side="left", padx=10)
        ctk.CTkRadioButton(
            conflict_frame, text="Écraser", variable=self.conflict_var, value="overwrite"
        ).pack(side="left", padx=10)
        ctk.CTkRadioButton(
            conflict_frame, text="Ignorer", variable=self.conflict_var, value="skip"
        ).pack(side="left", padx=10)

        # Options supplémentaires
        extra_frame = ctk.CTkFrame(section3, fg_color="transparent")
        extra_frame.pack(fill="x", padx=15, pady=(0, 15))

        self.preserve_structure_var = ctk.BooleanVar(value=False)
        cb1 = ctk.CTkCheckBox(
            extra_frame, text="Conserver la structure des dossiers",
            variable=self.preserve_structure_var
        )
        cb1.pack(anchor="w", pady=2)
        Tooltip(cb1, "Recrée les sous-dossiers dans la destination")

        self.ignore_locked_var = ctk.BooleanVar(value=True)
        cb2 = ctk.CTkCheckBox(
            extra_frame, text="Ignorer les fichiers verrouillés (~$...)",
            variable=self.ignore_locked_var
        )
        cb2.pack(anchor="w", pady=2)

        self.create_log_var = ctk.BooleanVar(value=True)
        cb3 = ctk.CTkCheckBox(
            extra_frame, text="Créer un rapport d'opération",
            variable=self.create_log_var
        )
        cb3.pack(anchor="w", pady=2)

        # === Section Actions ===
        action_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        action_frame.pack(fill="x", pady=(0, 10))

        btn_row = ctk.CTkFrame(action_frame, fg_color="transparent")
        btn_row.pack(fill="x", pady=(0, 10))

        self.preview_btn = ctk.CTkButton(
            btn_row, text="👁️ Prévisualiser", width=150,
            fg_color=COLORS["accent_primary"],
            command=self._preview_operation
        )
        self.preview_btn.pack(side="left", padx=(0, 10))

        self.execute_btn = ctk.CTkButton(
            btn_row, text="▶️ EXÉCUTER", width=150,
            font=("Segoe UI", 12, "bold"),
            fg_color=COLORS["success"],
            command=self._start_operation
        )
        self.execute_btn.pack(side="left", padx=(0, 10))

        self.cancel_btn = ctk.CTkButton(
            btn_row, text="⏹️ Annuler", width=100,
            fg_color=COLORS["error"], state="disabled",
            command=self._cancel_operation
        )
        self.cancel_btn.pack(side="left")

        # Progression
        self.progress_bar = ctk.CTkProgressBar(action_frame, height=8)
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            action_frame, text="Prêt", font=("Segoe UI", 10), text_color=COLORS["text_muted"]
        )
        self.status_label.pack(anchor="w", pady=(5, 0))

        # Statistiques
        self.stats_cards = StatCardGroup(scroll)
        self.stats_cards.pack(fill="x", pady=(0, 10))

        self.stats_cards.add_card("total", "Total", "0", "📊")
        self.stats_cards.add_card("success", "Succès", "0", "✓", color=COLORS["success"])
        self.stats_cards.add_card("errors", "Erreurs", "0", "✗", color=COLORS["error"])
        self.stats_cards.add_card("not_found", "Introuvables", "0", "❓", color=COLORS["warning"])

    def _create_preview_tab(self):
        """Onglet de prévisualisation"""
        tab = self.tabview.tab("Prévisualisation")

        self.preview_table = PreviewTable(
            tab,
            title="Fichiers à traiter",
            max_rows=1000
        )
        self.preview_table.pack(fill="both", expand=True, padx=10, pady=10)

    def _create_logs_tab(self):
        """Onglet des journaux"""
        tab = self.tabview.tab("Journaux")

        # Zone de log
        self.log_text = ctk.CTkTextbox(
            tab, font=("Consolas", 10), fg_color=("gray90", "gray20")
        )
        self.log_text.pack(fill="both", expand=True, padx=10, pady=10)

        # Boutons
        btn_frame = ctk.CTkFrame(tab, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkButton(
            btn_frame, text="📋 Copier", width=100,
            command=self._copy_logs
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_frame, text="💾 Exporter", width=100,
            command=self._export_logs
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_frame, text="🗑️ Effacer", width=100,
            fg_color=COLORS["error"],
            command=lambda: self.log_text.delete("1.0", "end")
        ).pack(side="left")

    def _on_list_loaded(self, df: pd.DataFrame):
        """Callback quand le fichier liste est chargé"""
        columns = list(df.columns)
        self.path_col_combo.configure(state="normal", values=columns)

        # Auto-sélection d'une colonne de chemin
        for col in columns:
            if any(kw in col.lower() for kw in ["chemin", "path", "fichier", "file", "source"]):
                self.path_col_combo.set(col)
                break
        else:
            self.path_col_combo.set(columns[0])

        self.log_info(f"Liste chargée: {len(df)} entrées")

    def _browse_destination(self):
        """Sélectionne le dossier de destination"""
        folder = filedialog.askdirectory(title="Sélectionner le dossier de destination")
        if folder:
            self.dest_path_var.set(folder)

    def _log_message(self, message: str, level: str = "info"):
        """Ajoute un message au journal"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        colors = {"info": "", "success": "✓ ", "warning": "⚠️ ", "error": "✗ "}
        prefix = colors.get(level, "")
        full_message = f"[{timestamp}] {prefix}{message}\n"

        self.log_text.insert("end", full_message)
        self.log_text.see("end")

    def _preview_operation(self):
        """Prévisualise l'opération"""
        if not self.list_selector.is_loaded():
            messagebox.showwarning("Attention", "Veuillez charger un fichier liste")
            return

        df = self.list_selector.get_dataframe()
        col = self.path_col_combo.get()

        if col not in df.columns:
            messagebox.showwarning("Attention", "Colonne invalide")
            return

        dest = self.dest_path_var.get()

        # Créer un DataFrame de prévisualisation
        preview_data = []
        for idx, row in df.iterrows():
            source_path = str(row[col]) if pd.notna(row[col]) else ""

            if not source_path:
                continue

            source = Path(source_path)
            exists = source.exists()
            locked = source.name.startswith("~$") if exists else False

            if dest:
                if self.preserve_structure_var.get() and source.parent != Path("."):
                    dest_path = Path(dest) / source.name
                else:
                    dest_path = Path(dest) / source.name
            else:
                dest_path = "Non défini"

            status = "✓ OK" if exists and not locked else ("🔒 Verrouillé" if locked else "❓ Introuvable")

            preview_data.append({
                "Fichier": source.name,
                "Source": str(source.parent),
                "Destination": str(dest_path.parent) if isinstance(dest_path, Path) else dest_path,
                "Taille": f"{source.stat().st_size / 1024:.1f} Ko" if exists else "-",
                "Statut": status
            })

        preview_df = pd.DataFrame(preview_data)
        self.preview_table.load_data(preview_df)
        self.tabview.set("Prévisualisation")

        # Mettre à jour les stats
        total = len(preview_data)
        ok = sum(1 for d in preview_data if "✓" in d["Statut"])
        not_found = sum(1 for d in preview_data if "❓" in d["Statut"])
        locked = sum(1 for d in preview_data if "🔒" in d["Statut"])

        self.stats_cards.update_card("total", str(total))
        self.stats_cards.update_card("success", str(ok))
        self.stats_cards.update_card("not_found", str(not_found))

        self._log_message(f"Prévisualisation: {total} fichiers, {ok} disponibles, {not_found} introuvables")

    def _start_operation(self):
        """Démarre l'opération"""
        if not self.list_selector.is_loaded():
            messagebox.showwarning("Attention", "Veuillez charger un fichier liste")
            return

        dest = self.dest_path_var.get()
        if not dest:
            messagebox.showwarning("Attention", "Veuillez sélectionner un dossier de destination")
            return

        # Confirmation
        op_type = "déplacer" if self.operation_var.get() == "move" else "copier"
        df = self.list_selector.get_dataframe()
        col = self.path_col_combo.get()
        count = df[col].dropna().count()

        if not messagebox.askyesno(
            "Confirmation",
            f"Voulez-vous {op_type} {count} fichiers vers:\n{dest}?"
        ):
            return

        self._stop_event.clear()
        self._stats = OperationStats()

        self.execute_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")
        self.progress_bar.set(0)
        self.log_text.delete("1.0", "end")

        thread = threading.Thread(target=self._do_operation, daemon=True)
        thread.start()

    def _cancel_operation(self):
        """Annule l'opération en cours"""
        self._stop_event.set()
        self._log_message("Annulation demandée...", "warning")

    def _do_operation(self):
        """Effectue l'opération (dans un thread)"""
        df = self.list_selector.get_dataframe()
        col = self.path_col_combo.get()
        dest = Path(self.dest_path_var.get())
        operation = self.operation_var.get()
        conflict = self.conflict_var.get()
        ignore_locked = self.ignore_locked_var.get()

        # Créer le dossier de destination
        dest.mkdir(parents=True, exist_ok=True)

        paths = df[col].dropna().tolist()
        total = len(paths)
        self._stats.total = total

        self.frame.after(0, lambda: self._log_message(f"Démarrage: {total} fichiers à traiter"))

        for idx, path_str in enumerate(paths):
            if self._stop_event.is_set():
                self.frame.after(0, lambda: self._log_message("Opération annulée", "warning"))
                break

            source = Path(path_str)

            # Mise à jour progression
            progress = (idx + 1) / total
            self.frame.after(0, lambda p=progress: self.progress_bar.set(p))
            self.frame.after(0, lambda n=source.name: self.status_label.configure(text=f"Traitement: {n}"))

            # Vérifications
            if not source.exists():
                self._stats.not_found += 1
                self.frame.after(0, lambda n=source.name: self._log_message(f"Introuvable: {n}", "warning"))
                continue

            if ignore_locked and source.name.startswith("~$"):
                self._stats.skipped += 1
                self.frame.after(0, lambda n=source.name: self._log_message(f"Ignoré (verrouillé): {n}", "warning"))
                continue

            # Déterminer la destination
            dest_file = dest / source.name

            # Gestion des conflits
            if dest_file.exists():
                if conflict == "skip":
                    self._stats.skipped += 1
                    self.frame.after(0, lambda n=source.name: self._log_message(f"Ignoré (existe): {n}", "warning"))
                    continue
                elif conflict == "rename":
                    counter = 1
                    while dest_file.exists():
                        dest_file = dest / f"{source.stem}_copy{counter}{source.suffix}"
                        counter += 1

            # Effectuer l'opération
            try:
                if operation == "move":
                    shutil.move(str(source), str(dest_file))
                else:
                    shutil.copy2(str(source), str(dest_file))

                self._stats.success += 1
                self.frame.after(0, lambda n=source.name: self._log_message(f"Traité: {n}", "success"))

            except Exception as e:
                self._stats.errors += 1
                self.frame.after(0, lambda n=source.name, err=str(e): self._log_message(f"Erreur {n}: {err}", "error"))

        # Mise à jour finale
        self.frame.after(0, self._finish_operation)

    def _finish_operation(self):
        """Termine l'opération"""
        self.execute_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        self.progress_bar.set(1.0)

        # Mettre à jour les statistiques
        self.stats_cards.update_card("total", str(self._stats.total))
        self.stats_cards.update_card("success", str(self._stats.success))
        self.stats_cards.update_card("errors", str(self._stats.errors))
        self.stats_cards.update_card("not_found", str(self._stats.not_found))

        self._log_message(
            f"Terminé: {self._stats.success}/{self._stats.total} fichiers traités, "
            f"{self._stats.errors} erreurs, {self._stats.not_found} introuvables",
            "success" if self._stats.errors == 0 else "warning"
        )

        self.status_label.configure(text="Terminé")
        self.tabview.set("Journaux")

        # Créer le rapport si demandé
        if self.create_log_var.get():
            self._create_report()

    def _create_report(self):
        """Crée un rapport d'opération"""
        dest = Path(self.dest_path_var.get())
        report_path = dest / f"rapport_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"

        try:
            content = self.log_text.get("1.0", "end")
            report_path.write_text(content, encoding='utf-8')
            self._log_message(f"Rapport créé: {report_path.name}", "success")
        except Exception as e:
            self._log_message(f"Erreur création rapport: {e}", "error")

    def _copy_logs(self):
        """Copie les logs dans le presse-papier"""
        content = self.log_text.get("1.0", "end")
        self.frame.clipboard_clear()
        self.frame.clipboard_append(content)
        messagebox.showinfo("Info", "Logs copiés dans le presse-papier")

    def _export_logs(self):
        """Exporte les logs vers un fichier"""
        filepath = filedialog.asksaveasfilename(
            title="Exporter les logs",
            defaultextension=".txt",
            filetypes=[("Fichiers texte", "*.txt")]
        )
        if filepath:
            try:
                content = self.log_text.get("1.0", "end")
                Path(filepath).write_text(content, encoding='utf-8')
                messagebox.showinfo("Succès", f"Logs exportés vers:\n{filepath}")
            except Exception as e:
                messagebox.showerror("Erreur", str(e))

    def validate_inputs(self) -> tuple[bool, str]:
        return True, ""

    def _execute_task(self) -> Dict[str, Any]:
        return {}

    def update_status(self, message: str, level: str = "info"):
        color = {"info": COLORS["text_muted"], "success": COLORS["success"],
                 "warning": COLORS["warning"], "error": COLORS["error"]}.get(level, COLORS["text_muted"])
        self.status_label.configure(text=message, text_color=color)

    def update_progress(self, progress: float):
        self.progress_bar.set(progress)

    def reset(self):
        super().reset()
        self._stats = OperationStats()
        self.list_selector.reset()
        self.dest_path_var.set("")
        self.preview_table.clear()
        self.log_text.delete("1.0", "end")
        self.stats_cards.reset_all()
