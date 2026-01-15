"""
Module d'extraction de code VBA
Extrait le code VBA depuis les fichiers Excel (.xlsm, .xls, .xlsb)
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
from typing import Dict, Any, Optional, List
from pathlib import Path
import threading
from datetime import datetime

from .base_module import BaseModule
from ..ui.components.tooltip import Tooltip
from ..ui.components.stat_card import StatCardGroup
from ..core.constants import COLORS


class VBAExtractorModule(BaseModule):
    """
    Module d'extraction de code VBA depuis fichiers Excel

    Fonctionnalités:
    - Extraction VBA via Win32COM (Windows + Excel)
    - Extraction VBA via oletools (multiplateforme)
    - Export en fichiers individuels (.bas, .cls, .frm)
    - Création d'un fichier concaténé unique
    - Détection automatique des dépendances
    """

    MODULE_ID = "vba_extractor"
    MODULE_NAME = "Extraction VBA"
    MODULE_DESCRIPTION = "Extrait le code VBA depuis les fichiers Excel"
    MODULE_ICON = "📜"

    def __init__(self, *args, **kwargs):
        self._has_win32com = False
        self._has_oletools = False
        self._check_dependencies()
        super().__init__(*args, **kwargs)

    def _check_dependencies(self):
        """Vérifie les dépendances disponibles"""
        try:
            import win32com.client
            self._has_win32com = True
        except ImportError:
            self._has_win32com = False

        try:
            from oletools.olevba import VBA_Parser
            self._has_oletools = True
        except ImportError:
            self._has_oletools = False

    def _create_interface(self):
        """Crée l'interface du module"""
        main_scroll = ctk.CTkScrollableFrame(self.frame, fg_color="transparent")
        main_scroll.pack(fill="both", expand=True, padx=10, pady=10)

        # === Section État des dépendances ===
        deps_section = ctk.CTkFrame(main_scroll, fg_color=COLORS["bg_card"], corner_radius=10)
        deps_section.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            deps_section, text="🔧 DÉPENDANCES", font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        deps_frame = ctk.CTkFrame(deps_section, fg_color="transparent")
        deps_frame.pack(fill="x", padx=15, pady=(0, 15))

        # Win32COM
        win32_status = "✓ Disponible" if self._has_win32com else "✗ Non installé"
        win32_color = COLORS["success"] if self._has_win32com else COLORS["error"]
        ctk.CTkLabel(
            deps_frame, text=f"Win32COM: {win32_status}",
            text_color=win32_color, font=("Segoe UI", 10)
        ).pack(anchor="w")

        # Oletools
        ole_status = "✓ Disponible" if self._has_oletools else "✗ Non installé"
        ole_color = COLORS["success"] if self._has_oletools else COLORS["error"]
        ctk.CTkLabel(
            deps_frame, text=f"Oletools: {ole_status}",
            text_color=ole_color, font=("Segoe UI", 10)
        ).pack(anchor="w")

        if not self._has_win32com and not self._has_oletools:
            ctk.CTkLabel(
                deps_frame,
                text="⚠️ Installez win32com ou oletools pour utiliser ce module",
                text_color=COLORS["warning"], font=("Segoe UI", 10, "bold")
            ).pack(anchor="w", pady=(5, 0))

        # === Section Fichier source ===
        file_section = ctk.CTkFrame(main_scroll, fg_color=COLORS["bg_card"], corner_radius=10)
        file_section.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            file_section, text="📁 FICHIER SOURCE", font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        file_frame = ctk.CTkFrame(file_section, fg_color="transparent")
        file_frame.pack(fill="x", padx=15, pady=(0, 10))

        self.file_path_var = ctk.StringVar()
        ctk.CTkEntry(
            file_frame, textvariable=self.file_path_var, width=350, state="disabled"
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            file_frame, text="📁 Parcourir", width=100,
            command=self._browse_file
        ).pack(side="left")

        self.file_info_label = ctk.CTkLabel(
            file_section, text="", text_color=COLORS["text_muted"]
        )
        self.file_info_label.pack(anchor="w", padx=15, pady=(0, 15))

        # === Section Dossier de sortie ===
        output_section = ctk.CTkFrame(main_scroll, fg_color=COLORS["bg_card"], corner_radius=10)
        output_section.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            output_section, text="📂 DOSSIER DE SORTIE", font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        output_frame = ctk.CTkFrame(output_section, fg_color="transparent")
        output_frame.pack(fill="x", padx=15, pady=(0, 15))

        self.output_path_var = ctk.StringVar()
        ctk.CTkEntry(
            output_frame, textvariable=self.output_path_var, width=350, state="disabled"
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            output_frame, text="📂 Parcourir", width=100,
            command=self._browse_output
        ).pack(side="left")

        # === Section Options ===
        options_section = ctk.CTkFrame(main_scroll, fg_color=COLORS["bg_card"], corner_radius=10)
        options_section.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            options_section, text="⚙️ OPTIONS", font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        # Méthode d'extraction
        method_frame = ctk.CTkFrame(options_section, fg_color="transparent")
        method_frame.pack(fill="x", padx=15, pady=(0, 10))

        ctk.CTkLabel(method_frame, text="Méthode:", width=100).pack(side="left")

        self.method_var = ctk.StringVar(value="auto")
        methods = [
            ("Automatique", "auto"),
            ("Win32COM", "win32com"),
            ("Oletools", "oletools")
        ]

        for text, value in methods:
            state = "normal"
            if value == "win32com" and not self._has_win32com:
                state = "disabled"
            elif value == "oletools" and not self._has_oletools:
                state = "disabled"

            rb = ctk.CTkRadioButton(
                method_frame, text=text, variable=self.method_var, value=value, state=state
            )
            rb.pack(side="left", padx=10)

        # Options de sauvegarde
        save_frame = ctk.CTkFrame(options_section, fg_color="transparent")
        save_frame.pack(fill="x", padx=15, pady=(0, 15))

        self.save_individual_var = ctk.BooleanVar(value=True)
        cb1 = ctk.CTkCheckBox(
            save_frame, text="Sauvegarder fichiers individuels",
            variable=self.save_individual_var
        )
        cb1.pack(anchor="w", pady=2)
        Tooltip(cb1, "Sauvegarde chaque module dans un fichier séparé (.bas, .cls, .frm)")

        self.save_combined_var = ctk.BooleanVar(value=True)
        cb2 = ctk.CTkCheckBox(
            save_frame, text="Créer fichier combiné",
            variable=self.save_combined_var
        )
        cb2.pack(anchor="w", pady=2)
        Tooltip(cb2, "Crée un fichier unique contenant tout le code VBA")

        # === Section Actions ===
        action_frame = ctk.CTkFrame(main_scroll, fg_color="transparent")
        action_frame.pack(fill="x", pady=(0, 10))

        can_extract = self._has_win32com or self._has_oletools
        self.extract_btn = ctk.CTkButton(
            action_frame,
            text="📜 EXTRAIRE LE CODE VBA",
            font=("Segoe UI", 14, "bold"),
            height=50,
            fg_color=COLORS["success"] if can_extract else COLORS["text_muted"],
            state="normal" if can_extract else "disabled",
            command=self._start_extraction
        )
        self.extract_btn.pack(fill="x", pady=(0, 10))

        # Progression
        self.progress_bar = ctk.CTkProgressBar(action_frame, height=8)
        self.progress_bar.pack(fill="x")
        self.progress_bar.set(0)

        self.status_label = ctk.CTkLabel(
            action_frame, text="Prêt", font=("Segoe UI", 10), text_color=COLORS["text_muted"]
        )
        self.status_label.pack(anchor="w", pady=(5, 0))

        # === Section Statistiques ===
        self.stats_cards = StatCardGroup(main_scroll)
        self.stats_cards.pack(fill="x", pady=(0, 10))

        self.stats_cards.add_card("modules", "Modules", "0", "📦")
        self.stats_cards.add_card("classes", "Classes", "0", "📐", color=COLORS["info"])
        self.stats_cards.add_card("forms", "Formulaires", "0", "📋", color=COLORS["warning"])
        self.stats_cards.add_card("lines", "Lignes de code", "0", "📝", color=COLORS["success"])

        # === Section Log ===
        log_section = ctk.CTkFrame(main_scroll, fg_color=COLORS["bg_card"], corner_radius=10)
        log_section.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            log_section, text="📋 JOURNAL", font=("Segoe UI", 12, "bold")
        ).pack(anchor="w", padx=15, pady=(15, 10))

        self.log_text = ctk.CTkTextbox(
            log_section, height=200, font=("Consolas", 10),
            fg_color=("gray90", "gray20")
        )
        self.log_text.pack(fill="x", padx=15, pady=(0, 15))

    def _browse_file(self):
        """Sélectionne un fichier Excel avec macros"""
        filepath = filedialog.askopenfilename(
            title="Sélectionner un fichier Excel avec macros",
            filetypes=[
                ("Fichiers Excel avec macros", "*.xlsm *.xls *.xlsb"),
                ("Tous les fichiers Excel", "*.xlsx *.xlsm *.xls *.xlsb")
            ]
        )
        if filepath:
            self.file_path_var.set(filepath)
            p = Path(filepath)
            self.file_info_label.configure(
                text=f"Fichier: {p.name} ({p.suffix.upper()[1:]})",
                text_color=COLORS["success"]
            )

            # Suggérer un dossier de sortie
            if not self.output_path_var.get():
                output_dir = p.parent / f"{p.stem}_VBA"
                self.output_path_var.set(str(output_dir))

    def _browse_output(self):
        """Sélectionne le dossier de sortie"""
        folder = filedialog.askdirectory(title="Sélectionner le dossier de sortie")
        if folder:
            self.output_path_var.set(folder)

    def _log_message(self, message: str, level: str = "info"):
        """Ajoute un message au journal"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = {"info": "ℹ️", "success": "✓", "warning": "⚠️", "error": "✗"}.get(level, "")
        full_message = f"[{timestamp}] {prefix} {message}\n"

        self.log_text.insert("end", full_message)
        self.log_text.see("end")

    def _start_extraction(self):
        """Démarre l'extraction dans un thread"""
        filepath = self.file_path_var.get()
        output_dir = self.output_path_var.get()

        if not filepath:
            messagebox.showwarning("Attention", "Veuillez sélectionner un fichier")
            return

        if not output_dir:
            messagebox.showwarning("Attention", "Veuillez sélectionner un dossier de sortie")
            return

        self.extract_btn.configure(state="disabled")
        self.progress_bar.set(0)
        self.log_text.delete("1.0", "end")

        thread = threading.Thread(target=self._do_extraction, daemon=True)
        thread.start()

    def _do_extraction(self):
        """Effectue l'extraction (dans un thread)"""
        filepath = self.file_path_var.get()
        output_dir = Path(self.output_path_var.get())
        method = self.method_var.get()

        try:
            # Créer le dossier de sortie
            output_dir.mkdir(parents=True, exist_ok=True)

            self.frame.after(0, lambda: self._log_message(f"Extraction depuis: {Path(filepath).name}"))
            self.frame.after(0, lambda: self._log_message(f"Dossier de sortie: {output_dir}"))

            # Choisir la méthode
            if method == "auto":
                if self._has_win32com:
                    method = "win32com"
                elif self._has_oletools:
                    method = "oletools"
                else:
                    raise Exception("Aucune méthode d'extraction disponible")

            self.frame.after(0, lambda: self._log_message(f"Méthode: {method}"))
            self.frame.after(0, lambda: self.progress_bar.set(0.1))

            # Extraire le code
            if method == "win32com":
                modules = self._extract_with_win32com(filepath)
            else:
                modules = self._extract_with_oletools(filepath)

            self.frame.after(0, lambda: self.progress_bar.set(0.5))

            # Sauvegarder les modules
            stats = {"modules": 0, "classes": 0, "forms": 0, "lines": 0}
            combined_content = []

            for module_name, module_code, module_type in modules:
                # Compter les lignes
                lines = len(module_code.strip().split('\n'))
                stats["lines"] += lines

                # Déterminer l'extension et le type
                if module_type == "Class":
                    ext = ".cls"
                    stats["classes"] += 1
                elif module_type == "Form":
                    ext = ".frm"
                    stats["forms"] += 1
                else:
                    ext = ".bas"
                    stats["modules"] += 1

                # Sauvegarder individuellement
                if self.save_individual_var.get():
                    file_path = output_dir / f"{module_name}{ext}"
                    file_path.write_text(module_code, encoding='utf-8')
                    self.frame.after(0, lambda n=module_name: self._log_message(f"Sauvegardé: {n}", "success"))

                # Ajouter au fichier combiné
                if self.save_combined_var.get():
                    combined_content.append(f"{'=' * 60}")
                    combined_content.append(f"' MODULE: {module_name} ({module_type})")
                    combined_content.append(f"{'=' * 60}")
                    combined_content.append(module_code)
                    combined_content.append("")

            self.frame.after(0, lambda: self.progress_bar.set(0.8))

            # Sauvegarder le fichier combiné
            if self.save_combined_var.get() and combined_content:
                combined_path = output_dir / "ALL_VBA_CODE.txt"
                combined_path.write_text('\n'.join(combined_content), encoding='utf-8')
                self.frame.after(0, lambda: self._log_message(f"Fichier combiné: ALL_VBA_CODE.txt", "success"))

            self.frame.after(0, lambda: self.progress_bar.set(1.0))

            # Mettre à jour les statistiques
            self.frame.after(0, lambda: self._update_stats(stats))
            self.frame.after(0, lambda: self._log_message(
                f"Extraction terminée: {stats['modules']} modules, {stats['classes']} classes, "
                f"{stats['forms']} formulaires, {stats['lines']} lignes",
                "success"
            ))

            self.frame.after(0, lambda: messagebox.showinfo(
                "Succès",
                f"Extraction terminée!\n\n"
                f"Modules: {stats['modules']}\n"
                f"Classes: {stats['classes']}\n"
                f"Formulaires: {stats['forms']}\n"
                f"Lignes de code: {stats['lines']}\n\n"
                f"Dossier: {output_dir}"
            ))

        except Exception as e:
            self.frame.after(0, lambda: self._log_message(str(e), "error"))
            self.frame.after(0, lambda: messagebox.showerror("Erreur", str(e)))

        finally:
            self.frame.after(0, lambda: self.extract_btn.configure(state="normal"))
            self.frame.after(0, lambda: self.status_label.configure(text="Terminé"))

    def _extract_with_win32com(self, filepath: str) -> List[tuple]:
        """Extrait le code VBA avec Win32COM"""
        import win32com.client

        modules = []
        excel = None

        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            wb = excel.Workbooks.Open(filepath)

            try:
                vb_project = wb.VBProject

                for component in vb_project.VBComponents:
                    name = component.Name
                    code = component.CodeModule.Lines(1, component.CodeModule.CountOfLines)

                    # Déterminer le type
                    comp_type = component.Type
                    if comp_type == 1:  # vbext_ct_StdModule
                        module_type = "Module"
                    elif comp_type == 2:  # vbext_ct_ClassModule
                        module_type = "Class"
                    elif comp_type == 3:  # vbext_ct_MSForm
                        module_type = "Form"
                    else:
                        module_type = "Other"

                    if code.strip():
                        modules.append((name, code, module_type))

            finally:
                wb.Close(False)

        finally:
            if excel:
                excel.Quit()

        return modules

    def _extract_with_oletools(self, filepath: str) -> List[tuple]:
        """Extrait le code VBA avec oletools"""
        from oletools.olevba import VBA_Parser

        modules = []
        vba_parser = VBA_Parser(filepath)

        try:
            if vba_parser.detect_vba_macros():
                for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                    if vba_code and vba_code.strip():
                        # Déterminer le type par le nom ou le contenu
                        name = vba_filename or stream_path.split('/')[-1]

                        if "Class" in name or vba_code.startswith("VERSION 1.0 CLASS"):
                            module_type = "Class"
                        elif "Form" in name or "UserForm" in vba_code:
                            module_type = "Form"
                        else:
                            module_type = "Module"

                        modules.append((name, vba_code, module_type))

        finally:
            vba_parser.close()

        return modules

    def _update_stats(self, stats: Dict):
        """Met à jour les cartes de statistiques"""
        self.stats_cards.update_card("modules", str(stats["modules"]))
        self.stats_cards.update_card("classes", str(stats["classes"]))
        self.stats_cards.update_card("forms", str(stats["forms"]))
        self.stats_cards.update_card("lines", str(stats["lines"]))

    def validate_inputs(self) -> tuple[bool, str]:
        return True, ""

    def _execute_task(self) -> Dict[str, Any]:
        return {}

    def update_status(self, message: str, level: str = "info"):
        color = COLORS.get(level, COLORS["text_muted"])
        self.status_label.configure(text=message, text_color=color)

    def update_progress(self, progress: float):
        self.progress_bar.set(progress)

    def reset(self):
        super().reset()
        self.file_path_var.set("")
        self.output_path_var.set("")
        self.file_info_label.configure(text="")
        self.log_text.delete("1.0", "end")
        self.stats_cards.reset_all()
