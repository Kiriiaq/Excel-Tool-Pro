"""
Application principale ExcelToolsPro
Interface unifiée avec navigation par onglets/modules
Tous les paramètres sont exposés et modifiables via l'IHM
"""

import customtkinter as ctk
from tkinter import messagebox, filedialog
import webbrowser
from typing import Dict, Optional
from pathlib import Path

from ..core.config import ConfigManager
from ..core.logger import Logger, get_logger, set_logger, LogLevel
from ..core.constants import COLORS, APP_INFO

from .components.tooltip import Tooltip
from .components.log_viewer import LogViewer
from .components.settings_panel import SettingsPanel, SettingDefinition, SettingType

from ..modules.merge_module import MergeModule
from ..modules.file_search_module import FileSearchModule
from ..modules.data_transfer_module import DataTransferModule
from ..modules.csv_converter_module import CSVConverterModule
from ..modules.compare_module import CompareModule
from ..modules.vba_extractor_module import VBAExtractorModule
from ..modules.file_manager_module import FileManagerModule
from ..modules.table_copy_module import TableCopyModule


class ExcelToolsProApp(ctk.CTk):
    """
    Application principale ExcelToolsPro

    Interface moderne unifiée intégrant tous les modules:
    - Fusion de documents
    - Recherche de données
    - Transfert de données
    - Conversion CSV/Excel
    - Comparaison de données
    - Extraction code VBA
    - Gestionnaire de fichiers
    - Copie de tableaux
    """

    def __init__(self):
        super().__init__()

        # Initialisation des composants
        self.config_manager = ConfigManager()
        self.logger = Logger()
        set_logger(self.logger)

        # Appliquer la configuration de l'interface
        config = self.config_manager.config
        ctk.set_appearance_mode(config.ui.theme)
        ctk.set_default_color_theme("blue")

        # Configuration de base
        self.title(f"{APP_INFO['name']} - Suite d'outils Excel professionnelle")
        self.geometry(f"{config.ui.window_width}x{config.ui.window_height}")
        self.minsize(config.ui.window_min_width, config.ui.window_min_height)

        # Définir l'icône de l'application
        self._set_app_icon()

        # Dictionnaire des modules
        self.modules: Dict[str, any] = {}
        self.current_module: Optional[str] = None

        # Créer l'interface
        self._create_interface()

        # Charger les modules
        self._load_modules()

        # Binding de fermeture
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

        # Raccourcis clavier
        self.bind("<Control-comma>", lambda e: self._show_settings())
        self.bind("<F1>", lambda e: self._show_help())
        self.bind("<Control-q>", lambda e: self._on_closing())
        self.bind("<Control-r>", lambda e: self._reset_current_module())
        self.bind("<Control-1>", lambda e: self._switch_to_module(0))
        self.bind("<Control-2>", lambda e: self._switch_to_module(1))
        self.bind("<Control-3>", lambda e: self._switch_to_module(2))
        self.bind("<Control-4>", lambda e: self._switch_to_module(3))
        self.bind("<Control-5>", lambda e: self._switch_to_module(4))
        self.bind("<Control-6>", lambda e: self._switch_to_module(5))
        self.bind("<Control-7>", lambda e: self._switch_to_module(6))
        self.bind("<Control-8>", lambda e: self._switch_to_module(7))

        # Log de démarrage
        self.logger.info("Application démarrée")

    def _set_app_icon(self):
        """Définit l'icône de l'application pour la fenêtre et la barre des tâches"""
        import sys
        import os

        # Trouver le chemin de l'icône
        if getattr(sys, 'frozen', False):
            # Mode exécutable PyInstaller
            base_path = sys._MEIPASS
        else:
            # Mode développement
            base_path = Path(__file__).parent.parent.parent

        icon_path = Path(base_path) / "ico" / "icone.ico"

        if icon_path.exists():
            try:
                # Définir l'icône de la fenêtre
                self.iconbitmap(str(icon_path))

                # Pour Windows: définir l'AppUserModelID pour la barre des tâches
                if sys.platform == 'win32':
                    try:
                        import ctypes
                        myappid = 'edvance.exceltoolspro.app.1.0.0'
                        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
                    except Exception:
                        pass
            except Exception as e:
                print(f"Impossible de charger l'icône: {e}")

    def _create_interface(self):
        """Crée l'interface principale"""
        # === HEADER ===
        self._create_header()

        # === CONTENEUR PRINCIPAL ===
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        # Layout: Sidebar + Content
        sidebar_width = self.config_manager.config.ui.sidebar_width
        self.sidebar = ctk.CTkFrame(main_container, width=sidebar_width, fg_color=COLORS["bg_card"])
        self.sidebar.pack(side="left", fill="y", padx=(0, 10))
        self.sidebar.pack_propagate(False)

        self.content_area = ctk.CTkFrame(main_container, fg_color="transparent")
        self.content_area.pack(side="left", fill="both", expand=True)

        # === SIDEBAR ===
        self._create_sidebar()

        # === CONTENT AREA ===
        self._create_content_area()

    def _create_header(self):
        """Crée le header de l'application"""
        header = ctk.CTkFrame(self, fg_color=COLORS["accent_primary"], height=60, corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        # Logo et titre
        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.pack(side="left", padx=20)

        ctk.CTkLabel(
            title_frame,
            text="📊 ExcelToolsPro",
            font=("Segoe UI", 22, "bold"),
            text_color="white"
        ).pack(side="left")

        ctk.CTkLabel(
            title_frame,
            text=f"v{APP_INFO['version']}",
            font=("Segoe UI", 10),
            text_color="#a0a0a0"
        ).pack(side="left", padx=(10, 0), pady=(8, 0))

        # Boutons à droite
        btn_frame = ctk.CTkFrame(header, fg_color="transparent")
        btn_frame.pack(side="right", padx=20)

        # Bouton paramètres
        settings_btn = ctk.CTkButton(
            btn_frame,
            text="⚙️",
            width=40,
            height=32,
            fg_color="transparent",
            hover_color=COLORS["accent_secondary"],
            command=self._show_settings
        )
        settings_btn.pack(side="left", padx=5)
        Tooltip(settings_btn, "Paramètres (Ctrl+,)")

        # Bouton aide
        help_btn = ctk.CTkButton(
            btn_frame,
            text="❓",
            width=40,
            height=32,
            fg_color="transparent",
            hover_color=COLORS["accent_secondary"],
            command=self._show_help
        )
        help_btn.pack(side="left", padx=5)
        Tooltip(help_btn, "Aide (F1)")

    def _create_sidebar(self):
        """Crée la barre latérale de navigation"""
        # Titre
        ctk.CTkLabel(
            self.sidebar,
            text="MODULES",
            font=("Segoe UI", 11, "bold"),
            text_color=COLORS["text_muted"]
        ).pack(anchor="w", padx=15, pady=(20, 10))

        # Container des boutons de navigation
        self.nav_buttons_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.nav_buttons_frame.pack(fill="x", padx=10)

        self.nav_buttons: Dict[str, ctk.CTkButton] = {}

        # Séparateur
        sep = ctk.CTkFrame(self.sidebar, height=2, fg_color=COLORS["border_light"])
        sep.pack(fill="x", padx=15, pady=20)

        # Section Logs
        ctk.CTkLabel(
            self.sidebar,
            text="JOURNAL",
            font=("Segoe UI", 11, "bold"),
            text_color=COLORS["text_muted"]
        ).pack(anchor="w", padx=15, pady=(0, 10))

        log_height = self.config_manager.config.ui.log_viewer_height
        self.log_viewer = LogViewer(
            self.sidebar,
            max_entries=self.config_manager.config.log.max_entries,
            show_filters=False,
            show_search=False,
            height=log_height
        )
        self.log_viewer.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Connecter le logger au viewer
        self.logger.add_callback(self.log_viewer.add_entry)

    def _create_content_area(self):
        """Crée la zone de contenu principale"""
        # Frame pour les modules
        self.modules_frame = ctk.CTkFrame(self.content_area, fg_color="transparent")
        self.modules_frame.pack(fill="both", expand=True)

    def _load_modules(self):
        """Charge et initialise tous les modules"""
        module_classes = [
            MergeModule,
            FileSearchModule,
            DataTransferModule,
            CSVConverterModule,
            CompareModule,
            VBAExtractorModule,
            FileManagerModule,
            TableCopyModule,
        ]

        for module_class in module_classes:
            try:
                # Créer l'instance du module
                module = module_class(
                    self.modules_frame,
                    config_manager=self.config_manager,
                    logger=self.logger
                )

                module_id = module_class.MODULE_ID
                self.modules[module_id] = module

                # Créer le bouton de navigation
                self._create_nav_button(module_class)

                self.logger.debug(f"Module chargé: {module_class.MODULE_NAME}")

            except Exception as e:
                self.logger.error(f"Erreur chargement module {module_class.MODULE_NAME}: {e}")

        # Afficher le premier module par défaut
        if self.modules:
            first_module_id = list(self.modules.keys())[0]
            self._switch_module(first_module_id)

    def _create_nav_button(self, module_class):
        """Crée un bouton de navigation pour un module"""
        module_id = module_class.MODULE_ID

        btn = ctk.CTkButton(
            self.nav_buttons_frame,
            text=f"{module_class.MODULE_ICON} {module_class.MODULE_NAME}",
            font=("Segoe UI", 12),
            height=40,
            anchor="w",
            fg_color="transparent",
            text_color=COLORS["text_primary"],
            hover_color=COLORS["accent_primary"],
            command=lambda mid=module_id: self._switch_module(mid)
        )
        btn.pack(fill="x", pady=2)

        Tooltip(btn, module_class.MODULE_DESCRIPTION)
        self.nav_buttons[module_id] = btn

    def _switch_module(self, module_id: str):
        """Change le module affiché"""
        if module_id not in self.modules:
            return

        # Masquer le module actuel
        if self.current_module and self.current_module in self.modules:
            self.modules[self.current_module].hide()

        # Mettre à jour les boutons de navigation
        for mid, btn in self.nav_buttons.items():
            if mid == module_id:
                btn.configure(fg_color=COLORS["accent_primary"])
            else:
                btn.configure(fg_color="transparent")

        # Afficher le nouveau module
        self.modules[module_id].show()
        self.current_module = module_id

        self.logger.debug(f"Module activé: {module_id}")

    def _switch_to_module(self, index: int):
        """Change vers le module à l'index donné (pour raccourcis clavier)"""
        module_ids = list(self.modules.keys())
        if 0 <= index < len(module_ids):
            self._switch_module(module_ids[index])

    def _reset_current_module(self):
        """Réinitialise le module actuel (si la méthode reset existe)"""
        if self.current_module and self.current_module in self.modules:
            module = self.modules[self.current_module]
            if hasattr(module, 'reset') and callable(getattr(module, 'reset')):
                module.reset()
                self.logger.info(f"Module {self.current_module} réinitialisé")

    def _show_settings(self):
        """Affiche la fenêtre de paramètres complète"""
        settings_window = ctk.CTkToplevel(self)
        settings_window.title("Paramètres - ExcelToolsPro")
        settings_window.geometry("750x850")
        settings_window.transient(self)
        settings_window.grab_set()

        # Centrer
        settings_window.update_idletasks()
        x = (settings_window.winfo_screenwidth() // 2) - 375
        y = (settings_window.winfo_screenheight() // 2) - 425
        settings_window.geometry(f"+{x}+{y}")

        config = self.config_manager.config

        # Définitions complètes des paramètres
        settings_defs = [
            # ===== APPARENCE =====
            SettingDefinition(
                key="ui.theme",
                label="Thème",
                type=SettingType.CHOICE,
                default=config.ui.theme,
                choices=["dark", "light", "system"],
                tooltip="Thème de l'interface (sombre, clair ou système)",
                category="Apparence"
            ),
            SettingDefinition(
                key="ui.font_size",
                label="Taille de police",
                type=SettingType.SLIDER,
                default=config.ui.font_size,
                min_value=10,
                max_value=16,
                step=1,
                tooltip="Taille de la police de l'interface",
                category="Apparence"
            ),
            SettingDefinition(
                key="ui.enable_animations",
                label="Activer les animations",
                type=SettingType.BOOLEAN,
                default=config.ui.enable_animations,
                tooltip="Active les animations de l'interface",
                category="Apparence"
            ),

            # ===== COMPORTEMENT =====
            SettingDefinition(
                key="continue_on_error",
                label="Continuer après erreur",
                type=SettingType.BOOLEAN,
                default=config.continue_on_error,
                tooltip="Continue le traitement même en cas d'erreur sur un fichier",
                category="Comportement"
            ),
            SettingDefinition(
                key="confirm_before_overwrite",
                label="Confirmer avant écrasement",
                type=SettingType.BOOLEAN,
                default=config.confirm_before_overwrite,
                tooltip="Demande confirmation avant d'écraser un fichier existant",
                category="Comportement"
            ),
            SettingDefinition(
                key="auto_save_config",
                label="Sauvegarde auto des paramètres",
                type=SettingType.BOOLEAN,
                default=config.auto_save_config,
                tooltip="Sauvegarde automatiquement les paramètres à chaque modification",
                category="Comportement"
            ),

            # ===== EXPORT EXCEL =====
            SettingDefinition(
                key="excel_export.freeze_header",
                label="Geler l'en-tête",
                type=SettingType.BOOLEAN,
                default=config.excel_export.freeze_header,
                tooltip="Gèle la première ligne dans les exports Excel",
                category="Export Excel"
            ),
            SettingDefinition(
                key="excel_export.auto_fit_columns",
                label="Ajuster colonnes automatiquement",
                type=SettingType.BOOLEAN,
                default=config.excel_export.auto_fit_columns,
                tooltip="Ajuste automatiquement la largeur des colonnes",
                category="Export Excel"
            ),
            SettingDefinition(
                key="excel_export.alternate_row_colors",
                label="Alternance de couleurs",
                type=SettingType.BOOLEAN,
                default=config.excel_export.alternate_row_colors,
                tooltip="Alterne les couleurs de fond des lignes",
                category="Export Excel"
            ),
            SettingDefinition(
                key="excel_export.add_borders",
                label="Ajouter des bordures",
                type=SettingType.BOOLEAN,
                default=config.excel_export.add_borders,
                tooltip="Ajoute des bordures fines autour des cellules",
                category="Export Excel"
            ),
            SettingDefinition(
                key="excel_export.header_bg_color",
                label="Couleur en-tête",
                type=SettingType.COLOR,
                default=config.excel_export.header_bg_color,
                tooltip="Couleur de fond de l'en-tête",
                category="Export Excel"
            ),
            SettingDefinition(
                key="excel_export.min_column_width",
                label="Largeur min. colonnes",
                type=SettingType.INTEGER,
                default=config.excel_export.min_column_width,
                min_value=5,
                max_value=30,
                tooltip="Largeur minimale des colonnes en caractères",
                category="Export Excel",
                advanced=True
            ),
            SettingDefinition(
                key="excel_export.max_column_width",
                label="Largeur max. colonnes",
                type=SettingType.INTEGER,
                default=config.excel_export.max_column_width,
                min_value=20,
                max_value=100,
                tooltip="Largeur maximale des colonnes en caractères",
                category="Export Excel",
                advanced=True
            ),

            # ===== LOGS =====
            SettingDefinition(
                key="log.level",
                label="Niveau de log",
                type=SettingType.CHOICE,
                default=config.log.level,
                choices=["DEBUG", "INFO", "WARNING", "ERROR"],
                tooltip="Niveau minimum des logs affichés",
                category="Logs"
            ),
            SettingDefinition(
                key="log.max_entries",
                label="Nombre max. d'entrées",
                type=SettingType.INTEGER,
                default=config.log.max_entries,
                min_value=100,
                max_value=2000,
                tooltip="Nombre maximum d'entrées dans le journal",
                category="Logs"
            ),
            SettingDefinition(
                key="log.show_timestamps",
                label="Afficher les horodatages",
                type=SettingType.BOOLEAN,
                default=config.log.show_timestamps,
                tooltip="Affiche l'heure dans les entrées de log",
                category="Logs"
            ),

            # ===== CHEMINS =====
            SettingDefinition(
                key="default_output_dir",
                label="Dossier de sortie par défaut",
                type=SettingType.DIRECTORY,
                default=config.default_output_dir,
                tooltip="Dossier utilisé par défaut pour les exports",
                category="Chemins"
            ),
            SettingDefinition(
                key="default_input_dir",
                label="Dossier d'entrée par défaut",
                type=SettingType.DIRECTORY,
                default=config.default_input_dir,
                tooltip="Dossier utilisé par défaut pour ouvrir les fichiers",
                category="Chemins"
            ),

            # ===== RECHERCHE =====
            SettingDefinition(
                key="search.max_results_display",
                label="Résultats max. affichés",
                type=SettingType.INTEGER,
                default=config.search.max_results_display,
                min_value=100,
                max_value=5000,
                tooltip="Nombre maximum de résultats affichés",
                category="Recherche"
            ),
            SettingDefinition(
                key="search.default_search_mode",
                label="Mode de recherche par défaut",
                type=SettingType.CHOICE,
                default=config.search.default_search_mode,
                choices=["contains", "exact", "starts", "ends", "regex"],
                tooltip="Mode de recherche utilisé par défaut",
                category="Recherche"
            ),
            SettingDefinition(
                key="search.highlight_matches",
                label="Surligner les correspondances",
                type=SettingType.BOOLEAN,
                default=config.search.highlight_matches,
                tooltip="Surligne les termes trouvés dans les résultats",
                category="Recherche"
            ),

            # ===== FUSION =====
            SettingDefinition(
                key="merge.default_output_sheet_name",
                label="Nom onglet par défaut",
                type=SettingType.STRING,
                default=config.merge.default_output_sheet_name,
                tooltip="Nom par défaut de l'onglet créé lors d'une fusion",
                category="Fusion"
            ),
            SettingDefinition(
                key="merge.default_add_match_column",
                label="Ajouter colonne MATCH",
                type=SettingType.BOOLEAN,
                default=config.merge.default_add_match_column,
                tooltip="Ajoute par défaut une colonne indiquant les correspondances",
                category="Fusion"
            ),

            # ===== TRANSFERT =====
            SettingDefinition(
                key="transfer.max_rows_to_scan",
                label="Lignes max. à scanner",
                type=SettingType.INTEGER,
                default=config.transfer.max_rows_to_scan,
                min_value=50,
                max_value=1000,
                tooltip="Nombre maximum de lignes analysées pour trouver les données",
                category="Transfert",
                advanced=True
            ),
            SettingDefinition(
                key="transfer.default_output_sheet_name",
                label="Nom onglet par défaut",
                type=SettingType.STRING,
                default=config.transfer.default_output_sheet_name,
                tooltip="Nom par défaut de l'onglet créé",
                category="Transfert"
            ),

            # ===== CSV =====
            SettingDefinition(
                key="csv.default_encoding",
                label="Encodage par défaut",
                type=SettingType.CHOICE,
                default=config.csv.default_encoding,
                choices=config.csv.available_encodings,
                tooltip="Encodage utilisé par défaut pour les fichiers CSV",
                category="CSV"
            ),
            SettingDefinition(
                key="csv.default_separator",
                label="Séparateur par défaut",
                type=SettingType.CHOICE,
                default=config.csv.default_separator,
                choices=[",", ";", "\\t", "|"],
                tooltip="Caractère de séparation par défaut",
                category="CSV"
            ),

            # ===== PERFORMANCE =====
            SettingDefinition(
                key="performance.preview_max_rows",
                label="Lignes max. en aperçu",
                type=SettingType.INTEGER,
                default=config.performance.preview_max_rows,
                min_value=10,
                max_value=200,
                tooltip="Nombre de lignes affichées dans les aperçus",
                category="Performance",
                advanced=True
            ),
            SettingDefinition(
                key="performance.use_threading",
                label="Utiliser le multithreading",
                type=SettingType.BOOLEAN,
                default=config.performance.use_threading,
                tooltip="Exécute les tâches lourdes en arrière-plan",
                category="Performance",
                advanced=True
            ),

            # ===== AVANCÉ =====
            SettingDefinition(
                key="debug_mode",
                label="Mode debug",
                type=SettingType.BOOLEAN,
                default=config.debug_mode,
                tooltip="Active les informations de débogage détaillées",
                category="Avancé",
                advanced=True
            ),
            SettingDefinition(
                key="max_recent_files",
                label="Fichiers récents max.",
                type=SettingType.INTEGER,
                default=config.max_recent_files,
                min_value=5,
                max_value=50,
                tooltip="Nombre de fichiers récents à mémoriser",
                category="Avancé",
                advanced=True
            ),
        ]

        # Créer le panneau de paramètres
        settings_panel = SettingsPanel(
            settings_window,
            settings=settings_defs,
            on_change=self._on_setting_change,
            show_advanced=config.show_advanced_options
        )
        settings_panel.pack(fill="both", expand=True, padx=20, pady=(20, 10))

        # Boutons en bas
        btn_frame = ctk.CTkFrame(settings_window, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=(0, 20))

        # Export/Import config
        ctk.CTkButton(
            btn_frame,
            text="📤 Exporter",
            width=100,
            command=self._export_config
        ).pack(side="left", padx=(0, 10))

        ctk.CTkButton(
            btn_frame,
            text="📥 Importer",
            width=100,
            command=self._import_config
        ).pack(side="left")

        ctk.CTkButton(
            btn_frame,
            text="Fermer",
            width=100,
            command=settings_window.destroy
        ).pack(side="right")

    def _on_setting_change(self, key: str, value):
        """Callback lors d'un changement de paramètre"""
        self.config_manager.set(key, value)

        # Actions spécifiques selon le paramètre
        if key == "ui.theme":
            ctk.set_appearance_mode(value)
        elif key == "log.level":
            # Mettre à jour le niveau de log
            self.logger.set_level(value)
        elif key == "show_advanced_options":
            self.config_manager.config.show_advanced_options = value

        self.logger.debug(f"Paramètre modifié: {key} = {value}")

    def _export_config(self):
        """Exporte la configuration vers un fichier"""
        filepath = filedialog.asksaveasfilename(
            title="Exporter la configuration",
            defaultextension=".json",
            filetypes=[("Fichiers JSON", "*.json")]
        )
        if filepath:
            if self.config_manager.export_config(Path(filepath)):
                messagebox.showinfo("Succès", f"Configuration exportée vers:\n{filepath}")
            else:
                messagebox.showerror("Erreur", "Erreur lors de l'export")

    def _import_config(self):
        """Importe la configuration depuis un fichier"""
        filepath = filedialog.askopenfilename(
            title="Importer la configuration",
            filetypes=[("Fichiers JSON", "*.json")]
        )
        if filepath:
            if self.config_manager.import_config(Path(filepath)):
                messagebox.showinfo(
                    "Succès",
                    "Configuration importée.\nRedémarrez l'application pour appliquer tous les changements."
                )
            else:
                messagebox.showerror("Erreur", "Erreur lors de l'import")

    def _show_help(self):
        """Affiche la fenêtre d'aide"""
        help_window = ctk.CTkToplevel(self)
        help_window.title("Aide - ExcelToolsPro")
        help_window.geometry("700x600")
        help_window.transient(self)

        # Centrer
        help_window.update_idletasks()
        x = (help_window.winfo_screenwidth() // 2) - 350
        y = (help_window.winfo_screenheight() // 2) - 300
        help_window.geometry(f"+{x}+{y}")

        # Contenu
        scroll = ctk.CTkScrollableFrame(help_window, fg_color="transparent")
        scroll.pack(fill="both", expand=True, padx=20, pady=20)

        help_text = f"""
ExcelToolsPro - Suite d'outils Excel professionnelle
{'═' * 50}

VERSION: {APP_INFO['version']}
AUTEUR: {APP_INFO['author']}

MODULES DISPONIBLES:

🔗 FUSION DE DOCUMENTS (Ctrl+1)
   Fusionne deux fichiers Excel sur une colonne commune.
   - Chargez un fichier source et un fichier référence
   - Sélectionnez les colonnes clés de jointure
   - Prévisualisez et exportez le résultat

🔍 RECHERCHE DE DONNÉES (Ctrl+2)
   Recherche avancée dans les fichiers Excel.
   - Recherche simple avec plusieurs modes (contient, exact, regex, fuzzy)
   - Recherche par liste de mots avec statistiques
   - Opérateurs logiques (AND, OR)
   - Export des résultats et statistiques

📋 TRANSFERT DE DONNÉES (Ctrl+3)
   Extrait et transfère des données entre fichiers.
   - Définition de champs à extraire
   - Traitement par lots
   - Création de feuilles formatées

🔄 CONVERSION & FUSION (Ctrl+4)
   Convertit et fusionne des fichiers.
   - Conversion CSV ↔ Excel
   - Fusion de plusieurs fichiers
   - Exploration de fichiers Excel

⚖️ COMPARAISON DE DONNÉES (Ctrl+5)
   Compare des données entre fichiers Excel et documents.
   - Comparaison Excel ↔ Excel (colonnes sélectionnables)
   - Comparaison Excel ↔ PDF/Word (extraction de texte)
   - Mode exact ou approximatif (fuzzy matching)
   - Export des trouvés/non-trouvés

📜 EXTRACTION VBA (Ctrl+6)
   Extrait le code VBA des fichiers Excel.
   - Support .xlsm, .xls, .xlsb
   - Extraction des modules, classes et formulaires
   - Sauvegarde fichiers séparés + fichier combiné
   - Statistiques détaillées

📂 GESTIONNAIRE DE FICHIERS (Ctrl+7)
   Déplace/copie des fichiers selon une liste Excel.
   - Parcours récursif des dossiers source
   - Gestion des conflits (renommer, écraser, ignorer)
   - Prévisualisation avant exécution
   - Rapport détaillé des opérations

📊 COPIE DE TABLEAUX (Ctrl+8)
   Copie des tableaux Excel complets.
   - Détection automatique des en-têtes
   - Création de tableaux Excel natifs avec filtres
   - Réorganisation des colonnes
   - Traitement par lots (Traités/Non Traités)

RACCOURCIS CLAVIER:
   Ctrl+1 à 8 : Accès rapide aux modules
   Ctrl+R     : Réinitialiser le module actuel
   Ctrl+Q     : Quitter l'application
   Ctrl+,     : Paramètres
   F1         : Aide

PARAMÈTRES:
   Tous les paramètres sont accessibles via le bouton ⚙️
   dans le coin supérieur droit. Les options incluent:
   - Apparence (thème, taille de police)
   - Export Excel (formatage, couleurs, bordures)
   - Comportement (gestion des erreurs)
   - Performance (limites d'affichage)

SUPPORT:
   Pour signaler un bug ou demander une fonctionnalité,
   visitez: {APP_INFO['github']}
"""

        text_widget = ctk.CTkTextbox(scroll, font=("Segoe UI", 11), height=400)
        text_widget.pack(fill="both", expand=True)
        text_widget.insert("1.0", help_text)
        text_widget.configure(state="disabled")

        # Boutons
        btn_frame = ctk.CTkFrame(scroll, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(20, 0))

        ctk.CTkButton(
            btn_frame,
            text="🌐 GitHub",
            width=100,
            command=lambda: webbrowser.open(APP_INFO['github'])
        ).pack(side="left")

        ctk.CTkButton(
            btn_frame,
            text="Fermer",
            width=100,
            command=help_window.destroy
        ).pack(side="right")

    def _on_closing(self):
        """Gestionnaire de fermeture de l'application"""
        # Sauvegarder la configuration
        self.config_manager.save()

        # Log de fermeture
        self.logger.info("Application fermée")

        # Fermer
        self.destroy()


def main():
    """Point d'entrée principal"""
    try:
        app = ExcelToolsProApp()
        app.mainloop()
    except Exception as e:
        import traceback
        error_msg = f"""
Erreur de démarrage de ExcelToolsPro

{str(e)}

{traceback.format_exc()}

Vérifiez que toutes les dépendances sont installées:
pip install customtkinter pandas openpyxl
"""
        print(error_msg)

        # Afficher dans une fenêtre Tkinter basique
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Erreur de démarrage", error_msg)


if __name__ == "__main__":
    main()
