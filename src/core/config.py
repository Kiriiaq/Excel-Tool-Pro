"""
Gestionnaire de configuration centralisé pour ExcelToolsPro
Sauvegarde/charge les préférences utilisateur et paramètres de l'application
Tous les paramètres sont exposés et modifiables via l'IHM
"""

import json
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import Dict, List, Any, Optional
from datetime import datetime


@dataclass
class ExcelExportConfig:
    """Configuration des exports Excel"""
    # Formatage
    freeze_header: bool = True
    auto_fit_columns: bool = True
    alternate_row_colors: bool = True
    add_borders: bool = True

    # Couleurs (format hex)
    header_bg_color: str = "#1F4E79"
    header_font_color: str = "#FFFFFF"
    alternate_row_color: str = "#F2F2F2"
    success_color: str = "#C6EFCE"
    error_color: str = "#FFC7CE"
    warning_color: str = "#FFEB9C"

    # Dimensions colonnes
    min_column_width: int = 10
    max_column_width: int = 50
    default_column_width: int = 15

    # Lignes scannées pour auto-fit (performance)
    autofit_sample_rows: int = 100

    # Police
    header_font_size: int = 11
    data_font_size: int = 10
    header_font_bold: bool = True


@dataclass
class SearchConfig:
    """Configuration du module de recherche"""
    # Limites
    max_results_display: int = 500
    max_column_name_length: int = 30

    # Comportement par défaut
    default_case_sensitive: bool = False
    default_and_mode: bool = False
    default_search_mode: str = "contains"  # contains, exact, starts, ends, regex

    # Highlight
    highlight_matches: bool = True


@dataclass
class MergeConfig:
    """Configuration du module de fusion"""
    # Colonnes clés suggérées (auto-détection)
    key_column_patterns: List[str] = field(default_factory=lambda: [
        "REF", "Référence", "Reference", "ID", "DOC_REF",
        "Legacy Number", "TC Reference", "Code"
    ])

    # Nom par défaut de l'onglet de sortie
    default_output_sheet_name: str = "Données_Fusionnées"

    # Options par défaut
    default_add_match_column: bool = True
    default_filter_last_only: bool = False
    default_export_matches_only: bool = False


@dataclass
class TransferConfig:
    """Configuration du module de transfert"""
    # Scan des fichiers
    max_rows_to_scan: int = 200
    adjacent_columns_to_check: int = 3

    # Feuille de sortie
    default_output_sheet_name: str = "Activité"

    # Styles
    header_title: str = "DONNÉES EXTRAITES"


@dataclass
class CSVConfig:
    """Configuration du module CSV"""
    # Encodages disponibles
    available_encodings: List[str] = field(default_factory=lambda: [
        "utf-8", "utf-8-sig", "latin-1", "cp1252", "iso-8859-1", "ascii"
    ])
    default_encoding: str = "utf-8"

    # Séparateurs
    available_separators: List[str] = field(default_factory=lambda: [
        ",", ";", "\\t", "|"
    ])
    default_separator: str = ","

    # Options
    default_skip_headers_on_merge: bool = True


@dataclass
class PerformanceConfig:
    """Configuration des performances"""
    # Preview
    preview_max_rows: int = 50
    preview_max_columns: int = 20

    # Traitement
    chunk_size: int = 10000
    use_threading: bool = True

    # Cache
    enable_file_cache: bool = True
    max_cache_size_mb: int = 100


@dataclass
class UIConfig:
    """Configuration de l'interface utilisateur"""
    # Thème
    theme: str = "dark"  # dark, light, system

    # Police
    font_family: str = "Segoe UI"
    font_size: int = 12
    monospace_font: str = "Consolas"

    # Fenêtre
    window_width: int = 1400
    window_height: int = 900
    window_min_width: int = 1200
    window_min_height: int = 800

    # Sidebar
    sidebar_width: int = 220
    log_viewer_height: int = 150

    # Animations
    enable_animations: bool = True
    tooltip_delay_ms: int = 500


@dataclass
class LogConfig:
    """Configuration des logs"""
    level: str = "INFO"  # DEBUG, INFO, WARNING, ERROR
    max_entries: int = 500
    show_timestamps: bool = True
    show_source: bool = True

    # Sauvegarde
    save_logs_to_file: bool = False
    log_file_path: str = ""
    max_log_file_size_mb: int = 10


@dataclass
class ModuleConfig:
    """Configuration d'un module spécifique"""
    enabled: bool = True
    last_used: Optional[str] = None
    settings: Dict[str, Any] = field(default_factory=dict)


@dataclass
class AppConfig:
    """Configuration globale de l'application"""
    # Sous-configurations
    excel_export: ExcelExportConfig = field(default_factory=ExcelExportConfig)
    search: SearchConfig = field(default_factory=SearchConfig)
    merge: MergeConfig = field(default_factory=MergeConfig)
    transfer: TransferConfig = field(default_factory=TransferConfig)
    csv: CSVConfig = field(default_factory=CSVConfig)
    performance: PerformanceConfig = field(default_factory=PerformanceConfig)
    ui: UIConfig = field(default_factory=UIConfig)
    log: LogConfig = field(default_factory=LogConfig)

    # Comportement global
    continue_on_error: bool = False
    show_advanced_options: bool = False
    auto_save_config: bool = True
    confirm_before_overwrite: bool = True

    # Chemins par défaut
    default_input_dir: str = ""
    default_output_dir: str = ""
    last_used_dir: str = ""

    # Configurations des modules
    modules: Dict[str, ModuleConfig] = field(default_factory=dict)

    # Historique
    recent_files: List[str] = field(default_factory=list)
    max_recent_files: int = 10

    # Debug
    debug_mode: bool = False


class ConfigManager:
    """Gestionnaire de configuration avec persistance JSON"""

    DEFAULT_CONFIG_PATH = Path.home() / ".exceltoolspro" / "config.json"

    def __init__(self, config_path: Optional[Path] = None):
        self.config_path = config_path or self.DEFAULT_CONFIG_PATH
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        self._config: Optional[AppConfig] = None
        self._callbacks: List[callable] = []

    @property
    def config(self) -> AppConfig:
        """Accès à la configuration, charge si nécessaire"""
        if self._config is None:
            self.load()
        return self._config

    def load(self) -> AppConfig:
        """Charge la configuration depuis le fichier JSON"""
        if self.config_path.exists():
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                self._config = self._dict_to_config(data)

            except (json.JSONDecodeError, TypeError, KeyError) as e:
                print(f"Erreur de configuration, utilisation des valeurs par défaut: {e}")
                self._config = AppConfig()
        else:
            self._config = AppConfig()

        return self._config

    def _dict_to_config(self, data: dict) -> AppConfig:
        """Convertit un dictionnaire en AppConfig avec sous-configs"""
        # Extraire les sous-configurations
        excel_export_data = data.pop('excel_export', {})
        search_data = data.pop('search', {})
        merge_data = data.pop('merge', {})
        transfer_data = data.pop('transfer', {})
        csv_data = data.pop('csv', {})
        performance_data = data.pop('performance', {})
        ui_data = data.pop('ui', {})
        log_data = data.pop('log', {})
        modules_data = data.pop('modules', {})

        # Créer les sous-configs
        excel_export = ExcelExportConfig(**excel_export_data) if excel_export_data else ExcelExportConfig()
        search = SearchConfig(**search_data) if search_data else SearchConfig()
        merge = MergeConfig(**merge_data) if merge_data else MergeConfig()
        transfer = TransferConfig(**transfer_data) if transfer_data else TransferConfig()
        csv = CSVConfig(**csv_data) if csv_data else CSVConfig()
        performance = PerformanceConfig(**performance_data) if performance_data else PerformanceConfig()
        ui = UIConfig(**ui_data) if ui_data else UIConfig()
        log = LogConfig(**log_data) if log_data else LogConfig()

        # Convertir les modules
        modules = {
            name: ModuleConfig(**mod_data) if isinstance(mod_data, dict) else mod_data
            for name, mod_data in modules_data.items()
        }

        return AppConfig(
            excel_export=excel_export,
            search=search,
            merge=merge,
            transfer=transfer,
            csv=csv,
            performance=performance,
            ui=ui,
            log=log,
            modules=modules,
            **data
        )

    def save(self) -> bool:
        """Sauvegarde la configuration dans le fichier JSON"""
        try:
            data = self._config_to_dict(self.config)

            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False, default=str)

            return True
        except Exception as e:
            print(f"Erreur lors de la sauvegarde de la configuration: {e}")
            return False

    def _config_to_dict(self, config: AppConfig) -> dict:
        """Convertit une AppConfig en dictionnaire"""
        return asdict(config)

    def get(self, key: str, default: Any = None) -> Any:
        """Récupère une valeur de configuration par clé (supporte la notation pointée)"""
        keys = key.split('.')
        value = self.config

        for k in keys:
            if hasattr(value, k):
                value = getattr(value, k)
            elif isinstance(value, dict) and k in value:
                value = value[k]
            else:
                return default

        return value

    def set(self, key: str, value: Any) -> None:
        """Définit une valeur de configuration (supporte la notation pointée)"""
        keys = key.split('.')

        if len(keys) == 1:
            if hasattr(self.config, key):
                setattr(self.config, key, value)
        else:
            # Navigation vers le bon sous-objet
            obj = self.config
            for k in keys[:-1]:
                if hasattr(obj, k):
                    obj = getattr(obj, k)
                else:
                    return

            if hasattr(obj, keys[-1]):
                setattr(obj, keys[-1], value)

        if self.config.auto_save_config:
            self.save()
        self._notify_change(key, value)

    def get_module_config(self, module_name: str) -> ModuleConfig:
        """Récupère la configuration d'un module spécifique"""
        if module_name not in self.config.modules:
            self.config.modules[module_name] = ModuleConfig()
        return self.config.modules[module_name]

    def set_module_setting(self, module_name: str, key: str, value: Any) -> None:
        """Définit un paramètre pour un module spécifique"""
        mod_config = self.get_module_config(module_name)
        mod_config.settings[key] = value
        mod_config.last_used = datetime.now().isoformat()
        if self.config.auto_save_config:
            self.save()

    def add_recent_file(self, file_path: str) -> None:
        """Ajoute un fichier à l'historique des fichiers récents"""
        if file_path in self.config.recent_files:
            self.config.recent_files.remove(file_path)

        self.config.recent_files.insert(0, file_path)

        # Limiter la taille de l'historique
        if len(self.config.recent_files) > self.config.max_recent_files:
            self.config.recent_files = self.config.recent_files[:self.config.max_recent_files]

        if self.config.auto_save_config:
            self.save()

    def on_change(self, callback: callable) -> None:
        """Enregistre un callback appelé lors d'un changement de configuration"""
        self._callbacks.append(callback)

    def _notify_change(self, key: str, value: Any) -> None:
        """Notifie les callbacks d'un changement"""
        for callback in self._callbacks:
            try:
                callback(key, value)
            except Exception as e:
                print(f"Erreur dans le callback de configuration: {e}")

    def reset_to_defaults(self) -> None:
        """Réinitialise la configuration aux valeurs par défaut"""
        self._config = AppConfig()
        self.save()

    def reset_section(self, section: str) -> None:
        """Réinitialise une section spécifique de la configuration"""
        defaults = {
            'excel_export': ExcelExportConfig(),
            'search': SearchConfig(),
            'merge': MergeConfig(),
            'transfer': TransferConfig(),
            'csv': CSVConfig(),
            'performance': PerformanceConfig(),
            'ui': UIConfig(),
            'log': LogConfig(),
        }

        if section in defaults:
            setattr(self.config, section, defaults[section])
            if self.config.auto_save_config:
                self.save()

    def export_config(self, export_path: Path) -> bool:
        """Exporte la configuration vers un fichier externe"""
        try:
            data = self._config_to_dict(self.config)
            with open(export_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False, default=str)
            return True
        except Exception as e:
            print(f"Erreur lors de l'export: {e}")
            return False

    def import_config(self, import_path: Path) -> bool:
        """Importe la configuration depuis un fichier externe"""
        try:
            with open(import_path, 'r', encoding='utf-8') as f:
                data = json.load(f)

            self._config = self._dict_to_config(data)
            self.save()
            return True
        except Exception as e:
            print(f"Erreur lors de l'import: {e}")
            return False

    def get_all_settings_flat(self) -> Dict[str, Any]:
        """Retourne tous les paramètres sous forme de dictionnaire plat"""
        result = {}

        def flatten(obj, prefix=''):
            if hasattr(obj, '__dataclass_fields__'):
                for field_name in obj.__dataclass_fields__:
                    value = getattr(obj, field_name)
                    key = f"{prefix}.{field_name}" if prefix else field_name
                    if hasattr(value, '__dataclass_fields__'):
                        flatten(value, key)
                    else:
                        result[key] = value

        flatten(self.config)
        return result
