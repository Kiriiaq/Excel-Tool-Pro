"""
Tests unitaires pour le gestionnaire de configuration
"""

import pytest
import sys
from pathlib import Path
import tempfile
import os
import json

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.core.config import (
    ConfigManager, AppConfig, ExcelExportConfig, SearchConfig,
    MergeConfig, TransferConfig, CSVConfig, PerformanceConfig,
    UIConfig, LogConfig, ModuleConfig
)


class TestExcelExportConfig:
    """Tests pour ExcelExportConfig"""

    def test_default_values(self):
        """Test valeurs par défaut"""
        config = ExcelExportConfig()
        assert config.freeze_header is True
        assert config.auto_fit_columns is True
        assert config.header_bg_color == "#1F4E79"
        assert config.min_column_width == 10
        assert config.max_column_width == 50


class TestSearchConfig:
    """Tests pour SearchConfig"""

    def test_default_values(self):
        """Test valeurs par défaut"""
        config = SearchConfig()
        assert config.max_results_display == 500
        assert config.default_case_sensitive is False
        assert config.default_search_mode == "contains"


class TestAppConfig:
    """Tests pour AppConfig"""

    def test_default_values(self):
        """Test valeurs par défaut"""
        config = AppConfig()
        assert config.continue_on_error is False
        assert config.auto_save_config is True
        assert isinstance(config.excel_export, ExcelExportConfig)
        assert isinstance(config.search, SearchConfig)

    def test_nested_configs(self):
        """Test configurations imbriquées"""
        config = AppConfig()
        assert config.ui.theme == "dark"
        assert config.log.level == "INFO"


class TestConfigManager:
    """Tests pour ConfigManager"""

    @pytest.fixture
    def temp_config_file(self):
        """Crée un fichier de config temporaire"""
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as tmp:
            filepath = Path(tmp.name)

        yield filepath

        # Nettoyage
        if filepath.exists():
            filepath.unlink()

    def test_create_manager(self, temp_config_file):
        """Test création du manager"""
        manager = ConfigManager(config_path=temp_config_file)
        assert manager.config is not None

    def test_save_and_load(self, temp_config_file):
        """Test sauvegarde et chargement"""
        manager = ConfigManager(config_path=temp_config_file)
        manager.config.ui.theme = "light"
        manager.save()

        # Recharger
        manager2 = ConfigManager(config_path=temp_config_file)
        assert manager2.config.ui.theme == "light"

    def test_get_simple_value(self, temp_config_file):
        """Test récupération valeur simple"""
        manager = ConfigManager(config_path=temp_config_file)
        value = manager.get("continue_on_error")
        assert value is False

    def test_get_nested_value(self, temp_config_file):
        """Test récupération valeur imbriquée"""
        manager = ConfigManager(config_path=temp_config_file)
        value = manager.get("ui.theme")
        assert value == "dark"

    def test_get_default_value(self, temp_config_file):
        """Test valeur par défaut"""
        manager = ConfigManager(config_path=temp_config_file)
        value = manager.get("nonexistent.key", default="fallback")
        assert value == "fallback"

    def test_set_simple_value(self, temp_config_file):
        """Test définition valeur simple"""
        manager = ConfigManager(config_path=temp_config_file)
        manager.config.auto_save_config = False  # Désactiver pour le test

        manager.set("continue_on_error", True)
        assert manager.config.continue_on_error is True

    def test_set_nested_value(self, temp_config_file):
        """Test définition valeur imbriquée"""
        manager = ConfigManager(config_path=temp_config_file)
        manager.config.auto_save_config = False

        manager.set("ui.theme", "light")
        assert manager.config.ui.theme == "light"

    def test_module_config(self, temp_config_file):
        """Test configuration de module"""
        manager = ConfigManager(config_path=temp_config_file)
        manager.config.auto_save_config = False

        mod_config = manager.get_module_config("test_module")
        assert isinstance(mod_config, ModuleConfig)
        assert mod_config.enabled is True

    def test_set_module_setting(self, temp_config_file):
        """Test définition paramètre module"""
        manager = ConfigManager(config_path=temp_config_file)
        manager.config.auto_save_config = False

        manager.set_module_setting("test_module", "custom_key", "custom_value")
        mod_config = manager.get_module_config("test_module")

        assert mod_config.settings["custom_key"] == "custom_value"
        assert mod_config.last_used is not None

    def test_add_recent_file(self, temp_config_file):
        """Test ajout fichier récent"""
        manager = ConfigManager(config_path=temp_config_file)
        manager.config.auto_save_config = False

        manager.add_recent_file("/path/to/file1.xlsx")
        manager.add_recent_file("/path/to/file2.xlsx")

        assert "/path/to/file2.xlsx" in manager.config.recent_files
        assert manager.config.recent_files[0] == "/path/to/file2.xlsx"

    def test_recent_files_limit(self, temp_config_file):
        """Test limite fichiers récents"""
        manager = ConfigManager(config_path=temp_config_file)
        manager.config.auto_save_config = False
        manager.config.max_recent_files = 3

        for i in range(5):
            manager.add_recent_file(f"/path/to/file{i}.xlsx")

        assert len(manager.config.recent_files) == 3

    def test_reset_to_defaults(self, temp_config_file):
        """Test réinitialisation"""
        manager = ConfigManager(config_path=temp_config_file)
        manager.config.ui.theme = "light"
        manager.config.continue_on_error = True

        manager.reset_to_defaults()

        assert manager.config.ui.theme == "dark"
        assert manager.config.continue_on_error is False

    def test_reset_section(self, temp_config_file):
        """Test réinitialisation section"""
        manager = ConfigManager(config_path=temp_config_file)
        manager.config.auto_save_config = False
        manager.config.ui.theme = "light"
        manager.config.ui.font_size = 20

        manager.reset_section("ui")

        assert manager.config.ui.theme == "dark"
        assert manager.config.ui.font_size == 12

    def test_export_config(self, temp_config_file, temp_directory):
        """Test export configuration"""
        manager = ConfigManager(config_path=temp_config_file)
        export_path = Path(temp_directory) / "exported_config.json"

        result = manager.export_config(export_path)
        assert result is True
        assert export_path.exists()

        # Vérifier le contenu
        with open(export_path, 'r') as f:
            data = json.load(f)
        assert "ui" in data
        assert "excel_export" in data

    def test_import_config(self, temp_config_file, temp_directory):
        """Test import configuration"""
        # Créer un fichier de config à importer
        import_data = {
            "ui": {"theme": "light", "font_size": 14},
            "continue_on_error": True,
            "excel_export": {},
            "search": {},
            "merge": {},
            "transfer": {},
            "csv": {},
            "performance": {},
            "log": {},
            "modules": {}
        }
        import_path = Path(temp_directory) / "import_config.json"
        with open(import_path, 'w') as f:
            json.dump(import_data, f)

        manager = ConfigManager(config_path=temp_config_file)
        result = manager.import_config(import_path)

        assert result is True
        assert manager.config.ui.theme == "light"
        assert manager.config.continue_on_error is True

    def test_callback_on_change(self, temp_config_file):
        """Test callback lors de changement"""
        manager = ConfigManager(config_path=temp_config_file)
        manager.config.auto_save_config = False

        callback_data = {"called": False, "key": None, "value": None}

        def callback(key, value):
            callback_data["called"] = True
            callback_data["key"] = key
            callback_data["value"] = value

        manager.on_change(callback)
        manager.set("continue_on_error", True)

        assert callback_data["called"] is True
        assert callback_data["key"] == "continue_on_error"
        assert callback_data["value"] is True

    def test_get_all_settings_flat(self, temp_config_file):
        """Test récupération paramètres plats"""
        manager = ConfigManager(config_path=temp_config_file)
        flat = manager.get_all_settings_flat()

        assert "ui.theme" in flat
        assert "excel_export.freeze_header" in flat
        assert flat["ui.theme"] == "dark"

    def test_corrupted_config_file(self, temp_config_file):
        """Test fichier config corrompu"""
        # Écrire un JSON invalide
        with open(temp_config_file, 'w') as f:
            f.write("{invalid json}")

        manager = ConfigManager(config_path=temp_config_file)
        # Doit charger les valeurs par défaut
        assert manager.config.ui.theme == "dark"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
