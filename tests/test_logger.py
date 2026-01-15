"""
Tests unitaires pour le système de logging
"""

import pytest
import sys
from pathlib import Path
import tempfile
import os

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.core.logger import Logger, LogLevel, LogEntry, get_logger, set_logger


class TestLogLevel:
    """Tests pour LogLevel"""

    def test_log_level_values(self):
        """Test valeurs des niveaux"""
        assert LogLevel.DEBUG.name_str == "DEBUG"
        assert LogLevel.INFO.name_str == "INFO"
        assert LogLevel.ERROR.name_str == "ERROR"

    def test_log_level_colors(self):
        """Test couleurs des niveaux"""
        assert LogLevel.SUCCESS.color == "#00bf63"
        assert LogLevel.ERROR.color == "#ff6b6b"


class TestLogEntry:
    """Tests pour LogEntry"""

    def test_create_entry(self):
        """Test création entrée"""
        from datetime import datetime
        entry = LogEntry(
            timestamp=datetime.now(),
            level=LogLevel.INFO,
            message="Test message"
        )
        assert entry.message == "Test message"
        assert entry.level == LogLevel.INFO

    def test_format_entry(self):
        """Test formatage entrée"""
        from datetime import datetime
        entry = LogEntry(
            timestamp=datetime.now(),
            level=LogLevel.WARNING,
            message="Warning message",
            source="TestModule"
        )
        formatted = entry.format()

        assert "[WARNING]" in formatted
        assert "[TestModule]" in formatted
        assert "Warning message" in formatted

    def test_format_without_timestamp(self):
        """Test formatage sans horodatage"""
        from datetime import datetime
        entry = LogEntry(
            timestamp=datetime.now(),
            level=LogLevel.INFO,
            message="Test"
        )
        formatted = entry.format(include_timestamp=False)

        # Ne doit pas contenir le format de l'heure
        assert "[INFO]" in formatted


class TestLogger:
    """Tests pour Logger"""

    @pytest.fixture
    def temp_log_dir(self):
        """Crée un répertoire de logs temporaire"""
        tmpdir = tempfile.mkdtemp()
        yield Path(tmpdir)

        # Nettoyage avec gestion des erreurs Windows
        import shutil
        import gc
        gc.collect()  # Libérer les handles de fichiers
        try:
            if os.path.exists(tmpdir):
                shutil.rmtree(tmpdir, ignore_errors=True)
        except (PermissionError, OSError):
            pass  # Ignorer les erreurs de permission sur Windows

    def test_create_logger(self, temp_log_dir):
        """Test création logger"""
        logger = Logger(log_dir=temp_log_dir)
        assert logger is not None
        assert logger.name == "ExcelToolsPro"

    def test_log_directory_created(self, temp_log_dir):
        """Test création répertoire logs"""
        log_subdir = temp_log_dir / "sublogs"
        logger = Logger(log_dir=log_subdir)
        assert log_subdir.exists()

    def test_log_info(self, temp_log_dir):
        """Test log INFO"""
        logger = Logger(log_dir=temp_log_dir)
        logger.info("Test info message")

        assert len(logger.entries) == 1
        assert logger.entries[0].level == LogLevel.INFO
        assert logger.entries[0].message == "Test info message"

    def test_log_debug(self, temp_log_dir):
        """Test log DEBUG"""
        logger = Logger(log_dir=temp_log_dir)
        logger.debug("Debug message")

        assert logger.entries[0].level == LogLevel.DEBUG

    def test_log_success(self, temp_log_dir):
        """Test log SUCCESS"""
        logger = Logger(log_dir=temp_log_dir)
        logger.success("Success message")

        assert logger.entries[0].level == LogLevel.SUCCESS

    def test_log_warning(self, temp_log_dir):
        """Test log WARNING"""
        logger = Logger(log_dir=temp_log_dir)
        logger.warning("Warning message")

        assert logger.entries[0].level == LogLevel.WARNING
        assert logger.warning_count == 1

    def test_log_error(self, temp_log_dir):
        """Test log ERROR"""
        logger = Logger(log_dir=temp_log_dir)
        logger.error("Error message")

        assert logger.entries[0].level == LogLevel.ERROR
        assert logger.error_count == 1

    def test_log_critical(self, temp_log_dir):
        """Test log CRITICAL"""
        logger = Logger(log_dir=temp_log_dir)
        logger.critical("Critical message")

        assert logger.entries[0].level == LogLevel.CRITICAL
        assert logger.error_count == 1  # Critical compte comme erreur

    def test_log_with_source(self, temp_log_dir):
        """Test log avec source"""
        logger = Logger(log_dir=temp_log_dir)
        logger.info("Test", source="TestModule")

        assert logger.entries[0].source == "TestModule"

    def test_max_entries_limit(self, temp_log_dir):
        """Test limite d'entrées"""
        logger = Logger(log_dir=temp_log_dir, max_entries=5)

        for i in range(10):
            logger.info(f"Message {i}")

        assert len(logger.entries) == 5

    def test_callback_called(self, temp_log_dir):
        """Test appel callback"""
        logger = Logger(log_dir=temp_log_dir)
        callback_entries = []

        def callback(entry):
            callback_entries.append(entry)

        logger.add_callback(callback)
        logger.info("Test callback")

        assert len(callback_entries) == 1
        assert callback_entries[0].message == "Test callback"

    def test_remove_callback(self, temp_log_dir):
        """Test suppression callback"""
        logger = Logger(log_dir=temp_log_dir)
        callback_entries = []

        def callback(entry):
            callback_entries.append(entry)

        logger.add_callback(callback)
        logger.remove_callback(callback)
        logger.info("Test")

        assert len(callback_entries) == 0

    def test_clear_callbacks(self, temp_log_dir):
        """Test suppression tous callbacks"""
        logger = Logger(log_dir=temp_log_dir)

        logger.add_callback(lambda e: None)
        logger.add_callback(lambda e: None)
        logger.clear_callbacks()

        assert len(logger._callbacks) == 0

    def test_get_entries_filtered_by_level(self, temp_log_dir):
        """Test récupération entrées filtrées par niveau"""
        logger = Logger(log_dir=temp_log_dir)
        logger.info("Info 1")
        logger.warning("Warning 1")
        logger.info("Info 2")
        logger.error("Error 1")

        info_entries = logger.get_entries(level=LogLevel.INFO)
        assert len(info_entries) == 2

    def test_get_entries_filtered_by_source(self, temp_log_dir):
        """Test récupération entrées filtrées par source"""
        logger = Logger(log_dir=temp_log_dir)
        logger.info("Test 1", source="ModuleA")
        logger.info("Test 2", source="ModuleB")
        logger.info("Test 3", source="ModuleA")

        module_a_entries = logger.get_entries(source="ModuleA")
        assert len(module_a_entries) == 2

    def test_get_errors(self, temp_log_dir):
        """Test récupération erreurs"""
        logger = Logger(log_dir=temp_log_dir)
        logger.info("Info")
        logger.error("Error 1")
        logger.warning("Warning")
        logger.critical("Critical")

        errors = logger.get_errors()
        assert len(errors) == 2

    def test_get_warnings(self, temp_log_dir):
        """Test récupération avertissements"""
        logger = Logger(log_dir=temp_log_dir)
        logger.warning("Warning 1")
        logger.error("Error")
        logger.warning("Warning 2")

        warnings = logger.get_warnings()
        assert len(warnings) == 2

    def test_clear_logs(self, temp_log_dir):
        """Test effacement logs"""
        logger = Logger(log_dir=temp_log_dir)
        logger.info("Test 1")
        logger.error("Error")
        logger.warning("Warning")

        logger.clear()

        assert len(logger.entries) == 0
        assert logger.error_count == 0
        assert logger.warning_count == 0

    def test_log_file_created(self, temp_log_dir):
        """Test création fichier log"""
        logger = Logger(log_dir=temp_log_dir)
        logger.info("Test file creation")

        assert logger.log_file.exists()

    def test_export_logs(self, temp_log_dir):
        """Test export logs"""
        logger = Logger(log_dir=temp_log_dir)
        logger.info("Export test 1")
        logger.warning("Export test 2")

        export_path = temp_log_dir / "exported_logs.txt"
        result = logger.export_logs(export_path)

        assert result is True
        assert export_path.exists()

        with open(export_path, 'r', encoding='utf-8') as f:
            content = f.read()
        assert "Export test 1" in content

    def test_save_error_report(self, temp_log_dir):
        """Test sauvegarde rapport d'erreurs"""
        logger = Logger(log_dir=temp_log_dir)
        logger.error("Test error")
        logger.warning("Test warning")

        report_path = logger.save_error_report()
        assert report_path is not None
        assert report_path.exists()

    def test_save_error_report_empty(self, temp_log_dir):
        """Test rapport erreurs vide"""
        logger = Logger(log_dir=temp_log_dir)
        logger.info("Only info")

        report_path = logger.save_error_report()
        assert report_path is None


class TestGlobalLogger:
    """Tests pour le logger global"""

    def test_get_logger(self):
        """Test récupération logger global"""
        logger = get_logger()
        assert logger is not None

    def test_set_logger(self, temp_directory):
        """Test définition logger global"""
        custom_logger = Logger(log_dir=Path(temp_directory))
        set_logger(custom_logger)

        retrieved = get_logger()
        assert retrieved is custom_logger


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
