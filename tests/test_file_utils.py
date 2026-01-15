"""
Tests unitaires pour les utilitaires de fichiers
"""

import pytest
import sys
from pathlib import Path
import tempfile
import os
import shutil
from datetime import datetime, timedelta

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.utils.file_utils import FileUtils


class TestFileUtilsUniqueName:
    """Tests pour la génération de noms uniques"""

    def test_unique_name_no_conflict(self, temp_directory):
        """Test nom unique sans conflit"""
        filepath = os.path.join(temp_directory, "test.xlsx")
        result = FileUtils.get_unique_filename(filepath)
        assert result == filepath

    def test_unique_name_with_conflict(self, temp_directory):
        """Test nom unique avec conflit"""
        filepath = os.path.join(temp_directory, "test.xlsx")

        # Créer le fichier existant
        with open(filepath, 'w') as f:
            f.write("test")

        result = FileUtils.get_unique_filename(filepath)
        assert result != filepath
        assert "_copy1" in result

    def test_multiple_conflicts(self, temp_directory):
        """Test noms uniques avec multiples conflits"""
        base = os.path.join(temp_directory, "test.xlsx")

        # Créer plusieurs fichiers
        for i in range(3):
            filepath = base if i == 0 else os.path.join(temp_directory, f"test_copy{i}.xlsx")
            with open(filepath, 'w') as f:
                f.write("test")

        result = FileUtils.get_unique_filename(base)
        assert "_copy3" in result


class TestFileUtilsDirectory:
    """Tests pour la gestion des répertoires"""

    def test_ensure_directory_new(self, temp_directory):
        """Test création nouveau répertoire"""
        new_dir = os.path.join(temp_directory, "new_folder")
        result = FileUtils.ensure_directory(new_dir)
        assert result is True
        assert os.path.exists(new_dir)

    def test_ensure_directory_existing(self, temp_directory):
        """Test répertoire existant"""
        result = FileUtils.ensure_directory(temp_directory)
        assert result is True

    def test_ensure_directory_nested(self, temp_directory):
        """Test création répertoires imbriqués"""
        nested = os.path.join(temp_directory, "a", "b", "c")
        result = FileUtils.ensure_directory(nested)
        assert result is True
        assert os.path.exists(nested)


class TestFileUtilsMoveFile:
    """Tests pour le déplacement de fichiers"""

    def test_move_file_success(self, temp_directory):
        """Test déplacement réussi"""
        # Créer fichier source
        source = os.path.join(temp_directory, "source.txt")
        with open(source, 'w') as f:
            f.write("content")

        dest_dir = os.path.join(temp_directory, "dest")

        success, new_path, error = FileUtils.move_file(source, dest_dir)
        assert success is True
        assert new_path is not None
        assert os.path.exists(new_path)
        assert not os.path.exists(source)

    def test_move_file_create_dir(self, temp_directory):
        """Test déplacement avec création répertoire"""
        source = os.path.join(temp_directory, "file.txt")
        with open(source, 'w') as f:
            f.write("test")

        dest_dir = os.path.join(temp_directory, "new_dest")
        success, new_path, error = FileUtils.move_file(source, dest_dir, create_dir=True)

        assert success is True
        assert os.path.exists(dest_dir)


class TestFileUtilsCopyFile:
    """Tests pour la copie de fichiers"""

    def test_copy_file_success(self, temp_directory):
        """Test copie réussie"""
        source = os.path.join(temp_directory, "source.txt")
        with open(source, 'w') as f:
            f.write("content")

        dest_dir = os.path.join(temp_directory, "dest")

        success, new_path, error = FileUtils.copy_file(source, dest_dir)
        assert success is True
        assert os.path.exists(new_path)
        assert os.path.exists(source)  # Source toujours présente

    def test_copy_preserves_content(self, temp_directory):
        """Test préservation du contenu"""
        source = os.path.join(temp_directory, "source.txt")
        content = "Test content 123"
        with open(source, 'w') as f:
            f.write(content)

        dest_dir = os.path.join(temp_directory, "dest")
        success, new_path, error = FileUtils.copy_file(source, dest_dir)

        with open(new_path, 'r') as f:
            assert f.read() == content


class TestFileUtilsListFiles:
    """Tests pour le listage de fichiers"""

    def test_list_files_empty_dir(self, temp_directory):
        """Test répertoire vide"""
        files = FileUtils.list_files(temp_directory)
        assert len(files) == 0

    def test_list_files_with_extension_filter(self, temp_directory):
        """Test filtre par extension"""
        # Créer fichiers de différents types
        for ext in [".xlsx", ".txt", ".csv"]:
            with open(os.path.join(temp_directory, f"file{ext}"), 'w') as f:
                f.write("test")

        xlsx_files = FileUtils.list_files(temp_directory, extensions=[".xlsx"])
        assert len(xlsx_files) == 1
        assert xlsx_files[0].suffix == ".xlsx"

    def test_list_files_exclude_temp(self, temp_directory):
        """Test exclusion fichiers temporaires"""
        # Créer fichier normal et temporaire
        with open(os.path.join(temp_directory, "normal.xlsx"), 'w') as f:
            f.write("test")
        with open(os.path.join(temp_directory, "~$temp.xlsx"), 'w') as f:
            f.write("temp")

        files = FileUtils.list_files(temp_directory, exclude_temp=True)
        assert len(files) == 1
        assert "~$" not in files[0].name

    def test_list_files_recursive(self, temp_directory):
        """Test recherche récursive"""
        # Créer structure de répertoires
        subdir = os.path.join(temp_directory, "subdir")
        os.makedirs(subdir)

        with open(os.path.join(temp_directory, "file1.txt"), 'w') as f:
            f.write("1")
        with open(os.path.join(subdir, "file2.txt"), 'w') as f:
            f.write("2")

        # Non récursif
        files_non_recursive = FileUtils.list_files(temp_directory, recursive=False)
        assert len(files_non_recursive) == 1

        # Récursif
        files_recursive = FileUtils.list_files(temp_directory, recursive=True)
        assert len(files_recursive) == 2


class TestFileUtilsListExcelFiles:
    """Tests pour le listage de fichiers Excel"""

    def test_list_excel_files(self, temp_directory):
        """Test listage fichiers Excel"""
        # Créer fichiers Excel et non-Excel
        for ext in [".xlsx", ".xls", ".xlsm", ".txt", ".csv"]:
            with open(os.path.join(temp_directory, f"file{ext}"), 'w') as f:
                f.write("test")

        excel_files = FileUtils.list_excel_files(temp_directory)
        assert len(excel_files) == 3
        for f in excel_files:
            assert f.suffix in [".xlsx", ".xls", ".xlsm"]


class TestFileUtilsFileInfo:
    """Tests pour les informations de fichier"""

    def test_get_file_info_existing(self, temp_csv_file):
        """Test infos fichier existant"""
        info = FileUtils.get_file_info(temp_csv_file)

        assert info["exists"] is True
        assert info["extension"] == ".csv"
        assert info["size"] >= 0
        assert "size_formatted" in info

    def test_get_file_info_nonexistent(self):
        """Test infos fichier inexistant"""
        info = FileUtils.get_file_info("fichier_inexistant.txt")
        assert info["exists"] is False


class TestFileUtilsFormatSize:
    """Tests pour le formatage de taille"""

    def test_format_size_bytes(self):
        """Test formatage octets"""
        assert "B" in FileUtils.format_size(500)

    def test_format_size_kilobytes(self):
        """Test formatage kilo-octets"""
        assert "KB" in FileUtils.format_size(2048)

    def test_format_size_megabytes(self):
        """Test formatage méga-octets"""
        assert "MB" in FileUtils.format_size(2 * 1024 * 1024)


class TestFileUtilsValidation:
    """Tests pour la validation de chemins"""

    def test_validate_path_existing(self, temp_csv_file):
        """Test chemin valide"""
        valid, error = FileUtils.validate_path(temp_csv_file)
        assert valid is True
        assert error is None

    def test_validate_path_nonexistent(self):
        """Test chemin inexistant"""
        valid, error = FileUtils.validate_path("inexistant.txt")
        assert valid is False

    def test_validate_path_empty(self):
        """Test chemin vide"""
        valid, error = FileUtils.validate_path("")
        assert valid is False
        assert "vide" in error

    def test_validate_directory(self, temp_directory):
        """Test répertoire valide"""
        valid, error = FileUtils.validate_directory(temp_directory)
        assert valid is True


class TestFileUtilsBackup:
    """Tests pour les sauvegardes"""

    def test_create_backup(self, temp_directory):
        """Test création sauvegarde"""
        source = os.path.join(temp_directory, "original.txt")
        with open(source, 'w') as f:
            f.write("original content")

        success, backup_path, error = FileUtils.create_backup(source)
        assert success is True
        assert os.path.exists(backup_path)
        assert "_backup_" in backup_path

    def test_create_backup_custom_dir(self, temp_directory):
        """Test sauvegarde dans répertoire personnalisé"""
        source = os.path.join(temp_directory, "file.txt")
        with open(source, 'w') as f:
            f.write("content")

        backup_dir = os.path.join(temp_directory, "backups")

        success, backup_path, error = FileUtils.create_backup(source, backup_dir)
        assert success is True
        assert backup_dir in backup_path


class TestFileUtilsCleanOldFiles:
    """Tests pour le nettoyage de fichiers anciens"""

    def test_clean_old_files(self, temp_directory):
        """Test nettoyage fichiers anciens"""
        # Créer un fichier
        filepath = os.path.join(temp_directory, "old_file.txt")
        with open(filepath, 'w') as f:
            f.write("test")

        # Modifier la date pour simuler un ancien fichier
        old_time = (datetime.now() - timedelta(days=60)).timestamp()
        os.utime(filepath, (old_time, old_time))

        deleted = FileUtils.clean_old_files(temp_directory, max_age_days=30)
        assert deleted == 1
        assert not os.path.exists(filepath)


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
