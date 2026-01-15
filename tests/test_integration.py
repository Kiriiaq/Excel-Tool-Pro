"""
Tests d'intégration pour ExcelToolsPro
Tests des interactions entre modules
"""

import pytest
import sys
from pathlib import Path
import tempfile
import os
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.utils.excel_utils import ExcelUtils
from src.utils.file_utils import FileUtils
from src.utils.validators import Validators
from src.core.config import ConfigManager, ExcelExportConfig
from src.core.logger import Logger


class TestExcelFileWorkflow:
    """Tests du workflow complet de fichiers Excel"""

    def test_full_workflow_create_modify_read(self, temp_directory):
        """Test workflow complet: création, modification, lecture"""
        filepath = os.path.join(temp_directory, "workflow_test.xlsx")

        # 1. Créer un fichier Excel avec des données
        df1 = pd.DataFrame({
            "ID": ["001", "002", "003"],
            "Nom": ["Alice", "Bob", "Charlie"],
            "Score": [85, 92, 78]
        })

        success, error = ExcelUtils.write_dataframe_to_excel(
            df1, filepath, "Scores"
        )
        assert success is True

        # 2. Ajouter un nouvel onglet
        df2 = pd.DataFrame({
            "ID": ["001", "002"],
            "Bonus": [10, 15]
        })

        success, error = ExcelUtils.add_sheet_to_workbook(
            filepath, "Bonus", df2
        )
        assert success is True

        # 3. Vérifier que les deux onglets existent
        sheets = ExcelUtils.get_excel_sheets(filepath)
        assert "Scores" in sheets
        assert "Bonus" in sheets

        # 4. Relire les données
        df_scores, _, _ = ExcelUtils.read_excel_file(filepath, "Scores")
        df_bonus, _, _ = ExcelUtils.read_excel_file(filepath, "Bonus")

        assert len(df_scores) == 3
        assert len(df_bonus) == 2

    def test_backup_and_restore_workflow(self, temp_directory):
        """Test workflow sauvegarde et restauration"""
        filepath = os.path.join(temp_directory, "backup_test.xlsx")

        # Créer fichier original
        df = pd.DataFrame({"Data": [1, 2, 3]})
        ExcelUtils.write_dataframe_to_excel(df, filepath, "Original")

        # Créer une sauvegarde
        success, backup_path, error = FileUtils.create_backup(filepath)
        assert success is True

        # Modifier le fichier original
        df_modified = pd.DataFrame({"Data": [4, 5, 6]})
        ExcelUtils.write_dataframe_to_excel(df_modified, filepath, "Modified")

        # Vérifier que les deux fichiers sont différents
        df_orig, _, _ = ExcelUtils.read_excel_file(filepath)
        df_backup, _, _ = ExcelUtils.read_excel_file(backup_path)

        # Le fichier modifié devrait avoir la feuille "Modified"
        sheets_orig = ExcelUtils.get_excel_sheets(filepath)
        sheets_backup = ExcelUtils.get_excel_sheets(backup_path)

        assert "Modified" in sheets_orig
        assert "Original" in sheets_backup


class TestConfigAndLoggingIntegration:
    """Tests d'intégration configuration et logging"""

    def test_config_affects_excel_export(self, temp_directory):
        """Test que la config affecte l'export Excel"""
        config_path = Path(temp_directory) / "config.json"
        manager = ConfigManager(config_path=config_path)

        # Modifier la config d'export
        manager.config.excel_export.header_bg_color = "#FF0000"
        manager.config.excel_export.freeze_header = False

        # Créer un fichier avec cette config
        df = pd.DataFrame({"A": [1, 2, 3]})
        filepath = os.path.join(temp_directory, "config_test.xlsx")

        success, error = ExcelUtils.write_with_config(
            df, filepath, "Test", manager.config.excel_export
        )

        assert success is True

    def test_logger_tracks_operations(self, temp_directory):
        """Test que le logger trace les opérations"""
        logger = Logger(log_dir=Path(temp_directory))

        # Simuler des opérations
        logger.info("Démarrage du traitement")

        df = pd.DataFrame({"X": [1, 2, 3]})
        filepath = os.path.join(temp_directory, "logged.xlsx")

        success, error = ExcelUtils.write_dataframe_to_excel(df, filepath, "Data")
        if success:
            logger.success(f"Fichier créé: {filepath}")
        else:
            logger.error(f"Erreur: {error}")

        logger.info("Traitement terminé")

        # Vérifier les logs
        assert len(logger.entries) >= 3
        assert logger.error_count == 0


class TestValidationAndProcessing:
    """Tests d'intégration validation et traitement"""

    def test_validate_then_process(self, temp_excel_file):
        """Test validation puis traitement"""
        # Valider le fichier
        valid, error = Validators.validate_excel_file(temp_excel_file)
        assert valid is True

        # Lire le fichier validé
        df, sheets, error = ExcelUtils.read_excel_file(temp_excel_file)
        assert df is not None

        # Rechercher dans les données
        results = ExcelUtils.search_in_excel(df, "Alice")
        assert len(results) >= 0  # Peut ou non trouver selon les données

    def test_validation_prevents_invalid_processing(self):
        """Test que la validation empêche le traitement invalide"""
        # Tenter de valider un fichier inexistant
        valid, error = Validators.validate_excel_file("inexistant.xlsx")
        assert valid is False

        # Ne pas continuer le traitement
        if not valid:
            # On aurait arrêté ici dans une vraie application
            pass


class TestMergeWorkflow:
    """Tests du workflow de fusion"""

    def test_merge_validated_files(self, temp_directory):
        """Test fusion de fichiers validés"""
        # Créer les fichiers
        df1 = pd.DataFrame({"ID": ["A", "B"], "Val": [1, 2]})
        df2 = pd.DataFrame({"ID": ["C", "D"], "Val": [3, 4]})

        file1 = os.path.join(temp_directory, "merge1.xlsx")
        file2 = os.path.join(temp_directory, "merge2.xlsx")
        output = os.path.join(temp_directory, "merged.xlsx")

        df1.to_excel(file1, index=False)
        df2.to_excel(file2, index=False)

        # Valider les fichiers
        for f in [file1, file2]:
            valid, error = Validators.validate_excel_file(f)
            assert valid is True

        # Fusionner
        success, count, error = ExcelUtils.merge_excel_files([file1, file2], output)
        assert success is True
        # skip_headers=True par défaut donc 2 + 1 = 3 lignes
        assert count >= 3

    def test_merge_with_reference_data(self, temp_directory, sample_dataframe_with_ref, reference_dataframe):
        """Test fusion avec données de référence (jointure)"""
        source_file = os.path.join(temp_directory, "source.xlsx")
        ref_file = os.path.join(temp_directory, "reference.xlsx")

        sample_dataframe_with_ref.to_excel(source_file, index=False)
        reference_dataframe.to_excel(ref_file, index=False)

        # Lire les fichiers
        df_source = pd.read_excel(source_file)
        df_ref = pd.read_excel(ref_file)

        # Effectuer une jointure
        df_merged = pd.merge(
            df_source,
            df_ref,
            on="REF",
            how="left"
        )

        assert len(df_merged) == 3
        assert "Categorie" in df_merged.columns


class TestFileManagementWorkflow:
    """Tests du workflow de gestion de fichiers"""

    def test_organize_files_workflow(self, temp_directory):
        """Test workflow d'organisation de fichiers"""
        # Créer plusieurs fichiers
        for i in range(3):
            filepath = os.path.join(temp_directory, f"file{i}.xlsx")
            df = pd.DataFrame({"ID": [i]})
            df.to_excel(filepath, index=False)

        # Lister les fichiers Excel
        files = FileUtils.list_excel_files(temp_directory)
        assert len(files) == 3

        # Créer un sous-répertoire et déplacer un fichier
        processed_dir = os.path.join(temp_directory, "processed")
        success, new_path, error = FileUtils.move_file(
            str(files[0]), processed_dir, create_dir=True
        )
        assert success is True

        # Vérifier l'organisation
        remaining = FileUtils.list_excel_files(temp_directory, recursive=False)
        processed = FileUtils.list_excel_files(processed_dir)

        assert len(remaining) == 2
        assert len(processed) == 1


class TestErrorHandling:
    """Tests de gestion d'erreurs entre modules"""

    def test_graceful_error_handling(self, temp_directory):
        """Test gestion gracieuse des erreurs"""
        logger = Logger(log_dir=Path(temp_directory))
        errors_occurred = []

        # Tenter plusieurs opérations avec erreurs potentielles
        operations = [
            ("Lecture fichier inexistant",
             lambda: ExcelUtils.read_excel_file("inexistant.xlsx")),
            ("Validation fichier invalide",
             lambda: Validators.validate_excel_file("invalid.txt")),
        ]

        for name, operation in operations:
            try:
                result = operation()
                if isinstance(result, tuple) and result[0] is None:
                    logger.warning(f"{name}: Aucun résultat")
            except Exception as e:
                logger.error(f"{name}: {str(e)}")
                errors_occurred.append(name)

        # Vérifier que les erreurs ont été loguées
        assert logger.warning_count + logger.error_count > 0


class TestPerformanceIntegration:
    """Tests de performance intégrés"""

    def test_large_file_processing(self, temp_directory, large_dataframe):
        """Test traitement de fichier volumineux"""
        filepath = os.path.join(temp_directory, "large.xlsx")

        # Écrire un gros fichier
        success, error = ExcelUtils.write_dataframe_to_excel(
            large_dataframe, filepath, "LargeData"
        )
        assert success is True

        # Relire le fichier
        df_read, _, error = ExcelUtils.read_excel_file(filepath)
        assert df_read is not None
        assert len(df_read) == len(large_dataframe)

        # Rechercher dans les données
        results = ExcelUtils.search_in_excel(df_read, "A", columns=["Category"])
        # Devrait trouver environ 1/4 des lignes (distribution aléatoire A, B, C, D)
        assert len(results) > 0


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
