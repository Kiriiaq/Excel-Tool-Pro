"""
Tests pour les utilitaires Excel
"""

import pytest
import pandas as pd
from pathlib import Path
import tempfile
import os

# Ajuster le path pour les imports
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from src.utils.excel_utils import ExcelUtils


class TestExcelUtils:
    """Tests pour la classe ExcelUtils"""

    def test_read_excel_file_not_found(self):
        """Test lecture fichier inexistant"""
        df, sheets, error = ExcelUtils.read_excel_file("fichier_inexistant.xlsx")
        assert df is None
        assert error is not None

    def test_write_and_read_excel(self, temp_directory):
        """Test écriture puis lecture d'un fichier Excel"""
        # Créer un DataFrame de test
        df = pd.DataFrame({
            "Nom": ["Alice", "Bob", "Charlie"],
            "Age": [25, 30, 35],
            "Ville": ["Paris", "Lyon", "Marseille"]
        })

        # Utiliser le répertoire temporaire de la fixture
        tmp_path = os.path.join(temp_directory, "test_write_read.xlsx")

        # Écrire
        success, error = ExcelUtils.write_dataframe_to_excel(
            df, tmp_path, "TestSheet"
        )
        assert success is True
        assert error is None

        # Relire
        df_read, sheets, error = ExcelUtils.read_excel_file(tmp_path)
        assert df_read is not None
        assert error is None
        assert len(df_read) == 3
        assert "TestSheet" in sheets

    def test_get_excel_sheets(self, temp_directory):
        """Test récupération des feuilles"""
        # Créer un fichier avec plusieurs feuilles
        df = pd.DataFrame({"A": [1, 2, 3]})
        tmp_path = os.path.join(temp_directory, "test_sheets.xlsx")

        with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Sheet1", index=False)
            df.to_excel(writer, sheet_name="Sheet2", index=False)

        sheets = ExcelUtils.get_excel_sheets(tmp_path)
        assert "Sheet1" in sheets
        assert "Sheet2" in sheets
        assert len(sheets) == 2

    def test_hex_to_rgb(self):
        """Test conversion hex vers RGB"""
        assert ExcelUtils._hex_to_rgb("#1F4E79") == "1F4E79"
        assert ExcelUtils._hex_to_rgb("ff0000") == "FF0000"

    def test_read_specific_sheet(self, temp_excel_file):
        """Test lecture d'un onglet spécifique"""
        df, sheets, error = ExcelUtils.read_excel_file(temp_excel_file, sheet_name="TestData")
        assert df is not None
        assert error is None

    def test_read_nonexistent_sheet(self, temp_excel_file):
        """Test lecture d'un onglet inexistant"""
        df, sheets, error = ExcelUtils.read_excel_file(temp_excel_file, sheet_name="NonExistent")
        assert df is None
        assert error is not None
        assert "introuvable" in error

    def test_write_with_formatting(self, temp_directory, sample_dataframe):
        """Test écriture avec formatage"""
        filepath = os.path.join(temp_directory, "formatted.xlsx")

        success, error = ExcelUtils.write_dataframe_to_excel(
            sample_dataframe,
            filepath,
            "Formatted",
            apply_formatting=True,
            freeze_header=True,
            auto_fit_columns=True,
            alternate_rows=True,
            add_borders=True
        )

        assert success is True
        assert os.path.exists(filepath)

    def test_write_without_formatting(self, temp_directory, sample_dataframe):
        """Test écriture sans formatage"""
        filepath = os.path.join(temp_directory, "unformatted.xlsx")

        success, error = ExcelUtils.write_dataframe_to_excel(
            sample_dataframe,
            filepath,
            "Unformatted",
            apply_formatting=False
        )

        assert success is True

    def test_get_sheet_names(self, multi_sheet_excel_file):
        """Test récupération noms d'onglets"""
        sheets, error = ExcelUtils.get_sheet_names(multi_sheet_excel_file)

        assert error is None
        assert len(sheets) == 2
        assert "Sheet1" in sheets
        assert "Sheet2" in sheets


class TestExcelUtilsMerge:
    """Tests pour la fusion de fichiers"""

    def test_merge_excel_files(self, temp_directory):
        """Test fusion de plusieurs fichiers Excel"""
        df1 = pd.DataFrame({"Col1": [1, 2], "Col2": ["A", "B"]})
        df2 = pd.DataFrame({"Col1": [3, 4], "Col2": ["C", "D"]})

        file1 = os.path.join(temp_directory, "file1.xlsx")
        file2 = os.path.join(temp_directory, "file2.xlsx")
        output = os.path.join(temp_directory, "merged.xlsx")

        # Créer les fichiers source
        df1.to_excel(file1, index=False)
        df2.to_excel(file2, index=False)

        # Fusionner
        success, count, error = ExcelUtils.merge_excel_files(
            [file1, file2], output
        )

        assert success is True
        # Note: skip_headers=True par défaut donc 2 + 1 = 3 lignes
        assert count >= 3
        assert error is None

        # Vérifier le résultat
        df_merged, _, _ = ExcelUtils.read_excel_file(output)
        assert len(df_merged) >= 3

    def test_merge_with_skip_headers(self, temp_directory):
        """Test fusion avec saut d'en-têtes"""
        df1 = pd.DataFrame({"A": ["h1", "v1"], "B": ["h2", "v2"]})
        df2 = pd.DataFrame({"A": ["h1", "v3"], "B": ["h2", "v4"]})

        file1 = os.path.join(temp_directory, "f1.xlsx")
        file2 = os.path.join(temp_directory, "f2.xlsx")
        output = os.path.join(temp_directory, "merged.xlsx")

        df1.to_excel(file1, index=False)
        df2.to_excel(file2, index=False)

        success, count, error = ExcelUtils.merge_excel_files(
            [file1, file2], output, skip_headers=True
        )

        assert success is True


class TestExcelUtilsAddSheet:
    """Tests pour l'ajout d'onglets"""

    def test_add_sheet_to_existing_workbook(self, temp_excel_file, sample_dataframe_with_ref):
        """Test ajout onglet à fichier existant"""
        success, error = ExcelUtils.add_sheet_to_workbook(
            temp_excel_file,
            "NewSheet",
            sample_dataframe_with_ref
        )

        assert success is True

        # Vérifier que le nouvel onglet existe
        sheets = ExcelUtils.get_excel_sheets(temp_excel_file)
        assert "NewSheet" in sheets

    def test_add_sheet_replace_existing(self, temp_excel_file, sample_dataframe):
        """Test remplacement onglet existant"""
        # D'abord ajouter un onglet
        ExcelUtils.add_sheet_to_workbook(temp_excel_file, "ReplaceMe", sample_dataframe)

        # Le remplacer avec des données différentes
        new_df = pd.DataFrame({"X": [1, 2, 3]})
        success, error = ExcelUtils.add_sheet_to_workbook(temp_excel_file, "ReplaceMe", new_df)

        assert success is True

        # Vérifier le contenu
        df_read, _, _ = ExcelUtils.read_excel_file(temp_excel_file, sheet_name="ReplaceMe")
        assert "X" in df_read.columns


class TestExcelUtilsSearch:
    """Tests pour la recherche dans Excel"""

    def test_search_basic(self, sample_dataframe):
        """Test recherche basique"""
        results = ExcelUtils.search_in_excel(sample_dataframe, "Alice")
        assert len(results) == 1

    def test_search_case_insensitive(self, sample_dataframe):
        """Test recherche insensible à la casse"""
        results = ExcelUtils.search_in_excel(sample_dataframe, "alice", case_sensitive=False)
        assert len(results) == 1

    def test_search_case_sensitive(self, sample_dataframe):
        """Test recherche sensible à la casse"""
        # Avec case_sensitive=True, "alice" ne devrait pas correspondre à "Alice"
        results = ExcelUtils.search_in_excel(sample_dataframe, "NOTEXIST", case_sensitive=True)
        assert len(results) == 0
        # Vérifier que "Alice" exact est trouvé en mode sensible à la casse
        results2 = ExcelUtils.search_in_excel(sample_dataframe, "Alice", case_sensitive=True)
        assert len(results2) == 1

    def test_search_exact_match(self, sample_dataframe):
        """Test correspondance exacte"""
        results = ExcelUtils.search_in_excel(sample_dataframe, "Paris", exact_match=True)
        assert len(results) == 1

    def test_search_specific_columns(self, sample_dataframe):
        """Test recherche dans colonnes spécifiques"""
        results = ExcelUtils.search_in_excel(
            sample_dataframe, "25", columns=["Age"]
        )
        assert len(results) == 1


class TestExcelUtilsStatistics:
    """Tests pour les statistiques"""

    def test_column_statistics(self, sample_dataframe):
        """Test statistiques de colonne"""
        stats = ExcelUtils.get_column_statistics(sample_dataframe, "Nom")

        assert "total" in stats
        assert stats["total"] == 5
        assert "uniques" in stats

    def test_column_statistics_nonexistent(self, sample_dataframe):
        """Test statistiques colonne inexistante"""
        stats = ExcelUtils.get_column_statistics(sample_dataframe, "NonExistent")
        assert stats == {}


class TestExcelUtilsStatusFills:
    """Tests pour les fills de statut"""

    def test_get_status_fills_default(self):
        """Test fills par défaut"""
        fills = ExcelUtils.get_status_fills()

        assert "success" in fills
        assert "error" in fills
        assert "warning" in fills


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
