"""
Tests unitaires pour les validateurs
"""

import pytest
import sys
from pathlib import Path
import tempfile
import os

sys.path.insert(0, str(Path(__file__).parent.parent))

from src.utils.validators import Validators


class TestValidatorsEmail:
    """Tests pour la validation d'email"""

    def test_valid_email(self):
        """Test email valide"""
        assert Validators.is_valid_email("test@example.com") is True
        assert Validators.is_valid_email("user.name@domain.org") is True
        assert Validators.is_valid_email("user+tag@domain.co.uk") is True

    def test_invalid_email(self):
        """Test email invalide"""
        assert Validators.is_valid_email("invalid") is False
        assert Validators.is_valid_email("@domain.com") is False
        assert Validators.is_valid_email("user@") is False
        assert Validators.is_valid_email("") is False


class TestValidatorsReference:
    """Tests pour la validation de références"""

    def test_valid_reference(self):
        """Test référence valide"""
        assert Validators.is_valid_reference("REF001") is True
        assert Validators.is_valid_reference("ABC-123") is True
        assert Validators.is_valid_reference("test_ref.v2") is True

    def test_invalid_reference(self):
        """Test référence invalide"""
        assert Validators.is_valid_reference("") is False
        assert Validators.is_valid_reference("   ") is False

    def test_custom_pattern(self):
        """Test avec pattern personnalisé"""
        assert Validators.is_valid_reference("ABC123", pattern=r"^[A-Z]{3}\d{3}$") is True
        assert Validators.is_valid_reference("AB123", pattern=r"^[A-Z]{3}\d{3}$") is False


class TestValidatorsExcelFile:
    """Tests pour la validation de fichiers Excel"""

    def test_valid_excel_file(self, temp_excel_file):
        """Test fichier Excel valide"""
        valid, error = Validators.validate_excel_file(temp_excel_file)
        assert valid is True
        assert error is None

    def test_nonexistent_file(self):
        """Test fichier inexistant"""
        valid, error = Validators.validate_excel_file("fichier_inexistant.xlsx")
        assert valid is False
        assert "n'existe pas" in error

    def test_empty_path(self):
        """Test chemin vide"""
        valid, error = Validators.validate_excel_file("")
        assert valid is False

    def test_invalid_extension(self, temp_csv_file):
        """Test extension invalide"""
        valid, error = Validators.validate_excel_file(temp_csv_file)
        assert valid is False
        assert "Extension invalide" in error

    def test_temp_file(self, temp_directory):
        """Test fichier temporaire Excel (préfixe ~$)"""
        temp_file = os.path.join(temp_directory, "~$test.xlsx")
        with open(temp_file, 'w') as f:
            f.write("temp")

        valid, error = Validators.validate_excel_file(temp_file)
        assert valid is False
        assert "temporaire" in error


class TestValidatorsCSVFile:
    """Tests pour la validation de fichiers CSV"""

    def test_valid_csv_file(self, temp_csv_file):
        """Test fichier CSV valide"""
        valid, error = Validators.validate_csv_file(temp_csv_file)
        assert valid is True
        assert error is None

    def test_nonexistent_csv(self):
        """Test fichier CSV inexistant"""
        valid, error = Validators.validate_csv_file("inexistant.csv")
        assert valid is False

    def test_empty_csv_path(self):
        """Test chemin CSV vide"""
        valid, error = Validators.validate_csv_file("")
        assert valid is False


class TestValidatorsColumnName:
    """Tests pour la validation de noms de colonnes"""

    def test_valid_column_name(self):
        """Test nom de colonne valide"""
        valid, error = Validators.validate_column_name("Nom_Colonne")
        assert valid is True
        assert error is None

    def test_empty_column_name(self):
        """Test nom de colonne vide"""
        valid, error = Validators.validate_column_name("")
        assert valid is False
        assert "vide" in error

    def test_column_name_too_long(self):
        """Test nom de colonne trop long"""
        long_name = "A" * 300
        valid, error = Validators.validate_column_name(long_name)
        assert valid is False
        assert "trop long" in error

    def test_column_name_invalid_chars(self):
        """Test caractères invalides dans nom de colonne"""
        valid, error = Validators.validate_column_name("Col[1]")
        assert valid is False
        assert "interdit" in error

    def test_duplicate_column(self):
        """Test colonne dupliquée"""
        existing = ["Col1", "Col2", "Col3"]
        valid, error = Validators.validate_column_name("Col1", existing_columns=existing)
        assert valid is False
        assert "existante" in error


class TestValidatorsSheetName:
    """Tests pour la validation de noms d'onglets"""

    def test_valid_sheet_name(self):
        """Test nom d'onglet valide"""
        valid, error = Validators.validate_sheet_name("MonOnglet")
        assert valid is True

    def test_empty_sheet_name(self):
        """Test nom d'onglet vide"""
        valid, error = Validators.validate_sheet_name("")
        assert valid is False

    def test_sheet_name_too_long(self):
        """Test nom d'onglet trop long (>31 caractères)"""
        long_name = "A" * 35
        valid, error = Validators.validate_sheet_name(long_name)
        assert valid is False
        assert "trop long" in error

    def test_sheet_name_invalid_chars(self):
        """Test caractères invalides"""
        for char in ['[', ']', ':', '*', '?', '/', '\\']:
            valid, error = Validators.validate_sheet_name(f"Sheet{char}Test")
            assert valid is False

    def test_sheet_name_apostrophe(self):
        """Test apostrophe en début/fin"""
        valid, error = Validators.validate_sheet_name("'Sheet")
        assert valid is False

        valid, error = Validators.validate_sheet_name("Sheet'")
        assert valid is False


class TestValidatorsNumeric:
    """Tests pour la validation numérique"""

    def test_valid_numeric(self):
        """Test valeur numérique valide"""
        valid, error, value = Validators.validate_numeric("123.45")
        assert valid is True
        assert value == 123.45

    def test_invalid_numeric(self):
        """Test valeur non numérique"""
        valid, error, value = Validators.validate_numeric("abc")
        assert valid is False
        assert value is None

    def test_numeric_min_value(self):
        """Test valeur minimale"""
        valid, error, value = Validators.validate_numeric("5", min_value=10)
        assert valid is False
        assert "trop petite" in error

    def test_numeric_max_value(self):
        """Test valeur maximale"""
        valid, error, value = Validators.validate_numeric("100", max_value=50)
        assert valid is False
        assert "trop grande" in error

    def test_numeric_allow_none(self):
        """Test valeur None autorisée"""
        valid, error, value = Validators.validate_numeric(None, allow_none=True)
        assert valid is True
        assert value is None


class TestValidatorsInteger:
    """Tests pour la validation d'entiers"""

    def test_valid_integer(self):
        """Test entier valide"""
        valid, error, value = Validators.validate_integer("42")
        assert valid is True
        assert value == 42

    def test_decimal_rejected(self):
        """Test décimal rejeté"""
        valid, error, value = Validators.validate_integer("42.5")
        assert valid is False


class TestValidatorsChoice:
    """Tests pour la validation de choix"""

    def test_valid_choice(self):
        """Test choix valide"""
        valid, error = Validators.validate_choice("A", ["A", "B", "C"])
        assert valid is True

    def test_invalid_choice(self):
        """Test choix invalide"""
        valid, error = Validators.validate_choice("D", ["A", "B", "C"])
        assert valid is False

    def test_case_insensitive(self):
        """Test insensibilité à la casse"""
        valid, error = Validators.validate_choice("a", ["A", "B", "C"], case_sensitive=False)
        assert valid is True


class TestValidatorsRegex:
    """Tests pour la validation regex"""

    def test_valid_regex_match(self):
        """Test correspondance regex"""
        valid, error = Validators.validate_regex("ABC123", r"^[A-Z]{3}\d{3}$")
        assert valid is True

    def test_invalid_regex_match(self):
        """Test non-correspondance regex"""
        valid, error = Validators.validate_regex("AB12", r"^[A-Z]{3}\d{3}$")
        assert valid is False


class TestValidatorsSanitize:
    """Tests pour le nettoyage de noms"""

    def test_sanitize_filename(self):
        """Test nettoyage nom de fichier"""
        assert Validators.sanitize_filename("file:name?.xlsx") == "file_name_.xlsx"
        assert Validators.sanitize_filename("  test  ") == "test"

    def test_sanitize_filename_length(self):
        """Test limitation longueur nom de fichier"""
        long_name = "A" * 250 + ".xlsx"
        sanitized = Validators.sanitize_filename(long_name)
        assert len(sanitized) <= 205  # 200 + ".xlsx"

    def test_sanitize_sheet_name(self):
        """Test nettoyage nom d'onglet"""
        assert Validators.sanitize_sheet_name("Sheet[1]") == "Sheet_1_"
        assert len(Validators.sanitize_sheet_name("A" * 50)) <= 31

    def test_sanitize_empty_sheet_name(self):
        """Test nettoyage nom d'onglet vide"""
        assert Validators.sanitize_sheet_name("") == "Sheet"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
