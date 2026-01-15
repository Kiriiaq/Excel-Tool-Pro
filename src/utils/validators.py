"""
Validateurs pour ExcelToolsPro
Validation des entrées utilisateur et des données
"""

import re
from typing import Any, Tuple, Optional, List
from pathlib import Path


class Validators:
    """Classe utilitaire pour la validation des données"""

    @staticmethod
    def is_valid_email(email: str) -> bool:
        """Vérifie si une adresse email est valide"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email))

    @staticmethod
    def is_valid_reference(ref: str, pattern: Optional[str] = None) -> bool:
        """
        Vérifie si une référence est valide

        Args:
            ref: Référence à valider
            pattern: Expression régulière personnalisée (optionnel)
        """
        if not ref or not ref.strip():
            return False

        if pattern:
            return bool(re.match(pattern, ref))

        # Pattern par défaut : alphanumérique avec tirets/underscores
        default_pattern = r'^[A-Za-z0-9][A-Za-z0-9\-_\.]+$'
        return bool(re.match(default_pattern, ref))

    @staticmethod
    def validate_excel_file(filepath: str) -> Tuple[bool, Optional[str]]:
        """
        Valide qu'un fichier est un fichier Excel valide

        Returns:
            Tuple (valide, message d'erreur ou None)
        """
        if not filepath:
            return False, "Aucun fichier spécifié"

        path = Path(filepath)

        if not path.exists():
            return False, f"Le fichier n'existe pas: {filepath}"

        if not path.is_file():
            return False, f"Ce n'est pas un fichier: {filepath}"

        valid_extensions = ['.xlsx', '.xls', '.xlsm', '.xlsb']
        if path.suffix.lower() not in valid_extensions:
            return False, f"Extension invalide. Attendu: {', '.join(valid_extensions)}"

        # Vérifier que ce n'est pas un fichier temporaire
        if path.name.startswith('~$'):
            return False, "Fichier temporaire Excel (fichier ouvert dans Excel)"

        return True, None

    @staticmethod
    def validate_csv_file(filepath: str) -> Tuple[bool, Optional[str]]:
        """
        Valide qu'un fichier est un fichier CSV valide

        Returns:
            Tuple (valide, message d'erreur ou None)
        """
        if not filepath:
            return False, "Aucun fichier spécifié"

        path = Path(filepath)

        if not path.exists():
            return False, f"Le fichier n'existe pas: {filepath}"

        if path.suffix.lower() not in ['.csv', '.txt']:
            return False, "Extension invalide. Attendu: .csv ou .txt"

        return True, None

    @staticmethod
    def validate_column_name(name: str, existing_columns: List[str] = None) -> Tuple[bool, Optional[str]]:
        """
        Valide un nom de colonne Excel

        Returns:
            Tuple (valide, message d'erreur ou None)
        """
        if not name or not name.strip():
            return False, "Nom de colonne vide"

        name = name.strip()

        # Longueur maximale
        if len(name) > 255:
            return False, "Nom de colonne trop long (max 255 caractères)"

        # Caractères interdits
        invalid_chars = ['[', ']', ':', '*', '?', '/', '\\']
        for char in invalid_chars:
            if char in name:
                return False, f"Caractère interdit dans le nom: {char}"

        # Vérifier les doublons
        if existing_columns and name in existing_columns:
            return False, f"Colonne déjà existante: {name}"

        return True, None

    @staticmethod
    def validate_sheet_name(name: str, existing_sheets: List[str] = None) -> Tuple[bool, Optional[str]]:
        """
        Valide un nom d'onglet Excel

        Returns:
            Tuple (valide, message d'erreur ou None)
        """
        if not name or not name.strip():
            return False, "Nom d'onglet vide"

        name = name.strip()

        # Longueur maximale
        if len(name) > 31:
            return False, "Nom d'onglet trop long (max 31 caractères)"

        # Caractères interdits
        invalid_chars = ['[', ']', ':', '*', '?', '/', '\\']
        for char in invalid_chars:
            if char in name:
                return False, f"Caractère interdit dans le nom: {char}"

        # Ne peut pas commencer ou finir par une apostrophe
        if name.startswith("'") or name.endswith("'"):
            return False, "Le nom ne peut pas commencer ou finir par une apostrophe"

        # Vérifier les doublons
        if existing_sheets and name in existing_sheets:
            return False, f"Onglet déjà existant: {name}"

        return True, None

    @staticmethod
    def validate_numeric(
        value: Any,
        min_value: Optional[float] = None,
        max_value: Optional[float] = None,
        allow_none: bool = False
    ) -> Tuple[bool, Optional[str], Optional[float]]:
        """
        Valide et convertit une valeur numérique

        Returns:
            Tuple (valide, message d'erreur ou None, valeur convertie ou None)
        """
        if value is None or (isinstance(value, str) and not value.strip()):
            if allow_none:
                return True, None, None
            return False, "Valeur requise", None

        try:
            num_value = float(value)
        except (ValueError, TypeError):
            return False, f"Valeur non numérique: {value}", None

        if min_value is not None and num_value < min_value:
            return False, f"Valeur trop petite (min: {min_value})", None

        if max_value is not None and num_value > max_value:
            return False, f"Valeur trop grande (max: {max_value})", None

        return True, None, num_value

    @staticmethod
    def validate_integer(
        value: Any,
        min_value: Optional[int] = None,
        max_value: Optional[int] = None,
        allow_none: bool = False
    ) -> Tuple[bool, Optional[str], Optional[int]]:
        """
        Valide et convertit une valeur entière

        Returns:
            Tuple (valide, message d'erreur ou None, valeur convertie ou None)
        """
        valid, error, num = Validators.validate_numeric(value, min_value, max_value, allow_none)

        if not valid:
            return valid, error, None

        if num is None:
            return True, None, None

        try:
            int_value = int(num)
            if int_value != num:
                return False, "Valeur décimale non autorisée", None
            return True, None, int_value
        except (ValueError, TypeError):
            return False, "Conversion en entier impossible", None

    @staticmethod
    def validate_choice(
        value: str,
        valid_choices: List[str],
        case_sensitive: bool = False
    ) -> Tuple[bool, Optional[str]]:
        """
        Valide qu'une valeur est dans une liste de choix

        Returns:
            Tuple (valide, message d'erreur ou None)
        """
        if not value:
            return False, "Valeur requise"

        if case_sensitive:
            if value not in valid_choices:
                return False, f"Valeur invalide. Choix possibles: {', '.join(valid_choices)}"
        else:
            value_lower = value.lower()
            valid_lower = [c.lower() for c in valid_choices]
            if value_lower not in valid_lower:
                return False, f"Valeur invalide. Choix possibles: {', '.join(valid_choices)}"

        return True, None

    @staticmethod
    def validate_regex(value: str, pattern: str) -> Tuple[bool, Optional[str]]:
        """
        Valide une valeur contre une expression régulière

        Returns:
            Tuple (valide, message d'erreur ou None)
        """
        if not value:
            return False, "Valeur requise"

        try:
            if re.match(pattern, value):
                return True, None
            return False, f"Format invalide (pattern attendu: {pattern})"
        except re.error as e:
            return False, f"Expression régulière invalide: {e}"

    @staticmethod
    def sanitize_filename(filename: str) -> str:
        """
        Nettoie un nom de fichier en supprimant les caractères invalides

        Returns:
            Nom de fichier nettoyé
        """
        # Caractères interdits sous Windows
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')

        # Supprimer les espaces en début/fin
        filename = filename.strip()

        # Limiter la longueur
        if len(filename) > 200:
            name, ext = filename.rsplit('.', 1) if '.' in filename else (filename, '')
            max_name_len = 200 - len(ext) - 1 if ext else 200
            filename = f"{name[:max_name_len]}.{ext}" if ext else name[:200]

        return filename

    @staticmethod
    def sanitize_sheet_name(name: str) -> str:
        """
        Nettoie un nom d'onglet Excel

        Returns:
            Nom d'onglet nettoyé
        """
        # Caractères interdits
        invalid_chars = '[]:\\/?*'
        for char in invalid_chars:
            name = name.replace(char, '_')

        # Supprimer les apostrophes en début/fin
        name = name.strip("'")

        # Limiter la longueur
        if len(name) > 31:
            name = name[:31]

        return name or "Sheet"
