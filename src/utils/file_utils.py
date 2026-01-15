"""
Utilitaires de gestion de fichiers pour ExcelToolsPro
"""

import os
import shutil
from pathlib import Path
from typing import List, Optional, Tuple, Generator
from datetime import datetime


class FileUtils:
    """Classe utilitaire pour les opérations sur les fichiers"""

    @staticmethod
    def get_unique_filename(filepath: str) -> str:
        """
        Génère un nom de fichier unique si le fichier existe déjà

        Args:
            filepath: Chemin du fichier

        Returns:
            Chemin unique (avec suffixe _copy1, _copy2, etc.)
        """
        if not os.path.exists(filepath):
            return filepath

        path = Path(filepath)
        base_dir = path.parent
        name = path.stem
        ext = path.suffix

        counter = 1
        while True:
            new_name = f"{name}_copy{counter}{ext}"
            new_path = base_dir / new_name
            if not os.path.exists(new_path):
                return str(new_path)
            counter += 1

    @staticmethod
    def ensure_directory(directory: str) -> bool:
        """
        Crée un répertoire s'il n'existe pas

        Returns:
            True si le répertoire existe ou a été créé
        """
        try:
            Path(directory).mkdir(parents=True, exist_ok=True)
            return True
        except Exception:
            return False

    @staticmethod
    def move_file(
        source: str,
        destination_dir: str,
        create_dir: bool = True
    ) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        Déplace un fichier vers un répertoire

        Returns:
            Tuple (succès, nouveau chemin ou None, erreur ou None)
        """
        try:
            if create_dir:
                FileUtils.ensure_directory(destination_dir)

            filename = os.path.basename(source)
            destination = os.path.join(destination_dir, filename)
            destination = FileUtils.get_unique_filename(destination)

            shutil.move(source, destination)
            return True, destination, None

        except Exception as e:
            return False, None, str(e)

    @staticmethod
    def copy_file(
        source: str,
        destination_dir: str,
        create_dir: bool = True
    ) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        Copie un fichier vers un répertoire

        Returns:
            Tuple (succès, nouveau chemin ou None, erreur ou None)
        """
        try:
            if create_dir:
                FileUtils.ensure_directory(destination_dir)

            filename = os.path.basename(source)
            destination = os.path.join(destination_dir, filename)
            destination = FileUtils.get_unique_filename(destination)

            shutil.copy2(source, destination)
            return True, destination, None

        except Exception as e:
            return False, None, str(e)

    @staticmethod
    def list_files(
        directory: str,
        extensions: Optional[List[str]] = None,
        recursive: bool = False,
        exclude_temp: bool = True
    ) -> List[Path]:
        """
        Liste les fichiers d'un répertoire

        Args:
            directory: Répertoire à scanner
            extensions: Liste des extensions à inclure (ex: ['.xlsx', '.xls'])
            recursive: Recherche récursive
            exclude_temp: Exclure les fichiers temporaires (~$)

        Returns:
            Liste des chemins de fichiers
        """
        directory = Path(directory)
        if not directory.exists():
            return []

        files = []
        pattern = "**/*" if recursive else "*"

        for path in directory.glob(pattern):
            if not path.is_file():
                continue

            # Exclure les fichiers temporaires
            if exclude_temp and path.name.startswith("~$"):
                continue

            # Filtrer par extension
            if extensions:
                if path.suffix.lower() not in [ext.lower() for ext in extensions]:
                    continue

            files.append(path)

        # Trier par date de modification (plus récent en premier)
        files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        return files

    @staticmethod
    def list_excel_files(
        directory: str,
        recursive: bool = False
    ) -> List[Path]:
        """
        Liste les fichiers Excel d'un répertoire

        Returns:
            Liste des chemins de fichiers Excel
        """
        return FileUtils.list_files(
            directory,
            extensions=['.xlsx', '.xls', '.xlsm'],
            recursive=recursive,
            exclude_temp=True
        )

    @staticmethod
    def get_file_info(filepath: str) -> dict:
        """
        Récupère les informations sur un fichier

        Returns:
            Dictionnaire d'informations
        """
        try:
            path = Path(filepath)
            stat = path.stat()

            return {
                "name": path.name,
                "stem": path.stem,
                "extension": path.suffix,
                "directory": str(path.parent),
                "size": stat.st_size,
                "size_formatted": FileUtils.format_size(stat.st_size),
                "created": datetime.fromtimestamp(stat.st_ctime),
                "modified": datetime.fromtimestamp(stat.st_mtime),
                "exists": True
            }
        except Exception:
            return {"exists": False}

    @staticmethod
    def format_size(size_bytes: int) -> str:
        """
        Formate une taille en octets en format lisible

        Args:
            size_bytes: Taille en octets

        Returns:
            Taille formatée (ex: "1.5 MB")
        """
        for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
            if size_bytes < 1024:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024
        return f"{size_bytes:.1f} PB"

    @staticmethod
    def validate_path(filepath: str, must_exist: bool = True) -> Tuple[bool, Optional[str]]:
        """
        Valide un chemin de fichier

        Returns:
            Tuple (valide, message d'erreur ou None)
        """
        if not filepath:
            return False, "Chemin vide"

        path = Path(filepath)

        if must_exist and not path.exists():
            return False, f"Le fichier n'existe pas: {filepath}"

        if must_exist and not path.is_file():
            return False, f"Ce n'est pas un fichier: {filepath}"

        return True, None

    @staticmethod
    def validate_directory(
        directory: str,
        must_exist: bool = True,
        must_be_writable: bool = True
    ) -> Tuple[bool, Optional[str]]:
        """
        Valide un répertoire

        Returns:
            Tuple (valide, message d'erreur ou None)
        """
        if not directory:
            return False, "Chemin vide"

        path = Path(directory)

        if must_exist and not path.exists():
            return False, f"Le répertoire n'existe pas: {directory}"

        if must_exist and not path.is_dir():
            return False, f"Ce n'est pas un répertoire: {directory}"

        if must_be_writable and path.exists():
            if not os.access(path, os.W_OK):
                return False, f"Répertoire non accessible en écriture: {directory}"

        return True, None

    @staticmethod
    def create_backup(filepath: str, backup_dir: Optional[str] = None) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        Crée une sauvegarde d'un fichier

        Args:
            filepath: Fichier à sauvegarder
            backup_dir: Répertoire de sauvegarde (None = même répertoire)

        Returns:
            Tuple (succès, chemin backup ou None, erreur ou None)
        """
        try:
            path = Path(filepath)
            if not path.exists():
                return False, None, "Fichier source introuvable"

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"{path.stem}_backup_{timestamp}{path.suffix}"

            if backup_dir:
                FileUtils.ensure_directory(backup_dir)
                backup_path = Path(backup_dir) / backup_name
            else:
                backup_path = path.parent / backup_name

            shutil.copy2(filepath, backup_path)
            return True, str(backup_path), None

        except Exception as e:
            return False, None, str(e)

    @staticmethod
    def clean_old_files(
        directory: str,
        max_age_days: int = 30,
        extensions: Optional[List[str]] = None
    ) -> int:
        """
        Supprime les fichiers plus anciens qu'un certain nombre de jours

        Returns:
            Nombre de fichiers supprimés
        """
        deleted = 0
        cutoff_date = datetime.now().timestamp() - (max_age_days * 24 * 60 * 60)

        for filepath in FileUtils.list_files(directory, extensions):
            try:
                if filepath.stat().st_mtime < cutoff_date:
                    filepath.unlink()
                    deleted += 1
            except Exception:
                pass

        return deleted
