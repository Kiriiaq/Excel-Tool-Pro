"""
Système de logging centralisé pour ExcelToolsPro
Gestion des logs en temps réel avec callback pour l'IHM
"""

import logging
import sys
from pathlib import Path
from datetime import datetime
from enum import Enum
from typing import List, Callable, Optional
from dataclasses import dataclass


class LogLevel(Enum):
    """Niveaux de log avec couleurs associées"""
    DEBUG = ("DEBUG", "#6c757d")
    INFO = ("INFO", "#4da6ff")
    SUCCESS = ("SUCCESS", "#00bf63")
    WARNING = ("WARNING", "#ffbd59")
    ERROR = ("ERROR", "#ff6b6b")
    CRITICAL = ("CRITICAL", "#dc3545")

    @property
    def name_str(self) -> str:
        return self.value[0]

    @property
    def color(self) -> str:
        return self.value[1]


@dataclass
class LogEntry:
    """Entrée de log structurée"""
    timestamp: datetime
    level: LogLevel
    message: str
    source: str = ""

    def format(self, include_timestamp: bool = True) -> str:
        """Formate l'entrée pour affichage"""
        parts = []
        if include_timestamp:
            parts.append(f"[{self.timestamp.strftime('%H:%M:%S')}]")
        parts.append(f"[{self.level.name_str}]")
        if self.source:
            parts.append(f"[{self.source}]")
        parts.append(self.message)
        return " ".join(parts)


class Logger:
    """
    Gestionnaire de logs centralisé avec support de callbacks pour l'IHM

    Caractéristiques:
    - Logging multi-destinations (fichier, console, IHM)
    - Callbacks pour mise à jour temps réel de l'IHM
    - Historique en mémoire
    - Export des logs
    - Support UTF-8
    """

    def __init__(
        self,
        log_dir: Optional[Path] = None,
        name: str = "ExcelToolsPro",
        level: LogLevel = LogLevel.INFO,
        max_entries: int = 10000
    ):
        self.name = name
        self.level = level
        self.max_entries = max_entries
        self.entries: List[LogEntry] = []
        self._callbacks: List[Callable[[LogEntry], None]] = []
        self._error_count = 0
        self._warning_count = 0

        # Configuration du répertoire de logs
        if log_dir is None:
            log_dir = Path.home() / ".exceltoolspro" / "logs"
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)

        # Fichiers de log
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        self.log_file = self.log_dir / f"exceltoolspro_{timestamp}.log"
        self.error_file = self.log_dir / f"errors_{timestamp}.txt"

        # Configuration du logger Python standard
        self._setup_python_logger()

    def _setup_python_logger(self):
        """Configure le logger Python standard"""
        self._logger = logging.getLogger(self.name)
        self._logger.setLevel(logging.DEBUG)

        # Nettoyer les handlers existants
        self._logger.handlers.clear()

        # Handler fichier avec UTF-8
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(file_formatter)
        self._logger.addHandler(file_handler)

        # Handler console (optionnel)
        if sys.stdout is not None:
            console_handler = logging.StreamHandler(sys.stdout)
            console_handler.setLevel(logging.INFO)
            console_formatter = logging.Formatter(
                '%(asctime)s - %(levelname)s - %(message)s',
                datefmt='%H:%M:%S'
            )
            console_handler.setFormatter(console_formatter)
            self._logger.addHandler(console_handler)

    def _log(self, level: LogLevel, message: str, source: str = ""):
        """Méthode interne de logging"""
        entry = LogEntry(
            timestamp=datetime.now(),
            level=level,
            message=message,
            source=source
        )

        # Ajouter à l'historique
        self.entries.append(entry)
        if len(self.entries) > self.max_entries:
            self.entries = self.entries[-self.max_entries:]

        # Compteurs
        if level == LogLevel.ERROR or level == LogLevel.CRITICAL:
            self._error_count += 1
        elif level == LogLevel.WARNING:
            self._warning_count += 1

        # Logger Python
        python_level = getattr(logging, level.name_str, logging.INFO)
        self._logger.log(python_level, message)

        # Notifier les callbacks (IHM)
        for callback in self._callbacks:
            try:
                callback(entry)
            except Exception:
                pass  # Ignorer les erreurs de callback

    def debug(self, message: str, source: str = ""):
        """Log de niveau DEBUG"""
        self._log(LogLevel.DEBUG, message, source)

    def info(self, message: str, source: str = ""):
        """Log de niveau INFO"""
        self._log(LogLevel.INFO, message, source)

    def success(self, message: str, source: str = ""):
        """Log de niveau SUCCESS"""
        self._log(LogLevel.SUCCESS, message, source)

    def warning(self, message: str, source: str = ""):
        """Log de niveau WARNING"""
        self._log(LogLevel.WARNING, message, source)

    def error(self, message: str, source: str = ""):
        """Log de niveau ERROR"""
        self._log(LogLevel.ERROR, message, source)

    def critical(self, message: str, source: str = ""):
        """Log de niveau CRITICAL"""
        self._log(LogLevel.CRITICAL, message, source)

    def add_callback(self, callback: Callable[[LogEntry], None]):
        """Ajoute un callback appelé à chaque nouveau log"""
        self._callbacks.append(callback)

    def remove_callback(self, callback: Callable[[LogEntry], None]):
        """Supprime un callback"""
        if callback in self._callbacks:
            self._callbacks.remove(callback)

    def clear_callbacks(self):
        """Supprime tous les callbacks"""
        self._callbacks.clear()

    def get_entries(
        self,
        level: Optional[LogLevel] = None,
        source: Optional[str] = None,
        limit: int = 100
    ) -> List[LogEntry]:
        """Récupère les entrées de log filtrées"""
        entries = self.entries

        if level:
            entries = [e for e in entries if e.level == level]

        if source:
            entries = [e for e in entries if e.source == source]

        return entries[-limit:]

    def get_errors(self) -> List[LogEntry]:
        """Récupère toutes les erreurs"""
        return [e for e in self.entries
                if e.level in (LogLevel.ERROR, LogLevel.CRITICAL)]

    def get_warnings(self) -> List[LogEntry]:
        """Récupère tous les avertissements"""
        return [e for e in self.entries if e.level == LogLevel.WARNING]

    @property
    def error_count(self) -> int:
        """Nombre total d'erreurs"""
        return self._error_count

    @property
    def warning_count(self) -> int:
        """Nombre total d'avertissements"""
        return self._warning_count

    def clear(self):
        """Efface l'historique des logs en mémoire"""
        self.entries.clear()
        self._error_count = 0
        self._warning_count = 0

    def save_error_report(self, stats: dict = None) -> Optional[Path]:
        """Génère un rapport d'erreurs formaté"""
        errors = self.get_errors()
        warnings = self.get_warnings()

        if not errors and not warnings:
            return None

        try:
            with open(self.error_file, 'w', encoding='utf-8') as f:
                f.write("=" * 80 + "\n")
                f.write("RAPPORT D'ERREURS - ExcelToolsPro\n")
                f.write(f"Date: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                f.write("=" * 80 + "\n\n")

                if stats:
                    f.write("STATISTIQUES:\n")
                    f.write("-" * 40 + "\n")
                    for key, value in stats.items():
                        f.write(f"  {key}: {value}\n")
                    f.write("\n")

                if errors:
                    f.write(f"ERREURS ({len(errors)}):\n")
                    f.write("-" * 40 + "\n")
                    for i, entry in enumerate(errors, 1):
                        f.write(f"[{i}] {entry.format()}\n")
                    f.write("\n")

                if warnings:
                    f.write(f"AVERTISSEMENTS ({len(warnings)}):\n")
                    f.write("-" * 40 + "\n")
                    for i, entry in enumerate(warnings, 1):
                        f.write(f"[{i}] {entry.format()}\n")
                    f.write("\n")

                f.write("=" * 80 + "\n")
                f.write(f"Log complet: {self.log_file}\n")

            return self.error_file

        except Exception as e:
            self.error(f"Impossible de sauvegarder le rapport d'erreurs: {e}")
            return None

    def export_logs(self, export_path: Path, include_debug: bool = False) -> bool:
        """Exporte tous les logs vers un fichier"""
        try:
            entries = self.entries if include_debug else [
                e for e in self.entries if e.level != LogLevel.DEBUG
            ]

            with open(export_path, 'w', encoding='utf-8') as f:
                f.write(f"Export des logs ExcelToolsPro\n")
                f.write(f"Date d'export: {datetime.now().isoformat()}\n")
                f.write(f"Nombre d'entrées: {len(entries)}\n")
                f.write("=" * 80 + "\n\n")

                for entry in entries:
                    f.write(entry.format() + "\n")

            return True
        except Exception:
            return False


# Instance globale du logger
_global_logger: Optional[Logger] = None


def get_logger() -> Logger:
    """Récupère ou crée l'instance globale du logger"""
    global _global_logger
    if _global_logger is None:
        _global_logger = Logger()
    return _global_logger


def set_logger(logger: Logger):
    """Définit l'instance globale du logger"""
    global _global_logger
    _global_logger = logger
