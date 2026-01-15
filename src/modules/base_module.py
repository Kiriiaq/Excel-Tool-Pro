"""
Classe de base pour tous les modules fonctionnels
"""

import customtkinter as ctk
from abc import ABC, abstractmethod
from typing import Optional, Dict, Any, List, Callable
import threading

from ..core.logger import Logger, LogLevel, get_logger
from ..core.config import ConfigManager
from ..core.constants import COLORS, StepStatus


class BaseModule(ABC):
    """
    Classe de base abstraite pour tous les modules de l'application

    Fournit:
    - Interface standard (frame)
    - Système de logging
    - Gestion de la configuration
    - Exécution asynchrone
    - Callbacks de progression
    """

    # Métadonnées du module (à surcharger)
    MODULE_ID = "base"
    MODULE_NAME = "Module de base"
    MODULE_DESCRIPTION = "Description du module"
    MODULE_ICON = "📦"

    def __init__(
        self,
        parent: ctk.CTkFrame,
        config_manager: Optional[ConfigManager] = None,
        logger: Optional[Logger] = None
    ):
        self.parent = parent
        self.config = config_manager
        self.logger = logger or get_logger()

        # Frame principal du module
        self.frame = ctk.CTkFrame(parent, fg_color="transparent")

        # État
        self.is_running = False
        self.should_cancel = False
        self.current_thread: Optional[threading.Thread] = None

        # Callbacks
        self._progress_callback: Optional[Callable[[float], None]] = None
        self._status_callback: Optional[Callable[[str, str], None]] = None
        self._complete_callback: Optional[Callable[[bool, Dict], None]] = None

        # Initialiser l'interface
        self._create_interface()

    @abstractmethod
    def _create_interface(self):
        """Crée l'interface du module (à implémenter)"""
        pass

    @abstractmethod
    def _execute_task(self) -> Dict[str, Any]:
        """
        Exécute la tâche principale du module (à implémenter)

        Returns:
            Dictionnaire de résultats
        """
        pass

    @abstractmethod
    def validate_inputs(self) -> tuple[bool, str]:
        """
        Valide les entrées utilisateur (à implémenter)

        Returns:
            Tuple (valide, message d'erreur si non valide)
        """
        pass

    def get_frame(self) -> ctk.CTkFrame:
        """Retourne le frame principal du module"""
        return self.frame

    def show(self):
        """Affiche le module"""
        self.frame.pack(fill="both", expand=True)

    def hide(self):
        """Cache le module"""
        self.frame.pack_forget()

    def log(self, message: str, level: LogLevel = LogLevel.INFO):
        """Enregistre un message de log"""
        self.logger._log(level, message, source=self.MODULE_ID)

    def log_info(self, message: str):
        self.log(message, LogLevel.INFO)

    def log_success(self, message: str):
        self.log(message, LogLevel.SUCCESS)

    def log_warning(self, message: str):
        self.log(message, LogLevel.WARNING)

    def log_error(self, message: str):
        self.log(message, LogLevel.ERROR)

    def set_progress_callback(self, callback: Callable[[float], None]):
        """Définit le callback de progression (0.0 à 1.0)"""
        self._progress_callback = callback

    def set_status_callback(self, callback: Callable[[str, str], None]):
        """Définit le callback de statut (message, niveau)"""
        self._status_callback = callback

    def set_complete_callback(self, callback: Callable[[bool, Dict], None]):
        """Définit le callback de fin (succès, résultats)"""
        self._complete_callback = callback

    def update_progress(self, progress: float):
        """Met à jour la progression"""
        if self._progress_callback:
            try:
                self._progress_callback(progress)
            except Exception:
                pass

    def update_status(self, message: str, level: str = "info"):
        """Met à jour le statut"""
        if self._status_callback:
            try:
                self._status_callback(message, level)
            except Exception:
                pass

    def start_execution(self):
        """Démarre l'exécution dans un thread séparé"""
        if self.is_running:
            self.log_warning("Une exécution est déjà en cours")
            return

        # Valider les entrées
        valid, error = self.validate_inputs()
        if not valid:
            self.log_error(f"Validation échouée: {error}")
            self.update_status(error, "error")
            return

        self.is_running = True
        self.should_cancel = False
        self.update_progress(0)
        self.update_status("Démarrage...", "info")

        # Lancer dans un thread
        self.current_thread = threading.Thread(target=self._run_task, daemon=True)
        self.current_thread.start()

    def _run_task(self):
        """Exécute la tâche dans le thread"""
        success = False
        results = {}

        try:
            self.log_info(f"Démarrage de {self.MODULE_NAME}")
            results = self._execute_task()
            success = not self.should_cancel

            if success:
                self.log_success(f"{self.MODULE_NAME} terminé avec succès")
                self.update_status("Terminé avec succès", "success")
            else:
                self.log_warning("Exécution annulée")
                self.update_status("Annulé", "warning")

        except Exception as e:
            self.log_error(f"Erreur: {str(e)}")
            self.update_status(f"Erreur: {str(e)}", "error")
            results["error"] = str(e)

        finally:
            self.is_running = False
            self.update_progress(1.0 if success else 0)

            if self._complete_callback:
                try:
                    self._complete_callback(success, results)
                except Exception:
                    pass

    def cancel_execution(self):
        """Annule l'exécution en cours"""
        if self.is_running:
            self.should_cancel = True
            self.log_warning("Annulation demandée...")
            self.update_status("Annulation en cours...", "warning")

    def is_cancelled(self) -> bool:
        """Vérifie si l'annulation a été demandée"""
        return self.should_cancel

    def get_config_value(self, key: str, default: Any = None) -> Any:
        """Récupère une valeur de configuration du module"""
        if self.config:
            mod_config = self.config.get_module_config(self.MODULE_ID)
            return mod_config.settings.get(key, default)
        return default

    def set_config_value(self, key: str, value: Any):
        """Définit une valeur de configuration du module"""
        if self.config:
            self.config.set_module_setting(self.MODULE_ID, key, value)

    def reset(self):
        """Réinitialise le module (à surcharger si nécessaire)"""
        self.should_cancel = False
        self.update_progress(0)
        self.update_status("", "info")

    @classmethod
    def get_metadata(cls) -> Dict[str, str]:
        """Retourne les métadonnées du module"""
        return {
            "id": cls.MODULE_ID,
            "name": cls.MODULE_NAME,
            "description": cls.MODULE_DESCRIPTION,
            "icon": cls.MODULE_ICON
        }
