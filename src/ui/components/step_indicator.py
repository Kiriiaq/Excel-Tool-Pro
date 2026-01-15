"""
Indicateur d'étapes de workflow avec états visuels
"""

import customtkinter as ctk
from typing import List, Optional, Callable
from dataclasses import dataclass

from .tooltip import Tooltip
from ...core.constants import COLORS, StepStatus, STATUS_ICONS


@dataclass
class WorkflowStep:
    """Représente une étape du workflow"""
    id: str
    name: str
    description: str = ""
    status: str = StepStatus.PENDING
    enabled: bool = True
    progress: float = 0.0
    error_message: str = ""


class StepIndicator(ctk.CTkFrame):
    """
    Widget d'affichage des étapes d'un workflow

    Fonctionnalités:
    - Liste des étapes avec icônes d'état
    - Progression visuelle
    - Activation/désactivation d'étapes
    - Tooltips informatifs
    """

    def __init__(
        self,
        parent,
        steps: List[WorkflowStep],
        on_step_toggle: Optional[Callable[[str, bool], None]] = None,
        **kwargs
    ):
        super().__init__(parent, **kwargs)

        self.steps = {step.id: step for step in steps}
        self.step_widgets = {}
        self.on_step_toggle = on_step_toggle

        self.configure(fg_color=COLORS["bg_card"], corner_radius=10)
        self._create_widgets()

    def _create_widgets(self):
        """Crée les widgets d'affichage"""
        # Titre
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(15, 10))

        ctk.CTkLabel(
            header,
            text="📋 Étapes du traitement",
            font=("Segoe UI", 14, "bold")
        ).pack(side="left")

        # Container des étapes avec scroll
        self.steps_container = ctk.CTkScrollableFrame(
            self,
            fg_color="transparent",
            height=300
        )
        self.steps_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # Créer les widgets pour chaque étape
        for step in self.steps.values():
            self._create_step_widget(step)

    def _create_step_widget(self, step: WorkflowStep):
        """Crée le widget pour une étape"""
        frame = ctk.CTkFrame(
            self.steps_container,
            fg_color=("gray85", "gray25"),
            corner_radius=8
        )
        frame.pack(fill="x", pady=3, padx=5)

        # Checkbox d'activation
        enabled_var = ctk.BooleanVar(value=step.enabled)
        checkbox = ctk.CTkCheckBox(
            frame,
            text="",
            variable=enabled_var,
            width=24,
            command=lambda: self._toggle_step(step.id, enabled_var.get())
        )
        checkbox.pack(side="left", padx=(10, 5), pady=8)
        Tooltip(checkbox, "Activer/désactiver cette étape")

        # Icône d'état
        status_label = ctk.CTkLabel(
            frame,
            text=STATUS_ICONS.get(step.status, "⏳"),
            font=("Segoe UI", 14),
            width=30
        )
        status_label.pack(side="left", padx=(0, 5))

        # Nom et description
        info_frame = ctk.CTkFrame(frame, fg_color="transparent")
        info_frame.pack(side="left", fill="x", expand=True, padx=5)

        name_label = ctk.CTkLabel(
            info_frame,
            text=step.name,
            font=("Segoe UI", 11, "bold"),
            anchor="w"
        )
        name_label.pack(anchor="w")

        if step.description:
            desc_label = ctk.CTkLabel(
                info_frame,
                text=step.description,
                font=("Segoe UI", 9),
                text_color=COLORS["text_muted"],
                anchor="w"
            )
            desc_label.pack(anchor="w")
            Tooltip(name_label, step.description)

        # Barre de progression (cachée par défaut)
        progress_bar = ctk.CTkProgressBar(
            frame,
            width=80,
            height=8
        )
        progress_bar.set(step.progress)
        progress_bar.pack(side="right", padx=10, pady=8)
        progress_bar.pack_forget()  # Caché par défaut

        # Stocker les références
        self.step_widgets[step.id] = {
            "frame": frame,
            "checkbox": checkbox,
            "enabled_var": enabled_var,
            "status_label": status_label,
            "name_label": name_label,
            "progress_bar": progress_bar
        }

    def _toggle_step(self, step_id: str, enabled: bool):
        """Callback de toggle d'une étape"""
        if step_id in self.steps:
            self.steps[step_id].enabled = enabled
            if self.on_step_toggle:
                self.on_step_toggle(step_id, enabled)

    def set_step_status(self, step_id: str, status: str, error_message: str = ""):
        """Met à jour le statut d'une étape"""
        if step_id not in self.steps:
            return

        step = self.steps[step_id]
        step.status = status
        step.error_message = error_message

        widgets = self.step_widgets.get(step_id)
        if not widgets:
            return

        # Mettre à jour l'icône
        widgets["status_label"].configure(text=STATUS_ICONS.get(status, "⏳"))

        # Couleur selon l'état
        color_map = {
            StepStatus.PENDING: COLORS["text_muted"],
            StepStatus.RUNNING: COLORS["info"],
            StepStatus.SUCCESS: COLORS["success"],
            StepStatus.ERROR: COLORS["error"],
            StepStatus.WARNING: COLORS["warning"],
            StepStatus.SKIPPED: COLORS["text_muted"],
        }
        color = color_map.get(status, COLORS["text_primary"])
        widgets["name_label"].configure(text_color=color)

        # Afficher/cacher la barre de progression
        if status == StepStatus.RUNNING:
            widgets["progress_bar"].pack(side="right", padx=10, pady=8)
        else:
            widgets["progress_bar"].pack_forget()

        # Tooltip d'erreur si applicable
        if error_message:
            Tooltip(widgets["frame"], f"Erreur: {error_message}")

    def set_step_progress(self, step_id: str, progress: float):
        """Met à jour la progression d'une étape (0.0 à 1.0)"""
        if step_id in self.steps:
            self.steps[step_id].progress = progress

        widgets = self.step_widgets.get(step_id)
        if widgets:
            widgets["progress_bar"].set(progress)

    def get_enabled_steps(self) -> List[str]:
        """Retourne la liste des IDs des étapes activées"""
        return [
            step_id for step_id, step in self.steps.items()
            if step.enabled
        ]

    def reset_all(self):
        """Réinitialise toutes les étapes à l'état initial"""
        for step_id in self.steps:
            self.set_step_status(step_id, StepStatus.PENDING)
            self.set_step_progress(step_id, 0.0)

    def mark_all_complete(self):
        """Marque toutes les étapes comme terminées"""
        for step_id, step in self.steps.items():
            if step.enabled:
                self.set_step_status(step_id, StepStatus.SUCCESS)

    def get_summary(self) -> dict:
        """Retourne un résumé des étapes"""
        summary = {
            "total": len(self.steps),
            "enabled": 0,
            "completed": 0,
            "errors": 0,
            "skipped": 0
        }

        for step in self.steps.values():
            if step.enabled:
                summary["enabled"] += 1
            if step.status == StepStatus.SUCCESS:
                summary["completed"] += 1
            elif step.status == StepStatus.ERROR:
                summary["errors"] += 1
            elif step.status == StepStatus.SKIPPED:
                summary["skipped"] += 1

        return summary
