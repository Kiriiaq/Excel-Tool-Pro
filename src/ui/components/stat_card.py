"""
Carte de statistique pour affichage de métriques
"""

import customtkinter as ctk
from typing import Optional

from .tooltip import Tooltip
from ...core.constants import COLORS


class StatCard(ctk.CTkFrame):
    """
    Widget de carte de statistique

    Affiche une métrique avec:
    - Icône
    - Titre
    - Valeur principale
    - Sous-texte optionnel
    """

    def __init__(
        self,
        parent,
        title: str,
        value: str = "0",
        icon: str = "📊",
        subtitle: str = "",
        tooltip: str = "",
        color: str = COLORS["accent_primary"],
        **kwargs
    ):
        super().__init__(parent, **kwargs)

        self.title = title
        self.color = color

        self.configure(
            fg_color=COLORS["bg_card"],
            corner_radius=10
        )

        self._create_widgets(title, value, icon, subtitle, tooltip)

    def _create_widgets(
        self,
        title: str,
        value: str,
        icon: str,
        subtitle: str,
        tooltip: str
    ):
        """Crée les widgets de la carte"""
        # Container principal
        container = ctk.CTkFrame(self, fg_color="transparent")
        container.pack(fill="both", expand=True, padx=15, pady=12)

        # Ligne du haut : icône + titre
        top_frame = ctk.CTkFrame(container, fg_color="transparent")
        top_frame.pack(fill="x")

        icon_label = ctk.CTkLabel(
            top_frame,
            text=icon,
            font=("Segoe UI", 18)
        )
        icon_label.pack(side="left")

        title_label = ctk.CTkLabel(
            top_frame,
            text=title,
            font=("Segoe UI", 11),
            text_color=COLORS["text_muted"]
        )
        title_label.pack(side="left", padx=(8, 0))

        if tooltip:
            Tooltip(container, tooltip)

        # Valeur principale
        self.value_label = ctk.CTkLabel(
            container,
            text=value,
            font=("Segoe UI", 28, "bold"),
            text_color=self.color
        )
        self.value_label.pack(anchor="w", pady=(5, 0))

        # Sous-texte optionnel
        if subtitle:
            self.subtitle_label = ctk.CTkLabel(
                container,
                text=subtitle,
                font=("Segoe UI", 10),
                text_color=COLORS["text_muted"]
            )
            self.subtitle_label.pack(anchor="w")
        else:
            self.subtitle_label = None

    def set_value(self, value: str):
        """Met à jour la valeur affichée"""
        self.value_label.configure(text=value)

    def set_subtitle(self, subtitle: str):
        """Met à jour le sous-texte"""
        if self.subtitle_label:
            self.subtitle_label.configure(text=subtitle)

    def set_color(self, color: str):
        """Met à jour la couleur de la valeur"""
        self.color = color
        self.value_label.configure(text_color=color)

    def highlight(self, success: bool = True):
        """Met en évidence la carte (succès ou erreur)"""
        color = COLORS["success"] if success else COLORS["error"]
        self.set_color(color)


class StatCardGroup(ctk.CTkFrame):
    """
    Groupe de cartes de statistiques disposées horizontalement
    """

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        self.cards = {}
        self.configure(fg_color="transparent")

    def add_card(
        self,
        card_id: str,
        title: str,
        value: str = "0",
        icon: str = "📊",
        subtitle: str = "",
        tooltip: str = "",
        color: str = COLORS["text_primary"]
    ) -> StatCard:
        """Ajoute une carte au groupe"""
        card = StatCard(
            self,
            title=title,
            value=value,
            icon=icon,
            subtitle=subtitle,
            tooltip=tooltip,
            color=color
        )
        card.pack(side="left", fill="both", expand=True, padx=5)

        self.cards[card_id] = card
        return card

    def get_card(self, card_id: str) -> Optional[StatCard]:
        """Récupère une carte par son ID"""
        return self.cards.get(card_id)

    def update_card(self, card_id: str, value: str, subtitle: str = None):
        """Met à jour une carte"""
        card = self.cards.get(card_id)
        if card:
            card.set_value(value)
            if subtitle is not None:
                card.set_subtitle(subtitle)

    def reset_all(self):
        """Réinitialise toutes les cartes à 0"""
        for card in self.cards.values():
            card.set_value("0")
            card.set_color(COLORS["text_primary"])
