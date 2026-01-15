"""
Composant Tooltip moderne pour CustomTkinter
Info-bulles avec style cohérent
"""

import tkinter as tk
from typing import Optional


class Tooltip:
    """
    Classe pour créer des info-bulles modernes.

    Usage:
        Tooltip(widget, "Texte de l'info-bulle")
    """

    def __init__(
        self,
        widget,
        text: str,
        delay: int = 500,
        wraplength: int = 300,
        bg_color: str = "#2d2d2d",
        fg_color: str = "#ffffff",
        font: tuple = ("Segoe UI", 10)
    ):
        """
        Initialise une info-bulle

        Args:
            widget: Widget auquel attacher l'info-bulle
            text: Texte à afficher
            delay: Délai avant affichage (ms)
            wraplength: Largeur max avant retour à la ligne
            bg_color: Couleur de fond
            fg_color: Couleur du texte
            font: Police à utiliser
        """
        self.widget = widget
        self.text = text
        self.delay = delay
        self.wraplength = wraplength
        self.bg_color = bg_color
        self.fg_color = fg_color
        self.font = font

        self.tooltip_window: Optional[tk.Toplevel] = None
        self.scheduled_id: Optional[str] = None

        # Bindings
        widget.bind("<Enter>", self._schedule_show)
        widget.bind("<Leave>", self._hide)
        widget.bind("<ButtonPress>", self._hide)

    def _schedule_show(self, event=None):
        """Planifie l'affichage de l'info-bulle"""
        self._cancel_schedule()
        self.scheduled_id = self.widget.after(self.delay, self._show)

    def _cancel_schedule(self):
        """Annule l'affichage planifié"""
        if self.scheduled_id:
            self.widget.after_cancel(self.scheduled_id)
            self.scheduled_id = None

    def _show(self, event=None):
        """Affiche l'info-bulle"""
        if self.tooltip_window or not self.text:
            return

        # Calculer la position
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5

        # Créer la fenêtre
        self.tooltip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tw.configure(bg=self.bg_color)

        # Cadre avec bordure subtile
        frame = tk.Frame(
            tw,
            bg=self.bg_color,
            padx=10,
            pady=6,
            highlightbackground="#404040",
            highlightthickness=1
        )
        frame.pack()

        # Label du texte
        label = tk.Label(
            frame,
            text=self.text,
            justify="left",
            bg=self.bg_color,
            fg=self.fg_color,
            font=self.font,
            wraplength=self.wraplength
        )
        label.pack()

        # S'assurer que la fenêtre reste visible à l'écran
        tw.update_idletasks()
        screen_width = tw.winfo_screenwidth()
        screen_height = tw.winfo_screenheight()
        tooltip_width = tw.winfo_width()
        tooltip_height = tw.winfo_height()

        if x + tooltip_width > screen_width:
            x = screen_width - tooltip_width - 10
        if y + tooltip_height > screen_height:
            y = self.widget.winfo_rooty() - tooltip_height - 5

        tw.wm_geometry(f"+{x}+{y}")

    def _hide(self, event=None):
        """Cache l'info-bulle"""
        self._cancel_schedule()
        if self.tooltip_window:
            self.tooltip_window.destroy()
            self.tooltip_window = None

    def update_text(self, new_text: str):
        """Met à jour le texte de l'info-bulle"""
        self.text = new_text

    def destroy(self):
        """Détruit proprement l'info-bulle"""
        self._hide()
        try:
            self.widget.unbind("<Enter>")
            self.widget.unbind("<Leave>")
            self.widget.unbind("<ButtonPress>")
        except Exception:
            pass
