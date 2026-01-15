"""
Panneau de paramètres avancé avec sections repliables
Génère automatiquement l'interface à partir des définitions de paramètres
"""

import customtkinter as ctk
from tkinter import filedialog, colorchooser
from typing import Dict, Any, Optional, Callable, List
from dataclasses import dataclass
from enum import Enum

from .tooltip import Tooltip
from ...core.constants import COLORS


class SettingType(Enum):
    """Types de paramètres supportés"""
    BOOLEAN = "boolean"
    INTEGER = "integer"
    FLOAT = "float"
    STRING = "string"
    CHOICE = "choice"
    FILE = "file"
    DIRECTORY = "directory"
    SLIDER = "slider"
    COLOR = "color"
    TEXT = "text"  # Multiline text
    LIST = "list"  # List of strings


@dataclass
class SettingDefinition:
    """Définition d'un paramètre"""
    key: str
    label: str
    type: SettingType
    default: Any = None
    tooltip: str = ""
    category: str = "Général"
    advanced: bool = False

    # Pour CHOICE
    choices: List[str] = None

    # Pour SLIDER et INTEGER
    min_value: int = 0
    max_value: int = 100
    step: int = 1

    # Pour FILE/DIRECTORY
    file_types: List[tuple] = None

    # Validation
    validator: Callable[[Any], bool] = None
    error_message: str = ""

    # Dépendances (clé du paramètre dont dépend celui-ci)
    depends_on: str = None
    depends_value: Any = True


class CollapsibleSection(ctk.CTkFrame):
    """Section repliable pour organiser les paramètres"""

    def __init__(
        self,
        parent,
        title: str,
        icon: str = "",
        initially_expanded: bool = True,
        **kwargs
    ):
        super().__init__(parent, fg_color=COLORS["bg_card"], corner_radius=10, **kwargs)

        self.title = title
        self.is_expanded = initially_expanded

        # Header cliquable
        self.header = ctk.CTkFrame(self, fg_color="transparent", cursor="hand2")
        self.header.pack(fill="x", padx=10, pady=(10, 5))
        self.header.bind("<Button-1>", self._toggle)

        # Icône d'expansion
        self.expand_label = ctk.CTkLabel(
            self.header,
            text="▼" if initially_expanded else "▶",
            font=("Segoe UI", 10),
            width=20
        )
        self.expand_label.pack(side="left")
        self.expand_label.bind("<Button-1>", self._toggle)

        # Titre
        title_text = f"{icon} {title}" if icon else title
        self.title_label = ctk.CTkLabel(
            self.header,
            text=title_text,
            font=("Segoe UI", 13, "bold"),
            anchor="w"
        )
        self.title_label.pack(side="left", padx=(5, 0))
        self.title_label.bind("<Button-1>", self._toggle)

        # Bouton reset section
        self.reset_btn = ctk.CTkButton(
            self.header,
            text="↺",
            width=25,
            height=22,
            font=("Segoe UI", 10),
            fg_color="transparent",
            hover_color=COLORS["warning"]
        )
        self.reset_btn.pack(side="right")
        Tooltip(self.reset_btn, "Réinitialiser cette section")

        # Contenu
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        if initially_expanded:
            self.content_frame.pack(fill="x", padx=15, pady=(0, 15))

    def _toggle(self, event=None):
        """Bascule l'état d'expansion"""
        self.is_expanded = not self.is_expanded

        if self.is_expanded:
            self.expand_label.configure(text="▼")
            self.content_frame.pack(fill="x", padx=15, pady=(0, 15))
        else:
            self.expand_label.configure(text="▶")
            self.content_frame.pack_forget()

    def get_content_frame(self) -> ctk.CTkFrame:
        """Retourne le frame de contenu pour y ajouter des widgets"""
        return self.content_frame

    def set_reset_command(self, command: Callable):
        """Définit la commande du bouton reset"""
        self.reset_btn.configure(command=command)


class SettingsPanel(ctk.CTkScrollableFrame):
    """
    Panneau de paramètres avec génération automatique de l'interface

    Organise les paramètres par catégories dans des sections repliables.
    Supporte les paramètres avancés (repliés par défaut).
    """

    def __init__(
        self,
        parent,
        settings: List[SettingDefinition],
        on_change: Optional[Callable[[str, Any], None]] = None,
        show_advanced: bool = False,
        **kwargs
    ):
        super().__init__(parent, fg_color="transparent", **kwargs)

        self.settings = settings
        self.settings_defs = {s.key: s for s in settings}
        self.on_change = on_change
        self.show_advanced = show_advanced

        self.values: Dict[str, Any] = {}
        self.widgets: Dict[str, Any] = {}
        self.sections: Dict[str, CollapsibleSection] = {}

        # Initialiser les valeurs par défaut
        for s in settings:
            self.values[s.key] = s.default

        self._build_interface()

    def _build_interface(self):
        """Construit l'interface des paramètres"""
        # Grouper par catégorie
        categories = {}
        for setting in self.settings:
            cat = setting.category
            if cat not in categories:
                categories[cat] = []
            categories[cat].append(setting)

        # Icônes par catégorie
        category_icons = {
            "Général": "⚙️",
            "Apparence": "🎨",
            "Interface": "🖥️",
            "Export Excel": "📊",
            "Recherche": "🔍",
            "Fusion": "🔗",
            "Transfert": "📋",
            "CSV": "📄",
            "Performance": "⚡",
            "Logs": "📝",
            "Chemins": "📁",
            "Avancé": "🔧",
            "Comportement": "⚙️",
        }

        # Créer les sections
        for category, cat_settings in categories.items():
            # Vérifier s'il y a des paramètres non-avancés
            has_basic = any(not s.advanced for s in cat_settings)

            if not has_basic and not self.show_advanced:
                continue

            # Créer la section
            icon = category_icons.get(category, "")
            is_advanced = all(s.advanced for s in cat_settings)
            initially_expanded = not is_advanced

            section = CollapsibleSection(
                self,
                title=category,
                icon=icon,
                initially_expanded=initially_expanded
            )
            section.pack(fill="x", pady=(0, 10))
            section.set_reset_command(lambda cat=category: self._reset_category(cat))
            self.sections[category] = section

            content = section.get_content_frame()

            # Ajouter les paramètres
            for setting in cat_settings:
                if setting.advanced and not self.show_advanced:
                    continue

                self._create_setting_widget(content, setting)

        # Boutons de contrôle en bas
        self._create_control_buttons()

    def _create_setting_widget(self, parent: ctk.CTkFrame, setting: SettingDefinition):
        """Crée le widget approprié pour un paramètre"""
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill="x", pady=5)

        # Label
        label = ctk.CTkLabel(
            frame,
            text=setting.label,
            font=("Segoe UI", 11),
            anchor="w",
            width=220
        )
        label.pack(side="left")

        if setting.tooltip:
            Tooltip(label, setting.tooltip)

        # Widget selon le type
        widget = None

        if setting.type == SettingType.BOOLEAN:
            var = ctk.BooleanVar(value=setting.default or False)
            widget = ctk.CTkSwitch(
                frame,
                text="",
                variable=var,
                width=50,
                command=lambda k=setting.key, v=var: self._on_value_change(k, v.get())
            )
            widget.pack(side="right")
            self.widgets[setting.key] = {"widget": widget, "var": var}

        elif setting.type == SettingType.INTEGER:
            var = ctk.StringVar(value=str(setting.default or 0))
            widget = ctk.CTkEntry(frame, textvariable=var, width=100)
            widget.pack(side="right")
            widget.bind("<FocusOut>", lambda e, k=setting.key, v=var:
                       self._on_entry_change(k, v.get(), int))
            self.widgets[setting.key] = {"widget": widget, "var": var}

        elif setting.type == SettingType.FLOAT:
            var = ctk.StringVar(value=str(setting.default or 0.0))
            widget = ctk.CTkEntry(frame, textvariable=var, width=100)
            widget.pack(side="right")
            widget.bind("<FocusOut>", lambda e, k=setting.key, v=var:
                       self._on_entry_change(k, v.get(), float))
            self.widgets[setting.key] = {"widget": widget, "var": var}

        elif setting.type == SettingType.STRING:
            var = ctk.StringVar(value=setting.default or "")
            widget = ctk.CTkEntry(frame, textvariable=var, width=200)
            widget.pack(side="right")
            widget.bind("<FocusOut>", lambda e, k=setting.key, v=var:
                       self._on_value_change(k, v.get()))
            self.widgets[setting.key] = {"widget": widget, "var": var}

        elif setting.type == SettingType.CHOICE:
            var = ctk.StringVar(value=setting.default or "")
            widget = ctk.CTkComboBox(
                frame,
                values=setting.choices or [],
                variable=var,
                width=200,
                command=lambda val, k=setting.key: self._on_value_change(k, val)
            )
            widget.pack(side="right")
            self.widgets[setting.key] = {"widget": widget, "var": var}

        elif setting.type == SettingType.SLIDER:
            slider_frame = ctk.CTkFrame(frame, fg_color="transparent")
            slider_frame.pack(side="right")

            var = ctk.IntVar(value=setting.default or setting.min_value)
            value_label = ctk.CTkLabel(slider_frame, text=str(var.get()), width=40)
            value_label.pack(side="right", padx=(10, 0))

            steps = max(1, (setting.max_value - setting.min_value) // setting.step)
            widget = ctk.CTkSlider(
                slider_frame,
                from_=setting.min_value,
                to=setting.max_value,
                number_of_steps=steps,
                variable=var,
                width=150,
                command=lambda val, k=setting.key, lbl=value_label:
                    self._on_slider_change(k, val, lbl)
            )
            widget.pack(side="right")
            self.widgets[setting.key] = {"widget": widget, "var": var, "label": value_label}

        elif setting.type == SettingType.COLOR:
            color_frame = ctk.CTkFrame(frame, fg_color="transparent")
            color_frame.pack(side="right")

            var = ctk.StringVar(value=setting.default or "#FFFFFF")

            color_preview = ctk.CTkLabel(
                color_frame,
                text="",
                width=30,
                height=25,
                fg_color=var.get(),
                corner_radius=5
            )
            color_preview.pack(side="right", padx=(5, 0))

            def choose_color(k=setting.key, v=var, preview=color_preview):
                color = colorchooser.askcolor(initialcolor=v.get())[1]
                if color:
                    v.set(color)
                    preview.configure(fg_color=color)
                    self._on_value_change(k, color)

            btn = ctk.CTkButton(
                color_frame,
                text="Choisir",
                width=70,
                height=25,
                command=choose_color
            )
            btn.pack(side="right")

            self.widgets[setting.key] = {"widget": btn, "var": var, "preview": color_preview}

        elif setting.type == SettingType.FILE:
            file_frame = ctk.CTkFrame(frame, fg_color="transparent")
            file_frame.pack(side="right")

            var = ctk.StringVar(value=setting.default or "")
            entry = ctk.CTkEntry(file_frame, textvariable=var, width=180, state="disabled")
            entry.pack(side="left")

            def browse_file(k=setting.key, v=var, ft=setting.file_types):
                filepath = filedialog.askopenfilename(
                    filetypes=ft or [("Tous", "*.*")]
                )
                if filepath:
                    v.set(filepath)
                    self._on_value_change(k, filepath)

            btn = ctk.CTkButton(file_frame, text="...", width=30, command=browse_file)
            btn.pack(side="left", padx=(5, 0))

            self.widgets[setting.key] = {"widget": entry, "var": var, "btn": btn}

        elif setting.type == SettingType.DIRECTORY:
            dir_frame = ctk.CTkFrame(frame, fg_color="transparent")
            dir_frame.pack(side="right")

            var = ctk.StringVar(value=setting.default or "")
            entry = ctk.CTkEntry(dir_frame, textvariable=var, width=180, state="disabled")
            entry.pack(side="left")

            def browse_dir(k=setting.key, v=var):
                dirpath = filedialog.askdirectory()
                if dirpath:
                    v.set(dirpath)
                    self._on_value_change(k, dirpath)

            btn = ctk.CTkButton(dir_frame, text="...", width=30, command=browse_dir)
            btn.pack(side="left", padx=(5, 0))

            self.widgets[setting.key] = {"widget": entry, "var": var, "btn": btn}

        elif setting.type == SettingType.TEXT:
            text_widget = ctk.CTkTextbox(frame, width=200, height=60)
            text_widget.insert("1.0", setting.default or "")
            text_widget.pack(side="right")

            def on_text_change(event, k=setting.key, w=text_widget):
                self._on_value_change(k, w.get("1.0", "end-1c"))

            text_widget.bind("<FocusOut>", on_text_change)
            self.widgets[setting.key] = {"widget": text_widget}

        elif setting.type == SettingType.LIST:
            list_frame = ctk.CTkFrame(frame, fg_color="transparent")
            list_frame.pack(side="right")

            default_list = setting.default if isinstance(setting.default, list) else []
            var = ctk.StringVar(value=", ".join(default_list))
            entry = ctk.CTkEntry(list_frame, textvariable=var, width=200)
            entry.pack(side="left")
            Tooltip(entry, "Valeurs séparées par des virgules")

            entry.bind("<FocusOut>", lambda e, k=setting.key, v=var:
                      self._on_value_change(k, [x.strip() for x in v.get().split(",") if x.strip()]))

            self.widgets[setting.key] = {"widget": entry, "var": var}

    def _create_control_buttons(self):
        """Crée les boutons de contrôle en bas du panneau"""
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", pady=(20, 10))

        # Toggle paramètres avancés
        self.advanced_var = ctk.BooleanVar(value=self.show_advanced)
        advanced_cb = ctk.CTkCheckBox(
            btn_frame,
            text="Afficher les options avancées",
            variable=self.advanced_var,
            command=self._toggle_advanced
        )
        advanced_cb.pack(side="left")

        # Bouton reset global
        reset_btn = ctk.CTkButton(
            btn_frame,
            text="Tout réinitialiser",
            width=120,
            fg_color=COLORS["error"],
            hover_color="#cc5555",
            command=self._reset_to_defaults
        )
        reset_btn.pack(side="right")
        Tooltip(reset_btn, "Réinitialiser tous les paramètres aux valeurs par défaut")

    def _on_value_change(self, key: str, value: Any):
        """Appelé lors d'un changement de valeur"""
        self.values[key] = value
        if self.on_change:
            self.on_change(key, value)

    def _on_entry_change(self, key: str, value: str, converter: type):
        """Callback pour les champs numériques avec validation"""
        try:
            converted = converter(value)
            setting = self.settings_defs.get(key)

            if setting:
                if setting.min_value is not None and converted < setting.min_value:
                    converted = setting.min_value
                if setting.max_value is not None and converted > setting.max_value:
                    converted = setting.max_value

                # Valider si un validateur est défini
                if setting.validator and not setting.validator(converted):
                    raise ValueError(setting.error_message or "Valeur invalide")

            self._on_value_change(key, converted)
        except ValueError:
            # Restaurer la valeur précédente
            if key in self.widgets:
                var = self.widgets[key].get("var")
                if var:
                    var.set(str(self.values.get(key, 0)))

    def _on_slider_change(self, key: str, value: float, label: ctk.CTkLabel):
        """Callback pour les sliders"""
        int_value = int(value)
        label.configure(text=str(int_value))
        self._on_value_change(key, int_value)

    def _toggle_advanced(self):
        """Bascule l'affichage des paramètres avancés"""
        self.show_advanced = self.advanced_var.get()

        # Reconstruire l'interface
        for widget in self.winfo_children():
            widget.destroy()

        self.widgets.clear()
        self.sections.clear()
        self._build_interface()

    def _reset_to_defaults(self):
        """Réinitialise tous les paramètres à leurs valeurs par défaut"""
        for setting in self.settings:
            self.set_value(setting.key, setting.default)
            self._on_value_change(setting.key, setting.default)

    def _reset_category(self, category: str):
        """Réinitialise les paramètres d'une catégorie"""
        for setting in self.settings:
            if setting.category == category:
                self.set_value(setting.key, setting.default)
                self._on_value_change(setting.key, setting.default)

    def get_value(self, key: str) -> Any:
        """Récupère la valeur d'un paramètre"""
        return self.values.get(key)

    def set_value(self, key: str, value: Any):
        """Définit la valeur d'un paramètre"""
        self.values[key] = value

        if key in self.widgets:
            widget_data = self.widgets[key]
            var = widget_data.get("var")

            if var:
                if isinstance(value, list):
                    var.set(", ".join(value))
                else:
                    var.set(value)

            # Cas spécial pour les couleurs
            preview = widget_data.get("preview")
            if preview:
                preview.configure(fg_color=value)

            # Cas spécial pour les textbox
            widget = widget_data.get("widget")
            if isinstance(widget, ctk.CTkTextbox):
                widget.delete("1.0", "end")
                widget.insert("1.0", str(value) if value else "")

            # Cas spécial pour les sliders
            label = widget_data.get("label")
            if label:
                label.configure(text=str(value))

    def get_all_values(self) -> Dict[str, Any]:
        """Récupère toutes les valeurs"""
        return self.values.copy()

    def set_all_values(self, values: Dict[str, Any]):
        """Définit toutes les valeurs"""
        for key, value in values.items():
            self.set_value(key, value)

    def reset_to_defaults(self):
        """Alias pour compatibilité"""
        self._reset_to_defaults()
