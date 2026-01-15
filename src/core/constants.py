"""
Constantes globales de l'application ExcelToolsPro
"""

# Informations de l'application
APP_INFO = {
    "name": "ExcelToolsPro",
    "version": "1.0.0",
    "author": "Edvance",
    "description": "Suite professionnelle d'outils Excel",
    "github": "https://github.com/edvance/exceltoolspro"
}

# Palette de couleurs moderne
COLORS = {
    # Fond
    "bg_dark": "#1a1a2e",
    "bg_card": "#16213e",
    "bg_secondary": "#0f3460",

    # Accents
    "accent_primary": "#0f3460",
    "accent_secondary": "#e94560",
    "accent_tertiary": "#533483",

    # États
    "success": "#00bf63",
    "warning": "#ffbd59",
    "error": "#ff6b6b",
    "info": "#4da6ff",

    # Texte
    "text_primary": "#eaeaea",
    "text_secondary": "#a0a0a0",
    "text_muted": "#6c757d",

    # Bordures
    "border_light": "#3d3d6a",
    "border_dark": "#1e1e2e",
}

# Types de fichiers supportés
FILE_TYPES = {
    "excel": [
        ("Fichiers Excel", "*.xlsx *.xls *.xlsm"),
        ("Excel 2007+ (.xlsx)", "*.xlsx"),
        ("Excel avec macros (.xlsm)", "*.xlsm"),
        ("Excel 97-2003 (.xls)", "*.xls"),
    ],
    "csv": [
        ("Fichiers CSV", "*.csv"),
        ("Tous les fichiers", "*.*"),
    ],
    "all": [
        ("Tous les fichiers", "*.*"),
    ]
}

# États des étapes de workflow
class StepStatus:
    PENDING = "pending"
    RUNNING = "running"
    SUCCESS = "success"
    ERROR = "error"
    SKIPPED = "skipped"
    WARNING = "warning"

# Icônes des états
STATUS_ICONS = {
    StepStatus.PENDING: "⏳",
    StepStatus.RUNNING: "🔄",
    StepStatus.SUCCESS: "✅",
    StepStatus.ERROR: "❌",
    StepStatus.SKIPPED: "⏭️",
    StepStatus.WARNING: "⚠️",
}

# Niveaux de log
LOG_LEVELS = {
    "DEBUG": 10,
    "INFO": 20,
    "WARNING": 30,
    "ERROR": 40,
    "CRITICAL": 50,
}
