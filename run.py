#!/usr/bin/env python3
"""
ExcelToolsPro - Point d'entrée principal
Lance l'application graphique unifiée
"""

import sys
from pathlib import Path

# Ajouter le dossier src au path
sys.path.insert(0, str(Path(__file__).parent))

from src.ui.main_app import main

if __name__ == "__main__":
    main()
