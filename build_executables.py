#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script de build pour ExcelToolsPro
Génère les exécutables Production et Debug
"""

import subprocess
import sys
import shutil
from pathlib import Path


def clean_build_dirs():
    """Nettoie les répertoires de build"""
    dirs_to_clean = ['build', 'dist']
    for dir_name in dirs_to_clean:
        dir_path = Path(dir_name)
        if dir_path.exists():
            print(f"Nettoyage de {dir_name}/...")
            shutil.rmtree(dir_path)


def build_executable(spec_file: str, name: str):
    """Build un exécutable à partir d'un fichier spec"""
    print(f"\n{'='*60}")
    print(f"Construction de {name}...")
    print(f"{'='*60}\n")

    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--clean',
        '--noconfirm',
        spec_file
    ]

    result = subprocess.run(cmd, capture_output=False)

    if result.returncode == 0:
        print(f"\n[OK] {name} créé avec succès!")
        return True
    else:
        print(f"\n[ERREUR] Échec de la création de {name}")
        return False


def main():
    """Point d'entrée principal"""
    print("=" * 60)
    print("ExcelToolsPro - Build Script v1.0.0")
    print("=" * 60)

    # Vérifier PyInstaller
    try:
        import PyInstaller
        print(f"PyInstaller version: {PyInstaller.__version__}")
    except ImportError:
        print("ERREUR: PyInstaller n'est pas installé.")
        print("Installez-le avec: pip install pyinstaller")
        sys.exit(1)

    # Nettoyer les anciens builds
    print("\nNettoyage des anciens builds...")
    clean_build_dirs()

    # Build Production
    success_prod = build_executable('ExcelToolsPro.spec', 'ExcelToolsPro (Production)')

    # Build Debug
    success_debug = build_executable('ExcelToolsPro_debug.spec', 'ExcelToolsPro_debug (Debug)')

    # Résumé
    print("\n" + "=" * 60)
    print("RÉSUMÉ DU BUILD")
    print("=" * 60)

    dist_dir = Path('dist')
    if dist_dir.exists():
        for exe in dist_dir.glob('*.exe'):
            size_mb = exe.stat().st_size / (1024 * 1024)
            print(f"  - {exe.name}: {size_mb:.2f} MB")

    print("\n" + "=" * 60)
    if success_prod and success_debug:
        print("BUILD TERMINÉ AVEC SUCCÈS!")
        print("\nLes exécutables sont disponibles dans le dossier 'dist/'")
        print("  - ExcelToolsPro.exe       : Version production (sans console)")
        print("  - ExcelToolsPro_debug.exe : Version debug (avec console)")
    else:
        print("BUILD TERMINÉ AVEC DES ERREURS")
        sys.exit(1)


if __name__ == '__main__':
    main()
