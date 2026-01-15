"""
Configuration pytest pour ExcelToolsPro
Fixtures partagées et configuration des tests
"""

import pytest
import pandas as pd
import tempfile
import os
import sys
from pathlib import Path

# Ajouter le répertoire racine au path
sys.path.insert(0, str(Path(__file__).parent.parent))


@pytest.fixture
def sample_dataframe():
    """DataFrame de test standard"""
    return pd.DataFrame({
        "ID": ["001", "002", "003", "004", "005"],
        "Nom": ["Alice", "Bob", "Charlie", "David", "Eve"],
        "Age": ["25", "30", "35", "28", "32"],
        "Ville": ["Paris", "Lyon", "Marseille", "Toulouse", "Bordeaux"],
        "Email": ["alice@test.com", "bob@test.com", "charlie@test.com", "david@test.com", "eve@test.com"]
    })


@pytest.fixture
def sample_dataframe_with_ref():
    """DataFrame avec colonne de référence pour tests de fusion"""
    return pd.DataFrame({
        "REF": ["REF001", "REF002", "REF003"],
        "Description": ["Produit A", "Produit B", "Produit C"],
        "Prix": ["100", "200", "150"]
    })


@pytest.fixture
def reference_dataframe():
    """DataFrame de référence pour tests de fusion"""
    return pd.DataFrame({
        "REF": ["REF001", "REF002", "REF004"],
        "Categorie": ["Cat1", "Cat2", "Cat3"],
        "Stock": ["50", "30", "20"],
        "LAST": ["Y", "Y", "N"]
    })


def _safe_remove(path):
    """Supprime un fichier en ignorant les erreurs de permission Windows"""
    import gc
    gc.collect()  # Libérer les handles
    try:
        if os.path.exists(path):
            os.remove(path)
    except (PermissionError, OSError):
        pass  # Ignorer les erreurs de permission sur Windows


def _safe_rmtree(path):
    """Supprime un répertoire en ignorant les erreurs de permission Windows"""
    import shutil
    import gc
    gc.collect()  # Libérer les handles
    try:
        if os.path.exists(path):
            shutil.rmtree(path, ignore_errors=True)
    except (PermissionError, OSError):
        pass


@pytest.fixture
def temp_excel_file(sample_dataframe):
    """Crée un fichier Excel temporaire"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        filepath = tmp.name

    sample_dataframe.to_excel(filepath, index=False, sheet_name="TestData")

    yield filepath

    _safe_remove(filepath)


@pytest.fixture
def temp_csv_file(sample_dataframe):
    """Crée un fichier CSV temporaire"""
    with tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode='w', encoding='utf-8') as tmp:
        filepath = tmp.name

    sample_dataframe.to_csv(filepath, index=False)

    yield filepath

    _safe_remove(filepath)


@pytest.fixture
def temp_directory():
    """Crée un répertoire temporaire"""
    tmpdir = tempfile.mkdtemp()
    yield tmpdir

    _safe_rmtree(tmpdir)


@pytest.fixture
def multi_sheet_excel_file(sample_dataframe, sample_dataframe_with_ref):
    """Crée un fichier Excel avec plusieurs feuilles"""
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        filepath = tmp.name

    with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
        sample_dataframe.to_excel(writer, sheet_name="Sheet1", index=False)
        sample_dataframe_with_ref.to_excel(writer, sheet_name="Sheet2", index=False)

    yield filepath

    _safe_remove(filepath)


@pytest.fixture
def empty_dataframe():
    """DataFrame vide"""
    return pd.DataFrame()


@pytest.fixture
def large_dataframe():
    """DataFrame volumineux pour tests de performance"""
    import random
    import string

    n_rows = 1000
    return pd.DataFrame({
        "ID": [f"ID{i:05d}" for i in range(n_rows)],
        "Name": [''.join(random.choices(string.ascii_letters, k=10)) for _ in range(n_rows)],
        "Value": [random.randint(1, 10000) for _ in range(n_rows)],
        "Category": [random.choice(["A", "B", "C", "D"]) for _ in range(n_rows)]
    })
