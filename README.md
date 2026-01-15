# Excel Tools Pro

**Suite d'outils Excel professionnelle** - Application Windows gratuite pour fusionner, rechercher, convertir et manipuler des fichiers Excel et CSV.

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)
![Version](https://img.shields.io/badge/Version-1.0.0-brightgreen.svg)

---

## Mots-clés

`excel tools` `fusionner excel` `csv to excel` `excel merger` `excel search tool` `data extraction excel` `outil excel gratuit` `excel automation` `bulk excel processing`

---

## Telechargement - Pret a l'emploi

**Aucune installation requise.** Telechargez et lancez directement :

| Fichier | Description | Utilisation |
|---------|-------------|-------------|
| **[ExcelToolsPro.exe](https://github.com/Kiriiaq/Excel-Tool-Pro/releases/latest)** | Version Production | Usage quotidien - Interface graphique sans console |
| **[ExcelToolsPro_debug.exe](https://github.com/Kiriiaq/Excel-Tool-Pro/releases/latest)** | Version Debug | Diagnostic - Console avec logs detailles |

> **Telecharger depuis** : [Releases GitHub](https://github.com/Kiriiaq/Excel-Tool-Pro/releases/latest)

---

## Probleme resolu

Excel Tools Pro simplifie les taches complexes de manipulation Excel :

- **Fusionner des fichiers Excel** sans formules compliquees ni macros
- **Rechercher dans des milliers de lignes** avec filtres avances et regex
- **Convertir CSV en Excel** (et inversement) en un clic
- **Transferer des donnees** entre fichiers avec mapping de colonnes

---

## Fonctionnalites

### Modules integres

| Module | Description |
|--------|-------------|
| **Fusion de Documents** | Fusionne deux fichiers Excel sur une colonne commune (jointure type VLOOKUP/RECHERCHEV) |
| **Recherche de Donnees** | Recherche avancee avec filtres, regex, operateurs logiques (AND/OR) |
| **Transfert de Donnees** | Extraction et transfert de donnees entre fichiers avec champs configurables |
| **Conversion CSV/Excel** | Conversion bidirectionnelle CSV - Excel, fusion de fichiers multiples |

### Caracteristiques techniques

- **Interface moderne** : GUI CustomTkinter avec theme sombre professionnel
- **Architecture modulaire** : Modules independants et extensibles
- **Configuration persistante** : Parametres sauvegardes entre sessions
- **Logging temps reel** : Journal d'activite avec niveaux de detail
- **Export professionnel** : Excel formate (en-tetes, couleurs alternees, colonnes auto-ajustees)
- **Gestion d'erreurs** : Robuste avec option de continuation sur erreur

---

## Captures d'ecran

### Interface principale

```
+--------------------------------------------------------------+
|  Excel Tools Pro v1.0.0                            [Config]  |
+--------------+-----------------------------------------------+
|  MODULES     |                                               |
|              |     [Zone de travail du module actif]         |
|  Fusion      |                                               |
|  Recherche   |     - Selection de fichiers                   |
|  Transfert   |     - Configuration des parametres            |
|  CSV         |     - Previsualisation des donnees            |
|              |     - Actions et export                       |
|--------------+                                               |
|  JOURNAL     |                                               |
|  [Logs...]   |                                               |
+--------------+-----------------------------------------------+
```

---

## Difference entre versions

| Aspect | Production (.exe) | Debug (.exe) |
|--------|-------------------|--------------|
| Console | Masquee | Visible avec logs |
| Logs | Fichier uniquement | Console + Fichier |
| Performance | Optimisee | Standard |
| Usage | Utilisateur final | Developpeur/Debug |

---

## Installation (Developpeurs)

### Prerequis

- Python 3.9 ou superieur
- pip (gestionnaire de paquets Python)

### Installation des dependances

```bash
# Cloner le repository
git clone https://github.com/Kiriiaq/Excel-Tool-Pro.git
cd Excel-Tool-Pro

# Installer les dependances
pip install -r requirements.txt
```

### Lancement depuis les sources

```bash
python run.py
```

### Execution des tests

```bash
python -m pytest tests/ -v
```

### Creation des executables

```bash
python build_executables.py
```

Les executables seront generes dans le dossier `dist/`.

---

## Architecture du projet

```
Excel-Tool-Pro/
|-- src/
|   |-- core/           # Configuration, constantes, logging
|   |-- ui/             # Interface graphique (CustomTkinter)
|   |-- modules/        # Modules fonctionnels
|   +-- utils/          # Utilitaires (Excel, fichiers, validation)
|-- tests/              # Tests unitaires et integration
|-- ico/                # Icone de l'application
|-- dist/               # Executables generes
|-- requirements.txt
|-- LICENSE
+-- README.md
```

---

## Utilisation rapide

### Fusionner deux fichiers Excel

1. Selectionnez le fichier source et le fichier de reference
2. Choisissez la colonne cle dans chaque fichier
3. Previsualiser le resultat
4. Exporter au format Excel

### Rechercher dans un fichier Excel

1. Chargez le fichier a analyser
2. Entrez les mots-cles (virgule = OR)
3. Choisissez le mode : Contient, Mot exact, Regex, etc.
4. Exportez les resultats filtres

### Convertir CSV vers Excel

1. Selectionnez le(s) fichier(s) CSV
2. Configurez l'encodage et le separateur
3. Cliquez sur Convertir
4. Recuperez le fichier Excel formate

---

## Configuration systeme requise

- **OS** : Windows 10 / Windows 11
- **Espace disque** : ~150 MB
- **RAM** : 4 GB minimum (8 GB recommande pour gros fichiers)
- **Dependances** : Aucune (Python embarque dans l'executable)

---

## Roadmap

- [ ] Export PDF des resultats
- [ ] Mode batch en ligne de commande
- [ ] Connexion base de donnees
- [ ] Plugins tiers

---

## Contribution

Les contributions sont les bienvenues !

1. Fork le projet
2. Creez une branche (`git checkout -b feature/nouvelle-fonctionnalite`)
3. Committez vos changements
4. Ouvrez une Pull Request

---

## Licence

Ce projet est sous licence MIT. Voir le fichier [LICENSE](LICENSE) pour plus de details.

---

## Credits

- **Developpement** : [Edvance](https://github.com/Kiriiaq)
- **Framework GUI** : [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter)
- **Traitement Excel** : [pandas](https://pandas.pydata.org/), [openpyxl](https://openpyxl.readthedocs.io/)

---

## Support

- [Issues GitHub](https://github.com/Kiriiaq/Excel-Tool-Pro/issues)

---

*Excel Tools Pro - Simplifiez vos manipulations Excel*
