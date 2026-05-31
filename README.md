# GameSalesDash

Automatisation d'un reporting Excel sur les ventes de jeux vidéo. Le script génère un classeur `.xlsx` structuré à partir du dataset [Video Game Sales](https://www.kaggle.com/datasets/gregorut/videogamesales) (Kaggle), comprenant deux dashboards interactifs avec filtres, tableaux TOP 5, graphiques et indicateurs clés.

Les données brutes sont hébergées sur MinIO et téléchargées automatiquement à chaque lancement. Le notebook d'exploration original est conservé tel quel dans `notebooks/`.

---

## Prérequis

- [uv](https://docs.astral.sh/uv/getting-started/installation/) installé sur la machine
- Python 3.11+

---

## Installation

```bash
git clone https://github.com/surybang/GameSalesDash.git
cd GameSalesDash
uv sync
```

---

## Utilisation

```bash
uv run make_reporting_videogames
```

Le fichier est produit dans `result/Dashboards_videogames.xlsx`. Une version générée est également disponible directement sur [MinIO](https://minio.lab.sspcloud.fr/fabienhos/GameSalesDash/Dashboards_videogames.xlsx).

---

## Structure du projet

```
GameSalesDash/
├── pyproject.toml
├── uv.lock
├── .python-version
│
├── data/                       # Données
├── result/                     # Fichiers Excel générés
├── img/                        # Captures d'écran utilisées
│
├── notebooks/
│   └── projet.ipynb            # Notebook d'exploration original
│
└── src/gamesalesdash/
    ├── __init__.py
    ├── config.py               # Constantes 
    ├── data.py                 # Chargement MinIO et nettoyage des données
    ├── dashboards.py           # Orchestration : assemble les onglets
    ├── main.py                 # Point d'entrée
    │
    ├── components/             # Briques réutilisables de construction Excel
    │   ├── filters.py          # Filtres déroulants (validation de données)
    │   ├── styles.py           # Bordures et styles partagés
    │   ├── tables.py           # Tableaux TOP 5 (titres, en-têtes, formules matricielles)
    │   ├── charts.py           # Graphiques (barres, courbes, camemberts)
    │   └── indicators.py       # Cartes indicateurs (COUNTIFS / SUMIFS)
    │
    └── sheets/                 # Construction par onglet
        ├── cleaned_data.py     # Onglet cleaned_data
        ├── resources.py        # Onglet Ressources (listes UNIQUE/SORT)
        ├── tdb1.py             # Onglet TDB_1 (dashboard principal)
        ├── calc_sheet.py       # Onglet calc_sheet (agrégats)
        └── tdb2.py             # Onglet TDB_2 (dashboard d'ensemble)
```

### Rôle de chaque module

**`config.py`** centralise toutes les valeurs dispersées dans le notebook : URL MinIO, noms d'onglets, palette de couleurs, mapping des colonnes de `cleaned_data` et valeurs par défaut des filtres.

**`data.py`** expose `load_data()` qui charge le CSV depuis MinIO, et `clean_data()` qui applique le nettoyage à la volée (suppression des lignes avec `year` ou `publisher` manquants, conversion `year` en entier).

**`components/`** regroupe les briques de construction Excel réutilisables : chaque module correspond à une responsabilité précise (filtres, styles, tables, graphiques, indicateurs) et ne connaît pas la structure des dashboards.

**`sheets/`** contient un module par onglet du classeur. Chaque module orchestre ses propres briques pour construire un onglet complet.

**`dashboards.py`** assemble les onglets dans l'ordre et retourne le classeur. Le workbook est construit entièrement en mémoire et sauvegardé une seule fois.

**`main.py`** est le point d'entrée CLI. Il appelle le pipeline complet : chargement, nettoyage, construction, sauvegarde.

---

## Données

Le dataset provient de [Kaggle – Video Game Sales](https://www.kaggle.com/datasets/gregorut/videogamesales). Il contient 16 598 jeux ayant dépassé 100 000 ventes, avec les champs suivants : rang, nom, plateforme, année, genre, éditeur, et les ventes par région (NA, EU, JP, Autres) ainsi que les ventes globales.

Le nettoyage supprime 307 lignes incomplètes (valeurs manquantes sur `year` ou `publisher`) et est appliqué à la volée à chaque run.