"""Constantes du projet : chemins, couleurs, références Excel, données de configuration."""

from pathlib import Path

# ---------------------------------------------------------------------------
# Chemins
# ---------------------------------------------------------------------------

DATA_URL: str = "https://minio.lab.sspcloud.fr/fabienhos/GameSalesDash/vgsales.csv"
OUTPUT_PATH: Path = Path("result/Dashboards_videogames3.xlsx")

# ---------------------------------------------------------------------------
# DEFAULT VAL FILTERS
# ---------------------------------------------------------------------------

DEFAULT_PLATFORM1 = "PS3"
DEFAULT_PLATFORM2 = "X360"

# ---------------------------------------------------------------------------
# Noms des onglets
# ---------------------------------------------------------------------------

SHEET_CLEANED = "cleaned_data"
SHEET_RESOURCES = "Ressources"
SHEET_TDB1 = "TDB_1"
SHEET_CALC = "calc_sheet"
SHEET_TDB2 = "TDB_2"

# ---------------------------------------------------------------------------
# Palette de couleurs
# ---------------------------------------------------------------------------

DARK_BLUE = "8DB4E2"
LIGHT_BLUE = "C5D9F1"
DARK_GREEN = "C4D79B"
LIGHT_GREEN = "EBF1DE"
GREY = "C0C0C0"

PIE_COLORS = [
    "8DB4E2",
    "C5D9F1",
    "C4D79B",
    "EBF1DE",
    "4F81BD",
    "76933C",
]

# ---------------------------------------------------------------------------
# Colonnes de cleaned_data (lettres Excel, pour les formules)
# ---------------------------------------------------------------------------

COL_RANK = "A"
COL_NAME = "B"
COL_PLATFORM = "C"
COL_YEAR = "D"
COL_GENRE = "E"
COL_PUBLISHER = "F"
COL_NA_SALES = "G"
COL_EU_SALES = "H"
COL_JP_SALES = "I"
COL_OTHER_SALES = "J"
COL_GLOBAL_SALES = "K"

# ---------------------------------------------------------------------------
# Références des cellules de filtre dans TDB_1
# ---------------------------------------------------------------------------

TDB1_YEAR_REF = "$C$1"
TDB1_PLATFORM1_REF = "$G$1"
TDB1_PLATFORM2_REF = "$G$3"
TDB1_GENRE_REF = "$K$1"

# ---------------------------------------------------------------------------
# Éditeurs affichés dans calc_sheet / TDB_2
# ---------------------------------------------------------------------------

PUBLISHERS = [
    "Activision",
    "Nintendo",
    "Take-Two Interactive",
    "Ubisoft",
    "Square Enix",
    "Electronic Arts",
]
