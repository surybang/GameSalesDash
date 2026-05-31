"""Chargement et nettoyage des données source."""

import pandas as pd
from loguru import logger

from gamesalesdash.config import DATA_URL


def load_data() -> pd.DataFrame:
    """Charge le CSV depuis l'URL publique MinIO définie dans la config.

    Returns:
        DataFrame brut.
    """
    logger.info(f"Chargement des données depuis {DATA_URL}")
    return pd.read_csv(DATA_URL)


def clean_data(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise et nettoie le DataFrame brut.

    Applique dans l'ordre :
    - Normalisation des noms de colonnes en minuscules.
    - Suppression des lignes avec des valeurs manquantes.
    - Conversion de la colonne year de float à int.

    Args:
        df: DataFrame brut chargé depuis le CSV source.

    Returns:
        DataFrame nettoyé, prêt à être écrit dans cleaned_data.
    """
    df = df.copy()
    df.columns = df.columns.str.lower()
    df = df.dropna()
    df["year"] = df["year"].astype(int)
    return df


def compute_list_lengths(df: pd.DataFrame) -> dict[str, int]:
    """Calcule le nombre de valeurs uniques par colonne pour les formules Excel.

    Le +1 est nécessaire car les formules UNIQUE démarrent à la ligne 1
    et la plage doit couvrir une ligne supplémentaire.

    Args:
        df: DataFrame nettoyé.

    Returns:
        Dictionnaire avec les clés len_platform, len_genre, len_publisher,
        len_year et leur valeur respective (nb valeurs uniques + 1).
    """
    cols = ["platform", "genre", "publisher", "year"]
    return {f"len_{col}": len(df[col].unique()) + 1 for col in cols}
