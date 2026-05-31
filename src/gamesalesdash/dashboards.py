"""Orchestration : assemblage du classeur Excel complet en mémoire."""

import pandas as pd
from openpyxl import Workbook

from gamesalesdash.sheets.calc_sheet import build_calc_sheet
from gamesalesdash.sheets.cleaned_data import build_cleaned_data
from gamesalesdash.sheets.resources import build_resources_sheet
from gamesalesdash.sheets.tdb1 import build_tdb1
from gamesalesdash.sheets.tdb2 import build_tdb2


def build_workbook(df: pd.DataFrame, len_dict: dict[str, int]) -> Workbook:
    """Construit le classeur Excel complet en mémoire.

    Les cinq onglets sont créés dans l'ordre : cleaned_data, Ressources,
    TDB_1, calc_sheet, TDB_2. La sauvegarde sur disque est laissée à l'appelant.

    Args:
        df: DataFrame nettoyé à écrire dans cleaned_data.
        len_dict: Tailles des listes uniques pour borner les formules.

    Returns:
        Le classeur openpyxl prêt à être sauvegardé.
    """
    wb = Workbook()
    wb.remove(wb.active)

    build_cleaned_data(wb, df)
    build_resources_sheet(wb, len_dict)
    build_tdb1(wb, len_dict)
    build_calc_sheet(wb)
    build_tdb2(wb)

    return wb
