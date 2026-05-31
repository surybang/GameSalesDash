"""Onglet cleaned_data : écriture du DataFrame nettoyé."""

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from gamesalesdash.config import SHEET_CLEANED


def build_cleaned_data(wb: Workbook, df: pd.DataFrame) -> None:
    """Écrit le DataFrame nettoyé dans l'onglet cleaned_data.

    Args:
        wb: Classeur cible.
        df: DataFrame nettoyé.
    """
    ws = wb.create_sheet(SHEET_CLEANED)
    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)
