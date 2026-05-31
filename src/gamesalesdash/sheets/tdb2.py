"""Onglet TDB_2 : tableau de bord d'ensemble (évolutions, genres, éditeurs)."""

from openpyxl import Workbook

from gamesalesdash.components.charts import create_bar_chart, create_line_chart, create_pie_chart
from gamesalesdash.config import SHEET_CALC, SHEET_TDB2


def build_tdb2(wb: Workbook) -> None:
    """Construit le second tableau de bord TDB_2.

    Six graphiques alimentés par calc_sheet : évolution des ventes et des
    sorties par année, répartition par genre, et parts des éditeurs.

    Args:
        wb: Classeur cible.
    """
    tdb2 = wb.create_sheet(SHEET_TDB2)
    calc = wb[SHEET_CALC]
    tdb2.sheet_view.showGridLines = False

    create_line_chart(
        tdb2, "Evolution du nombre de jeux sortis par année",
        {"min_col": 3, "min_row": 1, "max_row": 39, "max_col": 3},
        {"min_col": 1, "min_row": 2, "max_row": 39},
        "A1", y_axis_title="Nb jeux sortis",
        source_sheet=calc,
    )
    create_line_chart(
        tdb2, "Evolution des ventes de jeux par année",
        {"min_col": 2, "min_row": 1, "max_row": 39, "max_col": 2},
        {"min_col": 1, "min_row": 2, "max_row": 39},
        "A16", y_axis_title="Nb jeux vendus (en millions)",
        source_sheet=calc,
    )
    create_bar_chart(
        tdb2, "Nombre de jeux par genre",
        {"min_col": 5, "min_row": 1, "max_row": 13, "max_col": 5},
        {"min_col": 4, "min_row": 2, "max_row": 13},
        "J1", style=11,
        source_sheet=calc,
    )
    create_bar_chart(
        tdb2, "Ventes tot. par genre",
        {"min_col": 6, "min_row": 1, "max_row": 13, "max_col": 6},
        {"min_col": 4, "min_row": 2, "max_row": 13},
        "J16", style=11,
        source_sheet=calc,
    )
    create_pie_chart(
        tdb2, "Nombre de jeux sortis par éditeur",
        {"min_col": 8, "min_row": 1, "max_row": 7, "max_col": 8},
        {"min_col": 7, "min_row": 2, "max_row": 7},
        "S1",
        source_sheet=calc,
    )
    create_pie_chart(
        tdb2, "Nombre de jeux vendus par éditeur",
        {"min_col": 9, "min_row": 1, "max_row": 7, "max_col": 9},
        {"min_col": 7, "min_row": 2, "max_row": 7},
        "S16",
        source_sheet=calc,
    )
