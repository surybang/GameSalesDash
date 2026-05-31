"""Utilitaires de mise en forme : bordures et styles partagés."""

from openpyxl.styles import Border, Side
from openpyxl.worksheet.worksheet import Worksheet


def apply_border_style(
    sheet: Worksheet,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
    border_row: int,
) -> None:
    """Applique une bordure inférieure fine sur une ligne spécifique d'une plage.

    Args:
        sheet: Feuille Excel cible.
        start_row: Première ligne de la plage.
        end_row: Dernière ligne de la plage.
        start_col: Première colonne de la plage.
        end_col: Dernière colonne de la plage.
        border_row: Ligne sur laquelle appliquer la bordure inférieure.
            Les autres lignes de la plage voient leur bordure supprimée.
    """
    thin = Side(border_style="thin", color="4D4D4D")

    for row in sheet.iter_rows(
        min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col
    ):
        for cell in row:
            cell.border = Border(bottom=thin) if cell.row == border_row else None
