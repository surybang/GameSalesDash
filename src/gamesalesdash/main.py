"""Point d'entrée CLI : génération du reporting Excel."""

from loguru import logger

from gamesalesdash.config import OUTPUT_PATH
from gamesalesdash.dashboards import build_workbook
from gamesalesdash.data import clean_data, compute_list_lengths, load_data


def main() -> None:
    """Génère le classeur Excel de reporting à partir des données MinIO.

    Orchestre le pipeline complet : chargement, nettoyage, calcul des tailles
    de listes, construction du classeur puis sauvegarde dans OUTPUT_PATH.
    """
    logger.info("Chargement des données...")
    df_raw = load_data()

    logger.info("Nettoyage des données ({} lignes brutes)...", len(df_raw))
    df = clean_data(df_raw)
    logger.info("{} lignes conservées après nettoyage.", len(df))

    len_dict = compute_list_lengths(df)

    logger.info("Construction du classeur Excel...")
    wb = build_workbook(df, len_dict)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    logger.info("Sauvegarde dans {}...", OUTPUT_PATH)
    wb.save(OUTPUT_PATH)
    logger.success("Terminé.")


if __name__ == "__main__":
    main()
