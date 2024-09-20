from src.utils.utilsN2000 import create_table_infos


def process_n2000(ws, root, current_row):
    # Ajouter le premier tableau pour les informations générales
    current_row = create_table_infos(
        ws,
        current_row,
    )
    return current_row
