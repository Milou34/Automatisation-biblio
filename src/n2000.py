from src.utils import create_table


def process_n2000(ws, root, current_row):
    # Ajouter le premier tableau pour les informations générales
    current_row = create_table(
        ws,
        "Informations générales",
        [
            "Type de zone",
            "ID national",
            "Nom zone",
            "Surface totale ZNIEFF (Ha)",
            "Région biogéographique",
        ],
        current_row,
    )

    return current_row
