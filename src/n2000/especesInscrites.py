from src.utils.utils import extract_info


def process_especes_inscrites(species_row, ws, current_row):
    """
    Traite les habitats à partir des balises SPECIES_ROW et renvoie les valeurs des colonnes.

    Args:
        species_row (ET.Element): L'élément XML représentant une ligne d'information d'habitat.
        ws (openpyxl.worksheet.worksheet.Worksheet): La feuille Excel dans laquelle écrire les données.
        current_row (int): Le numéro de la ligne courante dans la feuille Excel.

    Returns:
        int: Le numéro de la ligne suivante après avoir écrit les données.
    """

    # Définir les chemins des balises à extraire pour la source, la surface et la période d'observation
    tag_paths = [
        ".//CD_UE",  # Code habitat
        ".//LB_HABDH_FR",  # Nom habitat
        ".//PF",  
        ".//AREA",      # Surface
        ".//COVER",     # % surface
        ".//CAVE",       # Nb de grottes
        ".//QUALITY",   # Qualité des observations        
        ".//REPRESENT", # Représentativité
        ".//REL_SURF",   # Surface relative
        ".//CONSERVE",  # Conservation
        ".//GLOBAL",    # Evaluation globale
    ]

    # Utiliser extract_info pour extraire les données
    extracted_values = extract_info(species_row, tag_paths)
    # Traiter les valeurs extraites
    cd_hab = extracted_values[0]
    nom_hab = extracted_values[1]