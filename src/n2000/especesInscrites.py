from src.utils.utils import extract_info


def process_especes_inscrites(species_row, ws, current_row):
    """
    Traite les espèces à partir des balises SPECIES_ROW et renvoie les valeurs des colonnes.

    Args:
        species_row (ET.Element): L'élément XML représentant une ligne d'information d'espèces.
        ws (openpyxl.worksheet.worksheet.Worksheet): La feuille Excel dans laquelle écrire les données.
        current_row (int): Le numéro de la ligne courante dans la feuille Excel.

    Returns:
        int: Le numéro de la ligne suivante après avoir écrit les données.
    """

    # Définir les chemins des balises à extraire pour la source, la surface et la période d'observation
    tag_paths = [
        ".//TAXGROUP",          # Groupe (str)
        ".//CODE_N2000",        # Code espèce site 2000 (int)
        ".//NOM",               # Nom scientifique (str)
        ".//TYPE",              # Type de comportement sur site (str)
        ".//SIZE_MIN",          # Taille min population sur site (int)
        ".//SIZE_MAX",          # Taille min population sur site (int)
        ".//UNIT",              # Unité de comptage (str)
        ".//CAT_POP",           # Catégories du point de vue de l’abondance (str)
        ".//QUALITY",           # Qualité des données (str)
        ".//POPULATION",        # Population (str)
        ".//CONSERVE",          # Conservation (str)
        ".//ISOLATION",         # Isolation (str)
        ".//GLOBAL",            # Evaluation globale (str)
    ]

    # Utiliser extract_info pour extraire les données
    extracted_values = extract_info(species_row, tag_paths)
    
    ws.append(extracted_values)
    current_row += 1
    
    return current_row
    
    