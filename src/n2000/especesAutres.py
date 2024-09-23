from src.utils.utils import extract_info


def process_especes_autres(species_other_row, ws, current_row):
    """
    Traite les espèces à partir des balises SPECIES_OTHER_ROW et renvoie les valeurs des colonnes.

    Args:
        species_other_row (ET.Element): L'élément XML représentant une ligne d'information d'espèces.
        ws (openpyxl.worksheet.worksheet.Worksheet): La feuille Excel dans laquelle écrire les données.
        current_row (int): Le numéro de la ligne courante dans la feuille Excel.

    Returns:
        int: Le numéro de la ligne suivante après avoir écrit les données.
    """

    # Définir les chemins des balises à extraire pour la source, la surface et la période d'observation
    tag_paths = [
        ".//TAXGROUP",          # Groupe (str)
        ".//LB_NOM",            # Nom scientifique (str)
        ".//SIZE_MIN",          # Taille min population sur site (int)
        ".//SIZE_MAX",          # Taille min population sur site (int)
        ".//UNIT",              # Unité de comptage (str)
        ".//CAT_POP",           # Catégories du point de vue de l’abondance (str)
        ".//ANNEX_IV",          # Espèce inscrite en annexe IV (boolean)
        ".//ANNEX_V",           # Espèce inscrite en annexe V (boolean)
        ".//A",                 # Espèce inscrite en liste rouge nationale (boolean)
        ".//B",                 # Espèce inscrite en espèce endémique (boolean)
        ".//C",                 # Espèce inscrite en conventions internationales (boolean)
        ".//D",                 # Espèce inscrite pour d'autres raisons (boolean) 
    ]
    

    # Utiliser extract_info pour extraire les données
    extracted_values = extract_info(species_other_row, tag_paths)
    
    # Traite pour l'Annexe IV
    if extracted_values[6] == 'true':
        extracted_values[6] = "X"
    else:
        extracted_values[6] = "-"
    
    # Traite pour l'Annexe V
    if extracted_values[7] == 'true':
        extracted_values[7] = "X"
    else:
        extracted_values[7] = "-"
    
    # Traite pour la liste rouge nationale
    if extracted_values[8] == 'true':
        extracted_values[8] = "X"
    else:
        extracted_values[8] = "-"
    
    # Traite pour les espèces endémiques
    if extracted_values[9] == 'true':
        extracted_values[9] = "X"
    else:
        extracted_values[9] = "-"
        
    # Traite pour les conventions internationales
    if extracted_values[10] == 'true':
        extracted_values[10] = "X"
    else:
        extracted_values[10] = "-"    
    
    # Traite pour les autres raisons
    if extracted_values[11] == 'true':
        extracted_values[11] = "X"
    else:
        extracted_values[11] = "-"    

    ws.append(extracted_values)
    current_row += 1
    
    return current_row