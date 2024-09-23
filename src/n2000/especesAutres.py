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
    # Traiter les valeurs extraites
    groupe = extracted_values[0] if extracted_values[0] else ""
    nom = extracted_values[1] if extracted_values[1] else ""
    min = int(extracted_values[2]) if extracted_values[2] else ""
    max = int(extracted_values[3]) if extracted_values[3] else ""
    unit = extracted_values[4] if extracted_values[4] else ""
    cat = extracted_values[5] if extracted_values[5] else ""
    
    annexe_IV = extracted_values[6]
    if annexe_IV == 'true':
        annexe_IV_text = "X"
    else:
        annexe_IV_text = ""
        
    annexe_V = extracted_values[7]
    if annexe_V == 'true':
        annexe_V_text = "X"
    else:
        annexe_V_text = ""
    
    a = extracted_values[8]
    if a == 'true':
        a_text = "X"
    else:
        a_text = ""
    
    b = extracted_values[9]
    if b == 'true':
        b_text = "X"
    else:
        b_text = ""
        
    c = extracted_values[10]
    if c == 'true':
        c_text = "X"
    else:
        c_text = ""    
    
    d = extracted_values[11]
    if d == 'true':
        d_text = "X"
    else:
        d_text = ""    
    
    especes_autres_values = [groupe, 
                             nom, 
                             min, 
                             max, 
                             unit, 
                             cat, 
                             annexe_IV_text, 
                             annexe_V_text, 
                             a_text, 
                             b_text, 
                             c_text, 
                             d_text]
    
    ws.append(especes_autres_values)
    current_row += 1
    
    return current_row