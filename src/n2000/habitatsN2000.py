from src.utils.utils import extract_info


def process_habitats_n2000(habit1_row, ws, current_row):
    """
    Traite les habitats à partir des balises HABIT1_ROW et renvoie les valeurs des colonnes.

    Args:
        habit1_row (ET.Element): L'élément XML représentant une ligne d'information d'habitat.
        ws (openpyxl.worksheet.worksheet.Worksheet): La feuille Excel dans laquelle écrire les données.
        current_row (int): Le numéro de la ligne courante dans la feuille Excel.

    Returns:
        int: Le numéro de la ligne suivante après avoir écrit les données.
    """

    # Définir les chemins des balises à extraire pour la source, la surface et la période d'observation
    tag_paths = [
        ".//CD_UE",             # Code habitat (int)
        ".//LB_HABDH_FR",       # Nom habitat (str)
        ".//PF",                # Forme prioritaire de l'habitat (boolean)
        ".//AREA",              # Surface (float)
        ".//COVER",             # % surface (float)
        ".//CAVE",              # Nb de grottes (int)
        ".//QUALITY",           # Qualité des observations (str)     
        ".//REPRESENT",         # Représentativité (str)
        ".//REL_SURF",          # Surface relative (str)
        ".//CONSERVE",          # Conservation (str)
        ".//GLOBAL",            # Evaluation globale (str)
    ]

    # Utiliser extract_info pour extraire les données
    extracted_values = extract_info(habit1_row, tag_paths)
    # Traiter les valeurs extraites
    cd_hab = extracted_values[0]
    nom_hab = extracted_values[1]
    habitat = cd_hab + " " + nom_hab

    pf = extracted_values[2]
    if pf == 'true':
        pf_text = "X"
    else:
        pf_text = ""
        
    surf = extracted_values[3]
    surf = float(surf.replace(",", ".")) if surf else ""
    
    surf_p = extracted_values[4]
    surf_p = float(surf_p.replace(",", ".")) if surf_p else ""
    
    grottes = extracted_values[5]
    grottes = int(grottes) if grottes else ""
    
    qualit = extracted_values[6]
    repr = extracted_values[7]
    surf_r = extracted_values[8]
    cons = extracted_values[9]
    e_globale = extracted_values[10]
    
    habitats_values = [habitat, 
                       pf_text,
                       surf,
                       surf_p,
                       grottes,
                       qualit,
                       repr,
                       surf_r,
                       cons,
                       e_globale
                       ]

    ws.append(habitats_values)
    current_row += 1
    
    return current_row

    


    
    
    
    

    

