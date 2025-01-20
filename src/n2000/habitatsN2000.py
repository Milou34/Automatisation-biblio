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
    
    # Extraire le code habitat et le nom de l'habitat
    cd_hab = str(extracted_values[0])
    nom_hab = extracted_values[1]

    # Traite pour PF
    if extracted_values[2] == 'true':
        extracted_values[2] = "X"
    else:
        extracted_values[2] = "-"

    # Préparer les données avec le code et le nom dans des cellules séparées
    extracted_values[0] = cd_hab
    extracted_values[1] = nom_hab

    # Ajouter les données à la feuille Excel
    ws.append(extracted_values)
    current_row += 1
    
    return current_row
