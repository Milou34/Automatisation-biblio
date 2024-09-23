from src.utils.utils import extract_info


def process_esp_d(esp_row, ws, current_row):
    """
    Traite une ligne d'espèce déterminante et ajoute les informations à la feuille Excel.

    Args:
        esp_row (ET.Element): L'élément XML représentant une ligne d'espèce.
        ws (openpyxl.worksheet.worksheet.Worksheet): La feuille Excel dans laquelle écrire les données.
        current_row (int): Le numéro de la ligne courante dans la feuille Excel.

    Returns:
        int: Le numéro de la ligne suivante après avoir écrit les données.
    """

    # Définir les chemins des balises XML à extraire
    tag_paths = [
        ".//GROUPE",            # Groupe (str)
        ".//CD_NOM",            # Code espèce (int)
        ".//NOM_COMPLET",       # Nom scientifique (str)
        ".//NOM_VERN",          # Nom vernaculaire (str)
        ".//STATUT_BIO_ESP",    # Statut biologique (str)
        ".//AUTEUR",            # Source (str)
        ".//CD_ABOND",          # Degré d'abondance (str)
        ".//NB_I_ABOND",        # Effectif inférieur estimé (int)
        ".//NB_S_ABOND",        # Effectif supérieur estimé (int)
        ".//AN_I_OBS",          # Année d'observation initiale (str)
        ".//AN_S_OBS",          # Année d'observation finale (str)
    ]

    # Utiliser extract_info pour extraire les valeurs
    extracted_values = extract_info(esp_row, tag_paths)

    # Traiter le statut biologique
    if extracted_values[4] == "R":
        extracted_values[4] = "Reproduction certaine ou probable"
    elif extracted_values[4] == "RI":
        extracted_values[4] = "Reproduction indéterminée"
    else:
        extracted_values[4] = "-"

    # Traiter le degré d'abondance
    if extracted_values[6] == "A":
        extracted_values[6] = "Fort"
    elif extracted_values[6] == "B":
        extracted_values[6] = "Moyen"
    elif extracted_values[6] == "C":
        extracted_values[6] = "Faible"
    else:
        extracted_values[6] = "-"

    # Combiner la période d'observation
    observation_I_text = str(extracted_values[9])
    observation_S_text = str(extracted_values[10])
    if observation_I_text == "-":
        observation_I_text = ""
    elif observation_S_text == "-":
        observation_S_text = ""
    observation = (observation_I_text + "-" + observation_S_text)
    if observation == "--":
        extracted_values[9] = "-"
    elif len(observation)<6:
        extracted_values[9] = (observation_I_text + observation_S_text)
    else:
        extracted_values[9] = observation
    extracted_values.pop(10)

    # Ajouter les données à la feuille Excel
    ws.append(extracted_values)
    
    # Mettre à jour current_row après avoir ajouté la ligne
    current_row += 1

    # Retourner la nouvelle valeur de current_row
    return current_row
