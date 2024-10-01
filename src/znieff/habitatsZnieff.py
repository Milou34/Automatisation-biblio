from src.utils.utils import extract_info


def process_habitats(typo_info_row, ws, current_row):
    """
    Traite les habitats à partir des balises TYPO_INFO_ROW et renvoie les valeurs des colonnes.

    Args:
        typo_info_row (ET.Element): L'élément XML représentant une ligne d'information d'habitat.
        ws (openpyxl.worksheet.worksheet.Worksheet): La feuille Excel dans laquelle écrire les données.
        current_row (int): Le numéro de la ligne courante dans la feuille Excel.

    Returns:
        int: Le numéro de la ligne suivante après avoir écrit les données.
    """

    # Définir les chemins des balises à extraire pour la source, la surface et la période d'observation
    tag_paths = [
        ".//AUTEUR",    # Auteur (str)
        ".//PC_TYPO",   # Surface (float)
        ".//AN_I_OBS",  # Début période observation (str)
        ".//AN_S_OBS",  # Fin période observation (str)
    ]

    # Utiliser extract_info pour extraire les données
    extracted_values = extract_info(typo_info_row, tag_paths)

    # Traiter les valeurs extraites
    source_text = extracted_values[0]

    # Gestion de la surface (convertir en float si disponible)
    surface = extracted_values[1]

    # Combiner la période d'observation
    observation_I_text = str(extracted_values[2])
    observation_S_text = str(extracted_values[3])
    if observation_I_text == "-":
        observation_I_text = ""
    elif observation_S_text == "-":
        observation_S_text = ""
    observation = (observation_I_text + "-" + observation_S_text)
    if observation == "--":
        observation = "-"
    elif len(observation)<6:
        observation = (observation_I_text + observation_S_text)
    else:
        observation = observation
        
    # Initialiser la liste des valeurs d'habitats
    habitats_values = ["", "", "", source_text, surface, observation]

    # Parcourir les balises TYPO_ROW pour extraire les informations supplémentaires
    for typo_row in typo_info_row.findall(".//TYPO_ROW"):
        lb_typo_element = typo_row.find("LB_TYPO")
        lb_typo = lb_typo_element.text if lb_typo_element is not None else ""
        # Récupérer tout le texte de la balise LB_HAB sans les balises de mise en forme comme <em>
        lb_hab_element = typo_row.find(".//LB_HAB")
        lb_hab = (
            "".join(lb_hab_element.itertext()) if lb_hab_element is not None else ""
        )

        # Extraire LB_CODE et combiner avec LB_HAB
        lb_code = (
            typo_row.find(".//LB_CODE").text
            if typo_row.find(".//LB_CODE") is not None
            else ""
        )
        lb_hab = lb_code + " " + lb_hab

        # Déterminer la colonne selon la valeur de LB_TYPO
        if lb_typo == "EUNIS 2012":
            habitats_values[0] = lb_hab  # Colonne 1 pour EUNIS
        elif lb_typo == "CORINE biotopes":
            habitats_values[1] = lb_hab  # Colonne 2 pour CORINE
        elif lb_typo == "Habitats d'intérêt communautaire (HIC)":
            habitats_values[2] = lb_hab  # Colonne 3 pour HIC
            
    # Ajouter les données à la feuille Excel
    ws.append(habitats_values)

    # Mettre à jour current_row après avoir ajouté la ligne
    current_row += 1

    # Retourner la nouvelle valeur de current_row
    return current_row
