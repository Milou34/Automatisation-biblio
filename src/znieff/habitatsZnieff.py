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
        ".//AUTEUR",  # Auteur
        ".//PC_TYPO",  # Surface
        ".//AN_I_OBS",  # Début période observation
        ".//AN_S_OBS",  # Fin période observation
    ]

    # Utiliser extract_info pour extraire les données
    extracted_values = extract_info(typo_info_row, tag_paths)

    # Traiter les valeurs extraites
    source_text = extracted_values[0]

    # Gestion de la surface (convertir en float si disponible)
    surface_text = extracted_values[1]
    surface_float = float(surface_text.replace(",", ".")) if surface_text else ""

    # Combiner la période d'observation
    observation_I_text = extracted_values[2]
    observation_S_text = extracted_values[3]
    observation = observation_I_text + " - " + observation_S_text

    # Initialiser la liste des valeurs d'habitats
    habitats_values = ["", "", "", source_text, surface_float, observation]

    # Parcourir les balises TYPO_ROW pour extraire les informations supplémentaires
    for typo_row in typo_info_row.findall(".//TYPO_ROW"):
        lb_typo = typo_row.find(".//LB_TYPO").text

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
            habitats_values[0] = lb_hab
        elif lb_typo == "CORINE biotopes":
            habitats_values[1] = lb_hab
        elif lb_typo == "Habitats d'intérêt communautaire (HIC)":
            habitats_values[2] = lb_hab

    # Ajouter les données à la feuille Excel
    ws.append(habitats_values)

    # Mettre à jour current_row après avoir ajouté la ligne
    current_row += 1

    # Retourner la nouvelle valeur de current_row
    return current_row