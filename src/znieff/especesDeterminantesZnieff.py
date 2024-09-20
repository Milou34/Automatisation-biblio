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
        ".//GROUPE",  # Groupe
        ".//CD_NOM",  # Code espèce
        ".//NOM_COMPLET",  # Nom scientifique
        ".//NOM_VERN",  # Nom vernaculaire
        ".//STATUT_BIO_ESP",  # Statut biologique
        ".//AUTEUR",  # Source
        ".//CD_ABOND",  # Degré d'abondance
        ".//NB_I_ABOND",  # Effectif inférieur estimé
        ".//NB_S_ABOND",  # Effectif supérieur estimé
        ".//AN_I_OBS",  # Année d'observation initiale
        ".//AN_S_OBS",  # Année d'observation finale
    ]

    # Utiliser extract_info pour extraire les valeurs
    extracted_values = extract_info(esp_row, tag_paths)

    # Traiter les valeurs extraites
    groupe = extracted_values[0]

    # Gestion du code espèce (convertir en int si possible)
    code_esp = extracted_values[1]
    code_esp = int(code_esp) if code_esp else ""

    nom_esp = extracted_values[2]
    nom_vern = extracted_values[3]

    # Traiter le statut biologique
    statut_bio_esp = extracted_values[4]
    if statut_bio_esp == "R":
        statut_bio_esp_txt = "Reproduction certaine ou probable"
    elif statut_bio_esp == "RI":
        statut_bio_esp_txt = "Reproduction indéterminée"
    else:
        statut_bio_esp_txt = ""

    source = extracted_values[5]

    # Traiter le degré d'abondance
    deg_abd = extracted_values[6]
    if deg_abd == "A":
        deg_abd_txt = "Fort"
    elif deg_abd == "B":
        deg_abd_txt = "Moyen"
    elif deg_abd == "C":
        deg_abd_txt = "Faible"
    else:
        deg_abd_txt = ""

    # Gestion de l'effectif inférieur
    eff_I = extracted_values[7]
    eff_I = int(eff_I) if eff_I else ""

    # Gestion de l'effectif supérieur
    eff_S = extracted_values[8]
    eff_S = int(eff_S) if eff_S else ""

    # Combiner la période d'observation
    observation_I_text = extracted_values[9]
    observation_S_text = extracted_values[10]
    observation = (
        observation_I_text + " - " + observation_S_text
        if observation_I_text or observation_S_text
        else ""
    )

    # Créer la liste des valeurs à ajouter à la feuille Excel
    espece_values = [
        groupe,
        code_esp,
        nom_esp,
        nom_vern,
        statut_bio_esp_txt,
        source,
        deg_abd_txt,
        eff_I,
        eff_S,
        observation,
    ]

    # Ajouter les données à la feuille Excel
    ws.append(espece_values)

    # Mettre à jour current_row après avoir ajouté la ligne
    current_row += 1

    # Retourner la nouvelle valeur de current_row
    return current_row
