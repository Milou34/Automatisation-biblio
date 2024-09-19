def process_esp_D(esp_row, ws, current_row):

    groupe = esp_row.find(".//GROUPE")
    groupe = groupe.text if groupe is not None and groupe.text is not None else ""

    code_esp = esp_row.find(".//CD_NOM")
    code_esp = (
        code_esp.text if code_esp is not None and code_esp.text is not None else ""
    )
    if code_esp:
        code_esp = int(code_esp)
    else:
        code_esp = ""

    nom_esp = esp_row.find(".//NOM_COMPLET")
    nom_esp = nom_esp.text if nom_esp is not None and nom_esp.text is not None else ""

    nom_vern = esp_row.find(".//NOM_VERN")
    nom_vern = (
        nom_vern.text if nom_vern is not None and nom_vern.text is not None else ""
    )

    statut_bio_esp = esp_row.find(".//STATUT_BIO_ESP")
    statut_bio_esp = (
        statut_bio_esp.text
        if statut_bio_esp is not None and statut_bio_esp.text is not None
        else ""
    )

    if statut_bio_esp == "R":
        statut_bio_esp_txt = "Reproduction certaine ou probable"
    elif statut_bio_esp == "RI":
        statut_bio_esp_txt = "Reproduction indéterminée"
    else:
        statut_bio_esp_txt = ""

    source = esp_row.find(".//AUTEUR")
    source = source.text if source is not None and source.text is not None else ""

    deg_abd = esp_row.find(".//CD_ABOND")
    deg_abd = deg_abd.text if deg_abd is not None and deg_abd.text is not None else ""

    if deg_abd == "A":
        deg_abd_txt = "Fort"
    elif deg_abd == "B":
        deg_abd_txt = "Moyen"
    elif deg_abd == "C":
        deg_abd_txt = "Faible"
    else:
        deg_abd_txt = ""

    eff_I = esp_row.find(".//NB_I_ABOND")
    eff_I = eff_I.text if eff_I is not None and eff_I.text is not None else ""
    if eff_I:
        eff_I = int(eff_I)
    else:
        eff_I = ""

    eff_S = esp_row.find(".//NB_S_ABOND")
    eff_S = eff_S.text if eff_S is not None and eff_S.text is not None else ""
    if eff_S:
        eff_S = int(eff_S)
    else:
        eff_S = ""

    observation_I = esp_row.find(".//AN_I_OBS")
    observation_I_text = (
        observation_I.text
        if observation_I is not None and observation_I.text is not None
        else ""
    )
    observation_S = esp_row.find(".//AN_S_OBS")
    observation_S_text = (
        observation_S.text
        if observation_S is not None and observation_S.text is not None
        else ""
    )
    observation = observation_I_text + " - " + observation_S_text

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


def merge_identical_groups(ws, start_row, end_row):
    """
    Fusionne les cellules de la colonne 'Groupe' si elles contiennent des valeurs identiques dans des lignes consécutives.
    """
    col_letter = "A"  # La colonne "Groupe" est la colonne A
    merge_start_row = start_row
    previous_value = ws[f"{col_letter}{start_row}"].value

    for row in range(start_row + 1, end_row + 1):
        current_value = ws[f"{col_letter}{row}"].value

        if current_value != previous_value:
            if row - 1 > merge_start_row:
                # Fusionner les cellules de la colonne A de la ligne merge_start_row à row-1
                ws.merge_cells(f"{col_letter}{merge_start_row}:{col_letter}{row - 1}")
            # Réinitialiser la valeur de départ
            merge_start_row = row
            previous_value = current_value

    # Fusionner les dernières cellules si nécessaire
    if end_row > merge_start_row:
        ws.merge_cells(f"{col_letter}{merge_start_row}:{col_letter}{end_row}")
