def process_habitats(typo_info_row, ws, current_row):
    """Traite les habitats à partir des balises TYPO_INFO_ROW et renvoie les valeurs des colonnes."""
    source = typo_info_row.find('.//AUTEUR')
    source_text = source.text if source is not None and source.text is not None else ''
    observation_I = typo_info_row.find('.//AN_I_OBS') 
    observation_I_text = observation_I.text if observation_I is not None and observation_I.text is not None else ''
    observation_S = typo_info_row.find('.//AN_S_OBS') 
    observation_S_text = observation_S.text if observation_S is not None and observation_S.text is not None else ''
    observation = observation_I_text + " - " + observation_S_text

    habitats_values = ["", "", "", source_text, observation]

    for typo_row in typo_info_row.findall('.//TYPO_ROW'):
        lb_typo = typo_row.find('.//LB_TYPO').text
        lb_hab = typo_row.find('.//LB_HAB').text

        # Déterminer la colonne selon la valeur de LB_TYPO
        if lb_typo == "EUNIS 2012":
            habitats_values[0] = lb_hab
        elif lb_typo == "CORINE biotopes":
            habitats_values[1] = lb_hab
        elif lb_typo == "Habitats d'intérêt communautaire":
            habitats_values[2] = lb_hab
            
    # Ajouter les données à la feuille Excel
    ws.append(habitats_values)
    
    # Mettre à jour current_row après avoir ajouté la ligne
    current_row += 1
    
    # Retourner la nouvelle valeur de current_row
    return current_row