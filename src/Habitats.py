def process_habitats(typo_info_row, ws, current_row):
    
    """Traite les habitats à partir des balises TYPO_INFO_ROW et renvoie les valeurs des colonnes."""
    ## L'auteur
    source = typo_info_row.find('.//AUTEUR')
    source_text = source.text if source is not None and source.text is not None else ''
    
    ## La surface
    surface = typo_info_row.find('.//PC_TYPO')
    surface_text = surface.text if surface is not None and surface.text is not None else ''
    if surface_text :
        surface_float = float(surface_text.replace(',', '.'))
    else :
        surface_float = ''
        
    ## La période d'observation
    observation_I = typo_info_row.find('.//AN_I_OBS') 
    observation_I_text = observation_I.text if observation_I is not None and observation_I.text is not None else ''
    observation_S = typo_info_row.find('.//AN_S_OBS') 
    observation_S_text = observation_S.text if observation_S is not None and observation_S.text is not None else ''
    observation = observation_I_text + " - " + observation_S_text

    habitats_values = ["", "", "", source_text, surface_float, observation]

    for typo_row in typo_info_row.findall('.//TYPO_ROW'):
        lb_typo = typo_row.find('.//LB_TYPO').text

        # Récupérer tout le texte de la balise LB_HAB sans les balises de mise en forme comme <em>
        lb_hab_element = typo_row.find('.//LB_HAB')
        lb_hab = ''.join(lb_hab_element.itertext()) if lb_hab_element is not None else ''
        lb_code = typo_row.find('.//LB_CODE')
        lb_code = lb_code.text if lb_code is not None else ''
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