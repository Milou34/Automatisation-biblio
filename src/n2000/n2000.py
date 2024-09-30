from src.n2000.legendes import legende_especes_autres, legende_especes_inscrites, legende_habitats
from src.n2000.especesAutres import process_especes_autres
from src.n2000.especesInscrites import process_especes_inscrites
from src.n2000.habitatsN2000 import process_habitats_n2000
from src.utils.utilsN2000 import create_table_hab_n2000, create_table_especes_inscrites, create_table_especes_autres
from src.utils.utils import extract_info, create_table


def process_n2000(ws, root, current_row, non_formated_cells):
    # Ajouter le premier tableau pour les informations générales
    current_row = create_table(
        ws,
        "Informations générales",
        ["Type de zone", "ID national", "Nom zone", "Surface totale (Ha)"],
        current_row,
    )
    
    # Liste des chemins de balises pour extraire les informations générales
    tag_paths = ["TYPE", "SITECODE", "SITE_NAME", "AREA"]

    # Extraction des données depuis le fichier XML
    extracted_value = extract_info(root, tag_paths)
    # Convertir les valeurs numériques pour l'affichage correct

    type_zone = extracted_value[0]
    if type_zone == 'A':
        extracted_value[0] = "p-SIC"
    elif type_zone == 'B':
        extracted_value[0] = "SIC"
    else:
        extracted_value[0] = "ZPS"
         
    ws.append(extracted_value)
    ws.append([])
    current_row += 2  # Mettre à jour la ligne après avoir ajouté le premier tableau
    
    
    # Ajouter le deuxième tableau pour les habitats
    if root.find(".//HABIT1_ROW") is not None:
        current_row = create_table_hab_n2000(ws, current_row)

        # Parcourir les balises HABIT1_ROW pour récupérer les habitats
        for habit1_row in root.findall(".//HABIT1_ROW"):
            current_row = process_habitats_n2000(habit1_row, ws, current_row)
        ws.append([])
        current_row += 1
        current_row, non_formated_cells = legende_habitats(ws, current_row, non_formated_cells)
        ws.append([])
        current_row += 1  # Saute une ligne et met à jour la ligne après avoir ajouté le deuxième tableau
    else:
        current_row = create_table(
            ws,
            "Types d’habitats inscrits à l’annexe I",
            [
                "Non renseigné",
            ],
            current_row,
        )
        
        # Ajouter le troisième tableau pour les espèces inscrites
    if root.find(".//SPECIES_ROW") is not None:
        current_row = create_table_especes_inscrites(ws, current_row)

        # # Parcourir les balises SPECIES_ROW pour récupérer les espèces
        for species_row in root.findall(".//SPECIES_ROW"):
            current_row = process_especes_inscrites(species_row, ws, current_row)
        ws.append([])
        current_row += 1
        current_row, non_formated_cells = legende_especes_inscrites(ws, current_row, non_formated_cells)
        ws.append([])
        current_row += 1  # Saute une ligne et met à jour la ligne après avoir ajouté le troisième tableau
    else:
        current_row = create_table(
            ws,
            "Espèces inscrites à l’annexe II de la directive 92/43/CEE et évaluation",
            [
                "Non renseigné",
            ],
            current_row,
        )
        
    # Ajouter le quatrième tableau pour les espèces autres
    if root.find(".//SPECIES_OTHER_ROW") is not None:
        current_row = create_table_especes_autres(ws, current_row)

        # # Parcourir les balises SPECIES_OTHER_ROW pour récupérer les espèces autres
        for species_other_row in root.findall(".//SPECIES_OTHER_ROW"):
            current_row = process_especes_autres(species_other_row, ws, current_row)
        ws.append([])
        current_row += 1
        current_row, non_formated_cells = legende_especes_autres(ws, current_row, non_formated_cells)
        ws.append([])
        current_row += 1  # Saute une ligne et met à jour la ligne après avoir ajouté le quatrième tableau
    else:
        current_row = create_table(
            ws,
            "Autres espèces importantes de faune et de flore",
            [
                "Non renseigné",
            ],
            current_row,
        )
    
    return current_row, non_formated_cells
