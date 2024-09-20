from src.n2000.habitatsN2000 import process_habitats_n2000
from src.utils.utilsN2000 import create_table_hab_n2000, create_table_especes_inscrites
from src.utils.utils import extract_info, create_table


def process_n2000(ws, root, current_row):
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
    
    extracted_value[3] = float(
        extracted_value[3].replace(",", ".")
    )  # Conversion de 'Surface'
    
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
        
        # Ajouter le troisième tableau pour les espèces
    if root.find(".//SPECIES_ROW") is not None:
        current_row = create_table_especes_inscrites(ws, current_row)

        # # Parcourir les balises HABIT1_ROW pour récupérer les habitats
        # for habit1_row in root.findall(".//HABIT1_ROW"):
        #     current_row = process_habitats_n2000(habit1_row, ws, current_row)
        # ws.append([])
        # current_row += 1  # Saute une ligne et met à jour la ligne après avoir ajouté le deuxième tableau
    else:
        current_row = create_table(
            ws,
            "Types d’habitats inscrits à l’annexe I",
            [
                "Non renseigné",
            ],
            current_row,
        )
        
    
    
    return current_row
