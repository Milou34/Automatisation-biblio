from src.znieff.especesDeterminantesZnieff import process_esp_d
from src.znieff.especesProtegeesZnieff import process_esp_p
from src.znieff.habitatsZnieff import process_habitats
from src.utils.utils import extract_info, merge_groups
from src.utils.utils import create_table


def process_znieff(ws, root, current_row):
    # Ajouter le premier tableau pour les informations générales
    current_row = create_table(
        ws,
        "Informations générales",
        ["ID national", "Nom ZNIEFF", "Surface totale ZNIEFF (Ha)"],
        current_row,
    )

    # Liste des chemins de balises pour extraire les informations générales
    tag_paths = ["NM_SFFZN", "LB_ZN", "SU_ZN"]

    # Extraction des données depuis le fichier XML
    infos_value = extract_info(root, tag_paths)
    # Convertir les valeurs numériques pour l'affichage correct
    ws.append(infos_value)
    ws.append([])
    current_row += 2  # Mettre à jour la ligne après avoir ajouté le premier tableau

    # Ajouter le deuxième tableau pour les habitats
    if root.find(".//TYPO_INFO_ROW") is not None:
        current_row = create_table(
            ws,
            "Habitats déterminants",
            [
                "EUNIS",
                "CORINE biotopes",
                "Habitats d’intérêt communautaire",
                "Source",
                "Surface en %",
                "Observation",
            ],
            current_row,
        )

        # Parcourir les balises TYPO_INFO_ROW pour récupérer les habitats
        for typo_info_row in root.findall(".//TYPO_INFO_ROW"):
            fg_typo = typo_info_row.find("FG_TYPO").text

            # On ne garde que les habitats avec FG_TYPO = "D"
            if fg_typo == "D":
                current_row = process_habitats(typo_info_row, ws, current_row)
        ws.append([])
        current_row += 1  # Saute une ligne et met à jour la ligne après avoir ajouté le deuxième tableau
    else:
        current_row = create_table(
            ws,
            "Habitats déterminants",
            [
                "Non renseigné",
            ],
            current_row,
        )
        ws.append([])
        current_row += 1

    # Ajoute le 3ème tableau
    if root.find(".//ESPECE_ROW") is not None:
        current_row = create_table(
            ws,
            "Espèces déterminantes",
            [
                "Groupe",
                "Code espèce",
                "Nom scientifique",
                "Nom vernaculaire",
                "Statut(s) biologique(s)",
                "Sources",
                "Degré d'abondance",
                "Effectif inférieur estimé",
                "Effectif supérieur estimé",
                "Année d'observation",
            ],
            current_row,
        )

        start_row_for_merge = current_row  # Enregistre la ligne de début pour la fusion
        for espece_row in root.findall(".//ESPECE_ROW"):
            fg_esp = espece_row.find("FG_ESP").text

            # On ne garde que les habitats avec FG_TYPO = "D"
            if fg_esp == "D":
                current_row = process_esp_d(espece_row, ws, current_row)
            else:
                pass
        # Fusionner les cellules de la colonne 'Groupe' après l'ajout des lignes
        merge_groups(ws, start_row_for_merge, current_row - 1, "A", "A")
        ws.append([])
        current_row += 1  # Saute une ligne et met à jour la ligne après avoir ajouté le troisième tableau
    else:
        current_row = create_table(
            ws,
            "Espèces déterminantes",
            [
                "Non renseigné",
            ],
            current_row,
        )
        ws.append([])
        current_row += 1

    # Ajoute le 4ème tableau
    if root.find(".//ESPECE_PROT_ROW") is not None:
        current_row = create_table(
            ws,
            "Espèces à statut réglementé",
            [
                "Groupe",
                "Code espèce",
                "Nom scientifique",
                "Statut de déterminance",
                "Réglementation",
            ],
            current_row,
        )

        start_row_for_merge = current_row  # Enregistre la ligne de début pour la fusion

        for espece_row in root.findall(".//ESPECE_PROT_ROW"):
            current_row = process_esp_p(espece_row, ws, current_row, root)
            
        # Fusionner les cellules pour les colonnes 'Code espèce', 'Nom scientifique', et 'Statut de déterminance' 
        # après l'ajout des lignes pour les espèces protégées
        merge_groups(ws, start_row_for_merge, current_row - 1, "A", "A")  # Fusionner les cellules de 'Groupe'
        merge_groups(ws, start_row_for_merge, current_row - 1, "D", "B")  # Fusionner les cellules de 'Statut de déterminance'
        merge_groups(ws, start_row_for_merge, current_row - 1, "C", "B")  # Fusionner les cellules de 'Nom scientifique'
        merge_groups(ws, start_row_for_merge, current_row - 1, "B", "B")  # Fusionner les cellules de 'Code espèce'
    else:
        current_row = create_table(
            ws,
            "Espèces à statut réglementé",
            [
                "Non renseigné",
            ],
            current_row,
        )
        ws.append([])
        current_row += 1

    ws.append([])
    current_row += 1
    
    return current_row
