import xml.etree.ElementTree as ET
from openpyxl import Workbook
import fnmatch
import os
from src.especesDeterminantes import process_esp_d
from src.especesProtegees import process_esp_p
from src.utils import adjust_columns, close_excel_file, extract_info, create_table, is_excel_file_open, merge_groups
from src.habitats import process_habitats


def main(folder_source):

    output_file = os.path.join(folder_source, "Récap.xlsx")

    # Vérifier si le fichier Excel est ouvert et le fermer si c'est le cas
    if is_excel_file_open(output_file):
        print(f"Le fichier {output_file} est ouvert. Tentative de fermeture...")
        close_excel_file(output_file)
        print(f"Fermeture du fichier {output_file} réussie.")
    # Liste pour stocker les fichiers trouvés
    fichiers_xml = []

    # Parcours des fichiers dans le dossier
    for chemin, sous_dossiers, fichiers in os.walk(folder_source):
        for fichier in fichiers:
            if fnmatch.fnmatch(fichier, "*.xml"):
                fichiers_xml.append(os.path.join(chemin, fichier))

    # Créer un nouveau fichier Excel
    wb = Workbook()

    # Créer des feuilles distinctes pour ZNIEFF 1 et ZNIEFF 2
    ws_znieff1 = wb.active
    ws_znieff1.title = "ZNIEFF 1"
    ws_znieff2 = wb.create_sheet(title="ZNIEFF 2")

    # Liste des chemins de balises pour extraire les informations générales
    tag_paths = ["NM_SFFZN", "LB_ZN", "SU_ZN"]

    # Initialiser la ligne actuelle pour chaque type de ZNIEFF
    current_row_znieff1 = 1
    current_row_znieff2 = 1

    # Parcourir les fichiers XML trouvés
    for fichier in fichiers_xml:
        # Parse le fichier XML
        tree = ET.parse(fichier)
        root = tree.getroot()

        # Déterminer le type de ZNIEFF à partir de la balise TY_ZONE
        type_znieff = int(
            root.find("TY_ZONE").text
        )  # On suppose que TY_ZONE existe et est valide

        # Sélectionner la feuille et la ligne en fonction du type de ZNIEFF
        if type_znieff == 1:
            ws = ws_znieff1
            current_row = current_row_znieff1
        elif type_znieff == 2:
            ws = ws_znieff2
            current_row = current_row_znieff2
        else:
            # Si TY_ZONE n'est pas 1 ou 2, on passe à l'itération suivante
            continue

        # Ajouter le premier tableau pour les informations générales
        current_row = create_table(
            ws,
            "Informations générales",
            ["ID national", "Nom ZNIEFF", "Surface totale ZNIEFF (Ha)"],
            current_row,
        )

        # Extraction des données depuis le fichier XML
        infos_value = extract_info(root, tag_paths)
        # Convertir les valeurs numériques pour l'affichage correct
        infos_value[2] = float(
            infos_value[2].replace(",", ".")
        )  # Conversion de 'Surface totale ZNIEFF'
        infos_value[0] = int(infos_value[0])  # Conversion de 'ID national'
        ws.append(infos_value)
        ws.append([])
        current_row += 2  # Mettre à jour la ligne après avoir ajouté le premier tableau

        # Ajouter le deuxième tableau pour les habitats
        if root.find('.//TYPO_INFO_ROW'): 
            current_row = create_table(
                ws,
                "Habitats déterminants",
                [
                    "EUNIS",
                    "CORINE biotopes",
                    "Habitats d’intérêt communautaire",
                    "Source",
                    "Surface en Ha",
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
        else :
            pass
                
        # Ajoute le 3ème tableau        
        if root.find('.//ESPECE_ROW'): 
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
            pass
        
        # Ajoute le 4ème tableau
        if root.find('.//ESPECE_PROT_ROW'): 
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
            merge_groups(ws, start_row_for_merge, current_row - 1, "D", "B")  # Fusionner les cellules de 'Statut de déterminance'
            merge_groups(ws, start_row_for_merge, current_row - 1, "C", "B")  # Fusionner les cellules de 'Nom scientifique'
            merge_groups(ws, start_row_for_merge, current_row - 1, "B", "B")  # Fusionner les cellules de 'Code espèce'
            merge_groups(ws, start_row_for_merge, current_row - 1, "A", "A")  # Fusionner les cellules de 'Groupe'
        else:
            pass
            


        # Ajoute une ligne vide entre chaque fichier XML
        ws.append([])
        current_row += 1  # Mettre à jour la ligne après avoir ajouté une ligne vide

        # Mettre à jour la ligne courante pour le type de ZNIEFF
        if type_znieff == 1:
            current_row_znieff1 = current_row
        elif type_znieff == 2:
            current_row_znieff2 = current_row

    # Sauvegarder le fichier Excel
    adjust_columns(wb)
    wb.save(os.path.join(folder_source, "Récap.xlsx"))
    print(f"Fichier Excel '{folder_source}/Récap.xlsx' généré avec succès !")
    
    # Ouvrir le fichier Excel généré
    os.startfile(output_file)


# Demander les chemins des fichiers à l'utilisateur
# folder_source = input("Entrez le chemin du dossier des XML sources, le Excel sera créé dans ce dossier : ")

# Lancer la conversion XML vers Excel
main("C:\\Users\\MarylouBERTIN\\OneDrive - Grive Environnement\\Bureau\\Test")
