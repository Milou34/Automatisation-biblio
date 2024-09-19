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
    
    # Parcours des fichiers dans le dossier
    for chemin, sous_dossiers, fichiers in os.walk(folder_source):
        # Parcourir les fichiers XML trouvés
        for fichier in fichiers:
            if fnmatch.fnmatch(fichier, "*.xml"):
                # Parse le fichier XML
                tree = ET.parse(os.path.join(chemin, fichier))
                root = tree.getroot()
                file_path = os.path.join(chemin, fichier)

                # Déterminer le type de ZNIEFF à partir de la balise TY_ZONE
                type_znieff = int(
                    root.find("TY_ZONE").text
                )  # On suppose que TY_ZONE existe et est valide
                
                code_zone = int(
                    root.find("NM_SFFZN").text
                )     
                
                # Sélectionner la feuille et la ligne en fonction du type de ZNIEFF
                if type_znieff == 1:
                    ws = ws_znieff1
                    current_row = current_row_znieff1
                    new_file_name = f"znieff1_{code_zone}.xml"
                    new_file_path = os.path.join(chemin, new_file_name)
                    # Renommer le fichier
                    os.rename(file_path, new_file_path)
                elif type_znieff == 2:
                    ws = ws_znieff2
                    current_row = current_row_znieff2
                    new_file_name = f"znieff2_{code_zone}.xml"
                    new_file_path = os.path.join(chemin, new_file_name)
                    # Renommer le fichier
                    os.rename(file_path, new_file_path)
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
                if root.find('.//TYPO_INFO_ROW') is not None: 
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
                    current_row = create_table(
                        ws,
                        "Habitats déterminants",
                        [
                            "Non renseigné",
                        ],
                        current_row,
                    )
                        
                # Ajoute le 3ème tableau        
                if root.find('.//ESPECE_ROW') is not None: 
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
                
                # Ajoute le 4ème tableau
                if root.find('.//ESPECE_PROT_ROW') is not None: 
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
                    current_row = create_table(
                        ws,
                        "Espèces à statut réglementé",
                        [
                            "Non renseigné",
                        ],
                        current_row,
                    )
                    
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
