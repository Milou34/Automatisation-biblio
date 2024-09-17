import xml.etree.ElementTree as ET
from openpyxl import Workbook
import fnmatch
import os
from src.infosGenerales import extract_general_info, create_table
from src.habitats import process_habitats

def main(folder_source):
    # Liste pour stocker les fichiers trouvés
    fichiers_xml = []

    # Parcours des fichiers dans le dossier
    for chemin, sous_dossiers, fichiers in os.walk(folder_source):
        for fichier in fichiers:
            if fnmatch.fnmatch(fichier, '*.xml'):
                fichiers_xml.append(os.path.join(chemin, fichier))

    # Créer un nouveau fichier Excel
    wb = Workbook()
    ws = wb.active
    
    # Liste des chemins de balises pour extraire les informations générales
    tag_paths = ['NM_SFFZN', 'LB_ZN', 'TY_ZONE', 'SU_ZN']

    # Déterminer la ligne actuelle pour créer le tableau
    current_row = 1

    # Parcourir les fichiers XML trouvés
    for fichier in fichiers_xml:
        # Parse le fichier XML
        tree = ET.parse(fichier)
        root = tree.getroot()

        # Ajouter le premier tableau
        current_row = create_table(ws, "Informations générales", 
                     ['ID national', 'Nom ZNIEFF', 'Type ZNIEFF', 'Surface totale ZNIEFF'],
                     current_row)
        
        # Extraction des données depuis le fichier XML
        infos_value = extract_general_info(root, tag_paths)
        # Convertir les valeurs numériques pour l'affichage correct
        infos_value[3] = float(infos_value[3].replace(',', '.'))  # Conversion de 'Surface totale ZNIEFF'
        infos_value[2] = int(infos_value[2])  # Conversion de 'Type ZNIEFF'
        infos_value[0] = int(infos_value[0])  # Conversion de 'ID national'
        ws.append(infos_value)
        ws.append([])
        current_row += 2  # Mettre à jour la ligne après avoir ajouté le premier tableau

        # Ajouter le deuxième tableau
        current_row = create_table(ws, "Habitats déterminants",
                     ['EUNIS', 'CORINE biotopes', 'Habitats d’intérêt communautaire', 'Source', 'Observation'],
                     current_row)

        # Parcourir les balises TYPO_INFO_ROW pour récupérer les habitats
        for typo_info_row in root.findall('.//TYPO_INFO_ROW'):
            fg_typo = typo_info_row.find('FG_TYPO').text
            
            # On ne garde que les habitats avec FG_TYPO = "D"
            if fg_typo == "D":
                current_row = process_habitats(typo_info_row, ws, current_row)

        # Ajoute une ligne vide entre chaque fichier XML
        ws.append([])
        current_row += 1  # Mettre à jour la ligne après avoir ajouté une ligne vide

        # Modifier le titre de la feuille en fonction du type ZNIEFF
        type_znieff = int(root.find('TY_ZONE').text)  # Assurez-vous que type_znieff est défini ici
        if type_znieff == 1:
            ws.title = "ZNIEFF 1"
        elif type_znieff == 2:
            ws.title = "ZNIEFF 2"

    # Sauvegarder le fichier Excel
    wb.save(os.path.join(folder_source, 'Récap.xlsx'))
    print(f"Fichier Excel '{folder_source}/Récap.xlsx' généré avec succès !")

# Demander les chemins des fichiers à l'utilisateur
# folder_source = input("Entrez le chemin du dossier des XML sources, le Excel sera créé dans ce dossier : ")


# Lancer la conversion XML vers Excel
main("C:\\Users\\MarylouBERTIN\\OneDrive - Grive Environnement\\Bureau\\Test 2")

# xml_to_excel(folder_source)
