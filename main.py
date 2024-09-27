import xml.etree.ElementTree as ET
from openpyxl import Workbook
import fnmatch
import os
from src.n2000.n2000 import process_n2000
from src.utils.utils import adjust_columns, apply_borders, close_excel_if_open
from src.znieff.znieff import process_znieff


def main(folder_source):
    non_formated_cells = []
    output_file = os.path.join(folder_source, "Récap.xlsx")

    # Vérifier si le fichier Excel est ouvert et le fermer si c'est le cas
    close_excel_if_open(output_file)

    # Créer un nouveau fichier Excel
    wb = Workbook()

    # Créer des feuilles distinctes pour ZNIEFF 1, ZNIEFF 2 et N2000
    ws_znieff1 = wb.active
    ws_znieff1.title = "ZNIEFF 1"
    ws_znieff2 = wb.create_sheet(title="ZNIEFF 2")
    ws_n2000 = wb.create_sheet(title="N2000")

    # Initialiser la ligne actuelle pour chaque type de zone
    current_row_znieff1 = 1
    current_row_znieff2 = 1
    current_row_n2000 = 1

    # Parcours des fichiers dans le dossier
    for chemin, sous_dossiers, fichiers in os.walk(folder_source):
        # Parcourir les fichiers XML trouvés
        for fichier in fichiers:
            if fnmatch.fnmatch(fichier, "*.xml"):
                # Parse le fichier XML
                tree = ET.parse(os.path.join(chemin, fichier))
                root = tree.getroot()
                file_path = os.path.join(chemin, fichier)

                # Sélectionner la feuille et la ligne en fonction du type de ZNIEFF
                if root.tag == "ZNIEFF":
                    # Déterminer le type de ZNIEFF à partir de la balise TY_ZONE
                    type_znieff = int(
                        root.find("TY_ZONE").text
                    )  # On suppose que TY_ZONE existe et est valide
                    code_zone = int(root.find("NM_SFFZN").text)
                    new_file_name = f"znieff{type_znieff}_{code_zone}.xml"
                    new_file_path = os.path.join(chemin, new_file_name)
                    # Renommer le fichier
                    os.rename(file_path, new_file_path)
                    if type_znieff == 1:
                        ws = ws_znieff1
                        current_row = process_znieff(ws, root, current_row_znieff1)
                    else:
                        ws = ws_znieff2
                        current_row = process_znieff(ws, root, current_row_znieff2)
                else:
                    ws = ws_n2000
                    current_row, non_formated_cells = process_n2000(ws, root, current_row_n2000, non_formated_cells)

                # Ajoute 2 lignes vides entre chaque fichier XML
                ws.append([])
                ws.append([])
                # Mettre à jour la ligne après avoir ajouté une ligne vide
                current_row += 2

                # Mettre à jour la ligne courante pour le type de ZNIEFF
                if root.tag == "ZNIEFF":
                    # Déterminer le type de ZNIEFF à partir de la balise TY_ZONE
                    type_znieff = int(
                        root.find("TY_ZONE").text
                    )  # On suppose que TY_ZONE existe et est valide
                    if type_znieff == 1:
                        current_row_znieff1 = current_row
                    else:
                        current_row_znieff2 = current_row
                else:
                    current_row_n2000 = current_row

    # Sauvegarder le fichier Excel
    adjust_columns(wb, non_formated_cells)
    apply_borders(wb)
    wb.save(os.path.join(folder_source, "Récap.xlsx"))

    # Ouvrir le fichier Excel généré
    print(f"Ouverture de {output_file} généré avec succès !")
    os.startfile(output_file)


# Demander les chemins des fichiers à l'utilisateur
# folder_source = input("Entrez le chemin du dossier des XML sources, le Excel sera créé dans ce dossier : ")
# TODO : rajouter demande du nom de l'Excel
# Créer le dossier de destination s'il n'existe pas
# os.makedirs(folder_source, exist_ok=True)

# Lancer la conversion XML vers Excel
main("C:\\Users\\MarylouBERTIN\\Documents\\Test")