import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Font
import fnmatch
import os

def xml_to_excel(folder_source):

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
    
    # Affiche les fichiers XML trouvés
    for fichier in fichiers_xml:
        # Parse le fichier XML
        tree = ET.parse(fichier)
        root = tree.getroot()

        # Ajout du titre "Informations générales"
        ws.append(["Informations générales"])  # Première ligne avec le titre
        title_cell = ws['A1']
        title_cell.font = Font(bold=True)  # Mettre le titre en gras

        # Fusionner les cellules pour le titre
        ws.merge_cells('A1:D1')

        # Ajouter les en-têtes du premier tableau
        headers = ['ID national', 'Nom ZNIEFF', 'Type ZNIEFF', 'Surface totale ZNIEFF']
        ws.append(headers)

        id_national = int(root.find('NM_SFFZN').text)
        nom_znieff = root.find('LB_ZN').text
        type_znieff = int(root.find('TY_ZONE').text)
        surface_znieff = float(root.find('SU_ZN').text.replace(',', '.'))
        headers_value = [id_national, nom_znieff, type_znieff, surface_znieff]
        ws.append(headers_value)
        ws.append([])

        if type_znieff == 1:
            ws.title = "ZNIEFF 1"
        elif type_znieff == 2:
            ws.title = "ZNIEFF 2"

    # Sauvegarder le fichier Excel
    wb.save(folder_source + '\Récap.xlsx')
    print(f"Fichier Excel '{folder_source}' généré avec succès !")

# Demander les chemins des fichiers à l'utilisateur
folder_source = input("Entrez le chemin du dossier des XML sources, le excel sera créé dans ce dossier : ")

# Lancer la conversion XML vers Excel
xml_to_excel(folder_source)
