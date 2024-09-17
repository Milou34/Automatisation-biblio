import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Font

# Dossier contenant les fichiers Excel
input_folder = r'C:\Users\MarylouBERTIN\OneDrive - Grive Environnement\Documents\Projets\Référente\106 - Féricy\Bibliographie\ZNIEFF 1'
output_file = r'C:\Users\MarylouBERTIN\OneDrive - Grive Environnement\Bureau\Final.xlsx'

# Les noms des tableaux à extraire
tableaux_recherches = ["Habitats déterminants", "Espèces déterminantes", "Espèces à statut réglementé"]

# Initialisation des listes pour stocker les données
result_tables = []
zone_names = []

# Parcourir tous les fichiers Excel dans le dossier
for filename in os.listdir(input_folder):
    if filename.endswith('.xlsx'):
        print(f"Chargement des exels sources : {filename}")
        # Charger chaque fichier Excel
        file_path = os.path.join(input_folder, filename)
        xls = pd.ExcelFile(file_path)
        sheet_page = 0
        # Parcourir toutes les feuilles du fichier
        for sheet_name in xls.sheet_names:
            sheet_page += 1
            sheet_data = xls.parse(sheet_name)

            # Chercher les tableaux à partir du nom de la colonne ou d'autres critères
            for tableau in tableaux_recherches:
                if any(sheet_data.columns.str.contains(tableau, case=False, na=False)) or \
                sheet_data.map(lambda x: tableau.lower() in str(x).lower()).any().any():
                    # Ajouter le nom de la zone et le tableau aux résultats
                    if sheet_page == 1:
                        zone_name = sheet_data.iloc[0, 0]  # Lire le nom de la zone (première cellule)
                        print(f"Zone : {zone_name}")
                        zone_names.append(zone_name)
                    result_tables.append(sheet_data)
                    break

# Créer un nouveau fichier Excel
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    wb = writer.book
    ws = wb.create_sheet(title='Résumé')

    # Ajouter les noms des zones et les données dans le fichier Excel
    row_start = 1
    for zone_name, df in zip(zone_names, result_tables):
        # Ajouter le nom de la zone en gras
        ws.cell(row=row_start, column=1, value=zone_name).font = Font(bold=True)
        row_start += 1

        # Ajouter les données du tableau
        for r_idx, row in df.iterrows():
            for c_idx, value in enumerate(row):
                ws.cell(row=row_start, column=c_idx+1, value=value)
            row_start += 1
        
        # Ajouter une ligne vide entre les tableaux
        row_start += 1

print(f"Les tableaux ont été combinés et enregistrés dans {output_file}.")
