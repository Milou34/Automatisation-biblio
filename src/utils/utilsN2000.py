from openpyxl.styles import Font, Alignment


def create_table_infos(ws, start_row):
    """Crée un tableau avec un titre et des en-têtes stylisés dans la feuille de calcul."""
    # Ajouter le titre du tableau et le mettre en gras
    ws.append(["Infos générales"])
    title_cell = ws.cell(row=start_row, column=1)
    title_cell.font = Font(bold=True)

    # Fusionner les cellules du titre sur la largeur des en-têtes
    ws.merge_cells(
        start_row=start_row, start_column=1, end_row=start_row, end_column=7
    )
        
    # Définir les titres pour chaque colonne de B à E
    titles = ["Type de zone", "ID national", "Nom zone", "Surface totale"]

    # Ajouter les titres et fusionner les cellules
    for col in range(2, 6):  # De B (2) à E (5)
        cell = ws.cell(row=start_row + 1, column=col)
        cell.value = titles[col - 2]  # Remplir avec le titre approprié
        ws.merge_cells(start_row=start_row + 1, start_column=col, end_row=start_row + 2, end_column=col)  # Fusionner les cellules
    
        # Styliser les en-têtes
        cell.font = Font(bold=True)
        new_cell = ws.cell(start_row + 1, col)
        new_cell.alignment = Alignment(vertical='center', horizontal='centerContinuous')
        
    # Fusionner les cellules F et G sur la même ligne de départ
    ws.merge_cells(start_row=start_row + 1, start_column=6, end_row=start_row + 1, end_column=7)
    ws.cell(row=start_row + 1, column=6, value="Région biogéographique").font = Font(bold=True)  # Titre dans la cellule fusionnée
    ws.cell(row=start_row + 1, column=6).alignment = Alignment(horizontal='center', vertical='center')


    # Ajouter des titres dans F2 et G2
    ws.cell(row=start_row + 2, column=6, value="Nom").font = Font(bold=True)
    ws.cell(row=start_row + 2, column=7, value="Pourcentage").font = Font(bold=True)
    
    for col in range(6, 8):
        ws.cell(row=start_row + 2, column=col).alignment = Alignment(horizontal='center', vertical='center')

    return start_row + 3

    