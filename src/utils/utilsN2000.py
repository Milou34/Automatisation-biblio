from openpyxl.styles import Font, Alignment

def create_table_hab_n2000(ws, current_row):
    """
    Crée un en-tête de tableau avec un titre en gras et centré avant, et des colonnes fusionnées.
    La fusion commence à partir de la colonne A.
    
    :param ws: La feuille de calcul (Worksheet)
    :param current_row: La ligne actuelle où l'en-tête doit être inséré
    :param main_title: Le titre principal du tableau
    :param column_names_A_F: Une liste de noms pour les colonnes A à F
    :param title_G_J: Le titre pour les colonnes G à J (fusion horizontale)
    :param sub_titles_H_J: Une liste de sous-titres pour les colonnes H à J
    """
    
    # Ajouter le titre principal et le fusionner de A à J (colonnes 1 à 10)
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
    title_cell = ws.cell(row=current_row, column=1)
    title_cell.value = "Types d’habitats présents sur le site et évaluations"
    title_cell.font = Font(bold=True, size=16)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Passer à la ligne suivante pour l'en-tête des colonnes
    current_row += 1
    
    # Fusionner les cellules de A à F sur la première ligne (fusion horizontale)
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
    title_cell = ws.cell(row=current_row, column=1)
    title_cell.value = "Types d’habitats inscrits à l’annexe I"
    title_cell.font = Font(bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Définir les titres pour chaque colonne de B à E
    titles = ["Habitat", "PF", "Superficie (Ha)", "Superficie (% de couverture)", "Grottes (nombre)", "Qualité des données"]

    # Fusionner les cellules de A à F sur les deux lignes suivantes (fusion verticale) et utiliser les noms fournis
    for col in range(1, 7):  # Colonnes de A (1) à F (6)
        ws.merge_cells(start_row=current_row + 1, start_column=col, end_row=current_row + 2, end_column=col)
        cell = ws.cell(row=current_row + 1, column=col)
        cell.value = titles[col - 1]  # Utiliser les noms de la liste
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Fusionner les cellules de G à J sur la première ligne (fusion horizontale)
    ws.merge_cells(start_row=current_row, start_column=7, end_row=current_row, end_column=10)
    title_cell_G_J = ws.cell(row=current_row, column=7)
    title_cell_G_J.value = "Évaluation du site"
    title_cell_G_J.font = Font(bold=True)
    title_cell_G_J.alignment = Alignment(horizontal="center", vertical="center")

    # Fusionner les cellules de la colonne G sur les deux lignes suivantes (fusion verticale)
    ws.merge_cells(start_row=current_row + 1, start_column=7, end_row=current_row + 2, end_column=7)
    cell_G = ws.cell(row=current_row + 1, column=7)
    cell_G.value = "Représentativité (A|B|C|D)"
    cell_G.font = Font(bold=True)
    cell_G.alignment = Alignment(horizontal="center", vertical="center")

    # Fusionner les cellules des colonnes H à J sur la deuxième ligne (fusion horizontale)
    ws.merge_cells(start_row=current_row + 1, start_column=8, end_row=current_row + 1, end_column=10)
    title_H_J = ws.cell(row=current_row + 1, column=8)
    title_H_J.value = "A|B|C"
    title_H_J.font = Font(bold=True)
    title_H_J.alignment = Alignment(horizontal="center", vertical="center")

    titles = ["Superficie relative", "Conservation", "Évaluation globale"]
    # Ajouter des titres dans les cellules H2 à J2 (troisième ligne)
    for col in range(8, 11):  # Colonnes H (8) à J (10)
        ws.cell(row=current_row + 2, column=col, value=titles[col - 8]).font = Font(bold=True)
        ws.cell(row=current_row + 2, column=col).alignment = Alignment(horizontal="center", vertical="center")

    # Retourner la nouvelle ligne pour la suite
    return current_row + 3


def create_table_especes_inscrites(ws, current_row):
    """Crée un en-tête de tableau avec un titre en gras et centré, et des colonnes fusionnées.
    
    :param ws: La feuille de calcul (Worksheet)
    :param current_row: La ligne actuelle où l'en-tête doit être inséré
    :param main_title: Le titre principal du tableau
    :param titles_col_A_C: Une liste de titres pour les colonnes A à C
    :param title_D_I: Le titre pour les colonnes D à I (fusion horizontale)
    :param title_G_J: Le titre pour les colonnes G à J (fusion verticale)
    :param title_K_M: Le titre pour les colonnes K à M (fusion horizontale)
    """
    
    # Ajouter le titre principal et le fusionner de A à J (colonnes 1 à 10)
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=10)
    title_cell = ws.cell(row=current_row, column=1)
    title_cell.value = "Espèces inscrites à l’annexe II de la directive 92/43/CEE et évaluation"
    title_cell.font = Font(bold=True, size=16)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    # Passer à la ligne suivante pour l'en-tête des colonnes
    current_row += 1
    
    # Fusionner les cellules de A à C sur la première ligne (fusion horizontale)
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
    title_cell = ws.cell(row=current_row, column=1)
    title_cell.value = "Espèce"
    title_cell.font = Font(bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    titles = ["Groupe", "Code", "Nom scientifique"]
    # Fusionner les cellules de A à C sur les deux lignes suivantes (fusion verticale) et utiliser les noms fournis
    for col in range(1, 4):  # Colonnes A (1) à C (3)
        ws.merge_cells(start_row=current_row + 1, start_column=col, end_row=current_row + 2, end_column=col)
        cell = ws.cell(row=current_row + 1, column=col)
        cell.value = titles[col - 1]  # Utiliser les noms de la liste
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Fusionner les cellules de D à I sur la première ligne (fusion horizontale)
    ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row, end_column=9)
    title_cell_D_I = ws.cell(row=current_row, column=4)
    title_cell_D_I.value = "Population présente sur le site"
    title_cell_D_I.font = Font(bold=True)
    title_cell_D_I.alignment = Alignment(horizontal="center", vertical="center")

    # Fusionner la colonne D verticalement
    ws.merge_cells(start_row=current_row + 1, start_column=4, end_row=current_row + 2, end_column=4)
    cell_D = ws.cell(row=current_row + 1, column=4)
    cell_D.value = "Type"
    cell_D.font = Font(bold=True)
    cell_D.alignment = Alignment(horizontal="center", vertical="center")

    # Fusionner les colonnes E à F sur la deuxième ligne (fusion horizontale)
    ws.merge_cells(start_row=current_row + 1, start_column=5, end_row=current_row + 1, end_column=6)
    title_E_F = ws.cell(row=current_row + 1, column=5)
    title_E_F.value = "Taille"
    title_E_F.font = Font(bold=True)
    title_E_F.alignment = Alignment(horizontal="center", vertical="center")

    # Titres dans la troisième ligne des colonnes E et F
    sub_titles_E_F = ["Min", "Max"]
    for col in range(5, 7):  # Colonnes E (5) et F (6)
        ws.cell(row=current_row + 2, column=col, value=sub_titles_E_F[col - 5]).font = Font(bold=True)
        ws.cell(row=current_row + 2, column=col).alignment = Alignment(horizontal="center")

    # Fusion verticale dans les colonnes G à J
    sub_titles_G_J = ["Unité", "Catégorie (C|R|V|P)", "Qualité des données", "Population (A|B|C|D)"]
    for col in range(7, 11):  # Colonnes G (7) à J (10)
        ws.merge_cells(start_row=current_row + 1, start_column=col, end_row=current_row + 2, end_column=col)
        cell_G = ws.cell(row=current_row + 1, column=col)
        cell_G.value = sub_titles_G_J[col - 7]  # Utiliser les noms de la liste
        cell_G.font = Font(bold=True)
        cell_G.alignment = Alignment(horizontal="center", vertical="center")

    # Fusion horizontale des colonnes J à M
    ws.merge_cells(start_row=current_row, start_column=10, end_row=current_row, end_column=13)
    title_K_M = ws.cell(row=current_row, column=10)
    title_K_M.value = "Évaluation du site"
    title_K_M.font = Font(bold=True)
    title_K_M.alignment = Alignment(horizontal="center", vertical="center")

    # Fusion horizontale sur la deuxième ligne des colonnes K à M
    ws.merge_cells(start_row=current_row + 1, start_column=11, end_row=current_row + 1, end_column=13)
    title_K_M_2 = ws.cell(row=current_row + 1, column=11)
    title_K_M_2.value = "A|B|C"
    title_K_M_2.font = Font(bold=True)
    title_K_M_2.alignment = Alignment(horizontal="center", vertical="center")

    # Titres dans la troisième ligne des colonnes K à M
    sub_titles_K_M = ["Conservation", "Isolement", "Évaluation globale"]
    for col in range(11, 14):  # Colonnes K (11) à M (13)
        ws.cell(row=current_row + 2, column=col, value=sub_titles_K_M[col - 11]).font = Font(bold=True)
        ws.cell(row=current_row + 2, column=col).alignment = Alignment(horizontal="center")

    # Retourner la nouvelle ligne pour la suite
    return current_row + 3