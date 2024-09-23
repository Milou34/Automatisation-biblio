from openpyxl.styles import Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter
import psutil
import os
import re

INT_PATTERN = r"^-?\d+$"
FLOAT_PATTERN = r"^-?\d+\,\d+$"
FLOAT_PATTERN2 = r"^-?\,\d+$"

def extract_info(root, tag_paths):
    """Extrait les valeurs des balises spécifiées depuis le XML et retourne une liste des valeurs en string.

    Args:
        root (ET.Element): La racine du document XML.
        tag_paths (list of str): Liste des chemins de balises à extraire.

    Returns:
        list of str: Liste des valeurs extraites sous forme de chaînes de caractères.
    """
    values = []
    
    for path in tag_paths:
        element = root.find(path)
        element = element.text if element is not None and element.text is not None else "-"
        if re.match(INT_PATTERN, element): 
            # Ajoute la valeur de l'élément ou une chaîne vide si l'élément est None
            values.append(int(element))
        elif re.match(FLOAT_PATTERN, element) or re.match(FLOAT_PATTERN2, element):
            element = float((element.replace(",", ".")))
            values.append(element)
        else:
            values.append(element)
            
    return values


def create_table(ws, title, headers, start_row):
    """Crée un tableau avec un titre et des en-têtes stylisés dans la feuille de calcul."""
    # Ajouter le titre du tableau et le mettre en gras
    ws.append([title])
    title_cell = ws.cell(row=start_row, column=1)
    title_cell.font = Font(bold=True, size=16)

    # Fusionner les cellules du titre sur la largeur des en-têtes
    ws.merge_cells(
        start_row=start_row, start_column=1, end_row=start_row, end_column=max(2, len(headers))
    )
    title_cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical='center')

    # Ajouter les en-têtes et les mettre en gras et centrés
    header_row = start_row + 1
    ws.append(headers)

    for col_num, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_num)
        cell.font = Font(bold=True)  # Mettre en gras
        cell.alignment = Alignment(horizontal="center", vertical='center')  # Centrer horizontalement

    # Retourner la ligne suivante après le tableau
    return start_row + 2


def merge_groups(ws, start_row, end_row, merge_column, check_column):
    """
    Fusionne les cellules d'une colonne spécifiée si elles contiennent des valeurs identiques
    dans une colonne de vérification spécifiée, sur des lignes consécutives.

    Args:
        ws (openpyxl.worksheet.worksheet.Worksheet): La feuille Excel dans laquelle effectuer la fusion.
        start_row (int): Le numéro de la ligne de début.
        end_row (int): Le numéro de la ligne de fin.
        merge_column (str): La lettre de la colonne dans laquelle fusionner les cellules.
        check_column (str): La lettre de la colonne dans laquelle vérifier les valeurs identiques.
    """
    merge_start_row = start_row
    previous_value = ws[f"{check_column}{start_row}"].value

    for row in range(start_row + 1, end_row + 1):
        current_value = ws[f"{check_column}{row}"].value

        if current_value != previous_value:
            if row - 1 > merge_start_row:
                # Fusionner les cellules de la colonne `merge_column` de la ligne merge_start_row à row-1
                ws.merge_cells(
                    f"{merge_column}{merge_start_row}:{merge_column}{row - 1}"
                )
            # Réinitialiser la valeur de départ pour la prochaine fusion
            merge_start_row = row
            previous_value = current_value

    # Fusionner les dernières cellules si nécessaire
    if end_row > merge_start_row:
        ws.merge_cells(f"{merge_column}{merge_start_row}:{merge_column}{end_row}")

    # Optionnel: Ajuster les alignements des cellules fusionnées
    for row in range(start_row, end_row + 1):
        cell = ws[f"{merge_column}{row}"]
        cell.alignment = cell.alignment.copy(horizontal="center")


def adjust_columns(wb):
    """
    Ajuste la largeur de chaque colonne en fonction de la longueur du texte le plus long dans chaque colonne
    et active le retour à la ligne automatique dans toutes les cellules pour toutes les feuilles du classeur.

    :param wb: Le classeur Excel (Workbook)
    """
    for ws in wb.worksheets:
        max_lengths = {}

        # Trouver la longueur maximale du texte pour chaque colonne
        for row in ws.iter_rows():
            for cell in row:
                col_letter = get_column_letter(
                    cell.column
                )  # Utiliser get_column_letter pour obtenir la lettre de la colonne
                if cell.value is not None:
                    cell_length = len(str(cell.value))
                else:
                    cell_length = 0
                if (
                    col_letter not in max_lengths
                    or cell_length > max_lengths[col_letter]
                ):
                    max_lengths[col_letter] = cell_length

        # Ajuster la largeur des colonnes en fonction des longueurs maximales trouvées
        for col_letter, length in max_lengths.items():
            # Utiliser un facteur d'ajustement pour éviter les colonnes trop larges
            if length > 45:
                adjusted_width = 45
            else:
                adjusted_width = length
            ws.column_dimensions[col_letter].width = max(
                10, adjusted_width
            )  # Largeur minimale de 10 pour éviter trop de réduction

        # Activer le retour à la ligne automatique et centrer horizontalement dans toutes les cellules
        for row in ws.iter_rows():
            for cell in row:
                # Vérifier si la cellule n'est pas en gras
                if not (cell.font.bold):
                    cell.alignment = Alignment(wrap_text=True, vertical='center')


def close_excel_if_open(file_path):
    """
    Vérifie si un fichier Excel est ouvert en recherchant les processus liés à Excel.

    Args:
        file_path (str): Le chemin du fichier à vérifier.

    Returns:
        bool: True si le fichier est ouvert, sinon False.
    """
    file_name = os.path.basename(file_path)

    for proc in psutil.process_iter(["pid", "name"]):
        try:
            if (
                proc.info["name"].lower() in ["excel.exe", "excel"]
                or "EXCEL" in proc.info["name"].upper()
            ):
                for open_file in proc.open_files():
                    if file_name in open_file.path:
                        print(
                            f"Le fichier {file_name} est ouvert. Tentative de fermeture..."
                        )
                        proc.terminate()
                        proc.wait()  # Assure que le processus est bien terminé
                        print(f"Fermeture du fichier {file_name} réussie.")
                        return

        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue


def extract_tables(wb):
    """
    Renvoie une liste de listes contenant les cellules qui constituent chaque tableau
    dans une feuille Excel, séparés par des lignes vides.
    
    :param ws: La feuille de calcul (Worksheet)
    :return: Une liste de listes de cellules pour chaque tableau.
    """
    all_tables = []

    for ws in wb:
        current_table = []
        # Itérer sur les lignes de la feuille
        for row in ws.iter_rows():
            # Vérifier si la ligne est vide
            if all(cell.value is None for cell in row):
                # Si on a une table en cours, l'ajouter
                if current_table:
                    all_tables.append(current_table)
                    current_table = []  # Réinitialiser pour le prochain tableau
            else:
                # Ajouter la ligne non vide à la table en cours
                current_table.append([cell.value for cell in row])

        # Ajouter le dernier tableau s'il y en a un
        if current_table:
            all_tables.append(current_table)
    return all_tables
                    

def apply_borders_to_tables(ws, tables):
    """
    Applique des bordures intérieures et extérieures à chaque tableau dans la liste de tableaux.
    
    :param ws: La feuille de calcul (Worksheet)
    :param tables: Liste de listes contenant les cellules de chaque tableau.
    """
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))

    for table in tables:
        # Déterminer les limites du tableau
        min_row = min(cell.row for row in table for cell in row if cell is not None)
        max_row = max(cell.row for row in table for cell in row if cell is not None)
        min_col = min(cell.column for row in table for cell in row if cell is not None)
        max_col = max(cell.column for row in table for cell in row if cell is not None)

        # Appliquer les bordures extérieures
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                cell = ws.cell(row=row, column=col)
                if row == min_row:  # Bordure supérieure
                    cell.border = Border(top=thin_border.top)
                if row == max_row:  # Bordure inférieure
                    cell.border = Border(bottom=thin_border.bottom)
                if col == min_col:  # Bordure gauche
                    cell.border = Border(left=thin_border.left)
                if col == max_col:  # Bordure droite
                    cell.border = Border(right=thin_border.right)

        # Appliquer les bordures intérieures
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                cell = ws.cell(row=row, column=col)
                if col < max_col:  # Bordure droite interne
                    cell.border = Border(right=thin_border.right)
                if row < max_row:  # Bordure inférieure interne
                    cell.border = Border(bottom=thin_border.bottom)