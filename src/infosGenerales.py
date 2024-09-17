import xml.etree.ElementTree as ET
from openpyxl import Workbook

def extract_general_info(root, tag_paths):
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
        # Ajoute la valeur de l'élément ou une chaîne vide si l'élément est None
        values.append(element.text if element is not None and element.text is not None else '')
    return values

def create_table(ws, title, headers, start_row):
    """Crée un tableau avec un titre et des en-têtes dans la feuille de calcul."""
    ws.append([title])
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=len(headers))
    ws.append(headers)
    print(f"{start_row}")
    return start_row + 2  # Retourne la ligne suivante après le tableau
    
