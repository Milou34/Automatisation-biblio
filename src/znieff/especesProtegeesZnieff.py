from optparse import Values
from src.utils.utils import extract_info


def process_esp_p(esp_row, ws, current_row, root):
    """
    Traite une ligne d'espèce déterminante et ajoute les informations à la feuille Excel.

    Args:
        esp_row (ET.Element): L'élément XML représentant une ligne d'espèce.
        ws (openpyxl.worksheet.worksheet.Worksheet): La feuille Excel dans laquelle écrire les données.
        current_row (int): Le numéro de la ligne courante dans la feuille Excel.

    Returns:
        int: Le numéro de la ligne suivante après avoir écrit les données.
    """

    # Définir les chemins des balises XML à extraire
    tag_paths = [
        ".//GROUPE",        # Groupe (str)
        ".//CD_NOM",        # Code espèce (int)
        ".//NOM_COMPLET",   # Nom scientifique (str)
        ".//FG_ESP",        # Source (str)
    ]

    # Utiliser extract_info pour extraire les valeurs
    extracted_values = extract_info(esp_row, tag_paths)

    # Traiter les valeurs extraites
    groupe = extracted_values[0]

    # Gestion du code espèce (convertir en int si possible)
    code_esp = extracted_values[1]

    nom_esp = extracted_values[2]

    # Traiter le statut biologique
    statut = extracted_values[3]
    if statut == "D":
        statut_txt = "Déterminante"
    elif statut == "E":
        statut_txt = "Enjeux"
    else:
        statut_txt = "Autre"

    # Extraction de la citation
    citation = esp_row.find(".//SHORT_CITATION")
    citation = (
        citation.text if citation is not None and citation.text is not None else ""
    )

    # Extraction de la citation
    url = esp_row.find(".//URL")
    url = url.text if url is not None and url.text is not None else ""

    # Insérer le lien hypertexte avec la citation comme texte cliquable
    if url:
        # Utiliser la fonction HYPERLINK pour faire de la citation un lien cliquable
        citation_text = f'=HYPERLINK("{url}", "{citation}")'

    # Créer la liste des valeurs à ajouter à la feuille Excel
    espece_values = [
        groupe,
        code_esp,
        nom_esp,
        statut_txt,
        citation_text,
    ]

    # Ajouter les données à la feuille Excel
    ws.append(espece_values)

    # if url:
    #     # Insérer un lien cliquable dans la cellule (colonne 5, la citation)
    #     ws.cell(row=current_row, column=5).value = f'=HYPERLINK("{url}", "lien")'

    # Mettre à jour current_row après avoir ajouté la ligne
    current_row += 1

    # Retourner la nouvelle valeur de current_row
    return current_row
