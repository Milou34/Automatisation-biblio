import os
from src.telechargement.utilsTelechargementXML import input_codes_n2000, input_codes_znieff, telechargement_xml


def input_telechargement_xml():
    # Demander à l'utilisateur le dossier de destination
    destination_folder = input(
        "Veuillez entrer le chemin du dossier de destination, il sera créé s'il n'existe pas : "
    )

    # Créer le dossier de destination s'il n'existe pas
    os.makedirs(destination_folder, exist_ok=True)

    valid_znieff_codes = input_codes_znieff()
    valid_n2000_codes = input_codes_n2000()

    # Télécharger les fichiers XML pour les codes ZNIEFF validés
    print("\nTéléchargement des fichiers ZNIEFF...")
    for code in valid_znieff_codes:
        telechargement_xml(code, destination_folder, "ZNIEFF")

    # Télécharger les fichiers XML pour les codes N2000 validés
    print("\nTéléchargement des fichiers N2000...")
    for code in valid_n2000_codes:
        telechargement_xml(code, destination_folder, "N2000")

    # Message final pour confirmer la fin des téléchargements
    print(
        f"\nTéléchargements terminés. Les fichiers ont été sauvegardés dans le dossier : {destination_folder}"
    )
    
    return destination_folder