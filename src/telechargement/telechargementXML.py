import os
from src.telechargement.utilsTelechargementXML import input_codes_n2000, input_codes_znieff, telechargement_xml


def input_telechargement_xml():

    while True:
        # Demander à l'utilisateur le dossier de destination
        destination_folder = input(
            "Veuillez entrer le chemin du dossier de destination, il sera créé s'il n'existe pas : "
        ).strip()

        if not destination_folder:
            print("Erreur : Le chemin du dossier ne peut pas être vide. Veuillez réessayer.")
            continue  # Redemande le chemin

        try:
            # Créer le dossier de destination s'il n'existe pas
            os.makedirs(destination_folder, exist_ok=True)

            # Vérifier si le chemin est valide et accessible
            if os.path.isdir(destination_folder):
                print(f"Chemin validé : {destination_folder}")
                break  # Sort de la boucle si le chemin est valide
            else:
                print("Erreur : Le chemin spécifié n'est pas un dossier valide. Veuillez réessayer.")

        except Exception as e:
            print(f"Une erreur est survenue : {e}. Veuillez entrer un chemin valide.")
            continue  # Redemande le chemin en cas d'erreur

    # Une fois que le chemin est validé, continuer le reste du processus
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
