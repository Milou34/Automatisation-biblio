import os
import requests
import re

# Fonction pour télécharger un XML depuis une URL
def download_xml(code, destination_folder, zone):
    # URL de base
    base_url_znieff = 'https://inpn.mnhn.fr/docs/ZNIEFF/znieffxml/'
    base_url_n2000 = 'https://inpn.mnhn.fr/docs/natura2000/fsdxml/'

    if zone == "ZNIEFF":
        base_url = base_url_znieff
    elif zone == "N2000":
        base_url = base_url_n2000
    else:
        print(f"Zone invalide !")
        return

    # Construire l'URL avec le code entré par l'utilisateur
    url = f'{base_url}{code}.xml'

    # Nom du fichier à sauvegarder dans le dossier de destination
    if base_url == base_url_znieff:
        output_file = os.path.join(destination_folder, f'znieff_{code}.xml')
    elif base_url == base_url_n2000:
        output_file = os.path.join(destination_folder, f'N2000_{code}.xml')

    try:
        # Télécharger le XML
        response = requests.get(url)

        # Vérifier si la requête est un succès
        if response.status_code == 200:
            # Sauvegarder le fichier XML localement
            with open(output_file, 'wb') as f:
                f.write(response.content)
            print(f"Le fichier {output_file} a été téléchargé avec succès.")
        else:
            print(f"Échec du téléchargement : le fichier {code}.xml n'existe pas.")
    except Exception as e:
        print(f"Erreur lors du téléchargement : {e}")

# Fonction pour valider un code ZNIEFF
def is_valid_znieff_code(code):
    return len(code) == 9 and code.isdigit()

# Fonction pour valider un code N2000
def is_valid_n2000_code(code):
    return re.match(r'^FR\d{7}$', code) is not None

# Demander à l'utilisateur le dossier de destination
destination_folder = input("Veuillez entrer le chemin du dossier de destination, il sera créé s'il n'existe pas : ")

# Créer le dossier de destination s'il n'existe pas
os.makedirs(destination_folder, exist_ok=True)

# Saisie et validation des codes ZNIEFF
while True:
    codes_znieff_input = input("Veuillez entrer les codes ZNIEFF, séparés par des virgules (ex : 740120259, 123456789) : ")
    codes_znieff = [code.strip() for code in codes_znieff_input.split(',') if code.strip()]

    # Validation des codes ZNIEFF
    valid_znieff_codes = []
    invalid_znieff_codes = []
    for code in codes_znieff:
        if is_valid_znieff_code(code):
            valid_znieff_codes.append(code)
        else:
            invalid_znieff_codes.append(code)
    
    if invalid_znieff_codes:
        print("\nCertains codes ZNIEFF sont invalides :")
        for code in invalid_znieff_codes:
            print(f"Code ZNIEFF invalide : {code}. Assurez-vous qu'il est composé de 9 chiffres.")
    else:
        print("\nTous les codes ZNIEFF sont valides ou aucun code ZNIEFF n'a été saisi.")
        break

# Saisie et validation des codes N2000
while True:
    codes_n2000_input = input("Veuillez entrer les codes N2000, séparés par des virgules (ex : FR1234567, FR7654321) : ")
    codes_n2000 = [code.strip() for code in codes_n2000_input.split(',') if code.strip()]

    # Validation des codes N2000
    valid_n2000_codes = []
    invalid_n2000_codes = []
    for code in codes_n2000:
        if is_valid_n2000_code(code):
            valid_n2000_codes.append(code)
        else:
            invalid_n2000_codes.append(code)
    
    if invalid_n2000_codes:
        print("\nCertains codes N2000 sont invalides :")
        for code in invalid_n2000_codes:
            print(f"Code N2000 invalide : {code}. Assurez-vous qu'il commence par 'FR' suivi de 7 chiffres.")
    else:
        print("\nTous les codes N2000 sont valides ou aucun code N2000 n'a été saisi.")
        break

# Télécharger les fichiers XML pour les codes ZNIEFF validés
print("\nTéléchargement des fichiers ZNIEFF...")
for code in valid_znieff_codes:
    download_xml(code, destination_folder, "ZNIEFF")

# Télécharger les fichiers XML pour les codes N2000 validés
print("\nTéléchargement des fichiers N2000...")
for code in valid_n2000_codes:
    download_xml(code, destination_folder, "N2000")

# Message final pour confirmer la fin des téléchargements
print(f"\nTéléchargements terminés. Les fichiers ont été sauvegardés dans le dossier : {destination_folder}")
