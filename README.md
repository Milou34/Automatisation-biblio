# Automatisation-biblio

<a href="https://github.com/Milou34/Automatisation-biblio/actions/workflows/build-executable-create-release.yml" target="_blank">![Build Status](https://github.com/Milou34/Automatisation-biblio/actions/workflows/build-executable-create-release.yml/badge.svg)</a>
<a href="https://github.com/Milou34/Automatisation-biblio/actions/workflows/pylint.yml" target="_blank">![Code Quality](https://github.com/Milou34/Automatisation-biblio/actions/workflows/pylint.yml/badge.svg)</a>
<a href="https://github.com/Milou34/Automatisation-biblio/blob/main/LICENSE.txt" target="_blank">![Licence](https://img.shields.io/badge/Licence-Apache_2.0-blue.svg)</a>

<a href="https://www.python.org/doc" target="_blank">![Python](https://img.shields.io/badge/Python-3.12-ffd343?logo=python)</a>
<a href="https://pypi.org/project/openpyxl" target="_blank">![Openpyxl](https://img.shields.io/badge/Openpyxl-3.1.5-ffd343?logo=pypi)</a>
<a href="https://pypi.org/project/requests" target="_blank">![Requests](https://img.shields.io/badge/Requests-2.32.3-ffd343?logo=pypi)</a>
<a href="https://pypi.org/project/psutil" target="_blank">![Psutil](https://img.shields.io/badge/Psutil-6.0.0-ffd343?logo=pypi)</a>

## Bibliographie Automatisée pour les Études Environnementales

Ce projet Python a pour objectif d'automatiser les tâches courantes de gestion de la bibliographie dans le cadre d'études environnementales. Il permet de centraliser la recherche, le formatage et l'organisation des sources, tout en facilitant l'intégration des références dans des feuilles de calcul.

## Fonctionnalités

- **Téléchargment automatisée** : Recherche et téléchargement des données ZNIEFF et N2000 au format XML depuis le site de l'INPN
- **Formatage des données** : Pipeline de récupération, préparation et formatage des données.
- **Création de plusieurs feuilles de calcul et intégration des données** : Création et insertion des données formatées dans un excel lisible et renommé.
- **Mise en forme du excel** : Formatage du style des cellules et des tableau.

## Utilisation

Avant de commencer, assurez-vous de suivre les étapes suivantes :


1. Télécharger <a href="https://github.com/Milou34/Automatisation-biblio/releases/latest" target="_blank">l'executable</a> (cliquer sur le .exe)
2. Préalablement au lancement du programme, pour le ou les projets dont vous souhaitez créer la bibliographie, assurez vous d'avoir bien créé la couche `zonages_aires_detude` à l'aide du modèle Zonages sur QGIS. Dans la table attributaire de cette couche, vous pourrez retrouver les `codes ZNIEFF et Natura 2000` demandés par le programme.
3. Lancez l'exécutable `v...-bibliographie-zonage` depuis votre dossier téléchargements.
4. A la première exécution du programme, cliquez sur `Informations complémentaires`, puis sur `Exécuter quand même`. 
5. Suivez les instructions qui s'affichent dans la console : 

    Renseignez le chemin du dossier de destination, où seront téléchargés les fichiers XML et où sera créé le Excel final.

    **ATTENTION! Ce dossier ne doit pas se trouver sur le OneDrive** (Sinon vous aurez une erreur).\
    **Tips** : Ouvrez l'explorateur de fichiers puis allez dans :
    Ce PC > Windows (C:) > Utilisateurs > PrénomNOM > 
    Puis créez un dossier `Documents` à cet endroit (en local). Vous travaillerez TOUJOURS depuis ce dossier pour créer vos Excels de bibliographie.

    A cette étape, vous pouvez donc copier/coller le chemin du dossier dans la console (clic droit sur le chemin en haut de la fenêtre, `copier l'adresse`) et écrire à la suite le nom du dossier à créer.\
    Exemple : `C:\Users\PrénomNOM\Documents\106-Féricy-Bibliographie`\
    Puis appuyer sur `Entrer`.

6. Entrer les codes des ZNIEFFs de types 1 et 2 présentes dans l'AER, qui sont notés dans la colonne `ID_MNHN` de la couche `zonages_aires_detude` dans QGIS, en les séparant par une virgule.\
Puis appuyer sur `Entrer`.

    S'il n'y a pas de code ZNIEFF à renseigner, appuyer seulement sur `Entrer` et poursuivez le programme.

    S'il y a une erreur sur l'un des codes (s'ils n'ont pas exactement 9 chiffres), les codes sont redemandés.\
    **Attention à bien renseigner uniquement des codes ZNIEFF.**

7. Entrer les codes des zones Natura 2000 présentes dans l'AEE, qui sont notés dans la colonne `SITE_CODE` de la couche `zonages_aires_detude` dans QGIS, en les séparant par une virgule.\
Puis appuyer sur `Entrer`.

    S'il n'y a pas de code Natura 2000 à renseigner, appuyer seulement sur `Entrer` et poursuivez le programme.

    S'il y a une erreur sur l'un des codes (s'ils ne commencent pas par FR suivi de 7 chiffres exactement), les codes sont redemandés.\
    **Attention à bien renseigner uniquement des codes Natura 2000.**

8. Entrer le numéro du projet, cela permettra de renommer automatiquement l'excel final en : `Bibliographie - n° projet`.\
Puis appuyer sur `Entrer`.

9. Le programme génère et ouvre l'excel final.

10. Si vous souhaitez poursuivre avec la bibliographie d'un autre projet, appuyez sur `Entrer`. Sinon, appuyez sur n'importe quelle autre touche pour quitter le programme.



## Structure du Projet

```
.
├── README.md               # Documentation du projet
├── requirements.txt        # Dépendances du projet
├── favicon.ico             # icon de l'executable
├── LICENCE.txt             # Condition d'utilisation
├── main.py                 # Script principal
├── .gitignore              # Ignorer les fichiers compilés
├── .github/                # Dossier de configuration du repository
└── src/                    # Dossier du code source
```

## Authors

- <a href="https://www.linkedin.com/in/marylou-bertin" target="_blank">**Marylou Bertin**</a> : Cheffe de Projet & Developpeuse
- <a href="https://www.linkedin.com/in/valentin-guyon" target="_blank">**Valentin Guyon**</a> : Lead Developper & DevOps

## Licence
Ce projet est sous licence Apache 2.0 Voir le fichier <a href="./LICENSE.txt" target="_blank">`LICENSE.txt`</a> pour plus de détails.
