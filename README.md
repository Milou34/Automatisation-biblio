# Automatisation-biblio

<a href="https://github.com/Milou34/Automatisation-biblio/actions/workflows/build-executable-create-release.yml" target="_blank">![Build Status](https://github.com/Milou34/Automatisation-biblio/actions/workflows/build-executable-create-release.yml/badge.svg)</a>
<a href="https://github.com/Milou34/Automatisation-biblio/actions/workflows/pylint.yml" target="_blank">![Code Quality](https://github.com/Milou34/Automatisation-biblio/actions/workflows/pylint.yml/badge.svg)</a>
<a href="https://github.com/Milou34/Automatisation-biblio/blob/main/LICENSE.txt" target="_blank">![Licence](https://img.shields.io/badge/Licence-Apache_2.0.-blue.svg)</a>

<a href="https://www.python.org/doc" target="_blank">![Python](https://img.shields.io/badge/Python-3.12-ffd343?logo=python)</a>
<a href="https://pypi.org/project/openpyxl" target="_blank">![Openpyxl](https://img.shields.io/badge/Openpyxl-3.1.5-ffd343?logo=pypi)</a>
<a href="https://pypi.org/project/requests" target="_blank">![Requests](https://img.shields.io/badge/Requests-2.32.3-ffd343?logo=pypi)</a>
<a href="https://pypi.org/project/psutil" target="_blank">![Psutil](https://img.shields.io/badge/Psutil-6.0.0-ffd343?logo=pypi)</a>

## Bibliographie Automatisée pour les Études Environnementales

Ce projet Python a pour objectif d'automatiser les tâches courantes de gestion de la bibliographie dans le cadre d'études environnementales. Il permet de centraliser la recherche, le formatage et l'organisation des sources, tout en facilitant l'intégration des références dans des feuilles de calcul.

## Fonctionnalités

- **Téléchargment automatisée** : Recherche et téléchargement des données au format XML depuis le site de l'INPN
- **Formatage des données** : Pipeline de récupération, préparation et formatage des données.
- **Création d'une feuille de calcul** : Création et insertion des données formatées dans un excel lisible.
- **Mise en forme du excel** : Formatage du style des cellules et des tableau.

## Utilisation

Avant de commencer, assurez-vous de suivre les étapes suivantes :

1. Télécharger <a href="https://github.com/Milou34/Automatisation-biblio/releases/latest" target="_blank">l'executable</a>
2. Lancer l'executable sur votre PC
3. Suivre les instructions qui s'affichent dans la console

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
Ce projet est sous licence Apache 2.0. Voir le fichier <a href="./LICENSE.txt" target="_blank">`LICENSE.txt`</a> pour plus de détails.
