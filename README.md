# Automatisation-biblio

[![Build Status](https://github.com/Milou34/Automatisation-biblio/actions/workflows/build-executable-create-release.yml/badge.svg)](https://github.com/Milou34/Automatisation-biblio/actions/workflows/build-executable-create-release.yml)
[![Code Quality](https://github.com/Milou34/Automatisation-biblio/actions/workflows/pylint.yml/badge.svg)](https://github.com/Milou34/Automatisation-biblio/actions/workflows/pylint.yml)


## Bibliographie Automatisée pour les Études Environnementales

Ce projet Python a pour objectif d'automatiser les tâches courantes de gestion de la bibliographie dans le cadre d'études environnementales. Il permet de centraliser la recherche, le formatage et l'organisation des sources, tout en facilitant l'intégration des références dans des feuilles de calcul.

## Fonctionnalités

- **Téléchargment automatisée** : Recherche et téléchargement des données au format XML depuis le site de l'INPN
- **Formatage des données** : Pipeline de récupération, préparation et formatage des données.
- **Création d'une feuille de calcul** : Création et insertion des données formatées dans un excel lisible.
- **Mise en forme du excel** : Formatage du style des cellules et des tableau.

## Utilisation

Avant de commencer, assurez-vous de suivre les étapes suivantes :

1. Télécharger [l'executable dans la release](https://github.com/Milou34/Automatisation-biblio/releases/latest)
2. Lancer l'executable sur votre PC
3. Suiver les instructions qui s'affichent dans la console

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

- **[Marylou Bertin](https://www.linkedin.com/in/marylou-bertin/)** : Cheffe de Projet & Developpeuse
- **[Valentin Guyon](https://www.linkedin.com/in/valentin-guyon/)** : Lead Developper

## Licence
Ce projet est sous licence Apache 2.0. Voir le fichier [LICENSE](./LICENSE.txt) pour plus de détails.
