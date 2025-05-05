# Segment Excel

Un utilitaire Python pour segmenter des fichiers Excel en fonction de critères de filtrage téléphonique et générer des fichiers de taille contrôlée.

## 📋 Description

Ce script permet de traiter des fichiers Excel contenant des données de contact et de :
- Filtrer les lignes contenant un numéro de téléphone (fixe ou mobile)
- Diviser les données filtrées en plusieurs fichiers Excel plus petits
- Organiser les fichiers de sortie dans une structure de dossiers cohérente

Idéal pour préparer des données pour des campagnes téléphoniques, des études de marché ou toute autre analyse nécessitant un traitement par lots.

## ✨ Fonctionnalités

- Filtrage sélectif sur les numéros de téléphone fixes ou mobiles
- Segmentation des données en fichiers de taille personnalisable
- Support multi-feuilles (traite toutes les feuilles du classeur)
- Système de journalisation pour suivre le processus
- Validation des entrées pour éviter les erreurs courantes
- Structure de dossiers organisée pour les fichiers de sortie

## 🚀 Installation

1. Clonez ce dépôt :
```bash
git clone https://github.com/morningstar-47/segment-excel.git
cd segment-excel
```

2. Installez les dépendances :
```bash
pip install -r requirements.txt
```

## 📝 Prérequis

- Python 3.6 ou supérieur
- Bibliothèque openpyxl

Pour installer les dépendances :
```bash
pip install openpyxl
```

## 🔧 Utilisation

### Syntaxe de base

```bash
python main.py chemin/vers/fichier.xlsx [-f | -m] taille_fichier [-o dossier_sortie] [-v]
```

### Arguments

- `chemin/vers/fichier.xlsx` : Chemin d'accès au fichier Excel à traiter
- `-f, --fixe` : Filtrer sur la colonne "Téléphone fixe"
- `-m, --mobile` : Filtrer sur la colonne "Numéro de téléphone"
- `taille_fichier` : Nombre maximal de lignes par fichier de sortie (sans compter l'en-tête)
- `-o, --output` : Dossier de sortie (par défaut: "output")
- `-v, --verbose` : Afficher les messages de débogage détaillés

### Exemples

Filtrer sur les numéros fixes et créer des fichiers de 100 lignes :
```bash
python main.py data.xlsx -f 100
```

Filtrer sur les numéros mobiles et créer des fichiers de 50 lignes avec un dossier de sortie personnalisé :
```bash
python main.py data.xlsx -m 50 -o resultats
```

Utiliser le mode verbeux pour afficher plus d'informations :
```bash
python main.py data.xlsx -f 200 -v
```

## 📁 Structure du projet

```
segment-excel/
├── main.py     # Script principal
├── requirements.txt     # Dépendances
└── README.md            # Documentation
```

## 📊 Structure des fichiers de sortie

```
output/                          # Dossier principal de sortie
├── Nom_de_feuille_1/            # Un dossier par feuille du classeur
│   ├── Nom_de_feuille_1_1.xlsx  # Premier segment
│   ├── Nom_de_feuille_1_2.xlsx  # Deuxième segment
│   └── ...
└── Nom_de_feuille_2/
    ├── Nom_de_feuille_2_1.xlsx
    └── ...
```

## 📋 Format de données attendu

Le script s'attend à ce que votre fichier Excel contienne au moins une des colonnes suivantes :
- "Téléphone fixe" - pour les numéros de téléphone fixes
- "Numéro de téléphone" - pour les numéros de téléphone mobiles

## 🤝 Contribution

Les contributions sont les bienvenues ! N'hésitez pas à :
1. Fork le projet
2. Créer une branche pour votre fonctionnalité (`git checkout -b feature/amazing-feature`)
3. Commit vos changements (`git commit -m 'Ajout d'une fonctionnalité incroyable'`)
4. Push sur la branche (`git push origin feature/amazing-feature`)
5. Ouvrir une Pull Request

## 📄 Licence

Ce projet est sous licence [MIT](LICENSE).

## 📧 Contact

MorningStar - 47

Lien du projet : [https://github.com/morningstar-47/segment-excel](https://github.com/morningstar-47/segment-excel)
