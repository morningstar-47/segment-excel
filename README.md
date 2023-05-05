# Segment_excel
Séparation des données de fichiers Excel
Ce script permet de séparer les données d'un fichier Excel en blocs de taille donnée et de filtrer ces données en fonction du type de numéro de téléphone (fixe ou mobile).

**Utilisation**

Pour utiliser ce script, exécutez la commande suivante dans un terminal :

``
python segment_excel.py path/to/excel/file.xlsx -f/-m size
``

où :

- `path/to/excel/file.xlsx` est le chemin d'accès au fichier Excel que vous souhaitez traiter;
- `'-f' '--fixe'` ou `'-m' '--mobile'` est l'option de filtrage pour les numéros de téléphone fixe ou mobile respectivement;
- `size` est la taille de chaque bloc de données de sortie.

**Dépendances**

Ce script nécessite les dépendances suivantes :

- argparse
- os
- openpyxl

Vous pouvez les installer en utilisant la commande suivante :
```
pip install argparse os openpyxl
```

**Exemple**

Supposons que vous avez un fichier Excel nommé data.xlsx dans votre répertoire courant, avec deux feuilles nommées Feuille 1 et Feuille 2. Vous voulez filtrer les données en fonction du type de téléphone mobile et créer des blocs de 50 lignes chacun.

Vous pouvez exécuter la commande suivante dans un terminal :

```
python segment_excel.py data.xlsx -m 50
```

Cela créera un dossier output dans votre répertoire courant, avec deux sous-dossiers nommés Feuille_1 et Feuille_2.
Chaque sous-dossier contiendra plusieurs fichiers Excel nommés Feuille_1_1.xlsx, Feuille_1_2.xlsx, etc. 
Ces fichiers contiendront des blocs de 50 lignes de données filtrées en fonction du type de téléphone mobile.
