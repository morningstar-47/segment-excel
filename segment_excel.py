import argparse
import os
import openpyxl


def main():
    # Parsing des arguments de la ligne de commande
    parser = argparse.ArgumentParser(
        description='Séparer les données de fichiers Excel')
    parser.add_argument('path', type=str,
                        help='le chemin d\'accès au fichier Excel')
    parser.add_argument('-f', '--fixe', action='store_true',
                        help='Filtrer sur Téléphone fixe')
    parser.add_argument('-m', '--mobile', action='store_true',
                        help='Filtrer sur Numéro de téléphone')
    parser.add_argument('size', type=int,
                        help='Définir la taille de ligne de sortie')
    args = parser.parse_args()

    # Vérification des options de filtrage
    if args.fixe and args.mobile:
        raise ValueError(
            'Vous ne pouvez spécifier qu\'un seul type de téléphone à filtrer.')
    elif not args.fixe and not args.mobile:
        raise ValueError(
            'Vous devez spécifier un type de téléphone à filtrer.')

    if not args.size:
        raise ValueError(
            'Vous devez spécifier le nombre de la ligne de sortie.')

    if not os.path.exists(args.path):
        raise FileNotFoundError(
            f"Le fichier '{args.path}' n'existe pas.")

    # Chargement du fichier Excel
    wb = openpyxl.load_workbook(args.path)

    for nom_feuille in wb.sheetnames:
        feuille = wb[nom_feuille]

        en_tete = [cell.value for cell in feuille[1]]

        if args.fixe:
            type_tel = "Téléphone fixe"
        else:
            type_tel = "Numéro de téléphone"

        if type_tel not in en_tete:
            raise ValueError(
                f"{type_tel} n'est pas présent dans la liste des entêtes.")

        index_telephone = en_tete.index(type_tel)

        donnees_filtrees = [en_tete]
        for ligne in feuille.iter_rows(min_row=2, values_only=True):
            if ligne[index_telephone]:
                donnees_filtrees.append(ligne)

        blocs = []
        bloc = [en_tete]
        for ligne in donnees_filtrees[1:]:
            bloc.append(ligne)
            if len(bloc) == args.size+1:
                blocs.append(bloc)
                bloc = [en_tete]
        if len(bloc) > 1:
            blocs.append(bloc)

        dossier_general = os.path.join(os.getcwd(), "output")
        if not os.path.exists(dossier_general):
            os.mkdir(dossier_general)

        nom_dossier = nom_feuille.replace(" ", "_")
        chemin_dossier = os.path.join(dossier_general, nom_dossier)
        if not os.path.exists(chemin_dossier):
            os.mkdir(chemin_dossier)

        for i, bloc in enumerate(blocs):
            nouveau_nom_fichier = os.path.join(
                chemin_dossier, nom_feuille.capitalize() + f"_{i+1}.xlsx")
            nouveau_fichier = openpyxl.Workbook()
            nouvelle_feuille = nouveau_fichier.active
            for ligne in bloc:
                nouvelle_feuille.append(ligne)
            nouveau_fichier.save(nouveau_nom_fichier)

            print("ok -", nom_feuille.capitalize() + f"_{i+1}.xlsx - ")


if __name__ == '__main__':
    main()
