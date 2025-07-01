import pandas as pd
import os

# Spécifiez le répertoire contenant vos fichiers Excel
input_folder = 'excels'  # Remplacez par le chemin du dossier contenant vos fichiers Excel
output_folder = 'Evolution_consommation_cilment_csv'  # Dossier où les fichiers CSV seront sauvegardés

# Créer le répertoire de sortie s'il n'existe pas
os.makedirs(output_folder, exist_ok=True)

# Liste tous les fichiers Excel dans le dossier
excel_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]

# Traiter chaque fichier Excel
for file_name in excel_files:
    file_path = os.path.join(input_folder, file_name)
    print(f"Traitement du fichier : {file_name}")

    # Lire le fichier Excel sans en-tête (header=None)
    df = pd.read_excel(file_path, header=None)

    # Supprimer la première ligne qui contient le titre
    df = df.drop(index=0)

    # Vérifier le nombre de colonnes
    print(f"Nombre de colonnes dans {file_name}: {df.shape[1]}")

    # Dynamiquement créer les noms de colonnes en fonction des années disponibles
    year_columns = [str(year) for year in range(2015, 2021)]  # Par défaut de 2015 à 2020

    if df.shape[1] > 7:  # Si le fichier contient plus de colonnes (par exemple, années jusqu'à 2023)
        year_columns = [str(year) for year in range(2015, 2021 + (df.shape[1] - 7))]  # Ajouter des années jusqu'à 2023

    # Réassigner les noms de colonnes
    df.columns = ['Régions'] + year_columns

    # Supprimer les lignes contenant des valeurs vides dans la colonne "Régions"
    df = df.dropna(subset=["Régions"])

    # Convertir toutes les colonnes de consommation en numérique, forcer les erreurs à NaN
    df[year_columns] = df[year_columns].apply(pd.to_numeric, errors='coerce')

    # Convertir les valeurs de consommation en entiers (en arrondissant les décimales)
    df[year_columns] = df[year_columns].round(0).astype(int)

    # Transformation en format long
    df_long = pd.melt(df, id_vars=["Régions"], var_name="Année", value_name="Consommation")

    # Trier les résultats par région pour avoir chaque région répétée pour chaque année
    df_long = df_long.sort_values(by=["Régions", "Année"])

    # Sauvegarder en CSV
    output_file_path = os.path.join(output_folder, file_name.replace('.xlsx', '_output.csv'))
    df_long.to_csv(output_file_path, index=False, sep=';', encoding='utf-8-sig')

    print(f"Fichier CSV généré avec succès : {output_file_path}")

print("Tous les fichiers ont été traités et sauvegardés.")