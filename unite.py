import pandas as pd
import os

input_dir = 'Unites_par_region_et_par_categorie_2016_2023'  # Dossier contenant les fichiers Excel
output_base_dir = 'Unites_par_region_et_par_categorie_2016_2023_csv_mod2'
os.makedirs(output_base_dir, exist_ok=True)

excel_files = [f for f in os.listdir(input_dir) if f.endswith('.xlsx')]

for excel_file in excel_files:
    file_path = os.path.join(input_dir, excel_file)
    df = pd.read_excel(file_path, header=0)

    records = []
    current_region = None

    for _, row in df.iterrows():
        region = row.iloc[0]
        category = row.iloc[1]

        if region == "Régions":
            continue

        if pd.isna(category):
            category = 0

        if pd.notna(region):
            current_region = region.strip()

        if current_region is None or pd.isna(category):
            continue

        year_cols = [2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023]
        col_idx = 2

        for year in year_cols:
            lots = pd.to_numeric(row.iloc[col_idx], errors='coerce') if col_idx < len(row) else None
            logements = pd.to_numeric(row.iloc[col_idx + 1], errors='coerce') if col_idx + 1 < len(row) else None
            restruct = pd.to_numeric(row.iloc[col_idx + 2], errors='coerce') if col_idx + 2 < len(row) else None

            lots = 0 if pd.isna(lots) else int(lots)
            logements = 0 if pd.isna(logements) else int(logements)
            restruct = 0 if pd.isna(restruct) else int(restruct)

            records.append({
                'Régions': current_region,
                'Catégories': category if isinstance(category, str) else str(category),
                'Année': year,
                'typologie': 'Lots',
                'nombre d\'unités': lots
            })
            records.append({
                'Régions': current_region,
                'Catégories': category if isinstance(category, str) else str(category),
                'Année': year,
                'typologie': 'Logements',
                'nombre d\'unités': logements
            })
            records.append({
                'Régions': current_region,
                'Catégories': category if isinstance(category, str) else str(category),
                'Année': year,
                'typologie': 'Unites de restructuration',
                'nombre d\'unités': restruct
            })

            col_idx += 4

    reshaped_df = pd.DataFrame(records)

    # Normaliser les noms de catégories (supprimer les espaces en trop)
    reshaped_df['Catégories'] = reshaped_df['Catégories'].str.strip()
    reshaped_df['Année'] = reshaped_df['Année'].astype(int)

    # Ordre des typologies
    typologie_order = {'Lots': 1, 'Logements': 2, 'Unites de restructuration': 3}
    reshaped_df['typologie_order'] = reshaped_df['typologie'].map(typologie_order)

    # Ordre personnalisé des catégories
    categorie_order = {
        'Opérations à faible valeur immobilière': 1,
        'Opérations économiques et sociales': 2,
        'Opérations économiques et sociales  ': 2,
        'Opérations économiques et sociales   ': 2,
        'App, en Immeubles (moyen et haut standing….)': 3,
        'Villas': 4,
        'Total': 5
    }
    reshaped_df['categorie_order'] = reshaped_df['Catégories'].map(categorie_order).fillna(999)

    # Tri final
    reshaped_df.sort_values(
        by=['Régions', 'Année', 'typologie_order', 'categorie_order'],
        inplace=True
    )

    reshaped_df.drop(columns=['typologie_order', 'categorie_order'], inplace=True)

    base_name = os.path.splitext(excel_file)[0].replace(" ", "").replace("-", "")
    output_dir = os.path.join(output_base_dir, base_name)
    os.makedirs(output_dir, exist_ok=True)

    regions = reshaped_df['Régions'].unique()
    regions = [region for region in regions if region != 'Total général' and region != 'Régions']

    for region_name in regions:
        safe_region = region_name.replace(" ", "").replace("-", "")
        file_name = os.path.join(output_dir, f"{safe_region}_csv.csv")
        region_df = reshaped_df[reshaped_df['Régions'] == region_name]
        region_df.to_csv(file_name, index=False, sep=';', encoding='utf-8-sig')

    total_general_df = reshaped_df[reshaped_df['Régions'] == 'Total général']
    if not total_general_df.empty:
        total_general_df.to_csv(os.path.join(output_dir, 'Total_general_csv.csv'), index=False, sep=';', encoding='utf-8-sig')

print("✅ Fichiers CSV générés avec 'nombre d'unités' renommé et tri : Région → Année → Typologie → Catégorie.")