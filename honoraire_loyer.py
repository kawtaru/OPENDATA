# src/extract_loyers_data.py
import pandas as pd
import os

INPUT_DIR = 'input_honoraire_loyer'
OUTPUT_DIR = 'montants-moyens-des-loyers_csv'

os.makedirs(OUTPUT_DIR, exist_ok=True)

def traiter_fichier_loyers(filepath):
    print(f"\n=== Traitement du fichier : {os.path.basename(filepath)} ===")

    df_raw = pd.read_excel(filepath, header=None)

    for i in range(min(10, len(df_raw))):
        row_data = []
        for j in range(min(50, len(df_raw.columns))):
            val = df_raw.iloc[i, j]
            row_data.append(f"Col{j}:{val}" if pd.notna(val) else f"Col{j}:NaN")
        print(f"Row {i}: {row_data[:15]}...")

    year_row_idx = None
    for i in range(len(df_raw)):
        row = df_raw.iloc[i]
        year_count = 0
        for j, val in enumerate(row):
            try:
                year_val = int(float(str(val).strip().replace(',', '.')))
                if 2000 <= year_val <= 2025:
                    year_count += 1
            except:
                continue
        if year_count >= 5:
            year_row_idx = i
            print(f"=== Ligne années trouvée à l’index {i} avec {year_count} années ===")
            break

    if year_row_idx is None:
        raise ValueError("Ligne des années introuvable dans le fichier")

    header_row = df_raw.iloc[year_row_idx]
    year_positions = [(i, int(float(str(val).replace(',', '.'))))
                      for i, val in enumerate(header_row)
                      if pd.notna(val) and str(val).strip().replace(',', '.').replace(' ', '').replace('\xa0', '').replace('.0', '').isdigit()
                      and 2000 <= int(float(str(val).replace(',', '.')))
                      <= 2025]

    split_idx = len(year_positions) // 2
    montant_positions = year_positions[:split_idx]
    indice_positions = year_positions[split_idx:]

    montant_map = {year: pos for pos, year in montant_positions}
    indice_map = {year: pos for pos, year in indice_positions}

    common_years = sorted(set(montant_map.keys()) & set(indice_map.keys()))
    if not common_years:
        common_years = sorted(montant_map.keys())

    data = []
    current_info = [None, None, None]

    for i in range(year_row_idx + 1, len(df_raw)):
        row = df_raw.iloc[i]
        for j in range(3):
            if pd.notna(row[j]) and str(row[j]).strip():
                current_info[j] = str(row[j]).strip()
        if not all(current_info):
            continue

        for year in common_years:
            try:
                montant_val = row[montant_map[year]] if year in montant_map else None
                indice_val = row[indice_map[year]] if year in indice_map else None

                if pd.isna(montant_val):
                    continue

                # Nettoyer et convertir le montant
                montant_str = str(montant_val).replace('\xa0', '').replace(',', '.').replace(' ', '')
                try:
                    montant_float = float(montant_str)
                    montant = "Non disponible" if montant_float == 0 else round(montant_float)
                except:
                    montant = "Non disponible"

                # Nettoyer et convertir l'indice
                indice = "Non disponible"
                if indice_val is not None and pd.notna(indice_val):
                    indice_str = str(indice_val).replace('\xa0', '').replace(',', '.').replace(' ', '')
                    try:
                        indice_float = float(indice_str)
                        if indice_float != 0:
                            indice = round(indice_float)
                    except:
                        pass

                data.append(current_info + [year, montant, indice])
            except:
                continue

    if data:
        df_out = pd.DataFrame(data, columns=[ 
            'Agglomération',
            'Type d\'habitat',
            'Envergure',
            'Année',
            'Montants moyens des loyers',
            'Indices des montants moyens des loyers'
        ])

        # Ajouter l'unité "dirham marocain (MAD)" à la colonne "Montants moyens des loyers"
        df_out = df_out.rename(columns={
            'Montants moyens des loyers': 'Montants moyens des loyers (MAD)'
        })

        df_out = df_out.sort_values(['Agglomération', 'Type d\'habitat', 'Envergure', 'Année'])

        base_name = os.path.splitext(os.path.basename(filepath))[0]
        output_file = os.path.join(OUTPUT_DIR, f"{base_name}.csv")
        df_out.to_csv(output_file, index=False, encoding='utf-8-sig', sep=';')

        print(f"-> Fichier sauvegardé : {output_file}")
        print(f"  Total records : {len(df_out)}")
    else:
        print("-> Aucun donnée extraite.")

for f in os.listdir(INPUT_DIR):
    if f.endswith('.xlsx'):
        full_path = os.path.join(INPUT_DIR, f)
        try:
            traiter_fichier_loyers(full_path)
        except Exception as e:
            print(f"Erreur avec le fichier {f} : {e}")
