import pandas as pd
import os

# === Config ===
INPUT_DIR = 'input_honoraire'     
OUTPUT_DIR = 'montants-moyens-des-cautions_csv'    

os.makedirs(OUTPUT_DIR, exist_ok=True)

# Fonction de traitement d’un fichier Excel (ton script complet adapté)
def traiter_fichier(filepath):
    print(f"\n=== Traitement du fichier : {os.path.basename(filepath)} ===")

    df_raw = pd.read_excel(filepath, header=None)

    # --- Debug structure (affiche les 10 premières lignes) ---
    for i in range(min(10, len(df_raw))):
        row_data = []
        for j in range(min(50, len(df_raw.columns))):
            val = df_raw.iloc[i, j]
            if pd.notna(val):
                row_data.append(f"Col{j}:{val}")
            else:
                row_data.append(f"Col{j}:NaN")
        print(f"Row {i}: {row_data[:15]}...")

    # --- Trouver la ligne des années ---
    year_row_idx = None
    for i in range(len(df_raw)):
        row = df_raw.iloc[i]
        year_count = 0
        years_found = []

        for j, val in enumerate(row):
            if pd.notna(val):
                str_val = str(val).strip()
                try:
                    if '.' in str_val and str_val.endswith('.0'):
                        year_val = int(float(str_val))
                    elif str_val.isdigit():
                        year_val = int(str_val)
                    else:
                        continue
                    if 2000 <= year_val <= 2025:
                        year_count += 1
                        years_found.append((j, year_val))
                except:
                    continue

        if year_count >= 5:
            year_row_idx = i
            print(f"=== Ligne années trouvée à l’index {i} avec {year_count} années ===")
            break

    if year_row_idx is None:
        raise ValueError("Ligne des années introuvable dans le fichier")

    header_row = df_raw.iloc[year_row_idx]

    # --- Identifier les positions des années ---
    year_positions = []
    for i, val in enumerate(header_row):
        if pd.notna(val):
            str_val = str(val).strip()
            try:
                if '.' in str_val and str_val.endswith('.0'):
                    year_val = int(float(str_val))
                elif str_val.isdigit():
                    year_val = int(str_val)
                else:
                    continue
                if 2000 <= year_val <= 2025:
                    year_positions.append((i, year_val))
            except:
                continue

    # --- Identifier les sections (montant et indice) ---
    sections = []
    current_section = []
    year_positions.sort(key=lambda x: x[0])

    for i, (pos, year) in enumerate(year_positions):
        if i == 0:
            current_section.append((pos, year))
            continue
        prev_pos, prev_year = year_positions[i-1]
        if (pos - prev_pos > 5) or (year < prev_year):
            if current_section:
                sections.append(current_section)
            current_section = [(pos, year)]
        else:
            current_section.append((pos, year))
    if current_section:
        sections.append(current_section)

    # Gestion si pas exactement 2 sections
    if len(sections) != 2:
        if len(sections) == 1:
            section = sections[0]
            mid_point = len(section) // 2
            sections = [section[:mid_point], section[mid_point:]]
        elif len(sections) > 2:
            sections = [sections[0], sections[-1]]

    montant_positions = sections[0] if len(sections) >= 1 else []
    indice_positions = sections[1] if len(sections) >= 2 else []

    montant_map = {year: pos for pos, year in montant_positions}
    indice_map = {year: pos for pos, year in indice_positions}

    common_years = sorted(set(montant_map.keys()) & set(indice_map.keys()))
    if not common_years:
        common_years = sorted(montant_map.keys())

    # --- Extraction des données ---
    data = []
    current_info = [None, None, None]  # Agglo, Type, Envergure

    for i in range(year_row_idx + 1, len(df_raw)):
        row = df_raw.iloc[i]

        info_updated = False
        for j in range(3):
            if pd.notna(row[j]) and str(row[j]).strip():
                current_info[j] = str(row[j]).strip()
                info_updated = True

        if not all(current_info):
            continue

        row_data_count = 0
        for year in common_years:
            try:
                montant_pos = montant_map.get(year)
                indice_pos = indice_map.get(year)

                if montant_pos is None:
                    continue

                montant_val = row[montant_pos]
                if pd.isna(montant_val):
                    continue

                indice_val = None
                if indice_pos is not None:
                    indice_val = row[indice_pos]

                montant_clean = str(montant_val).replace('\xa0', '').replace(',', '.').replace(' ', '')
                if montant_clean in ['', 'nan']:
                    continue
                montant = round(float(montant_clean))

                indice = None
                if indice_val is not None and pd.notna(indice_val):
                    indice_clean = str(indice_val).replace('\xa0', '').replace(',', '.').replace(' ', '')
                    if indice_clean not in ['', 'nan']:
                        indice = round(float(indice_clean))

                data.append(current_info + [year, montant, indice])
                row_data_count += 1

            except:
                continue

    # --- Export CSV ---
    if data:
        df_out = pd.DataFrame(data, columns=[
            'Agglomération',
            'Type d\'habitat',
            'Envergure',
            'Année',
            'Montants moyens des honoraires des agents immobiliers',
            'Indices des montants moyens des honoraires des agents immobiliers'
        ])

        df_out = df_out.sort_values(['Agglomération', 'Type d\'habitat', 'Envergure', 'Année'])

        # Nom fichier de sortie
        base_name = os.path.splitext(os.path.basename(filepath))[0]
        output_file = os.path.join(OUTPUT_DIR, f"{base_name}.csv")

        df_out.to_csv(output_file, index=False, encoding='utf-8-sig', sep=';')

        print(f"-> Fichier sauvegardé : {output_file}")
        print(f"  Total records : {len(df_out)}")
    else:
        print("-> Aucun donnée extraite.")

# === Boucle sur tous les fichiers Excel dans le dossier input ===
for f in os.listdir(INPUT_DIR):
    if f.endswith('.xlsx'):
        full_path = os.path.join(INPUT_DIR, f)
        try:
            traiter_fichier(full_path)
        except Exception as e:
            print(f"Erreur avec le fichier {f} : {e}")