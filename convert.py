import pandas as pd
import re
import os
from collections import defaultdict

input_folder = 'Prix et indices des matériaux de construction_xlsx'
excel_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]

for xlsx_file in excel_files:
    xlsx_path = os.path.join(input_folder, xlsx_file)
    workbook = pd.ExcelFile(xlsx_path)
    sheet_names = workbook.sheet_names
    print(f"=== Processing File: {xlsx_file} ===")
    print("Sheets Found:", sheet_names)

    sheet_groups = defaultdict(lambda: defaultdict(list))
    for sheet in sheet_names:
        parts = sheet.strip().split()
        if len(parts) >= 2:
            prefix = parts[0].strip().upper()
            region = parts[1].strip().upper()
            sheet_groups[region][prefix].append(sheet)

    base_output_dir = f"{os.path.splitext(xlsx_file)[0]}_CSV"

    for region, prefix_dict in sheet_groups.items():
        csv_output_dir = os.path.join(base_output_dir, region)
        os.makedirs(csv_output_dir, exist_ok=True)

        collected_data = {}

        for prefix in ['T1', 'T2', 'T3']:
            if prefix not in prefix_dict:
                continue

            data = []
            grouped_sheets = prefix_dict[prefix]
            for sheet in grouped_sheets:
                print(f"  -> {sheet}")
                df_all = pd.read_excel(xlsx_path, sheet_name=sheet, header=None)
                try:
                    header_idx = df_all[df_all.apply(lambda row: row.astype(str).str.contains('Corps').any(), axis=1)].index[0]
                except IndexError:
                    print(f"     'Corps' header not found in sheet {sheet}, skipping...")
                    continue

                header_row = df_all.iloc[header_idx].astype(str).str.strip().tolist()

                id_columns = ['Corps', 'Activité', 'Produit', 'Variété']
                nb_cols_produit = 0
                for col_name in id_columns:
                    if any(col_name.lower() == str(h).lower() for h in header_row):
                        nb_cols_produit += 1
                    else:
                        if col_name != 'Variété':
                            print(f"     Missing mandatory column '{col_name}' in sheet {sheet}, skipping...")
                            nb_cols_produit = 0
                            break
                if nb_cols_produit == 0:
                    continue

                df_raw = pd.read_excel(xlsx_path, sheet_name=sheet, skiprows=header_idx + 1, header=None)
                year_row = df_all.iloc[header_idx + 1]

                try:
                    year_start_col = next(i for i, val in enumerate(year_row) if re.match(r"\d{4}", str(val)))
                except StopIteration:
                    print(f"     No year found in the expected row in sheet {sheet}, skipping...")
                    continue

                years = [int(float(year_row[i])) for i in range(year_start_col, len(year_row)) if re.match(r"\d{4}", str(year_row[i]))]

                current_row = [None] * nb_cols_produit

                for _, row in df_raw.iterrows():
                    for i in range(nb_cols_produit):
                        if pd.notna(row[i]) and str(row[i]).strip():
                            current_row[i] = str(row[i]).strip()

                    if nb_cols_produit == 4:
                        if not all(current_row[:3]):
                            continue
                    else:
                        if not all(current_row):
                            continue

                    produit_info = current_row.copy()
                    values = row[year_start_col:year_start_col + len(years)]

                    for i, year in enumerate(years):
                        montant_raw = str(values.iloc[i]).replace('\xa0', '').replace(',', '.').replace(' ', '')
                        try:
                            montant = round(float(montant_raw), 2)
                        except:
                            continue
                        data.append(produit_info + [year, montant])

            if data:
                columns = ['Corps', 'Activité', 'Produit']
                if nb_cols_produit == 4:
                    columns.append('Variété')
                columns += ['Année', 'Montant']

                df = pd.DataFrame(data, columns=columns)
                df['Ville'] = region
                collected_data[prefix] = df

        if 'T1' in collected_data:
            df_t1 = collected_data['T1']
            if 'T2' in collected_data:
                df_t2 = collected_data['T2']
                df_t2_grouped = df_t2.groupby(['Corps', 'Activité', 'Produit', 'Année']).agg({'Montant': 'sum'}).reset_index()
                df_t2_grouped['Indice'] = df_t2_grouped['Montant'].apply(lambda x: f"{x:.1f}")
                df_t1 = df_t1.merge(df_t2_grouped[['Corps', 'Activité', 'Produit', 'Année', 'Indice']], on=['Corps', 'Activité', 'Produit', 'Année'], how='left')
            df_t1 = df_t1[['Ville', 'Corps', 'Activité', 'Produit', 'Variété', 'Année', 'Montant', 'Indice'] if 'Variété' in df_t1.columns else ['Ville', 'Corps', 'Activité', 'Produit', 'Année', 'Montant', 'Indice']]
            df_t1.to_csv(os.path.join(csv_output_dir, f"{region.lower()}_T1.csv"), index=False, encoding='utf-8-sig', sep=';')
            print(f"Saved: {region.lower()}_T1.csv")

        if 'T3' in collected_data:
            df_t3 = collected_data['T3']
            df_t3.rename(columns={'Montant': 'Indice'}, inplace=True)
            df_t3['Indice'] = df_t3['Indice'].apply(lambda x: f"{x:.1f}")
            df_t3 = df_t3[['Ville', 'Corps', 'Activité', 'Produit', 'Variété', 'Année', 'Indice'] if 'Variété' in df_t3.columns else ['Ville', 'Corps', 'Activité', 'Produit', 'Année', 'Indice']]
            df_t3.to_csv(os.path.join(csv_output_dir, f"{region.lower()}_T3.csv"), index=False, encoding='utf-8-sig', sep=';')
            print(f"Saved: {region.lower()}_T3.csv")