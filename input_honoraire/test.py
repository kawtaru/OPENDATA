import pandas as pd
import os

# Définir les chemins d'entrée
input_dir = 'prix et indices des materiaux CSV/prix-et-indices-des-materiaux-construction-region-rabat-sale-kenitra_CSV/RABAT'  # à adapter si besoin
t1_path = os.path.join(input_dir, 'rabat_T1.csv')
t3_path = os.path.join(input_dir, 'rabat_T3.csv')

# Charger les fichiers
df_t1 = pd.read_csv(t1_path, sep=';', encoding='utf-8-sig')
df_t3 = pd.read_csv(t3_path, sep=';', encoding='utf-8-sig')

# S'assurer que la colonne 'Variété' existe dans df_t1
if 'Variété' not in df_t1.columns:
    df_t1['Variété'] = ''

# Copier le nom du produit dans la colonne 'Variété' pour les lignes globales
df_t3['Variété'] = df_t3['Produit']

# Définir les colonnes cibles pour la sortie
fusion_cols = ['Ville', 'Corps', 'Activité', 'Produit', 'Variété', 'Année', 'Indice']
for df in [df_t1, df_t3]:
    for col in fusion_cols:
        if col not in df.columns:
            df[col] = ''

# Fusionner les deux DataFrames
df_fusion = pd.concat([
    df_t1[fusion_cols],
    df_t3[fusion_cols]
], ignore_index=True)

# Tri facultatif pour meilleure lisibilité
df_fusion.sort_values(by=['Produit', 'Variété'], inplace=True)

# Enregistrer le résultat
output_path = os.path.join(input_dir, 'rabat_fusion.csv')
df_fusion.to_csv(output_path, sep=';', encoding='utf-8-sig', index=False)

print(f"[✓] Fichier fusionné sauvegardé : {output_path}")
