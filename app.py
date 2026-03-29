import pandas as pd
import numpy as np
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ===================== CONFIGURATION ROBUSTE =====================
# On cherche le fichier dans plusieurs emplacements possibles
possible_paths = [
    "Ville opti.xlsx",
    "/home/workdir/Ville opti.xlsx",
    "../Ville opti.xlsx",
    "/mount/src/ville/Ville opti.xlsx"
]

EXCEL_FILE = None
for path in possible_paths:
    if os.path.exists(path):
        EXCEL_FILE = path
        break

if EXCEL_FILE is None:
    raise FileNotFoundError("Impossible de trouver 'Ville opti.xlsx'. Vérifiez le nom et l'emplacement.")

OUTPUT_FILE = "Resultat_Optimisation_Ville.xlsx"

print(f"✅ Fichier trouvé : {EXCEL_FILE}")
print(f"Répertoire de travail : {os.getcwd()}")

# ===================== LECTURE DES DONNÉES =====================
print("\nChargement du fichier Excel...")

terrain_df = pd.read_excel(EXCEL_FILE, sheet_name="Terrain", header=None)
bat_df_raw = pd.read_excel(EXCEL_FILE, sheet_name="Batiments")
actuel_df = pd.read_excel(EXCEL_FILE, sheet_name="Actuel", header=None)

# Nettoyage onglet Batiments
bat_df = bat_df_raw.iloc[1:].reset_index(drop=True)
bat_df.columns = ["Nom", "Longueur", "Largeur", "Nombre", "Type", "Culture", 
                  "Rayonnement", "Boost 25%", "Boost 50%", "Boost 100%", 
                  "Production", "Quantite"]

terrain = terrain_df.to_numpy()
grid_actuel = actuel_df.to_numpy()

print(f"✅ Chargement réussi !")
print(f"   Terrain shape     : {terrain.shape}")
print(f"   Bâtiments types   : {len(bat_df)}")
print(f"   Grille actuelle   : {grid_actuel.shape}")

# ===================== FONCTION INFO BÂTIMENT =====================
def get_building_info(nom):
    nom = str(nom).strip()
    row = bat_df[bat_df['Nom'].str.strip() == nom]
    if row.empty:
        return None
    row = row.iloc[0]
    return {
        'nom': row['Nom'],
        'longueur': int(row['Longueur']),
        'largeur': int(row['Largeur']),
        'type': str(row['Type']).strip(),
        'culture': float(row['Culture']) if pd.notna(row['Culture']) else 0.0,
        'rayonnement': int(row['Rayonnement']) if pd.notna(row['Rayonnement']) else 0,
        'production': str(row['Production']) if pd.notna(row['Production']) else "Rien",
        'quantite': float(row['Quantite']) if pd.notna(row['Quantite']) else 0.0,
        'seuil_25': float(row['Boost 25%']) if pd.notna(row['Boost 25%']) else 0.0,
        'seuil_50': float(row['Boost 50%']) if pd.notna(row['Boost 50%']) else 0.0,
        'seuil_100': float(row['Boost 100%']) if pd.notna(row['Boost 100%']) else 0.0,
    }

print("\nScript chargé avec succès. Le fichier Excel est bien lu.")

# ===================== À DÉVELOPPER ENSUITE =====================
if __name__ == "__main__":
    print("\nLe script fonctionne maintenant.")
    print(f"Le résultat sera sauvegardé dans : {OUTPUT_FILE}")
    print("\nProchaine étape : je vais vous fournir la version complète avec optimisation, boosts et séquence de déplacements.")
