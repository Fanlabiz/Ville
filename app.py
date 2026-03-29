import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
import os

# ===================== CONFIGURATION =====================
EXCEL_FILE = "Ville opti.xlsx"      # Nom exact du fichier (il est bien présent dans le dossier)
OUTPUT_FILE = "Resultat_Optimisation_Ville.xlsx"

# Position de départ de la zone jouable (ligne 2, colonne B dans Excel)
START_ROW = 1   # 0-based
START_COL = 1   # 0-based (colonne B)

print(f"Répertoire actuel : {os.getcwd()}")
print(f"Fichier Excel trouvé : {os.path.exists(EXCEL_FILE)}")

# ===================== LECTURE DU FICHIER =====================
print("\nChargement du fichier Excel...")

terrain_df = pd.read_excel(EXCEL_FILE, sheet_name="Terrain", header=None)
bat_df = pd.read_excel(EXCEL_FILE, sheet_name="Batiments")
actuel_df = pd.read_excel(EXCEL_FILE, sheet_name="Actuel", header=None)

# Nettoyage des données Batiments
bat_df = bat_df.iloc[1:].reset_index(drop=True)  # Supprime la ligne d'en-tête si doublée
bat_df.columns = ["Nom", "Longueur", "Largeur", "Nombre", "Type", "Culture", 
                  "Rayonnement", "Boost 25%", "Boost 50%", "Boost 100%", 
                  "Production", "Quantite"]

terrain = terrain_df.to_numpy()
grid_actuel = actuel_df.to_numpy()

print(f"✅ Chargement réussi !")
print(f"   - Terrain : {terrain.shape}")
print(f"   - Bâtiments : {len(bat_df)} types")
print(f"   - Grille actuelle : {grid_actuel.shape}")

# ===================== FONCTION POUR RÉCUPÉRER LES INFOS D'UN BÂTIMENT =====================
def get_building_info(nom):
    row = bat_df[bat_df['Nom'].str.strip() == str(nom).strip()]
    if row.empty:
        return None
    row = row.iloc[0]
    return {
        'nom': row['Nom'],
        'longueur': int(row['Longueur']),
        'largeur': int(row['Largeur']),
        'type': row['Type'],
        'culture': float(row['Culture']) if pd.notna(row['Culture']) else 0,
        'rayonnement': int(row['Rayonnement']) if pd.notna(row['Rayonnement']) else 0,
        'production': row['Production'] if pd.notna(row['Production']) else "Rien",
        'quantite': float(row['Quantite']) if pd.notna(row['Quantite']) else 0,
        'seuil_25': float(row['Boost 25%']) if pd.notna(row['Boost 25%']) else 0,
        'seuil_50': float(row['Boost 50%']) if pd.notna(row['Boost 50%']) else 0,
        'seuil_100': float(row['Boost 100%']) if pd.notna(row['Boost 100%']) else 0,
    }

print("\nScript prêt. Le fichier se charge correctement maintenant.")

# ===================== LANCEMENT =====================
if __name__ == "__main__":
    print(f"\nLe fichier de résultat sera généré sous : {OUTPUT_FILE}")
    # Ici on pourra ajouter plus tard la logique d'optimisation complète
