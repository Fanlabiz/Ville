import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
import os

# ===================== CONFIGURATION =====================
# ←←← MODIFIEZ ICI SI NÉCESSAIRE ←←←
EXCEL_FILE = "Ville opti.xlsx"          # Nom exact de votre fichier

# Si le fichier n'est pas dans le même dossier que le script, mettez le chemin complet :
# EXCEL_FILE = "/chemin/complet/vers/Ville opti.xlsx"

OUTPUT_FILE = "Resultat_Optimisation_Ville.xlsx"

# Position de la zone jouable (ligne 2, colonne B dans Excel)
START_ROW = 1   # index 0-based → ligne 2 Excel
START_COL = 1   # index 0-based → colonne B Excel

print(f"Répertoire actuel : {os.getcwd()}")
print(f"Fichier recherché : {EXCEL_FILE}")

# Vérification de l'existence du fichier
if not os.path.exists(EXCEL_FILE):
    raise FileNotFoundError(f"""
ERREUR : Le fichier '{EXCEL_FILE}' n'a pas été trouvé !
Répertoire actuel : {os.getcwd()}
Fichiers présents : {os.listdir('.')}
    
Solutions :
1. Placez 'Ville opti.xlsx' dans le même dossier que le script.
2. Ou modifiez la variable EXCEL_FILE avec le chemin complet.
""")

# ===================== LECTURE DES DONNÉES =====================
print("Lecture du fichier Excel...")

try:
    # Terrain
    terrain_df = pd.read_excel(EXCEL_FILE, sheet_name="Terrain", header=None)
    terrain = terrain_df.to_numpy()

    # Batiments
    bat_df = pd.read_excel(EXCEL_FILE, sheet_name="Batiments")
    if bat_df.columns[0] == "Nom":
        bat_df = bat_df.iloc[1:].reset_index(drop=True)
    bat_df.columns = ["Nom", "Longueur", "Largeur", "Nombre", "Type", "Culture", 
                      "Rayonnement", "Boost 25%", "Boost 50%", "Boost 100%", 
                      "Production", "Quantite"]

    # Configuration actuelle
    actuel_df = pd.read_excel(EXCEL_FILE, sheet_name="Actuel", header=None)
    grid = actuel_df.to_numpy()

    print(f"✅ Fichier chargé avec succès !")
    print(f"   Terrain : {terrain.shape}")
    print(f"   {len(bat_df)} types de bâtiments")
    print(f"   Grille actuelle : {grid.shape}")

except Exception as e:
    print(f"❌ Erreur lors de la lecture du fichier : {e}")
    raise

# ===================== FONCTIONS DE BASE =====================
def get_building_info(name):
    row = bat_df[bat_df['Nom'].str.strip() == name.strip()].iloc[0]
    return {
        'long': int(row['Longueur']),
        'larg': int(row['Largeur']),
        'type': row['Type'],
        'culture': float(row['Culture']) if pd.notna(row['Culture']) else 0,
        'ray': int(row['Rayonnement']) if pd.notna(row['Rayonnement']) else 0,
        'prod': row['Production'] if pd.notna(row['Production']) else "Rien",
        'qty': float(row['Quantite']) if pd.notna(row['Quantite']) else 0,
        'boost25': float(row['Boost 25%']) if pd.notna(row['Boost 25%']) else 0,
        'boost50': float(row['Boost 50%']) if pd.notna(row['Boost 50%']) else 0,
        'boost100': float(row['Boost 100%']) if pd.notna(row['Boost 100%']) else 0,
    }

# ===================== LANCEMENT =====================
if __name__ == "__main__":
    print("\nScript prêt à être étendu pour l'optimisation complète.")
    print("Pour l'instant, le fichier est bien chargé.")
    print(f"Fichier de sortie sera : {OUTPUT_FILE}")
