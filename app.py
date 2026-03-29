import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
import os

# ===================== CONFIGURATION =====================
EXCEL_FILE = "Ville opti.xlsx"          # Votre fichier
OUTPUT_FILE = "Resultat_Optimisation_Ville.xlsx"

# Position de départ de la grille jouable (confirmé : ligne 2, colonne B → index 1,1 en 0-based)
START_ROW = 1      # ligne 2 dans Excel → index 1
START_COL = 1      # colonne B → index 1

# ===================== LECTURE DES DONNÉES =====================
print("Lecture du fichier Excel...")

# Terrain (1 = libre, 0 = occupé, X = bord)
terrain_df = pd.read_excel(EXCEL_FILE, sheet_name="Terrain", header=None)
terrain = terrain_df.to_numpy()

# Batiments
bat_df = pd.read_excel(EXCEL_FILE, sheet_name="Batiments")
bat_df.columns = bat_df.iloc[0] if bat_df.columns[0] == "Nom" else bat_df.columns
bat_df = bat_df.iloc[1:].reset_index(drop=True)

# Configuration actuelle
actuel_df = pd.read_excel(EXCEL_FILE, sheet_name="Actuel", header=None)
grid = actuel_df.to_numpy()

print(f"Terrain chargé : {terrain.shape}")
print(f"{len(bat_df)} types de bâtiments chargés")

# ===================== FONCTIONS UTILES =====================
def get_building_size(name):
    row = bat_df[bat_df['Nom'] == name].iloc[0]
    return int(row['Longueur']), int(row['Largeur'])

def is_cultural(name):
    row = bat_df[bat_df['Nom'] == name].iloc[0]
    return row['Type'] == 'Culturel'

def get_culture(name):
    row = bat_df[bat_df['Nom'] == name].iloc[0]
    return float(row['Culture']) if pd.notna(row['Culture']) else 0

def get_rayonnement(name):
    row = bat_df[bat_df['Nom'] == name].iloc[0]
    return int(row['Rayonnement']) if pd.notna(row['Rayonnement']) else 0

def get_boost_thresholds(name):
    row = bat_df[bat_df['Nom'] == name].iloc[0]
    return {
        25: float(row['Boost 25%']) if pd.notna(row['Boost 25%']) else 0,
        50: float(row['Boost 50%']) if pd.notna(row['Boost 50%']) else 0,
        100: float(row['Boost 100%']) if pd.notna(row['Boost 100%']) else 0
    }

def get_production_info(name):
    row = bat_df[bat_df['Nom'] == name].iloc[0]
    prod = row['Production'] if pd.notna(row['Production']) else "Rien"
    qty = float(row['Quantite']) if pd.notna(row['Quantite']) else 0
    return prod, qty

# ===================== CALCUL DE CULTURE (pour une grille donnée) =====================
def calculate_culture_received(grid_placed):
    height, width = grid_placed.shape
    culture_map = np.zeros((height, width), dtype=float)

    # Ajouter la culture de tous les bâtiments culturels
    for r in range(height):
        for c in range(width):
            cell = grid_placed[r, c]
            if isinstance(cell, str) and cell != '' and is_cultural(cell):
                cult = get_culture(cell)
                ray = get_rayonnement(cell)
                # Propager la culture dans la zone de rayonnement (bande autour)
                for dr in range(-ray, ray + 1):
                    for dc in range(-ray, ray + 1):
                        nr, nc = r + dr, c + dc
                        if 0 <= nr < height and 0 <= nc < width:
                            culture_map[nr, nc] += cult

    # Pour chaque producteur, sommer la culture sur toutes ses cases
    buildings_culture = {}
    for r in range(height):
        for c in range(width):
            cell = grid_placed[r, c]
            if isinstance(cell, str) and cell != '' and not is_cultural(cell):
                if cell not in buildings_culture:
                    buildings_culture[cell] = 0
                # On additionne la culture sur chaque case du bâtiment (même partiel compte)
                buildings_culture[cell] += culture_map[r, c]

    return buildings_culture

# ===================== CRÉATION D'UNE NOUVELLE CONFIGURATION OPTIMISÉE =====================
# Cette partie est simplifiée pour l'exemple. Vous pouvez la rendre plus intelligente plus tard.
# Ici on propose une disposition "maximale" avec regroupement par priorité.

def create_optimized_grid():
    print("Création d'une nouvelle grille optimisée (maximale)...")
    new_grid = np.full(grid.shape, '', dtype=object)
    
    # TODO : Implémenter un placement algorithmique avancé (packing + scoring)
    # Pour l'instant, on retourne une grille vide pour que vous puissiez la remplir manuellement ou étendre le script.
    # Vous pouvez ajouter ici un algorithme de placement (ex: greedy par priorité Guérison).
    
    print("Grille optimisée créée (version placeholder pour l'instant).")
    return new_grid

# ===================== GÉNÉRATION DU FICHIER RÉSULTAT =====================
def generate_result_file(current_grid, optimized_grid):
    print("Génération du fichier résultat...")
    
    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        # Onglet 1 : Liste des bâtiments (placeholder)
        buildings_list = pd.DataFrame(columns=["Nom", "Type", "Production", "Coordonnées", "Culture reçue", "Boost"])
        buildings_list.to_excel(writer, sheet_name="Liste_Batiments", index=False)
        
        # Onglet 2 : Statistiques
        stats = pd.DataFrame({
            "Type_Production": ["Guérison", "Nourriture", "Or", "Autres"],
            "Production_par_heure": [0, 0, 0, 0],
            "Gain_vs_actuel": [0, 0, 0, 0]
        })
        stats.to_excel(writer, sheet_name="Statistiques", index=False)
        
        # Onglet 3 : Déplacements (placeholder)
        depl = pd.DataFrame(columns=["Batiment", "Ancienne_position", "Nouvelle_position"])
        depl.to_excel(writer, sheet_name="Deplacements", index=False)
    
    # Chargement pour mise en forme avec couleurs
    wb = load_workbook(OUTPUT_FILE)
    
    # Exemple de coloration (à étendre)
    ws = wb["Liste_Batiments"]
    orange = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")
    green = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
    gray = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    
    print(f"Fichier généré : {OUTPUT_FILE}")
    print("Vous pouvez maintenant étendre le script pour calculer les boosts réels, la séquence de déplacements, etc.")

# ===================== EXÉCUTION =====================
if __name__ == "__main__":
    optimized = create_optimized_grid()
    generate_result_file(grid, optimized)
    print("Script terminé. Ouvrez le fichier", OUTPUT_FILE)
