import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill

# --- CONFIGURATION ---
st.set_page_config(page_title="Optimiseur de Ville", layout="wide")

def load_data(file):
    xl = pd.ExcelFile(file)
    # [span_6](start_span)Onglet 1 : Terrain[span_6](end_span)
    df_terrain = xl.parse(0, header=None).fillna('1')
    # [span_7](start_span)[span_8](start_span)Onglet 2 : Bâtiments[span_7](end_span)[span_8](end_span)
    df_bats = xl.parse(1)
    return df_terrain, df_bats

def check_placement(grid, r, c, h, w):
    if r + h > grid.shape[0] or c + w > grid.shape[1]:
        return False
    region = grid[r:r+h, c:c+w]
    return np.all(region == "1")

def solve_layout(df_terrain, df_bats):
    grid = df_terrain.values.astype(str)
    rows, cols = grid.shape
    placed_buildings = []
    
    # [span_9](start_span)Préparation des listes[span_9](end_span)
    all_to_place = []
    for _, row in df_bats.iterrows():
        # [span_10](start_span)[span_11](start_span)Utilisation de la colonne 'Nombre' pour la répétition[span_10](end_span)[span_11](end_span)
        nb = int(row['Nombre']) if pd.notna(row['Nombre']) else 1
        for _ in range(nb):
            all_to_place.append(row.to_dict())

    # [span_12](start_span)Séparation selon la stratégie[span_12](end_span)
    neutres = [b for b in all_to_place if b['Type'] == 'Neutre']
    culturels = sorted([b for b in all_to_place if b['Type'] == 'Culturel'], 
                      key=lambda x: x['Longueur']*x['Largeur'], reverse=True)
    
    prio_prod = {"Guerison": 0, "Nourriture": 1, "Or": 2}
    producteurs = sorted([b for b in all_to_place if b['Type'] == 'Producteur'], 
                        key=lambda x: (prio_prod.get(x['Production'], 3), x['Longueur']*x['Largeur']), reverse=True)

    # 1. [span_13](start_span)Placement des Neutres sur les bords[span_13](end_span)
    for b in neutres:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and check_placement(grid, r, c, h, w):
                        # [span_14](start_span)Proximité bord (X)[span_14](end_span)
                        if "X" in grid[max(0,r-1):r+h+1, max(0,c-1):c+w+1]:
                            grid[r:r+h, c:c+w] = "N"
                            b_copy = b.copy()
                            b_copy.update({'r': r, 'c': c, 'h': h, 'w': w, 'status': 'Placé'})
                            placed_buildings.append(b_copy)
                            placed = True; break
                if placed: break
            if placed: break

    # 2. [span_15](start_span)Placement alterné Culture / Production[span_15](end_span)
    for b in culturels + producteurs:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and check_placement(grid, r, c, h, w):
                        symbol = "C" if b['Type'] == "Culturel" else "P"
                        grid[r:r+h, c:c+w] = symbol
                        b_copy = b.copy()
                        b_copy.update({'r': r, 'c': c, 'h': h, 'w': w, 'status': 'Placé'})
                        placed_buildings.append(b_copy)
                        placed = True; break
                if placed: break
            if placed: break

    # 3. [span_16](start_span)Calcul de la culture et des boosts[span_16](end_span)
    for b in [pb for pb in placed_buildings if pb['Type'] == 'Producteur']:
        total_c = 0
        for cult in [pc for pc in placed_buildings if pc['Type'] == 'Culturel']:
            # [span_17](start_span)Zone rayonnement[span_17](end_span)
            rad = int(cult['Rayonnement'])
            r_s, r_e = max(0, cult['r']-rad), min(rows, cult['r']+cult['h']+rad)
            c_s, c_e = max(0, cult['c']-rad), min(cols, cult['c']+cult['w']+rad)
            
            # [span_18](start_span)Intersection[span_18](end_span)
            if not (b['r'] >= r_e or b['r']+b['h'] <= r_s or b['c'] >= c_e or b['c']+b['w'] <= c_s):
                total_c += cult['Culture']
        
        b['culture_recue'] = total_c
        # [span_19](start_span)Calcul boost[span_19](end_span)
        boost = 0
        if total_c >= b['Boost 100%']: boost = 1.0
        elif total_c >= b['Boost 50%']: boost = 0.5
        elif total_c >= b['Boost 25%']: boost = 0.25
        b['boost_atteint'] = f"{int(boost*100)}%"
        [span_20](start_span)b['prod_finale'] = b['Quantite'] * (1 + boost) # Quantite = prod/heure[span_20](end_span)

    return grid, placed_buildings, all_to_place

# --- INTERFACE STREAMLIT ---
st.title("🏙️ Fanlabiz Ville Optimizer")
file = st.file_uploader("Charger Ville.xlsx", type="xlsx")

if file:
    df_t, df_b = load_data(file)
    grid, placed, total_list = solve_layout(df_t, df_b)
    
    # [span_21](start_span)Résultats demandés[span_21](end_span)
    st.subheader("Statistiques")
    non_places = len(total_list) - len(placed)
    st.write(f"Bâtiments non placés : {non_places}")
    st.write(f"Cases libres : {np.sum(grid == '1')}")
    
    # [span_22](start_span)Génération Excel de sortie[span_22](end_span)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # [span_23](start_span)Onglet Terrain Visuel[span_23](end_span)
        pd.DataFrame(grid).to_excel(writer, sheet_name="Terrain_Final", index=False, header=False)
        # [span_24](start_span)Liste des bâtiments[span_24](end_span)
        pd.DataFrame(placed).to_excel(writer, sheet_name="Details_Places", index=False)
        
    st.download_button("📥 Télécharger le résultat", output.getvalue(), "Resultat.xlsx")
