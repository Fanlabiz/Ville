import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment

# --- CONFIGURATION STREAMLIT ---
st.set_page_config(page_title="Fanlabiz Optimizer", layout="wide")

def load_data(file):
    xl = pd.ExcelFile(file)
    # [span_2](start_span)Onglet 1: Terrain - On remplace les vides par '1' (libre)[span_2](end_span)
    df_terrain = xl.parse(0, header=None).fillna('1')
    
    # [span_3](start_span)Onglet 2: Bâtiments - Nettoyage des noms de colonnes (supprime les espaces)[span_3](end_span)
    df_bats = xl.parse(1)
    df_bats.columns = df_bats.columns.str.strip()
    
    return df_terrain, df_bats

def check_placement(grid, r, c, h, w):
    [span_4](start_span)[span_5](start_span)"""Vérifie si l'espace est libre (valeur '1')[span_4](end_span)[span_5](end_span)"""
    if r + h > grid.shape[0] or c + w > grid.shape[1]:
        return False
    region = grid[r:r+h, c:c+w]
    return np.all(region == "1")

def solve_layout(df_terrain, df_bats):
    grid = df_terrain.values.astype(str)
    rows, cols = grid.shape
    placed_buildings = []
    
    # [span_6](start_span)Préparation de la liste complète des bâtiments à placer[span_6](end_span)
    all_to_place = []
    for _, row in df_bats.iterrows():
        qty = int(row['Nombre']) if pd.notna(row['Nombre']) else 1
        for _ in range(qty):
            all_to_place.append(row.to_dict())

    # [span_7](start_span)Stratégie de tri[span_7](end_span)
    neutres = [b for b in all_to_place if str(b['Type']).lower() == 'neutre']
    culturels = sorted([b for b in all_to_place if str(b['Type']).lower() == 'culturel'], 
                      key=lambda x: x['Longueur']*x['Largeur'], reverse=True)
    
    prio_prod = {"Guerison": 0, "Nourriture": 1, "Or": 2}
    producteurs = sorted([b for b in all_to_place if str(b['Type']).lower() == 'producteur'], 
                        key=lambda x: (prio_prod.get(x['Production'], 3), x['Longueur']*x['Largeur']), reverse=True)

    # 1. [span_8](start_span)Placement des Neutres sur les bords[span_8](end_span)
    for b in neutres:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and check_placement(grid, r, c, h, w):
                        # [span_9](start_span)Vérification si touche un bord 'X'[span_9](end_span)
                        if "X" in grid[max(0,r-1):r+h+1, max(0,c-1):c+w+1]:
                            grid[r:r+h, c:c+w] = "N"
                            b_copy = b.copy()
                            b_copy.update({'r': r, 'c': c, 'h': h, 'w': w})
                            placed_buildings.append(b_copy)
                            placed = True; break
                if placed: break
            if placed: break

    # 2. [span_10](start_span)Placement alterné Culture / Production[span_10](end_span)
    for b in culturels + producteurs:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and check_placement(grid, r, c, h, w):
                        symbol = "C" if str(b['Type']).lower() == "culturel" else "P"
                        grid[r:r+h, c:c+w] = symbol
                        b_copy = b.copy()
                        b_copy.update({'r': r, 'c': c, 'h': h, 'w': w})
                        placed_buildings.append(b_copy)
                        placed = True; break
                if placed: break
            if placed: break

    # 3. [span_11](start_span)Calcul Culture et Boosts[span_11](end_span)
    for b in [pb for pb in placed_buildings if str(pb['Type']).lower() == 'producteur']:
        total_c = 0
        for cult in [pc for pc in placed_buildings if str(pc['Type']).lower() == 'culturel']:
            rad = int(cult['Rayonnement'])
            r_s, r_e = max(0, cult['r']-rad), min(rows, cult['r']+cult['h']+rad)
            c_s, c_e = max(0, cult['c']-rad), min(cols, cult['c']+cult['w']+rad)
            
            # [span_12](start_span)Si intersection entre producteur et zone rayonnement[span_12](end_span)
            if not (b['r'] >= r_e or b['r']+b['h'] <= r_s or b['c'] >= c_e or b['c']+b['w'] <= c_s):
                total_c += cult['Culture']
        
        b['culture_recue'] = total_c
        
        # [span_13](start_span)Calcul du boost avec sécurité si colonnes vides[span_13](end_span)
        boost = 0
        try:
            if pd.notna(b.get('Boost 100%')) and total_c >= b['Boost 100%']: boost = 1.0
            elif pd.notna(b.get('Boost 50%')) and total_c >= b['Boost 50%']: boost = 0.5
            elif pd.notna(b.get('Boost 25%')) and total_c >= b['Boost 25%']: boost = 0.25
        except: pass
        
        b['boost_final'] = f"{int(boost*100)}%"
        b['prod_h_finale'] = b['Quantite'] * (1 + boost) if pd.notna(b['Quantite']) else 0

    return grid, placed_buildings, all_to_place

# --- INTERFACE ---
st.title("🏗️ Fanlabiz Optimizer - V2")

uploaded_file = st.file_uploader("Choisissez le fichier Ville.xlsx", type="xlsx")

if uploaded_file:
    df_t, df_b = load_data(uploaded_file)
    grid_res, placed_res, total_list = solve_layout(df_t, df_b)
    
    # [span_14](start_span)Affichage des résultats[span_14](end_span)
    st.success(f"Placement terminé ! {len(placed_res)} bâtiments placés sur {len(total_list)}.")
    
    # [span_15](start_span)Préparation du fichier Excel de sortie[span_15](end_span)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. [span_16](start_span)Terrain Visuel[span_16](end_span)
        df_grid = pd.DataFrame(grid_res)
        df_grid.to_excel(writer, sheet_name="Terrain", index=False, header=False)
        
        # 2. [span_17](start_span)Liste détaillée[span_17](end_span)
        cols_final = ['Nom', 'Type', 'Production', 'culture_recue', 'boost_final', 'prod_h_finale']
        pd.DataFrame(placed_res)[cols_final].to_excel(writer, sheet_name="Details", index=False)
        
        # 3. [span_18](start_span)Bâtiments non placés[span_18](end_span)
        noms_places = [b['Nom'] for b in placed_res]
        non_places = [b for b in total_list if b['Nom'] not in noms_places]
        pd.DataFrame(non_places).to_excel(writer, sheet_name="Non_Places", index=False)

    st.download_button(
        label="📥 Télécharger le résultat pour iPad",
        data=output.getvalue(),
        file_name="Resultat_Ville.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
