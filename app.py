import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font

# --- CONFIGURATION ---
st.set_page_config(page_title="Fanlabiz Optimizer", layout="wide")

def load_data(file):
    xl = pd.ExcelFile(file)
    # Lecture Terrain : on remplace les NaN par '1' (libre) et on nettoie les espaces
    df_terrain = xl.parse(0, header=None).fillna('1').astype(str)
    
    # Lecture Bâtiments
    df_bats = xl.parse(1)
    df_bats.columns = df_bats.columns.str.strip()
    return df_terrain, df_bats

def solve_layout(df_terrain, df_bats):
    grid = df_terrain.values
    rows, cols = grid.shape
    placed_buildings = []
    
    # Préparation
    all_to_place = []
    for _, row in df_bats.iterrows():
        try:
            nb = int(row['Nombre']) if pd.notna(row['Nombre']) else 1
            for _ in range(nb):
                all_to_place.append(row.to_dict())
        except: continue

    # Tri selon stratégie
    neutres = [b for b in all_to_place if str(b['Type']).lower() == 'neutre']
    culturels = sorted([b for b in all_to_place if str(b['Type']).lower() == 'culturel'], 
                      key=lambda x: x['Longueur']*x['Largeur'], reverse=True)
    
    prio_prod = {"Guerison": 0, "Nourriture": 1, "Or": 2}
    producteurs = sorted([b for b in all_to_place if str(b['Type']).lower() == 'producteur'], 
                        key=lambda x: (prio_prod.get(x['Production'], 3), x['Longueur']*x['Largeur']), reverse=True)

    def can_fit(r, c, h, w):
        if r + h > rows or c + w > cols: return False
        return np.all(grid[r:r+h, c:c+w] == '1')

    # 1. Placement Neutres (Bords)
    for b in neutres:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and can_fit(r, c, h, w):
                        # Check voisinage X
                        zone = grid[max(0,r-1):min(rows,r+h+1), max(0,c-1):min(cols,c+w+1)]
                        if 'X' in zone:
                            grid[r:r+h, c:c+w] = 'N'
                            b_copy = b.copy()
                            b_copy.update({'r': r, 'c': c, 'h': h, 'w': w})
                            placed_buildings.append(b_copy)
                            placed = True; break
                if placed: break
            if placed: break

    # 2. Placement Culturels & Producteurs
    for b in culturels + producteurs:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and can_fit(r, c, h, w):
                        grid[r:r+h, c:c+w] = 'C' if str(b['Type']).lower() == 'culturel' else 'P'
                        b_copy = b.copy()
                        b_copy.update({'r': r, 'c': c, 'h': h, 'w': w})
                        placed_buildings.append(b_copy)
                        placed = True; break
                if placed: break
            if placed: break

    # 3. Boosts
    for b in [pb for pb in placed_buildings if str(pb['Type']).lower() == 'producteur']:
        c_recue = 0
        for cult in [pc for pc in placed_buildings if str(pc['Type']).lower() == 'culturel']:
            rad = int(cult['Rayonnement'])
            # Zone rayonnement
            rs, re = max(0, cult['r']-rad), min(rows, cult['r']+cult['h']+rad)
            cs, ce = max(0, cult['c']-rad), min(cols, cult['c']+cult['w']+rad)
            if not (b['r'] >= re or b['r']+b['h'] <= rs or b['c'] >= ce or b['c']+b['w'] <= cs):
                c_recue += cult['Culture']
        
        b['culture_recue'] = c_recue
        boost = 0
        if pd.notna(b.get('Boost 100%')) and c_recue >= b['Boost 100%']: boost = 100
        elif pd.notna(b.get('Boost 50%')) and c_recue >= b['Boost 50%']: boost = 50
        elif pd.notna(b.get('Boost 25%')) and c_recue >= b['Boost 25%']: boost = 25
        b['boost_atteint'] = boost
        b['prod_finale'] = b['Quantite'] * (1 + boost/100) if pd.notna(b['Quantite']) else 0

    return grid, placed_buildings, all_to_place

# --- UI ---
st.title("🏗️ Fanlabiz Optimizer - Version Finale")

file = st.file_uploader("Charger Ville.xlsx", type="xlsx")

if file:
    df_t, df_b = load_data(file)
    grid_res, placed, total = solve_layout(df_t, df_b)
    
    # Métriques
    c1, c2, c3 = st.columns(3)
    c1.metric("Placés", f"{len(placed)} / {len(total)}")
    c2.metric("Cases Libres", np.sum(grid_res == '1'))
    
    # Génération Excel
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    # Onglet Terrain
    ws1 = wb.active
    ws1.title = "Terrain Final"
    f_c = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    f_p = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    f_n = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    for r in range(grid_res.shape[0]):
        for c in range(grid_res.shape[1]):
            val = grid_res[r,c]
            cell = ws1.cell(row=r+1, column=c+1, value=val if val in ['X','1','0'] else "")
            if val == 'C': cell.fill = f_c
            elif val == 'P': cell.fill = f_p
            elif val == 'N': cell.fill = f_n

    # Onglet Détails
    ws2 = wb.create_sheet("Synthèse")
    ws2.append(["Nom", "Type", "Production", "Culture reçue", "Boost (%)", "Prod/h Finale"])
    for b in placed:
        ws2.append([b['Nom'], b['Type'], b.get('Production',''), b.get('culture_recue',0), b.get('boost_atteint',0), b.get('prod_finale',0)])

    wb.save(output)
    st.download_button("📥 Télécharger Resultat_Ville.xlsx", output.getvalue(), "Resultat_Ville.xlsx")
