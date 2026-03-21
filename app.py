import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font

st.set_page_config(page_title="Fanlabiz Optimizer V4", layout="wide")

def load_data(file):
    xl = pd.ExcelFile(file)
    # On charge le terrain tel quel pour garder les X
    df_terrain = xl.parse(0, header=None).fillna('1').astype(str)
    df_bats = xl.parse(1)
    df_bats.columns = df_bats.columns.str.strip()
    return df_terrain, df_bats

def solve_layout(df_terrain, df_bats):
    grid = df_terrain.values.copy()
    rows, cols = grid.shape
    placed_buildings = []
    
    all_to_place = []
    for _, row in df_bats.iterrows():
        for _ in range(int(row['Nombre'] if pd.notna(row['Nombre']) else 1)):
            all_to_place.append(row.to_dict())

    # Séparation et Tri
    neutres = [b for b in all_to_place if str(b['Type']).lower() == 'neutre']
    culturels = sorted([b for b in all_to_place if str(b['Type']).lower() == 'culturel'], 
                      key=lambda x: x['Longueur']*x['Largeur'], reverse=True)
    
    prio_prod = {"Guerison": 0, "Nourriture": 1, "Or": 2}
    producteurs = sorted([b for b in all_to_place if str(b['Type']).lower() == 'producteur'], 
                        key=lambda x: (prio_prod.get(x['Production'], 3), x['Longueur']*x['Largeur']), reverse=True)

    def can_fit(r, c, h, w):
        if r + h > rows or c + w > cols: return False
        # STRICT : On ne peut placer QUE sur des cases '1'
        region = grid[r:r+h, c:c+w]
        return np.all(region == '1')

    # 1. Placement Neutres (Bords mais sans écraser X)
    for b in neutres:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and can_fit(r, c, h, w):
                        # Doit être adjacent à un X
                        zone_autour = grid[max(0,r-1):min(rows,r+h+1), max(0,c-1):min(cols,c+w+1)]
                        if 'X' in zone_autour:
                            grid[r:r+h, c:c+w] = b['Nom']
                            b.update({'r': r, 'c': c, 'h': h, 'w': w, 'status': 'Placé'})
                            placed_buildings.append(b); placed = True; break
                if placed: break
            if placed: break

    # 2. Placement Culturels & Producteurs
    for b in culturels + producteurs:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and can_fit(r, c, h, w):
                        grid[r:r+h, c:c+w] = b['Nom']
                        b.update({'r': r, 'c': c, 'h': h, 'w': w, 'status': 'Placé'})
                        placed_buildings.append(b); placed = True; break
                if placed: break
            if placed: break

    # 3. Calcul des Boosts
    for b in [pb for pb in placed_buildings if str(pb['Type']).lower() == 'producteur']:
        c_recue = 0
        for cult in [pc for pc in placed_buildings if str(pc['Type']).lower() == 'culturel']:
            rad = int(cult['Rayonnement'])
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

def generate_excel(grid, placed, total):
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    fill_c = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # Orange
    fill_p = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid") # Vert
    fill_n = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid") # Gris
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws_t = wb.active
    ws_t.title = "Terrain"

    # Dictionnaire pour retrouver les infos du bâtiment par son nom à une coordonnée
    # (nécessaire pour la coloration de TOUTES les cases)
    lookup = {}
    for b in placed:
        for rr in range(b['r'], b['r'] + b['h']):
            for cc in range(b['c'], b['c'] + b['w']):
                lookup[(rr, cc)] = b

    for r in range(grid.shape[0]):
        for c in range(grid.shape[1]):
            val = grid[r,c]
            cell = ws_t.cell(row=r+1, column=c+1, value=val)
            
            if (r, c) in lookup:
                b_info = lookup[(r, c)]
                # Texte : Uniquement sur la case en haut à gauche pour ne pas surcharger
                if r == b_info['r'] and c == b_info['c']:
                    cell.value = f"{b_info['Nom']}\n+{b_info.get('boost_atteint', 0)}%" if b_info['Type'] == 'Producteur' else b_info['Nom']
                else:
                    cell.value = "" # Case vide mais colorée
                
                cell.alignment = center
                # COLORATION DE TOUTE LA ZONE
                if b_info['Type'] == 'Culturel': cell.fill = fill_c
                elif b_info['Type'] == 'Producteur': cell.fill = fill_p
                else: cell.fill = fill_n
            elif val == 'X':
                cell.value = 'X' # On garde le X visible

    # Onglet Synthèse
    ws_s = wb.create_sheet("Resultats")
    ws_s.append(["Nom", "Type", "Production", "Coords", "Culture Reçue", "Boost %", "Prod Totale/h"])
    for b in placed:
        ws_s.append([b['Nom'], b['Type'], b.get('Production','-'), f"{b['r']},{b['c']}", b.get('culture_recue',0), b.get('boost_atteint',0), b.get('prod_finale',0)])

    # Stats finales
    ws_s.append([]); ws_s.append(["STATISTIQUES"])
    noms_placed = [b['Nom'] for b in placed]
    non_p = [b for b in total if b['Nom'] not in noms_placed]
    ws_s.append(["Bâtiments non placés:", len(non_p)])
    ws_s.append(["Cases non utilisées (libres):", np.sum(grid == '1')])
    ws_s.append(["Surface totale non placée:", sum(b['Longueur']*b['Largeur'] for b in non_p)])

    wb.save(output)
    return output.getvalue()

# --- UI ---
st.title("🏗️ Fanlabiz Optimizer V4")
uploaded = st.file_uploader("Fichier Ville.xlsx", type="xlsx")

if uploaded:
    df_t, df_b = load_data(uploaded)
    grid, placed, total = solve_layout(df_t, df_b)
    st.success(f"Terminé ! {len(placed)} bâtiments placés.")
    
    st.download_button("📥 Télécharger le résultat corrigé", generate_excel(grid, placed, total), "Resultat_Final_V4.xlsx")
