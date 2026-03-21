import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment

st.set_page_config(page_title="Fanlabiz Optimizer V7", layout="wide")

def load_data(file):
    xl = pd.ExcelFile(file)
    df_terrain = xl.parse(0, header=None)
    df_bats = xl.parse(1)
    df_bats.columns = df_bats.columns.str.strip()
    return df_terrain, df_bats

def solve_layout(df_terrain, df_bats):
    # Remplissage initial : tout ce qui est vide est 'OUT'
    raw_grid = df_terrain.fillna('OUT').values.astype(str)
    rows, cols = raw_grid.shape
    
    # Création de la grille logique de travail
    # On identifie le périmètre des X
    logic_grid = np.full((rows, cols), 'OUT')
    x_coords = np.argwhere(raw_grid == 'X')
    
    if len(x_coords) > 0:
        min_r, min_c = x_coords.min(axis=0)
        max_r, max_c = x_coords.max(axis=0)
        # On définit l'intérieur comme constructible ('1') sauf si c'est déjà un '0' ou 'X'
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                val = raw_grid[r, c]
                if val in ['X', '0']:
                    logic_grid[r, c] = val
                else:
                    logic_grid[r, c] = '1'

    placed_buildings = []
    all_to_place = []
    for _, row in df_bats.iterrows():
        qty = int(row['Nombre']) if pd.notna(row['Nombre']) else 1
        for _ in range(qty):
            all_to_place.append(row.to_dict())

    # Stratégie de tri (Guérison en priorité)
    neutres = [b for b in all_to_place if str(b['Type']).lower() == 'neutre']
    culturels = sorted([b for b in all_to_place if str(b['Type']).lower() == 'culturel'], 
                      key=lambda x: x['Longueur']*x['Largeur'], reverse=True)
    prio_prod = {"Guerison": 0, "Nourriture": 1, "Or": 2}
    producteurs = sorted([b for b in all_to_place if str(b['Type']).lower() == 'producteur'], 
                        key=lambda x: (prio_prod.get(x['Production'], 3), x['Longueur']*x['Largeur']), reverse=True)

    def can_fit(r, c, h, w):
        if r + h > rows or c + w > cols: return False
        return np.all(logic_grid[r:r+h, c:c+w] == '1')

    # Placement
    for b in neutres + culturels + producteurs:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and can_fit(r, c, h, w):
                        logic_grid[r:r+h, c:c+w] = 'BUSY'
                        b_copy = b.copy()
                        b_copy.update({'r': r, 'c': c, 'h': h, 'w': w, 'used_h': h, 'used_w': w})
                        placed_buildings.append(b_copy)
                        placed = True; break
                if placed: break
            if placed: break

    # Calcul des Boosts
    for b in [pb for pb in placed_buildings if str(pb['Type']).lower() == 'producteur']:
        c_recue = 0
        for cult in [pc for pc in placed_buildings if str(pc['Type']).lower() == 'culturel']:
            rad = int(cult['Rayonnement'])
            rs, re = max(0, cult['r']-rad), min(rows, cult['r']+cult['used_h']+rad)
            cs, ce = max(0, cult['c']-rad), min(cols, cult['c']+cult['used_w']+rad)
            if not (b['r'] >= re or b['r']+b['used_h'] <= rs or b['c'] >= ce or b['c']+b['used_w'] <= cs):
                c_recue += cult['Culture']
        
        b['culture_recue'] = c_recue
        boost = 0
        if pd.notna(b.get('Boost 100%')) and c_recue >= b['Boost 100%']: boost = 100
        elif pd.notna(b.get('Boost 50%')) and c_recue >= b['Boost 50%']: boost = 50
        elif pd.notna(b.get('Boost 25%')) and c_recue >= b['Boost 25%']: boost = 25
        b['boost_atteint'] = boost
        b['prod_finale'] = b['Quantite'] * (1 + boost/100) if pd.notna(b['Quantite']) else 0

    return raw_grid, placed_buildings, all_to_place

def generate_excel(raw_grid, placed, total):
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws_t = wb.active
    ws_t.title = "Terrain"
    
    # Styles de couleurs fixes (solid)
    fill_c = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    fill_p = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    fill_n = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1. Dessiner le fond (X et 0)
    for r in range(raw_grid.shape[0]):
        for c in range(raw_grid.shape[1]):
            val = raw_grid[r, c]
            if val != 'OUT':
                ws_t.cell(row=r+1, column=c+1, value=val if val in ['X', '0'] else "")

    # 2. Dessiner et colorer TOUTES les cases des bâtiments
    for b in placed:
        # Sélection de la couleur
        t = str(b['Type']).lower()
        color = fill_c if t == 'culturel' else fill_p if t == 'producteur' else fill_n
        
        for rr in range(b['r'], b['r'] + b['used_h']):
            for cc in range(b['c'], b['c'] + b['used_w']):
                cell = ws_t.cell(row=rr+1, column=cc+1)
                cell.fill = color
                cell.alignment = center
                # Texte uniquement sur la cellule supérieure gauche
                if rr == b['r'] and cc == b['c']:
                    txt = b['Nom']
                    if t == 'producteur':
                        txt += f"\n+{b.get('boost_atteint', 0)}%"
                    cell.value = txt

    # 3. Onglet Synthèse
    ws_s = wb.create_sheet("Resultats")
    ws_s.append(["Nom", "Type", "Production", "Culture", "Boost %", "Total/h"])
    for b in placed:
        ws_s.append([b['Nom'], b['Type'], b.get('Production',''), b.get('culture_recue',0), b.get('boost_atteint',0), b.get('prod_finale',0)])

    # Stats finales
    ws_s.append([])
    ws_s.append(["STATISTIQUES"])
    noms_placed = [b['Nom'] for b in placed]
    non_p = [b for b in total if b['Nom'] not in noms_placed]
    ws_s.append(["Bâtiments non placés", len(non_p)])
    ws_s.append(["Surface non placée", sum(b['Longueur']*b['Largeur'] for b in non_p)])

    wb.save(output)
    return output.getvalue()

# --- STREAMLIT ---
uploaded = st.file_uploader("Charger Ville.xlsx", type="xlsx")
if uploaded:
    df_t, df_b = load_data(uploaded)
    try:
        raw_grid, placed, total = solve_layout(df_t, df_b)
        st.success(f"Optimisation terminée : {len(placed)} bâtiments placés.")
        st.download_button("📥 Télécharger Resultat_V7.xlsx", generate_excel(raw_grid, placed, total), "Resultat_V7.xlsx")
    except Exception as e:
        st.error(f"Erreur : {e}")
