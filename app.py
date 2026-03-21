import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font

st.set_page_config(page_title="Fanlabiz Optimizer V6", layout="wide")

def load_data(file):
    xl = pd.ExcelFile(file)
    # On charge le terrain. On garde les NaN pour l'instant.
    df_terrain = xl.parse(0, header=None)
    df_bats = xl.parse(1)
    df_bats.columns = df_bats.columns.str.strip()
    return df_terrain, df_bats

def solve_layout(df_terrain, df_bats):
    # Conversion en matrice de chaînes de caractères pour simplifier
    grid = df_terrain.fillna('EMPTY').values.astype(str)
    rows, cols = grid.shape
    
    # Étape cruciale : Transformer les '1' et les cases vides ('EMPTY') en zones constructibles
    # mais UNIQUEMENT si elles sont dans le périmètre des X.
    logic_grid = np.full((rows, cols), 'OUT')
    
    # On trouve les limites du terrain (le rectangle des X)
    x_coords = np.argwhere(grid == 'X')
    if len(x_coords) > 0:
        min_r, min_c = x_coords.min(axis=0)
        max_r, max_c = x_coords.max(axis=0)
        for r in range(min_r, max_r + 1):
            for c in range(min_c, max_c + 1):
                val = grid[r, c]
                if val == 'X' or val == '0':
                    logic_grid[r, c] = val
                else:
                    logic_grid[r, c] = '1' # Case libre
    
    placed_buildings = []
    all_to_place = []
    for _, row in df_bats.iterrows():
        nb = int(row['Nombre']) if pd.notna(row['Nombre']) else 1
        for _ in range(nb):
            all_to_place.append(row.to_dict())

    # Tri : Neutres, puis Culturels (Taille), puis Producteurs (Guerison > Taille)
    neutres = [b for b in all_to_place if str(b['Type']).lower() == 'neutre']
    culturels = sorted([b for b in all_to_place if str(b['Type']).lower() == 'culturel'], 
                      key=lambda x: x['Longueur']*x['Largeur'], reverse=True)
    prio_prod = {"Guerison": 0, "Nourriture": 1, "Or": 2}
    producteurs = sorted([b for b in all_to_place if str(b['Type']).lower() == 'producteur'], 
                        key=lambda x: (prio_prod.get(x['Production'], 3), x['Longueur']*x['Largeur']), reverse=True)

    def can_fit(r, c, h, w):
        if r + h > rows or c + w > cols: return False
        region = logic_grid[r:r+h, c:c+w]
        return np.all(region == '1')

    # Placement
    for b in neutres + culturels + producteurs:
        placed = False
        for r in range(rows):
            for c in range(cols):
                # On teste les deux rotations
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and can_fit(r, c, h, w):
                        logic_grid[r:r+h, c:c+w] = 'BUSY'
                        b_copy = b.copy()
                        b_copy.update({'r': r, 'c': c, 'h': h, 'w': w})
                        placed_buildings.append(b_copy)
                        placed = True; break
                if placed: break
            if placed: break

    # Calcul Boosts
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
    
    # Styles
    f_c = PatternFill("FFA500","FFA500","solid") # Orange
    f_p = PatternFill("90EE90","90EE90","solid") # Vert
    f_n = PatternFill("D3D3D3","D3D3D3","solid") # Gris
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Onglet Terrain
    ws_t = wb.active
    ws_t.title = "Terrain"
    
    # On dessine la grille de base
    for r in range(grid.shape[0]):
        for c in range(grid.shape[1]):
            val = grid[r, c]
            if val != 'EMPTY':
                ws_t.cell(row=r+1, column=c+1, value=val)

    # On dessine les bâtiments
    for b in placed:
        color = f_c if b['Type'].lower() == 'culturel' else f_p if b['Type'].lower() == 'producteur' else f_n
        for rr in range(b['r'], b['r'] + b['h']):
            for cc in range(b['c'], b['c'] + b['w']):
                cell = ws_t.cell(row=rr+1, column=cc+1)
                cell.fill = color
                if rr == b['r'] and cc == b['c']:
                    cell.value = f"{b['Nom']}\n+{b.get('boost_atteint', 0)}%" if b['Type'].lower() == 'producteur' else b['Nom']
                cell.alignment = center

    # Onglet Synthèse (Point 5, 6, 7 de la demande)
    ws_s = wb.create_sheet("Resultats")
    ws_s.append(["Nom", "Type", "Production", "Culture", "Boost %", "Total/h"])
    if placed:
        for b in placed:
            ws_s.append([b['Nom'], b['Type'], b.get('Production',''), b.get('culture_recue',0), b.get('boost_atteint',0), b.get('prod_finale',0)])

    # Stats
    noms_placed = [b['Nom'] for b in placed]
    non_p = [b for b in total if b['Nom'] not in noms_placed]
    ws_s.append([]); ws_s.append(["STATISTIQUES"])
    ws_s.append(["Bâtiments non placés", len(non_p)])
    ws_s.append(["Surface non placée", sum(b['Longueur']*b['Largeur'] for b in non_p)])

    wb.save(output)
    return output.getvalue()

# --- UI STREAMLIT ---
uploaded = st.file_uploader("Fichier Ville.xlsx", type="xlsx")
if uploaded:
    df_t, df_b = load_data(uploaded)
    try:
        grid_vis, placed_list, total_list = solve_layout(df_t, df_b)
        st.success(f"Réussi : {len(placed_list)} bâtiments placés.")
        st.download_button("📥 Télécharger Resultat_V6.xlsx", generate_excel(grid_vis, placed_list, total_list), "Resultat_V6.xlsx")
    except Exception as e:
        st.error(f"Erreur lors du calcul : {e}")
