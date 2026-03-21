import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment

st.set_page_config(page_title="Fanlabiz Optimizer V5", layout="wide")

def load_data(file):
    xl = pd.ExcelFile(file)
    # On remplace les cases vides par 'None' pour identifier le vrai hors-terrain
    df_terrain = xl.parse(0, header=None).fillna('OUT') 
    df_bats = xl.parse(1)
    df_bats.columns = df_bats.columns.str.strip()
    return df_terrain, df_bats

def solve_layout(df_terrain, df_bats):
    # Conversion en matrice. '1' = libre, '0'/'X' = occupé/bord, 'OUT' = hors limite
    grid = df_terrain.values.copy()
    rows, cols = grid.shape
    placed_buildings = []
    
    all_to_place = []
    for _, row in df_bats.iterrows():
        for _ in range(int(row['Nombre'] if pd.notna(row['Nombre']) else 1)):
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
        region = grid[r:r+h, c:c+w]
        # On ne peut poser QUE sur des cases qui contiennent exactement '1' ou 1
        # On exclut 'X', '0', 'OUT', et les cases vides
        for cell in region.flatten():
            if str(cell) != '1':
                return False
        return True

    # 1. Placement Neutres (doivent toucher un 'X')
    for b in neutres:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and can_fit(r, c, h, w):
                        zone_autour = grid[max(0,r-1):min(rows,r+h+1), max(0,c-1):min(cols,c+w+1)]
                        if 'X' in zone_autour.flatten():
                            grid[r:r+h, c:c+w] = 'OCCUPIED'
                            b.update({'r': r, 'c': c, 'h': h, 'w': w})
                            placed_buildings.append(b); placed = True; break
                if placed: break
            if placed: break

    # 2. Placement Culturels et Producteurs
    for b in culturels + producteurs:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and can_fit(r, c, h, w):
                        grid[r:r+h, c:c+w] = 'OCCUPIED'
                        b.update({'r': r, 'c': c, 'h': h, 'w': w})
                        placed_buildings.append(b); placed = True; break
                if placed: break
            if placed: break

    # 3. Calcul des Boosts (identique)
    for b in [pb for pb in placed_buildings if str(pb['Type']).lower() == 'producteur']:
        c_recue = 0
        for cult in [pc for pc in placed_buildings if str(pc['Type']).lower() == 'culturel']:
            rad = int(cult['Rayonnement'])
            rs, re, cs, ce = max(0, cult['r']-rad), min(rows, cult['r']+cult['h']+rad), max(0, cult['c']-rad), min(cols, cult['c']+cult['w']+rad)
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

def generate_excel(df_orig_terrain, placed, total):
    output = BytesIO()
    wb = openpyxl.Workbook()
    ws_t = wb.active
    ws_t.title = "Terrain"
    
    # Recréer le terrain de base avec les X et les 0
    for r in range(df_orig_terrain.shape[0]):
        for c in range(df_orig_terrain.shape[1]):
            val = df_orig_terrain.iloc[r, c]
            if val != 'OUT': # On ne remplit pas ce qui est hors terrain
                ws_t.cell(row=r+1, column=c+1, value=val if val in ['X', '0', 0] else "")

    # Styles
    f_c, f_p, f_n = PatternFill("FFA500","FFA500","solid"), PatternFill("90EE90","90EE90","solid"), PatternFill("D3D3D3","D3D3D3","solid")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Coloration et noms
    for b in placed:
        color = f_c if b['Type'] == 'Culturel' else f_p if b['Type'] == 'Producteur' else f_n
        for rr in range(b['r'], b['r'] + b['h']):
            for cc in range(b['c'], b['c'] + b['w']):
                cell = ws_t.cell(row=rr+1, column=cc+1)
                cell.fill = color
                if rr == b['r'] and cc == b['c']: # Nom uniquement sur la première case
                    cell.value = f"{b['Nom']}\n+{b.get('boost_atteint', 0)}%" if b['Type'] == 'Producteur' else b['Nom']
                cell.alignment = center

    # Onglet Résultats
    ws_s = wb.create_sheet("Synthèse")
    ws_s.append(["Nom", "Type", "Production", "Culture reçue", "Boost %", "Prod Totale/h"])
    for b in placed:
        ws_s.append([b['Nom'], b['Type'], b.get('Production','-'), b.get('culture_recue',0), b.get('boost_atteint',0), b.get('prod_finale',0)])

    wb.save(output)
    return output.getvalue()

# --- STREAMLIT ---
uploaded = st.file_uploader("Fichier Ville.xlsx", type="xlsx")
if uploaded:
    df_t, df_b = load_data(uploaded)
    grid_logic, placed_list, total_list = solve_layout(df_t, df_b)
    st.success(f"Optimisation terminée : {len(placed_list)} bâtiments placés.")
    st.download_button("📥 Télécharger Resultat_V5.xlsx", generate_excel(df_t, placed_list, total_list), "Resultat_V5.xlsx")
