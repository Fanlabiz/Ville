import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font

st.set_page_config(page_title="Fanlabiz Optimizer V3", layout="wide")

def load_data(file):
    xl = pd.ExcelFile(file)
    df_terrain = xl.parse(0, header=None).fillna('1').astype(str)
    df_bats = xl.parse(1)
    df_bats.columns = df_bats.columns.str.strip()
    return df_terrain, df_bats

def solve_layout(df_terrain, df_bats):
    grid = df_terrain.values.copy()
    rows, cols = grid.shape
    placed_buildings = []
    
    # Préparation
    all_to_place = []
    for _, row in df_bats.iterrows():
        for _ in range(int(row['Nombre'] if pd.notna(row['Nombre']) else 1)):
            all_to_place.append(row.to_dict())

    # Séparation et Tri (Point 2 de la demande)
    neutres = [b for b in all_to_place if str(b['Type']).lower() == 'neutre']
    culturels = sorted([b for b in all_to_place if str(b['Type']).lower() == 'culturel'], 
                      key=lambda x: x['Longueur']*x['Largeur'], reverse=True)
    
    # Priorité Guérison (Point 1 de la demande)
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
                        zone = grid[max(0,r-1):min(rows,r+h+1), max(0,c-1):min(cols,c+w+1)]
                        if 'X' in zone:
                            grid[r:r+h, c:c+w] = b['Nom']
                            b.update({'r': r, 'c': c, 'h': h, 'w': w, 'status': 'Placé'})
                            placed_buildings.append(b); placed = True; break
                if placed: break
            if placed: break

    # 2. Placement Alterné (Culture/Production)
    # On essaie de placer un culturel, puis de remplir son rayonnement avec des producteurs
    to_place_cp = culturels + producteurs
    for b in to_place_cp:
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

# --- GENERATION EXCEL ---
def generate_excel(grid, placed, total):
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    # Styles
    fill_c = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # Orange
    fill_p = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid") # Vert
    fill_n = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid") # Gris
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1. FEUILLE TERRAIN (Point 4)
    ws_t = wb.active
    ws_t.title = "Terrain"
    for r in range(grid.shape[0]):
        for c in range(grid.shape[1]):
            val = grid[r,c]
            cell = ws_t.cell(row=r+1, column=c+1, value=val)
            if val == 'X': continue
            # Coloration et Texte
            b_info = next((item for item in placed if item['r'] == r and item['c'] == c), None)
            if b_info:
                cell.value = f"{b_info['Nom']}\n+{b_info.get('boost_atteint', 0)}%" if b_info['Type'] == 'Producteur' else b_info['Nom']
                cell.alignment = center
                if b_info['Type'] == 'Culturel': cell.fill = fill_c
                elif b_info['Type'] == 'Producteur': cell.fill = fill_p
                else: cell.fill = fill_n

    # 2. FEUILLE SYNTHESE (Points 1, 2, 3, 5, 6, 7)
    ws_s = wb.create_sheet("Resultats")
    
    # Liste des placés
    ws_s.append(["LISTE DES BATIMENTS PLACES"])
    ws_s.append(["Nom", "Type", "Production", "Coords", "Culture Reçue", "Boost %", "Prod Totale/h"])
    for b in placed:
        ws_s.append([b['Nom'], b['Type'], b.get('Production','-'), f"{b['r']},{b['c']}", b.get('culture_recue',0), b.get('boost_atteint',0), b.get('prod_finale',0)])

    # Synthèse par Production
    ws_s.append([]); ws_s.append(["SYNTHESE PAR PRODUCTION"])
    df_p = pd.DataFrame([b for b in placed if b['Type'] == 'Producteur'])
    if not df_p.empty:
        synth = df_p.groupby('Production').agg({'culture_recue': 'sum', 'prod_finale': 'sum'}).reset_index()
        for _, row in synth.iterrows():
            ws_s.append([row['Production'], "Culture Totale:", row['culture_recue'], "Prod/h Totale:", row['prod_finale']])

    # Non placés et Statistiques
    ws_s.append([]); ws_s.append(["STATISTIQUES ET NON PLACES"])
    noms_placed = [b['Nom'] for b in placed]
    non_p = [b for b in total if b['Nom'] not in noms_placed]
    ws_s.append(["Bâtiments non placés:", len(non_p)])
    ws_s.append(["Cases non utilisées:", np.sum(grid == '1')])
    ws_s.append(["Cases représentées par les non placés:", sum(b['Longueur']*b['Largeur'] for b in non_p)])
    
    for b in non_p:
        ws_s.append([b['Nom'], f"{b['Longueur']}x{b['Largeur']}", b['Type']])

    wb.save(output)
    return output.getvalue()

# --- STREAMLIT UI ---
st.title("🏗️ Fanlabiz Ville Optimizer - COMPLET")
file = st.file_uploader("Charger Ville.xlsx", type="xlsx")

if file:
    df_t, df_b = load_data(file)
    grid, placed, total = solve_layout(df_t, df_b)
    st.success("Optimisation terminée avec succès !")
    
    excel_out = generate_excel(grid, placed, total)
    st.download_button("📥 Télécharger le rapport complet", excel_out, "Resultat_Placement_Complet.xlsx")
