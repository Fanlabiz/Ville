import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment

# --- CONFIGURATION STREAMLIT ---
st.set_page_config(page_title="Optimiseur de Terrain", layout="wide")
st.title("🏗️ Optimiseur de Placement de Bâtiments")

def load_data(file):
    xl = pd.ExcelFile(file)
    # Onglet 1: Terrain (0=occupé, 1=libre, X=bord)
    df_terrain = xl.parse(0, header=None).fillna(' ')
    # Onglet 2: Bâtiments
    df_bats = xl.parse(1)
    # Onglet 3: Actuel (Optionnel pour cette V1)
    return df_terrain, df_bats

def check_placement(grid, r, c, h, w):
    """Vérifie si un bâtiment de taille h x w peut être placé à (r, c)"""
    if r + h > grid.shape[0] or c + w > grid.shape[1]:
        return False
    region = grid[r:r+h, c:c+w]
    return np.all(region == "1")

def get_radiation_zone(r, c, h, w, rad, grid_shape):
    """Définit les coordonnées de la zone de rayonnement entourant le bâtiment"""
    r_start = max(0, r - rad)
    r_end = min(grid_shape[0], r + h + rad)
    c_start = max(0, c - rad)
    c_end = min(grid_shape[1], c + w + rad)
    return r_start, r_end, c_start, c_end

def solve_layout(df_terrain, df_bats):
    grid = df_terrain.values.astype(str)
    rows, cols = grid.shape
    placed_buildings = []
    
    # 1. Préparation des listes par type
    all_bats = []
    for _, row in df_bats.iterrows():
        for _ in range(int(row['Nombre'])):
            all_bats.append(row.to_dict())

    neutres = sorted([b for b in all_bats if b['Type'] == 'Neutre'], 
                    key=lambda x: x['Longueur']*x['Largeur'], reverse=True)
    culturels = sorted([b for b in all_bats if b['Type'] == 'Culturel'], 
                      key=lambda x: x['Longueur']*x['Largeur'], reverse=True)
    
    # Priorité Production: Guérison > Nourriture > Or > Autres
    prio_map = {"Guerison": 0, "Nourriture": 1, "Or": 2}
    producteurs = sorted([b for b in all_bats if b['Type'] == 'Producteur'], 
                        key=lambda x: (prio_map.get(x['Production'], 3), x['Longueur']*x['Largeur']), reverse=True)

    # --- ÉTAPE 1: PLACEMENT DES NEUTRES (BORDS) ---
    for b in neutres:
        placed = False
        # On scanne les bords d'abord
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and check_placement(grid, r, c, h, w):
                        # Simple heuristique: est-ce proche d'un 'X'?
                        if "X" in grid[max(0,r-1):r+h+1, max(0,c-1):c+w+1]:
                            grid[r:r+h, c:c+w] = "N" # N pour Neutre
                            b_placed = b.copy()
                            b_placed.update({'r': r, 'c': c, 'h': h, 'w': w, 'final_boost': 0, 'culture_total': 0})
                            placed_buildings.append(b_placed)
                            placed = True
                            break
                if placed: break
            if placed: break

    # --- ÉTAPE 2: PLACEMENT CULTURELS & PRODUCTEURS (ALTERNÉ) ---
    # Pour cette version, on place les culturels au centre pour maximiser le rayonnement
    for b in culturels + producteurs:
        placed = False
        for r in range(rows):
            for c in range(cols):
                for h, w in [(b['Longueur'], b['Largeur']), (b['Largeur'], b['Longueur'])]:
                    if not placed and check_placement(grid, r, c, h, w):
                        symbol = "C" if b['Type'] == "Culturel" else "P"
                        grid[r:r+h, c:c+w] = symbol
                        b_placed = b.copy()
                        b_placed.update({'r': r, 'c': c, 'h': h, 'w': w})
                        placed_buildings.append(b_placed)
                        placed = True
                        break
                if placed: break
            if placed: break

    # --- ÉTAPE 3: CALCUL DES BOOSTS ---
    for b in placed_buildings:
        if b['Type'] == 'Producteur':
            total_cult = 0
            for c_bat in [pb for pb in placed_buildings if pb['Type'] == 'Culturel']:
                r_s, r_e, c_s, c_e = get_radiation_zone(c_bat['r'], c_bat['c'], c_bat['h'], c_bat['w'], c_bat['Rayonnement'], (rows, cols))
                # Si le producteur touche la zone de rayonnement
                if not (b['r'] >= r_e or b['r']+b['h'] <= r_s or b['c'] >= c_e or b['c']+b['w'] <= c_s):
                    total_cult += c_bat['Culture']
            
            b['culture_total'] = total_cult
            boost = 0
            if pd.notna(b['Boost 100%']) and total_cult >= b['Boost 100%']: boost = 100
            elif pd.notna(b['Boost 50%']) and total_cult >= b['Boost 50%']: boost = 50
            elif pd.notna(b['Boost 25%']) and total_cult >= b['Boost 25%']: boost = 25
            b['final_boost'] = boost

    return grid, placed_buildings, all_bats

def generate_excel(grid, placed_list, all_bats):
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    # Couleurs
    fill_cult = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # Orange
    fill_prod = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid") # Vert
    fill_neut = PatternFill(start_color="808080", end_color="808080", fill_type="solid") # Gris

    # 1. Onglet Terrain
    ws_map = wb.active
    ws_map.title = "Terrain Final"
    for r_idx, row in enumerate(grid, 1):
        for c_idx, val in enumerate(row, 1):
            cell = ws_map.cell(row=r_idx, column=c_idx, value=val)
            if val == "C": cell.fill = fill_cult
            elif val == "P": cell.fill = fill_prod
            elif val == "N": cell.fill = fill_neut

    # 2. Onglet Statistiques
    ws_stats = wb.create_sheet("Resultats")
    headers = ["Nom", "Type", "Production", "Coordonnées", "Culture Reçue", "Boost %", "Prod/Heure Final"]
    ws_stats.append(headers)
    
    prod_totale = {}
    for b in placed_list:
        p_val = b['Quantite'] * (1 + b.get('final_boost', 0)/100) if b['Type'] == 'Producteur' else 0
        ws_stats.append([b['Nom'], b['Type'], b.get('Production','-'), f"({b['r']},{b['c']})", b.get('culture_total',0), b.get('final_boost',0), p_val])
        
        if b['Type'] == 'Producteur':
            p_type = b['Production']
            prod_totale[p_type] = prod_totale.get(p_type, 0) + p_val

    # 3. Résumé par type
    ws_stats.append([])
    ws_stats.append(["Production Totale par Heure"])
    for k, v in prod_totale.items():
        ws_stats.append([k, v])

    wb.save(output)
    return output.getvalue()

# --- INTERFACE UTILISATEUR ---
uploaded_file = st.file_uploader("Chargez votre fichier Ville.xlsx", type="xlsx")

if uploaded_file:
    df_t, df_b = load_data(uploaded_file)
    grid_final, placed_list, all_original = solve_layout(df_t, df_b)
    
    st.success("Optimisation terminée !")
    
    # Métriques rapides
    c1, c2, c3 = st.columns(3)
    c1.metric("Bâtiments placés", len(placed_list))
    c2.metric("Cases libres", np.sum(grid_final == "1"))
    
    # Téléchargement
    excel_data = generate_excel(grid_final, placed_list, all_original)
    st.download_button(
        label="📥 Télécharger le résultat (Excel iPad)",
        data=excel_data,
        file_name="Resultat_Placement.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
