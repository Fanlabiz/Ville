import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font

# --- CONFIGURATION INTERFACE ---
st.set_page_config(page_title="Optimiseur de Ville - iPad", layout="wide")
st.title("🏗️ Générateur d'Aménagement de Terrain")

def check_collision(grid, r, c, h, w):
    """Vérifie si l'espace est libre (1) et dans les limites."""
    if r + h > grid.shape[0] or c + w > grid.shape[1]:
        return False
    return np.all(grid[r:r+h, c:c+w] == 1)

def get_radiance_zone(r, c, h, w, rad):
    [span_2](start_span)"""Définit les coordonnées de la zone de rayonnement[span_2](end_span)."""
    return (r - rad, c - rad, r + h + rad, c + w + rad)

def solve_layout(terrain_matrix, df_bat):
    # [span_3](start_span)Initialisation : X ou 0 = Bloqué, 1 = Libre[span_3](end_span)
    grid = np.where((terrain_matrix == 1), 1, 0)
    rows, cols = grid.shape
    
    placed = []
    unplaced = []
    
    # 1. [span_4](start_span)Stratégie : Séparation par Type[span_4](end_span)
    neutres = df_bat[df_bat['Type'] == 'Neutre'].copy()
    culturels = df_bat[df_bat['Type'] == 'Culturel'].copy()
    producteurs = df_bat[df_bat['Type'] == 'Producteur'].copy()
    
    # [span_5](start_span)Tri des producteurs par priorité : Guérison > Nourriture > Or[span_5](end_span)
    prio = {'Guerison': 1, 'Nourriture': 2, 'Or': 3}
    producteurs['prio_val'] = producteurs['Production'].map(lambda x: prio.get(x, 4))
    producteurs = producteurs.sort_values(['prio_val', 'Longueur'], ascending=[True, False])
    culturels = culturels.sort_values(['Longueur'], ascending=False)

    # Fonction interne de placement
    def try_place(row_data, is_edge=False):
        for r in range(rows):
            for c in range(cols):
                if is_edge and not (r <= 2 or r >= rows-3 or c <= 2 or c >= cols-3):
                    continue
                if check_collision(grid, r, c, row_data['Longueur'], row_data['Largeur']):
                    grid[r:r+int(row_data['Longueur']), c:c+int(row_data['Largeur'])] = 0
                    return (r, c)
        return None

    # [span_6](start_span)Placement 1 : Neutres sur les bords[span_6](end_span)
    for _, b in neutres.iterrows():
        for _ in range(int(b['Nombre'])):
            pos = try_place(b, is_edge=True)
            if pos: placed.append({**b.to_dict(), 'Row': pos[0], 'Col': pos[1]})
            else: unplaced.append(b.to_dict())

    # [span_7](start_span)Placement 2 : Alternance Culturel / Production[span_7](end_span)
    # On place d'abord les culturels pour créer des zones, puis les producteurs autour
    for _, b in pd.concat([culturels, producteurs]).iterrows():
        for _ in range(int(b['Nombre'])):
            pos = try_place(b)
            if pos: placed.append({**b.to_dict(), 'Row': pos[0], 'Col': pos[1]})
            else: unplaced.append(b.to_dict())

    df_placed = pd.DataFrame(placed)
    
    # 3. [span_8](start_span)Calcul de la culture reçue et des boosts[span_8](end_span)
    if not df_placed.empty and 'Type' in df_placed.columns:
        for i, row in df_placed.iterrows():
            if row['Type'] == 'Producteur':
                cult_totale = 0
                for _, cult in df_placed[df_placed['Type'] == 'Culturel'].iterrows():
                    # [span_9](start_span)Zone de rayonnement[span_9](end_span)
                    z = get_radiance_zone(cult['Row'], cult['Col'], cult['Longueur'], cult['Largeur'], cult['Rayonnement'])
                    # [span_10](start_span)Si intersection[span_10](end_span)
                    if not (row['Row'] >= z[2] or row['Row'] + row['Longueur'] <= z[0] or 
                            row['Col'] >= z[3] or row['Col'] + row['Largeur'] <= z[1]):
                        cult_totale += cult['Culture']
                
                df_placed.at[i, 'Culture reçue'] = cult_totale
                # [span_11](start_span)Calcul du Boost[span_11](end_span)
                boost = 0
                if cult_totale >= row['Boost 100%']: boost = 1.0
                elif cult_totale >= row['Boost 50%']: boost = 0.5
                elif cult_totale >= row['Boost 25%']: boost = 0.25
                
                df_placed.at[i, 'Boost'] = f"{int(boost*100)}%"
                df_placed.at[i, 'Prod réelle'] = row['Quantite'] * (1 + boost)
            else:
                df_placed.at[i, 'Culture reçue'] = 0
                df_placed.at[i, 'Boost'] = "0%"
                df_placed.at[i, 'Prod réelle'] = 0

    return df_placed, pd.DataFrame(unplaced), grid

def create_excel(placed, unplaced, final_grid):
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    # 1. [span_12](start_span)Onglet Résumé[span_12](end_span)
    ws_res = wb.active
    ws_res.title = "Résumé"
    stats = [
        ("Cases libres", np.sum(final_grid == 1)[span_13](start_span)),[span_13](end_span)
        ("Bâtiments non placés", len(unplaced)),
        ("Cases non placées", (unplaced['Longueur'] * unplaced['Largeur'])[span_14](start_span).sum() if not unplaced.empty else 0)[span_14](end_span)
    ]
    for r_idx, val in enumerate(stats, 1):
        ws_res.cell(row=r_idx, column=1, value=val[0])
        ws_res.cell(row=r_idx, column=2, value=val[1])

    # 2. [span_15](start_span)Onglet Production Totale[span_15](end_span)
    ws_prod = wb.create_sheet("Production totale")
    if not placed.empty:
        prod_sum = placed.groupby('Production')['Prod réelle'].sum().reset_index()
        for r_idx, row in enumerate(prod_sum.values, 1):
            ws_prod.cell(row=r_idx, column=1, value=row[0])
            ws_prod.cell(row=r_idx, column=2, value=row[1])

    # 3. [span_16](start_span)Onglet Terrain (Visuel avec couleurs)[span_16](end_span)
    ws_terr = wb.create_sheet("Terrain")
    fill_cult = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # Orange
    fill_prod = PatternFill(start_color="008000", end_color="008000", fill_type="solid") # Vert
    fill_neut = PatternFill(start_color="808080", end_color="808080", fill_type="solid") # Gris

    for _, b in placed.iterrows():
        color = fill_neut
        if b['Type'] == 'Culturel': color = fill_cult
        elif b['Type'] == 'Producteur': color = fill_prod
        
        for r in range(int(b['Longueur'])):
            for c in range(int(b['Largeur'])):
                cell = ws_terr.cell(row=int(b['Row'])+r+1, column=int(b['Col'])+c+1)
                if r == 0 and c == 0:
                    cell.value = f"{b['Nom']}\n{b['Boost']}"
                    cell.alignment = Alignment(wrapText=True, horizontal='center')
                cell.fill = color

    # 4. [span_17](start_span)Bâtiments Placés (Détails)[span_17](end_span)
    ws_list = wb.create_sheet("Bâtiments placés")
    # ... (Code standard export DataFrame vers Excel)

    wb.save(output)
    return output.getvalue()

# --- APP STREAMLIT ---
uploaded = st.file_uploader("Charger Ville.xlsx", type="xlsx")
if uploaded:
    with pd.ExcelFile(uploaded) as xls:
        df_terrain = pd.read_excel(xls, sheet_name=0, header=None)
        df_bat = pd.read_excel(xls, sheet_name=1)
    
    if st.button("Calculer l'agencement"):
        placed, unplaced, grid = solve_layout(df_terrain.values, df_bat)
        
        st.success("Traitement terminé")
        result_xlsx = create_excel(placed, unplaced, grid)
        st.download_button("Télécharger le Resultat.xlsx", result_xlsx, "Resultat.xlsx")
