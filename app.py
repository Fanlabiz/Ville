import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment

# --- CONFIGURATION ---
st.set_page_config(page_title="Optimiseur de Ville", layout="wide")
st.title("🏗️ Optimiseur de Placement de Bâtiments")

def can_place(grid, r, c, h, w):
    if r + h > grid.shape[0] or c + w > grid.shape[1]:
        return False
    return np.all(grid[r:r+h, c:c+w] == 1)

def solve_layout(terrain_matrix, df_bat):
    # Conversion du terrain : 1 = libre, 0 = occupé (X ou 0)
    grid = np.where(terrain_matrix == 1, 1, 0)
    rows, cols = grid.shape
    
    placed = []
    unplaced = []
    
    # [span_2](start_span)Séparation et Tri selon la Stratégie[span_2](end_span)
    neutres = df_bat[df_bat['Type'] == 'Neutre'].copy()
    culturels = df_bat[df_bat['Type'] == 'Culturel'].sort_values('Longueur', ascending=False)
    
    # [span_3](start_span)Priorité Producteurs : Guérison > Nourriture > Or[span_3](end_span)
    prio_map = {'Guerison': 1, 'Nourriture': 2, 'Or': 3}
    producteurs = df_bat[df_bat['Type'] == 'Producteur'].copy()
    producteurs['p_val'] = producteurs['Production'].map(lambda x: prio_map.get(x, 4))
    producteurs = producteurs.sort_values(['p_val', 'Longueur'], ascending=[True, False])

    def try_place_building(row_data, edge_only=False):
        h, w = int(row_data['Longueur']), int(row_data['Largeur'])
        # [span_4](start_span)Test des deux rotations[span_4](end_span)
        for bh, bw in [(h, w), (w, h)]:
            for r in range(rows - bh + 1):
                for c in range(cols - bw + 1):
                    if edge_only and not (r == 0 or r == rows-bh or c == 0 or c == cols-bw):
                        continue
                    if can_place(grid, r, c, bh, bw):
                        grid[r:r+bh, c:c+bw] = 0
                        return r, c, bh, bw, ("Oui" if bh == w else "Non")
        return None

    # 1. [span_5](start_span)Neutres sur les bords[span_5](end_span)
    for _, b in neutres.iterrows():
        for _ in range(int(b['Nombre'])):
            res = try_place_building(b, edge_only=True)
            if res:
                placed.append({**b.to_dict(), 'Row': res[0], 'Col': res[1], 'H_eff': res[2], 'W_eff': res[3], 'Rotation': res[4]})
            else:
                unplaced.append(b.to_dict())

    # 2. Culturels et Producteurs alternés
    for _, b in pd.concat([culturels, producteurs]).iterrows():
        for _ in range(int(b['Nombre'])):
            res = try_place_building(b)
            if res:
                placed.append({**b.to_dict(), 'Row': res[0], 'Col': res[1], 'H_eff': res[2], 'W_eff': res[3], 'Rotation': res[4]})
            else:
                unplaced.append(b.to_dict())

    df_placed = pd.DataFrame(placed)
    
    # 3. [span_6](start_span)Calcul de la culture et boosts[span_6](end_span)
    if not df_placed.empty:
        for i, row in df_placed.iterrows():
            if row['Type'] == 'Producteur':
                total_cult = 0
                for _, cult in df_placed[df_placed['Type'] == 'Culturel'].iterrows():
                    rad = cult['Rayonnement']
                    # [span_7](start_span)Vérification si le producteur est dans la zone de rayonnement[span_7](end_span)
                    if not (row['Row'] >= cult['Row'] + cult['H_eff'] + rad or 
                            row['Row'] + row['H_eff'] <= cult['Row'] - rad or 
                            row['Col'] >= cult['Col'] + cult['W_eff'] + rad or 
                            row['Col'] + row['W_eff'] <= cult['Col'] - rad):
                        total_cult += cult['Culture']
                
                df_placed.at[i, 'Culture reçue'] = total_cult
                boost = 0
                if total_cult >= row['Boost 100%']: boost = 1.0
                elif total_cult >= row['Boost 50%']: boost = 0.5
                elif total_cult >= row['Boost 25%']: boost = 0.25
                df_placed.at[i, 'Boost'] = f"{int(boost*100)}%"
                df_placed.at[i, 'Prod_reelle'] = row['Quantite'] * (1 + boost)
            else:
                df_placed.at[i, 'Culture reçue'] = 0
                df_placed.at[i, 'Boost'] = "0%"
                df_placed.at[i, 'Prod_reelle'] = 0

    return df_placed, pd.DataFrame(unplaced), grid

def create_excel(placed, unplaced, final_grid):
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    # [span_8](start_span)Onglet Résumé[span_8](end_span)
    ws_res = wb.active
    ws_res.title = "Résumé"
    ws_res.append(["Indicateur", "Valeur"])
    ws_res.append(["Cases libres", int(np.sum(final_grid == 1))])
    ws_res.append(["Bât non placés", len(unplaced)])
    ws_res.append(["Cases non placées", int((unplaced['Longueur'] * unplaced['Largeur']).sum()) if not unplaced.empty else 0])

    # [span_9](start_span)Onglet Terrain (Couleurs demandées)[span_9](end_span)
    ws_terr = wb.create_sheet("Terrain")
    colors = {
        'Culturel': PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid"), # Orange
        'Producteur': PatternFill(start_color="008000", end_color="008000", fill_type="solid"), # Vert
        'Neutre': PatternFill(start_color="808080", end_color="808080", fill_type="solid")     # Gris
    }

    for _, b in placed.iterrows():
        fill = colors.get(b['Type'], None)
        for r in range(int(b['H_eff'])):
            for c in range(int(b['W_eff'])):
                cell = ws_terr.cell(row=int(b['Row'])+r+1, column=int(b['Col'])+c+1)
                if r == 0 and c == 0:
                    cell.value = f"{b['Nom']} ({b['Boost']})"
                    cell.alignment = Alignment(wrapText=True, horizontal='center')
                if fill: cell.fill = fill

    # [span_10](start_span)Onglet Production totale[span_10](end_span)
    ws_prod = wb.create_sheet("Production totale")
    ws_prod.append(["Ressource", "Prod totale /h"])
    if not placed.empty:
        summary = placed.groupby('Production')['Prod_reelle'].sum().reset_index()
        for _, r in summary.iterrows():
            ws_prod.append([r['Production'], r['Prod_reelle']])

    # [span_11](start_span)Onglet Bâtiments placés[span_11](end_span)
    ws_list = wb.create_sheet("Bâtiments placés")
    cols = ["Nom", "Type", "Production", "Row", "Col", "Rotation", "Culture reçue", "Boost", "Prod_reelle"]
    ws_list.append(cols)
    for _, b in placed.iterrows():
        ws_list.append([b.get(c, "") for c in cols])

    wb.save(output)
    return output.getvalue()

# --- INTERFACE STREAMLIT ---
uploaded = st.file_uploader("Charger Ville.xlsx", type="xlsx")
if uploaded:
    xls = pd.ExcelFile(uploaded)
    df_terrain = pd.read_excel(xls, sheet_name="Terrain", header=None)
    df_bat = pd.read_excel(xls, sheet_name="Batiments")
    
    if st.button("Calculer l'agencement"):
        placed, unplaced, grid = solve_layout(df_terrain.values, df_bat)
        st.success("Traitement terminé")
        xlsx_data = create_excel(placed, unplaced, grid)
        st.download_button("📥 Télécharger Resultat.xlsx", xlsx_data, "Resultat.xlsx")
