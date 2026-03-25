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
    """Vérifie si l'espace est libre (1) et dans les limites du terrain."""
    if r + h > grid.shape[0] or c + w > grid.shape[1]:
        return False
    # On vérifie que toutes les cases du rectangle sont à 1 (libres)
    area = grid[r:r+h, c:c+w]
    return np.all(area == 1)

def solve_layout(terrain_matrix, df_bat):
    # Initialisation : On transforme le terrain en matrice numérique
    # 1 = libre, 0 = occupé (X ou déjà pris)
    grid = np.where(terrain_matrix == '1', 1, 0)
    rows, cols = grid.shape
    
    placed = []
    unplaced = []
    
    # 1. [span_2](start_span)Séparation par Type et Tri selon la Stratégie[span_2](end_span)
    neutres = df_bat[df_bat['Type'] == 'Neutre'].copy()
    culturels = df_bat[df_bat['Type'] == 'Culturel'].copy()
    producteurs = df_bat[df_bat['Type'] == 'Producteur'].copy()
    
    # [span_3](start_span)Priorité Production : Guérison > Nourriture > Or[span_3](end_span)
    prio_map = {'Guerison': 1, 'Nourriture': 2, 'Or': 3}
    producteurs['prio_val'] = producteurs['Production'].map(lambda x: prio_map.get(x, 4))
    
    # [span_4](start_span)Tri par taille décroissante[span_4](end_span)
    producteurs = producteurs.sort_values(['prio_val', 'Longueur'], ascending=[True, False])
    culturels = culturels.sort_values(['Longueur'], ascending=False)

    def try_place(row_data, is_edge=False):
        h, w = int(row_data['Longueur']), int(row_data['Largeur'])
        # [span_5](start_span)On teste les deux rotations possibles[span_5](end_span)
        for orientation in [(h, w), (w, h)]:
            curr_h, curr_w = orientation
            for r in range(rows - curr_h + 1):
                for c in range(cols - curr_w + 1):
                    # Si is_edge, on ne regarde que le périmètre
                    if is_edge and not (r == 0 or r == rows-curr_h or c == 0 or c == cols-curr_w):
                        continue
                    if check_collision(grid, r, c, curr_h, curr_w):
                        grid[r:r+curr_h, c:c+curr_w] = 0
                        return r, c, (curr_h, curr_w)
        return None

    # [span_6](start_span)Placement 1 : Neutres sur les bords[span_6](end_span)
    for _, b in neutres.iterrows():
        for _ in range(int(b['Nombre'])):
            res = try_place(b, is_edge=True)
            if res:
                placed.append({**b.to_dict(), 'Row': res[0], 'Col': res[1], 
                               'H_eff': res[2][0], 'W_eff': res[2][1]})
            else:
                unplaced.append(b.to_dict())

    # [span_7](start_span)Placement 2 : Culturels et Producteurs[span_7](end_span)
    for _, b in pd.concat([culturels, producteurs]).iterrows():
        for _ in range(int(b['Nombre'])):
            res = try_place(b)
            if res:
                placed.append({**b.to_dict(), 'Row': res[0], 'Col': res[1], 
                               'H_eff': res[2][0], 'W_eff': res[2][1]})
            else:
                unplaced.append(b.to_dict())

    df_placed = pd.DataFrame(placed)
    
    # 3. [span_8](start_span)Calcul de la culture reçue (Rayonnement)[span_8](end_span)
    if not df_placed.empty:
        for i, row in df_placed.iterrows():
            if row['Type'] == 'Producteur':
                total_c = 0
                for _, cult in df_placed[df_placed['Type'] == 'Culturel'].iterrows():
                    # [span_9](start_span)Zone de rayonnement = bande autour du bâtiment[span_9](end_span)
                    rad = cult['Rayonnement']
                    if not (row['Row'] >= cult['Row'] + cult['H_eff'] + rad or 
                            row['Row'] + row['H_eff'] <= cult['Row'] - rad or 
                            row['Col'] >= cult['Col'] + cult['W_eff'] + rad or 
                            row['Col'] + row['W_eff'] <= cult['Col'] - rad):
                        [span_10](start_span)total_c += cult['Culture'][span_10](end_span)
                
                df_placed.at[i, 'Culture reçue'] = total_c
                # [span_11](start_span)Calcul du Boost[span_11](end_span)
                boost = 0
                if total_c >= row['Boost 100%']: boost = 1.0
                elif total_c >= row['Boost 50%']: boost = 0.5
                elif total_c >= row['Boost 25%']: boost = 0.25
                
                df_placed.at[i, 'Boost'] = f"{int(boost*100)}%"
                df_placed.at[i, 'Prod_reelle'] = row['Quantite'] * (1 + boost)
            else:
                df_placed.at[i, 'Boost'] = "0%"
                df_placed.at[i, 'Prod_reelle'] = 0

    return df_placed, pd.DataFrame(unplaced), grid

def create_excel(placed, unplaced, final_grid):
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    # [span_12](start_span)Onglet Résumé[span_12](end_span)
    ws_res = wb.active
    ws_res.title = "Résumé"
    ws_res.append(["Indicateur", "Valeur"])
    ws_res.append(["Cases libres", int(np.sum(final_grid == 1))])
    ws_res.append(["Bât non placés", len(unplaced)])
    if not unplaced.empty:
        ws_res.append(["Cases non placées", int((unplaced['Longueur'] * unplaced['Largeur']).sum())])

    # [span_13](start_span)Onglet Terrain[span_13](end_span)
    ws_terr = wb.create_sheet("Terrain")
    f_cult = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # Orange
    f_prod = PatternFill(start_color="008000", end_color="008000", fill_type="solid") # Vert
    f_neut = PatternFill(start_color="808080", end_color="808080", fill_type="solid") # Gris

    for _, b in placed.iterrows():
        color = f_neut
        if b['Type'] == 'Culturel': color = f_cult
        elif b['Type'] == 'Producteur': color = f_prod
        
        for r in range(int(b['H_eff'])):
            for c in range(int(b['W_eff'])):
                cell = ws_terr.cell(row=int(b['Row'])+r+1, column=int(b['Col'])+c+1)
                if r == 0 and c == 0:
                    cell.value = f"{b['Nom']}\n{b['Boost']}"
                    cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
                cell.fill = color

    # [span_14](start_span)Onglet Production totale[span_14](end_span)
    ws_prod = wb.create_sheet("Production totale")
    ws_prod.append(["Ressource", "Prod totale /h"])
    if not placed.empty:
        summary = placed.groupby('Production')['Prod_reelle'].sum().reset_index()
        for _, r in summary.iterrows():
            ws_prod.append([r['Production'], r['Prod_reelle']])

    # [span_15](start_span)Onglet Bâtiments placés[span_15](end_span)
    ws_list = wb.create_sheet("Bâtiments placés")
    cols = ["Nom", "Type", "Production", "Row", "Col", "Culture reçue", "Boost", "Prod_reelle"]
    ws_list.append(cols)
    for _, b in placed.iterrows():
        ws_list.append([b.get(c, "") for c in cols])

    wb.save(output)
    return output.getvalue()

# --- STREAMLIT ---
uploaded = st.file_uploader("Fichier Ville.xlsx", type="xlsx")
if uploaded:
    xls = pd.ExcelFile(uploaded)
    # [span_16](start_span)On charge le terrain et les définitions[span_16](end_span)
    df_terr = pd.read_excel(xls, sheet_name=0, header=None).astype(str)
    df_bat = pd.read_excel(xls, sheet_name=1)
    
    if st.button("Lancer l'Optimisation"):
        placed, unplaced, grid = solve_layout(df_terr.values, df_bat)
        st.success("Calcul terminé !")
        
        xlsx_data = create_excel(placed, unplaced, grid)
        st.download_button("📥 Télécharger le Résultat", xlsx_data, "Resultat.xlsx")
