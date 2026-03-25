import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment

# --- CONFIGURATION INTERFACE ---
st.set_page_config(page_title="Optimiseur de Ville", layout="wide")
st.title("🏗️ Optimiseur de Placement de Bâtiments")

def can_place(grid, r, c, h, w):
    """Vérifie si le bâtiment tient sur le terrain sans collision."""
    if r + h > grid.shape[0] or c + w > grid.shape[1]:
        return False
    return np.all(grid[r:r+h, c:c+w] == 1)

def mark_placement(grid, r, c, h, w):
    """Marque les cases comme occupées (0)."""
    grid[r:r+h, c:c+w] = 0

def is_in_radiance(prod_rect, cult_rect, radiation):
    """
    Vérifie si un producteur est dans la zone de rayonnement d'un culturel.
    La zone est une bande de largeur 'radiation' autour du culturel.
    """
    pr, pc, ph, pw = prod_rect
    cr, cc, ch, cw = cult_rect
    
    # Extension du rectangle culturel par le rayonnement
    ext_cr = cr - radiation
    ext_cc = cc - radiation
    ext_ch = ch + (2 * radiation)
    ext_cw = cw + (2 * radiation)
    
    # Vérification d'intersection standard entre deux rectangles
    return not (pr + ph <= ext_cr or pr >= ext_cr + ext_ch or 
                pc + pw <= ext_cc or pc >= ext_cc + ext_cw)

def solve_layout(terrain_init, df_buildings):
    rows, cols = terrain_init.shape
    # [span_3](start_span)1 = libre, 0 = occupé ou bordure X[span_3](end_span)
    grid = np.where(terrain_init == 1, 1, 0)
    
    placed_list = []
    unplaced_list = []
    
    # 1. [span_4](start_span)Séparation et Tri selon la Stratégie[span_4](end_span)
    neutres = df_buildings[df_buildings['Type'] == 'Neutre'].copy()
    culturels = df_buildings[df_buildings['Type'] == 'Culturel'].sort_values(['Longueur', 'Largeur'], ascending=False)
    
    # [span_5](start_span)Priorité : Guérison > Nourriture > Or > Autres[span_5](end_span)
    prio_map = {'Guerison': 1, 'Nourriture': 2, 'Or': 3}
    producteurs = df_buildings[df_buildings['Type'] == 'Producteur'].copy()
    producteurs['prio'] = producteurs['Production'].map(lambda x: prio_map.get(x, 4))
    producteurs = producteurs.sort_values(['prio', 'Longueur'], ascending=[True, False])

    # 2. [span_6](start_span)Placement des Neutres sur les bords[span_6](end_span)
    for _, b in neutres.iterrows():
        found = False
        for r in range(rows):
            for c in range(cols):
                # On privilégie les bords (r=0, r=max, c=0, c=max)
                if (r == 0 or r == rows-b['Longueur'] or c == 0 or c == cols-b['Largeur']):
                    if can_place(grid, r, c, b['Longueur'], b['Largeur']):
                        mark_placement(grid, r, c, b['Longueur'], b['Largeur'])
                        b_dict = b.to_dict()
                        b_dict.update({'Row': r, 'Col': c, 'Culture reçue': 0, 'Boost': '0%'})
                        placed_list.append(b_dict)
                        found = True; break
            if found: break
        if not found: unplaced_list.append(b.to_dict())

    # 3. [span_7](start_span)Placement Alterné Culturels / Producteurs[span_7](end_span)
    # Pour ce script, nous plaçons les culturels stratégiquement puis les producteurs
    all_to_place = pd.concat([culturels, producteurs])
    for _, b in all_to_place.iterrows():
        found = False
        for r in range(rows):
            for c in range(cols):
                if can_place(grid, r, c, b['Longueur'], b['Largeur']):
                    mark_placement(grid, r, c, b['Longueur'], b['Largeur'])
                    b_dict = b.to_dict()
                    b_dict.update({'Row': r, 'Col': c, 'Culture reçue': 0, 'Boost': '0%'})
                    placed_list.append(b_dict)
                    found = True; break
            if found: break
        if not found: unplaced_list.append(b.to_dict())

    # 4. [span_8](start_span)Calcul de la Culture et Boosts[span_8](end_span)
    for p in [x for x in placed_list if x['Type'] == 'Producteur']:
        total_cult = 0
        p_rect = (p['Row'], p['Col'], p['Longueur'], p['Largeur'])
        for c in [x for x in placed_list if x['Type'] == 'Culturel']:
            c_rect = (c['Row'], c['Col'], c['Longueur'], c['Largeur'])
            if is_in_radiance(p_rect, c_rect, c['Rayonnement']):
                total_cult += c['Culture']
        
        p['Culture reçue'] = total_cult
        # [span_9](start_span)Calcul du Boost atteint[span_9](end_span)
        if total_cult >= p['Boost 100%']: p['Boost'] = "100%"
        elif total_cult >= p['Boost 50%']: p['Boost'] = "50%"
        elif total_cult >= p['Boost 25%']: p['Boost'] = "25%"
        else: p['Boost'] = "0%"

    return pd.DataFrame(placed_list), pd.DataFrame(unplaced_list), grid

# -[span_10](start_span)-- EXPORT EXCEL COULEURS ---[span_10](end_span)
def get_excel_download(placed, unplaced, grid):
    output = BytesIO()
    wb = Workbook()
    
    # [span_11](start_span)Feuille 1 : Visuel du Terrain[span_11](end_span)
    ws1 = wb.active
    ws1.title = "Terrain Final"
    f_cult = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # Orange
    f_prod = PatternFill(start_color="008000", end_color="008000", fill_type="solid") # Vert
    f_neut = PatternFill(start_color="808080", end_color="808080", fill_type="solid") # Gris

    if not placed.empty:
        for _, row in placed.iterrows():
            fill = f_neut
            if row['Type'] == 'Culturel': fill = f_cult
            elif row['Type'] == 'Producteur': fill = f_prod
            
            for r in range(int(row['Longueur'])):
                for c in range(int(row['Largeur'])):
                    cell = ws1.cell(row=int(row['Row'])+r+1, column=int(row['Col'])+c+1)
                    cell.value = f"{row['Nom']} ({row['Boost']})"
                    cell.fill = fill
                    cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

    # [span_12](start_span)Feuille 2 : Détails techniques[span_12](end_span)
    ws2 = wb.create_sheet("Liste Bâtiments")
    # ... (code standard d'écriture du dataframe dans ws2)
    
    wb.save(output)
    return output.getvalue()

# --- STREAMLIT APP ---
uploaded_file = st.file_uploader("Fichier Ville.xlsx", type="xlsx")

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    terrain_df = pd.read_excel(xls, sheet_name=0, header=None)
    batiments_df = pd.read_excel(xls, sheet_name=1)
    
    if st.button("Lancer l'Optimisation"):
        # Nettoyage des données (NaN en 0 pour les boosts)
        batiments_df = batiments_df.fillna(0)
        
        placed, unplaced, final_grid = solve_layout(terrain_df.values, batiments_df)
        
        st.success("Calcul réussi !")
        st.metric("Cases non utilisées", int(np.sum(final_grid == 1)))
        
        # [span_13](start_span)Téléchargement[span_13](end_span)
        excel_data = get_excel_download(placed, unplaced, final_grid)
        st.download_button("📥 Télécharger le Résultat iPad", excel_data, "Resultat.xlsx")
        st.dataframe(placed[['Nom', 'Type', 'Production', 'Culture reçue', 'Boost']])
