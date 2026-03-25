import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment

# --- CONFIGURATION ET INTERFACE ---
st.set_page_config(page_title="Optimiseur de Terrain", layout="wide")
st.title("🏗️ Planificateur de Terrain Optimisé")

def solve_layout(terrain_init, df_buildings, deja_places):
    rows, cols = terrain_init.shape
    # 0 = occupé ou bordure X, 1 = libre
    grid = np.where(terrain_init == 1, 1, 0) 
    
    # Marquer les bâtiments déjà placés (3ème onglet)
    for index, row in deja_places.iterrows():
        # Logique de marquage ici
        pass

    placed_list = []
    unplaced_list = []
    
    # [span_1](start_span)Séparation et Tri selon la Stratégie[span_1](end_span)
    neutres = df_buildings[df_buildings['Type'] == 'Neutre']
    culturels = df_buildings[df_buildings['Type'] == 'Culturel'].sort_values(['Longueur', 'Largeur'], ascending=False)
    
    # [span_2](start_span)Priorité Producteurs : Guérison > Nourriture > Or > Autres[span_2](end_span)
    prod_priority = {'Guerison': 1, 'Nourriture': 2, 'Or': 3}
    producteurs = df_buildings[df_buildings['Type'] == 'Producteur'].copy()
    producteurs['rank'] = producteurs['Production'].map(lambda x: prod_priority.get(x, 4))
    producteurs = producteurs.sort_values(['rank', 'Longueur'], ascending=[True, False])

    # 1. [span_3](start_span)Placement des Neutres sur les bords[span_3](end_span)
    for _, b in neutres.iterrows():
        for _ in range(int(b['Nombre'])):
            placed = try_place_on_perimeter(grid, b, placed_list, rows, cols)
            if not placed: unplaced_list.append(b.to_dict())

    # 2. [span_4](start_span)Placement Alterné Culturels / Producteurs[span_4](end_span)
    # On itère sur les producteurs prioritaires et on cherche à les entourer de culture
    for _, p in producteurs.iterrows():
        for _ in range(int(p['Nombre'])):
            # [span_5](start_span)Chercher une place optimale pour maximiser le rayonnement[span_5](end_span)
            placed = find_best_spot(grid, p, placed_list, rows, cols)
            if not placed: unplaced_list.append(p.to_dict())

    # 3. [span_6](start_span)Calcul de la Culture et Boosts[span_6](end_span)
    final_placed = calculate_boosts(placed_list, rows, cols)
    
    return final_placed, pd.DataFrame(unplaced_list), grid

def calculate_boosts(placed_list, rows, cols):
    # Pour chaque producteur, on calcule la somme de culture des bâtiments culturels 
    # [span_7](start_span)[span_8](start_span)dont la zone de rayonnement (bande autour) le touche[span_7](end_span)[span_8](end_span)
    for b in placed_list:
        if b['Type'] == 'Producteur':
            culture_totale = 0
            for c in placed_list:
                if c['Type'] == 'Culturel':
                    if is_in_radiance(b, c):
                        culture_totale += c['Culture']
            b['Culture reçue'] = culture_totale
            # [span_9](start_span)Détermination du palier de boost[span_9](end_span)
            b['Boost'] = "0%"
            if culture_totale >= b['Boost 100%']: b['Boost'] = "100%"
            elif culture_totale >= b['Boost 50%']: b['Boost'] = "50%"
            elif culture_totale >= b['Boost 25%']: b['Boost'] = "25%"
    return pd.DataFrame(placed_list)

# -[span_10](start_span)-- FONCTION EXCEL AVEC COULEURS[span_10](end_span) ---
def generate_excel_output(placed_df, unplaced_df, final_grid):
    output = BytesIO()
    wb = Workbook()
    
    # [span_11](start_span)Onglet 4: Terrain Visuel[span_11](end_span)
    ws_visuel = wb.active
    ws_visuel.title = "Terrain et Bâtiments"
    
    # [span_12](start_span)Fills pour les couleurs demandées[span_12](end_span)
    fill_culture = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid") # Orange
    fill_prod = PatternFill(start_color="008000", end_color="008000", fill_type="solid")    # Vert
    fill_neutre = PatternFill(start_color="808080", end_color="808080", fill_type="solid")  # Gris

    for _, row in placed_df.iterrows():
        f = fill_neutre
        if row['Type'] == 'Culturel': f = fill_culture
        elif row['Type'] == 'Producteur': f = fill_prod
        
        # [span_13](start_span)Coloration des cellules sur le terrain[span_13](end_span)
        for r in range(row['Row'], row['Row'] + row['Longueur']):
            for c in range(row['Col'], row['Col'] + row['Largeur']):
                cell = ws_visuel.cell(row=r+1, column=c+1, value=f"{row['Nom']} ({row['Boost']})")
                cell.fill = f
    
    # [span_14](start_span)Autres onglets (Stats, Listes)[span_14](end_span)
    # ... (Code pour les autres onglets)
    
    wb.save(output)
    return output.getvalue()

# --- MAIN STREAMLIT ---
file = st.file_uploader("Charger le fichier Ville.xlsx (iPad)", type="xlsx")
if file:
    # [span_15](start_span)[span_16](start_span)Lecture des 3 onglets[span_15](end_span)[span_16](end_span)
    xls = pd.ExcelFile(file)
    terrain_in = pd.read_excel(xls, sheet_name=0, header=None)
    batiments_in = pd.read_excel(xls, sheet_name=1)
    deja_places_in = pd.read_excel(xls, sheet_name=2)

    if st.button("Calculer l'agencement optimal"):
        placed, unplaced, grid = solve_layout(terrain_in.values, batiments_in, deja_places_in)
        
        # [span_17](start_span)Calcul des indicateurs finaux[span_17](end_span)
        cases_vides = np.sum(grid == 1)
        surf_non_placee = unplaced['Longueur'] * unplaced['Largeur'] if not unplaced.empty else 0
        
        [span_18](start_span)st.write(f"Nombre de cases non utilisées : {cases_vides}")[span_18](end_span)
        
        excel_file = generate_excel_output(placed, unplaced, grid)
        st.download_button("Télécharger le Résultat", excel_file, "Resultat_Placement.xlsx")
