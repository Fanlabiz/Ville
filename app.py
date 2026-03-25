import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font

# --- CONFIGURATION INTERFACE ---
st.set_page_config(page_title="Optimiseur de Terrain Pro", layout="wide")
st.title("🏗️ Moteur de Placement de Bâtiments")

def can_place(grid, r, c, h, w):
    """Vérifie la disponibilité spatiale."""
    if r + h > grid.shape[0] or c + w > grid.shape[1]:
        return False
    return np.all(grid[r:r+h, c:c+w] == 1)

def solve_layout(terrain_matrix, df_bat):
    # Prétraitement du terrain : 1 = libre, 0 = occupé
    # On gère les 'X' et les '0' comme obstacles
    grid = np.where((terrain_matrix == 1) | (terrain_matrix == "1"), 1, 0)
    rows, cols = grid.shape
    
    placed = []
    unplaced = []
    
    # --- PRÉPARATION DES DONNÉES (STRATÉGIE DEMANDE) ---
    # 1. Neutres (Gris)
    neutres = df_bat[df_bat['Type'] == 'Neutre'].copy()
    
    # 2. Producteurs Prioritaires (Guérison > Reste)
    prods = df_bat[df_bat['Type'] == 'Producteur'].copy()
    prio_map = {'Guerison': 1, 'Nourriture': 2, 'Or': 3}
    prods['prio'] = prods['Production'].map(lambda x: prio_map.get(x, 4))
    
    # 3. Culturels (Orange)
    cults = df_bat[df_bat['Type'] == 'Culturel'].copy()

    # Tri global par taille décroissante (Surface) pour maximiser le remplissage
    def get_sorted_list(df):
        df['surface'] = df['Longueur'] * df['Largeur']
        return df.sort_values(['surface'], ascending=False)

    # Fonction de tentative de placement intensive
    def try_place(b_row, edge_only=False):
        h, w = int(b_row['Longueur']), int(b_row['Largeur'])
        # On teste les deux sens : Normal et Pivoté (Rotation)
        for bh, bw, rot in [(h, w, "Non"), (w, h, "Oui")]:
            for r in range(rows - bh + 1):
                for c in range(cols - bw + 1):
                    if edge_only:
                        # On vérifie si on touche un bord du terrain
                        if not (r == 0 or r == rows - bh or c == 0 or c == cols - bw):
                            continue
                    if can_place(grid, r, c, bh, bw):
                        grid[r:r+bh, c:c+bw] = 0
                        return r, c, bh, bw, rot
        return None

    # --- ÉTAPE 1 : PLACEMENT DES NEUTRES SUR LES BORDS ---
    for _, b in get_sorted_list(neutres).iterrows():
        for _ in range(int(b['Nombre'])):
            res = try_place(b, edge_only=True)
            if not res: res = try_place(b, edge_only=False) # Si bord plein, n'importe où
            if res:
                placed.append({**b.to_dict(), 'Row': res[0], 'Col': res[1], 'H_eff': res[2], 'W_eff': res[3], 'Rotation': res[4]})
            else:
                unplaced.append(b.to_dict())

    # --- ÉTAPE 2 : PLACEMENT ALTERNÉ (STRATÉGIE) ---
    # On combine Culturels et Producteurs par ordre de priorité de taille
    others = pd.concat([get_sorted_list(prods), get_sorted_list(cults)])
    others = others.sort_values(['prio', 'surface'] if 'prio' in others.columns else ['surface'], ascending=[True, False])

    for _, b in others.iterrows():
        for _ in range(int(b['Nombre'])):
            res = try_place(b)
            if res:
                placed.append({**b.to_dict(), 'Row': res[0], 'Col': res[1], 'H_eff': res[2], 'W_eff': res[3], 'Rotation': res[4]})
            else:
                unplaced.append(b.to_dict())

    df_placed = pd.DataFrame(placed)
    
    # --- ÉTAPE 3 : CALCUL CALCUL DU RAYONNEMENT (SOMME) ---
    if not df_placed.empty:
        for i, row in df_placed.iterrows():
            if row['Type'] == 'Producteur':
                total_c = 0
                for _, c in df_placed[df_placed['Type'] == 'Culturel'].iterrows():
                    rad = c['Rayonnement']
                    # Vérification d'intersection avec la zone de rayonnement (bande)
                    if not (row['Row'] >= c['Row'] + c['H_eff'] + rad or 
                            row['Row'] + row['H_eff'] <= c['Row'] - rad or 
                            row['Col'] >= c['Col'] + c['W_eff'] + rad or 
                            row['Col'] + row['W_eff'] <= c['Col'] - rad):
                        total_c += c['Culture']
                
                df_placed.at[i, 'Culture reçue'] = total_c
                # Calcul du Boost
                boost_pct = 0
                if total_c >= row['Boost 100%']: boost_pct = 1.0
                elif total_c >= row['Boost 50%']: boost_pct = 0.5
                elif total_c >= row['Boost 25%']: boost_pct = 0.25
                
                df_placed.at[i, 'Boost'] = f"{int(boost_pct*100)}%"
                df_placed.at[i, 'Prod_Heure'] = row['Quantite'] * (1 + boost_pct)
            else:
                df_placed.at[i, 'Culture reçue'] = 0
                df_placed.at[i, 'Boost'] = "0%"
                df_placed.at[i, 'Prod_Heure'] = 0

    return df_placed, pd.DataFrame(unplaced), grid

def create_excel(placed, unplaced, final_grid):
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    # 1. Résumé
    ws = wb.active
    ws.title = "Résumé"
    ws.append(["Indicateur", "Valeur"])
    ws.append(["Cases vides sur terrain", int(np.sum(final_grid == 1))])
    ws.append(["Bâtiments non placés", len(unplaced)])
    ws.append(["Surface totale perdue (non placés)", int((unplaced['Longueur'] * unplaced['Largeur']).sum()) if not unplaced.empty else 0])

    # 2. Terrain Visuel
    ws_t = wb.create_sheet("Terrain")
    # Couleurs conformes à la demande
    fills = {
        'Culturel': PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid"), # Orange
        'Producteur': PatternFill(start_color="008000", end_color="008000", fill_type="solid"), # Vert
        'Neutre': PatternFill(start_color="808080", end_color="808080", fill_type="solid")      # Gris
    }

    for _, b in placed.iterrows():
        f = fills.get(b['Type'])
        for r in range(int(b['H_eff'])):
            for c in range(int(b['W_eff'])):
                cell = ws_t.cell(row=int(b['Row'])+r+1, column=int(b['Col'])+c+1)
                if r == 0 and c == 0:
                    cell.value = f"{b['Nom']}\n{b['Boost']}"
                    cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
                    cell.font = Font(size=8)
                if f: cell.fill = f

    # 3. Production Totale
    ws_p = wb.create_sheet("Production totale")
    ws_p.append(["Ressource", "Quantité Totale /h"])
    if not placed.empty:
        res = placed.groupby('Production')['Prod_Heure'].sum().reset_index()
        for _, row in res.iterrows():
            ws_p.append([row['Production'], row['Prod_Heure']])

    # 4. Détails Placés
    ws_d = wb.create_sheet("Bâtiments placés")
    cols = ["Nom", "Type", "Production", "Row", "Col", "Rotation", "Culture reçue", "Boost", "Prod_Heure"]
    ws_d.append(cols)
    for _, b in placed.iterrows():
        ws_d.append([b.get(c, "") for c in cols])

    wb.save(output)
    return output.getvalue()

# --- APP STREAMLIT ---
uploaded = st.file_uploader("📁 Charger Ville.xlsx (depuis iPad)", type="xlsx")

if uploaded:
    xls = pd.ExcelFile(uploaded)
    df_t = pd.read_excel(xls, sheet_name=0, header=None)
    df_b = pd.read_excel(xls, sheet_name=1)
    
    if st.button("🚀 Lancer l'Aménagement Obligatoire"):
        with st.spinner("Calcul des combinaisons..."):
            placed, unplaced, grid = solve_layout(df_t.values, df_b)
            
            st.success(f"Terminé ! {len(placed)} bâtiments placés.")
            if not unplaced.empty:
                st.warning(f"Attention : {len(unplaced)} bâtiments n'ont pas pu entrer.")
            
            data = create_excel(placed, unplaced, grid)
            st.download_button("📥 Télécharger Resultat.xlsx", data, "Resultat.xlsx")
