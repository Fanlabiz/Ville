import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font

# --- CONFIGURATION ---
st.set_page_config(page_title="Optimiseur de Ville Stratégique", layout="wide")
st.title("🏗️ Optimiseur de Terrain : Placement Stratégique")

def can_place(grid, r, c, h, w):
    if r + h > grid.shape[0] or c + w > grid.shape[1]:
        return False
    return np.all(grid[r:r+h, c:c+w] == 1)

def solve_layout(terrain_matrix, df_bat):
    # 1 = libre, 0 = occupé (X ou 0)
    grid = np.where((terrain_matrix == 1) | (terrain_matrix == "1"), 1, 0)
    rows, cols = grid.shape
    
    placed = []
    unplaced = []
    
    # Séparation des groupes
    neutres = df_bat[df_bat['Type'] == 'Neutre'].copy()
    culturels = df_bat[df_bat['Type'] == 'Culturel'].copy()
    producteurs = df_bat[df_bat['Type'] == 'Producteur'].copy()
    
    # Tri par taille décroissante pour optimiser l'espace (Obligation de placement)
    neutres = neutres.sort_values(['Longueur', 'Largeur'], ascending=False)
    culturels = culturels.sort_values(['Longueur', 'Largeur'], ascending=False)
    
    prio_map = {'Guerison': 1, 'Nourriture': 2, 'Or': 3}
    producteurs['prio'] = producteurs['Production'].map(lambda x: prio_map.get(x, 4))
    producteurs = producteurs.sort_values(['prio', 'Longueur'], ascending=[True, False])

    def try_place_logic(b_row, mode):
        h, w = int(b_row['Longueur']), int(b_row['Largeur'])
        rotations = [(h, w, "Non"), (w, h, "Oui")]
        
        # MODE 1 : NEUTRES -> Chercher UNIQUEMENT sur les bords d'abord
        if mode == 'neutre':
            for bh, bw, rot in rotations:
                for r in range(rows - bh + 1):
                    for c in range(cols - bw + 1):
                        # Contrainte : Doit toucher un bord (X)
                        if (r == 0 or r == rows - bh or c == 0 or c == cols - bw):
                            if can_place(grid, r, c, bh, bw):
                                grid[r:r+bh, c:c+bw] = 0
                                return r, c, bh, bw, rot
            # Si aucun bord n'est libre, on cherche ailleurs pour respecter l'obligation de placement
            mode = 'standard'

        # MODE 2 : CULTURELS -> Éviter les bords et l'accolement
        if mode == 'culturel':
            for bh, bw, rot in rotations:
                # On commence la recherche à partir de l'index 2 pour éviter les bords
                for r in range(min(2, rows-bh), rows - bh - 1):
                    for c in range(min(2, cols-bw), cols - bw - 1):
                        # Vérifier si un autre culturel est déjà à côté (distance de 1 case)
                        too_close = False
                        for p in placed:
                            if p['Type'] == 'Culturel':
                                if not (r > p['Row'] + p['H_eff'] + 1 or r + bh < p['Row'] - 1 or 
                                        c > p['Col'] + p['W_eff'] + 1 or c + bw < p['Col'] - 1):
                                    too_close = True
                                    break
                        if not too_close and can_place(grid, r, c, bh, bw):
                            grid[r:r+bh, c:c+bw] = 0
                            return r, c, bh, bw, rot
            mode = 'standard' # Si trop serré, on ignore la contrainte d'accolement pour placer

        # MODE 3 : STANDARD (Pour les producteurs ou en secours)
        if mode == 'standard':
            for bh, bw, rot in rotations:
                for r in range(rows - bh + 1):
                    for c in range(cols - bw + 1):
                        if can_place(grid, r, c, bh, bw):
                            grid[r:r+bh, c:c+bw] = 0
                            return r, c, bh, bw, rot
        return None

    # EXECUTION DU PLACEMENT DANS L'ORDRE STRATÉGIQUE
    # 1. Neutres d'abord (pour boucher les bords)
    for _, b in neutres.iterrows():
        for _ in range(int(b['Nombre'])):
            res = try_place_logic(b, 'neutre')
            if res: placed.append({**b.to_dict(), 'Row': res[0], 'Col': res[1], 'H_eff': res[2], 'W_eff': res[3], 'Rotation': res[4]})
            else: unplaced.append(b.to_dict())

    # 2. Culturels ensuite (au centre, espacés)
    for _, b in culturels.iterrows():
        for _ in range(int(b['Nombre'])):
            res = try_place_logic(b, 'culturel')
            if res: placed.append({**b.to_dict(), 'Row': res[0], 'Col': res[1], 'H_eff': res[2], 'W_eff': res[3], 'Rotation': res[4]})
            else: unplaced.append(b.to_dict())

    # 3. Producteurs (vont combler les espaces entre les culturels)
    for _, b in producteurs.iterrows():
        for _ in range(int(b['Nombre'])):
            res = try_place_logic(b, 'standard')
            if res: placed.append({**b.to_dict(), 'Row': res[0], 'Col': res[1], 'H_eff': res[2], 'W_eff': res[3], 'Rotation': res[4]})
            else: unplaced.append(b.to_dict())

    # CALCUL FINAL DES BOOSTS (Somme Cumulative)
    df_p = pd.DataFrame(placed)
    if not df_p.empty:
        for i, row in df_p.iterrows():
            if row['Type'] == 'Producteur':
                total_c = 0
                for _, c in df_p[df_p['Type'] == 'Culturel'].iterrows():
                    rad = c['Rayonnement']
                    if not (row['Row'] >= c['Row'] + c['H_eff'] + rad or row['Row'] + row['H_eff'] <= c['Row'] - rad or 
                            row['Col'] >= c['Col'] + c['W_eff'] + rad or row['Col'] + row['W_eff'] <= c['Col'] - rad):
                        total_c += c['Culture']
                df_p.at[i, 'Culture reçue'] = total_c
                # Logique Boost
                b25, b50, b100 = row['Boost 25%'], row['Boost 50%'], row['Boost 100%']
                pct = 1.0 if total_c >= b100 else 0.5 if total_c >= b50 else 0.25 if total_c >= b25 else 0
                df_p.at[i, 'Boost'] = f"{int(pct*100)}%"
                df_p.at[i, 'Prod_Heure'] = row['Quantite'] * (1 + pct)
            else:
                df_p.at[i, 'Boost'] = "0%"
                df_p.at[i, 'Prod_Heure'] = 0

    return df_p, pd.DataFrame(unplaced), grid

# --- FONCTION EXPORT EXCEL ---
def create_excel(placed, unplaced, final_grid):
    output = BytesIO()
    wb = openpyxl.Workbook()
    
    # Onglet Terrain
    ws_t = wb.active
    ws_t.title = "Terrain"
    fills = {
        'Culturel': PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid"),
        'Producteur': PatternFill(start_color="008000", end_color="008000", fill_type="solid"),
        'Neutre': PatternFill(start_color="808080", end_color="808080", fill_type="solid")
    }

    for _, b in placed.iterrows():
        f = fills.get(b['Type'])
        for r in range(int(b['H_eff'])):
            for c in range(int(b['W_eff'])):
                cell = ws_t.cell(row=int(b['Row'])+1, column=int(b['Col'])+1)
                cell = ws_t.cell(row=int(b['Row'])+r+1, column=int(b['Col'])+c+1)
                if r == 0 and c == 0:
                    cell.value = f"{b['Nom']}\n{b['Boost']}"
                    cell.alignment = Alignment(wrapText=True, horizontal='center', vertical='center')
                if f: cell.fill = f

    # Résumé et autres onglets (identiques à l'exemple)
    ws_r = wb.create_sheet("Résumé")
    ws_r.append(["Cases Libres", int(np.sum(final_grid == 1))])
    ws_r.append(["Bâtiments Non Placés", len(unplaced)])
    
    wb.save(output)
    return output.getvalue()

# --- STREAMLIT ---
uploaded = st.file_uploader("Charger Ville.xlsx", type="xlsx")
if uploaded:
    xls = pd.ExcelFile(uploaded)
    df_t = pd.read_excel(xls, sheet_name=0, header=None)
    df_b = pd.read_excel(xls, sheet_name=1)
    
    if st.button("Lancer l'Optimisation Stratégique"):
        p, u, g = solve_layout(df_t.values, df_b)
        st.success("Placement optimisé généré.")
        st.download_button("Télécharger Resultat.xlsx", create_excel(p, u, g), "Resultat.xlsx")
