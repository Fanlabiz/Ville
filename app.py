# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font

st.set_page_config(page_title="Optimiseur de Ville", layout="wide")
st.title("🚀 Optimiseur de placement de bâtiments")

uploaded_file = st.file_uploader("Choisir ton fichier Excel (Ville.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.info("Charge ton fichier pour commencer.")
    st.stop()

try:
    xls = pd.ExcelFile(uploaded_file)
    
    df_terrain = pd.read_excel(xls, sheet_name="Terrain", header=None)
    df_bat     = pd.read_excel(xls, sheet_name="Batiments")
    df_actuel  = pd.read_excel(xls, sheet_name="Actuel", header=None)  # non utilisé pour l'instant

    # ====================== CHARGEMENT TERRAIN ======================
    grid = df_terrain.fillna('').replace('X', -1).replace('', 0).astype(int).values
    # Nettoyage des colonnes/lignes vides
    while len(grid) > 0 and np.all(grid[-1] == 0): grid = grid[:-1]
    while grid.shape[1] > 0 and np.all(grid[:, -1] == 0): grid = np.delete(grid, -1, axis=1)
    
    rows, cols = grid.shape
    st.success(f"Terrain chargé : {rows} × {cols} cases")

    # ====================== PRÉPARATION BÂTIMENTS ======================
    batiments = []
    for _, row in df_bat.iterrows():
        if pd.isna(row.get('Nom')): continue
        b = row.to_dict()
        b['Longueur'] = int(b['Longueur'])
        b['Largeur']  = int(b['Largeur'])
        b['Nombre']   = int(b['Nombre'])
        b['placed']   = False
        for _ in range(b['Nombre']):
            batiments.append(b.copy())

    # Séparation par type
    neutres      = [b for b in batiments if b.get('Type') == 'Neutre']
    culturels    = [b for b in batiments if b.get('Type') == 'Culturel']
    producteurs  = [b for b in batiments if b.get('Type') == 'Producteur']

    # Tri par surface décroissante
    for lst in (neutres, culturels, producteurs):
        lst.sort(key=lambda b: b['Longueur'] * b['Largeur'], reverse=True)

    # Priorité Guérison en premier
    prod_guerison = [p for p in producteurs if str(p.get('Production','')).strip() == 'Guerison']
    prod_autres   = [p for p in producteurs if str(p.get('Production','')).strip() != 'Guerison']
    producteurs = prod_guerison + prod_autres

    # ====================== FONCTIONS PLACEMENT ======================
    def can_place(g, r, c, h, w, rot=False):
        if rot: h, w = w, h
        if r + h > g.shape[0] or c + w > g.shape[1]: return False
        return np.all(g[r:r+h, c:c+w] == 1)

    def place(g, r, c, h, w, value, rot=False):
        if rot: h, w = w, h
        g[r:r+h, c:c+w] = value

    def find_positions(g, h, w, prefer_border=False):
        positions = []
        for r in range(g.shape[0]):
            for c in range(g.shape[1]):
                if can_place(g, r, c, h, w, False):
                    positions.append((r, c, False))
                if h != w and can_place(g, r, c, h, w, True):
                    positions.append((r, c, True))
        if prefer_border:
            positions.sort(key=lambda p: min(p[0], rows-1-p[0], p[1], cols-1-p[1]))
        return positions

    # ====================== PLACEMENT ======================
    grid_work = grid.copy().astype(int)
    placed = []

    # 1. Neutres sur les bords
    st.write("Placement des bâtiments neutres sur les bords...")
    for b in neutres:
        for r, c, rot in find_positions(grid_work, b['Longueur'], b['Largeur'], prefer_border=True)[:60]:
            if can_place(grid_work, r, c, b['Longueur'], b['Largeur'], rot):
                place(grid_work, r, c, b['Longueur'], b['Largeur'], -2, rot)  # -2 = neutre
                placed.append({**b, 'row': r, 'col': c, 'rotation': rot, 'placed': True})
                b['placed'] = True
                break

    # 2. Placement alterné culturels + producteurs
    remaining = culturels + producteurs
    remaining.sort(key=lambda b: b['Longueur'] * b['Largeur'], reverse=True)

    st.write("Placement des culturels et producteurs...")
    i = 0
    while i < len(remaining):
        b = remaining[i]
        placed_ok = False
        for r, c, rot in find_positions(grid_work, b['Longueur'], b['Largeur'], prefer_border=False)[:120]:
            if can_place(grid_work, r, c, b['Longueur'], b['Largeur'], rot):
                value = -3 if b['Type'] == 'Culturel' else -4
                place(grid_work, r, c, b['Longueur'], b['Largeur'], value, rot)
                placed.append({**b, 'row': r, 'col': c, 'rotation': rot, 'placed': True})
                b['placed'] = True
                placed_ok = True
                break
        if not placed_ok:
            i += 1
        else:
            remaining = [x for x in remaining if not x.get('placed', False)]
            remaining.sort(key=lambda x: x['Longueur'] * x['Largeur'], reverse=True)
            i = 0

    # ====================== CALCUL CULTURE & BOOST ======================
    def compute_culture_coverage(grid, placed_list):
        rmax, cmax = grid.shape
        culture_map = np.zeros((rmax, cmax), dtype=float)
        
        for b in placed_list:
            if b.get('Type') != 'Culturel' or not b.get('placed'): continue
            r, c, rot = b['row'], b['col'], b['rotation']
            bh = b['Largeur'] if rot else b['Longueur']
            bw = b['Longueur'] if rot else b['Largeur']
            rad = int(b.get('Rayonnement', 0))
            cult = float(b.get('Culture', 0))
            
            r1 = max(0, r - rad)
            r2 = min(rmax, r + bh + rad)
            c1 = max(0, c - rad)
            c2 = min(cmax, c + bw + rad)
            
            culture_map[r1:r2, c1:c2] += cult
        return culture_map

    culture_map = compute_culture_coverage(grid_work, placed)

    # Calcul boost
    for b in placed:
        if b.get('Type') != 'Producteur': continue
        r, c, rot = b['row'], b['col'], b['rotation']
        h = b['Largeur'] if rot else b['Longueur']
        w = b['Longueur'] if rot else b['Largeur']
        
        cult_received = float(np.mean(culture_map[r:r+h, c:c+w]))
        
        b25 = b.get('Boost 25%')
        b50 = b.get('Boost 50%')
        b100 = b.get('Boost 100%')
        
        if pd.notna(b100) and cult_received >= float(b100):
            boost = 1.00
        elif pd.notna(b50) and cult_received >= float(b50):
            boost = 0.50
        elif pd.notna(b25) and cult_received >= float(b25):
            boost = 0.25
        else:
            boost = 0.00
        
        b['culture_recue'] = round(cult_received, 1)
        b['boost'] = boost
        b['prod_reelle'] = round(float(b.get('Quantite', 0)) * (1 + boost), 1)

    # ====================== STATISTIQUES ======================
    from collections import defaultdict
    stats = defaultdict(float)
    for b in placed:
        prod = b.get('Production', '')
        if b.get('Type') == 'Producteur' and prod and prod != 'Rien':
            stats[prod] += b['prod_reelle']

    non_places = [b['Nom'] for b in batiments if not b.get('placed', False)]
    cases_libres = int(np.sum(grid_work == 1))
    cases_non_place = sum(b['Longueur'] * b['Largeur'] for b in batiments if not b.get('placed', False))

    # ====================== EXPORT EXCEL ======================
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Bâtiments placés
        df_placed = pd.DataFrame([{
            'Nom': b['Nom'],
            'Type': b.get('Type',''),
            'Production': b.get('Production',''),
            'Row': b.get('row',''),
            'Col': b.get('col',''),
            'Rotation': 'Oui' if b.get('rotation') else 'Non',
            'Culture reçue': b.get('culture_recue', 0),
            'Boost': f"{int(b.get('boost',0)*100)}%",
            'Prod/heure réelle': b.get('prod_reelle', 0)
        } for b in placed if b.get('placed')])
        df_placed.to_excel(writer, sheet_name='Bâtiments placés', index=False)

        # 2. Non placés
        pd.DataFrame({'Bâtiment non placé': non_places}).to_excel(writer, sheet_name='Non placés', index=False)

        # 3. Production totale
        pd.DataFrame({
            'Ressource': list(stats.keys()),
            'Production totale / heure': [round(v, 1) for v in stats.values()]
        }).to_excel(writer, sheet_name='Production totale', index=False)

        # 4. Terrain avec fusion et couleurs
        wb = writer.book
        ws = wb.create_sheet('Terrain')

        # Fond gris pour les X (murs)
        gray_wall = PatternFill(start_color="666666", end_color="666666", fill_type="solid")
        for ri in range(rows):
            for ci in range(cols):
                cell = ws.cell(row=ri+1, column=ci+1)
                if grid_work[ri, ci] == -1:
                    cell.value = 'X'
                    cell.fill = gray_wall

        # Styles bâtiments
        orange_fill  = PatternFill(start_color="FF9900", end_color="FF9900", fill_type="solid")  # Culturel
        green_fill   = PatternFill(start_color="00CC44", end_color="00CC44", fill_type="solid")  # Producteur
        gray_fill    = PatternFill(start_color="AAAAAA", end_color="AAAAAA", fill_type="solid")  # Neutre

        center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
        bold_font    = Font(bold=True, size=10)

        for b in placed:
            if not b.get('placed'): continue
            r, c, rot = b['row'], b['col'], b['rotation']
            h = b['Largeur'] if rot else b['Longueur']
            w = b['Longueur'] if rot else b['Largeur']

            # Couleur selon type
            if b['Type'] == 'Culturel':
                fill_color = orange_fill
            elif b['Type'] == 'Producteur':
                fill_color = green_fill
            else:
                fill_color = gray_fill

            # Texte
            name = str(b['Nom'])[:16]
            if b['Type'] == 'Producteur' and 'boost' in b:
                text = f"{name}\n{int(b['boost']*100)}%"
            else:
                text = name

            # Fusion
            top_left_cell = f"{openpyxl.utils.get_column_letter(c+1)}{r+1}"
            bottom_right_cell = f"{openpyxl.utils.get_column_letter(c+w)}{r+h}"
            ws.merge_cells(f"{top_left_cell}:{bottom_right_cell}")

            # Style
            cell = ws[top_left_cell]
            cell.value = text
            cell.fill = fill_color
            cell.alignment = center_align
            cell.font = bold_font

            # Appliquer la couleur à toutes les cellules de la zone (important pour l'affichage)
            for ri in range(r, r + h):
                for ci in range(c, c + w):
                    ws.cell(row=ri+1, column=ci+1).fill = fill_color

        # Ajustement visuel
        for col_letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
            if ws.column_dimensions[col_letter].width is None:
                ws.column_dimensions[col_letter].width = 5.2
        for row_idx in range(1, rows + 10):
            ws.row_dimensions[row_idx].height = 28

        # 5. Résumé
        pd.DataFrame({
            'Indicateur': [
                'Cases libres',
                'Bâtiments non placés',
                'Cases des bâtiments non placés',
                'Production Guérison /h',
                'Production Nourriture /h',
                'Production Or /h',
                'Production totale autre /h'
            ],
            'Valeur': [
                cases_libres,
                len(non_places),
                cases_non_place,
                round(stats.get('Guerison', 0), 1),
                round(stats.get('Nourriture', 0), 1),
                round(stats.get('Or', 0), 1),
                round(sum(v for k,v in stats.items() if k not in ['Guerison','Nourriture','Or']), 1)
            ]
        }).to_excel(writer, sheet_name='Résumé', index=False)

    output.seek(0)

    st.success("Traitement terminé ✓")
    st.download_button(
        label="📥 Télécharger le résultat (Ville_resultat.xlsx)",
        data=output,
        file_name="Ville_resultat.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Production estimée")
        st.write(dict(stats))
    with col2:
        st.subheader("État final")
        st.write(f"Cases libres : **{cases_libres}**")
        st.write(f"Bâtiments non placés : **{len(non_places)}**")

except Exception as e:
    st.error("Une erreur est survenue pendant le traitement")
    st.exception(e)
