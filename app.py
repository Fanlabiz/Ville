# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

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
    while np.all(grid[-1] == 0): grid = grid[:-1]
    while np.all(grid[:, -1] == 0): grid = np.delete(grid, -1, axis=1)
    
    rows, cols = grid.shape
    st.success(f"Terrain chargé : {rows} × {cols} cases")

    # ====================== PRÉPARATION BÂTIMENTS ======================
    batiments = []
    for _, row in df_bat.iterrows():
        if pd.isna(row['Nom']): continue
        b = row.to_dict()
        b['Longueur'] = int(b['Longueur'])
        b['Largeur']  = int(b['Largeur'])
        b['Nombre']   = int(b['Nombre'])
        b['placed']   = False
        for _ in range(b['Nombre']):
            batiments.append(b.copy())

    # Séparation par type
    neutres      = [b for b in batiments if b['Type'] == 'Neutre']
    culturels    = [b for b in batiments if b['Type'] == 'Culturel']
    producteurs  = [b for b in batiments if b['Type'] == 'Producteur']

    # Tri par surface décroissante
    for lst in (neutres, culturels, producteurs):
        lst.sort(key=lambda b: b['Longueur'] * b['Largeur'], reverse=True)

    # Priorité Guérison en premier
    prod_guerison = [p for p in producteurs if str(p['Production']).strip() == 'Guerison']
    prod_autres   = [p for p in producteurs if str(p['Production']).strip() != 'Guerison']
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
    grid_work = grid.copy()
    placed = []

    # 1. Neutres sur les bords
    st.write("Placement des bâtiments neutres sur les bords...")
    for b in neutres:
        for r, c, rot in find_positions(grid_work, b['Longueur'], b['Largeur'], prefer_border=True)[:50]:
            if can_place(grid_work, r, c, b['Longueur'], b['Largeur'], rot):
                place(grid_work, r, c, b['Longueur'], b['Largeur'], -2, rot)  # -2 = neutre
                placed.append({**b, 'row': r, 'col': c, 'rotation': rot, 'placed': True})
                b['placed'] = True
                break

    # 2. Placement alterné culturels + producteurs (tri taille décroissante)
    remaining = culturels + producteurs
    remaining.sort(key=lambda b: b['Longueur'] * b['Largeur'], reverse=True)

    st.write("Placement des culturels et producteurs...")
    i = 0
    while i < len(remaining):
        b = remaining[i]
        placed_ok = False
        for r, c, rot in find_positions(grid_work, b['Longueur'], b['Largeur'], prefer_border=False)[:100]:
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
            # On recommence depuis le début pour prioriser les gros restants
            remaining = [x for x in remaining if not x.get('placed', False)]
            remaining.sort(key=lambda x: x['Longueur'] * x['Largeur'], reverse=True)
            i = 0

    # ====================== CALCUL CULTURE & BOOST ======================
    def compute_culture_coverage(grid, placed_list):
        rows, cols = grid.shape
        culture_map = np.zeros((rows, cols), dtype=float)
        
        for b in placed_list:
            if b['Type'] != 'Culturel' or not b.get('placed'): continue
            r, c, rot = b['row'], b['col'], b['rotation']
            bh = b['Largeur'] if rot else b['Longueur']
            bw = b['Longueur'] if rot else b['Largeur']
            rad = int(b.get('Rayonnement', 0))
            cult = float(b.get('Culture', 0))
            
            r1 = max(0, r - rad)
            r2 = min(rows, r + bh + rad)
            c1 = max(0, c - rad)
            c2 = min(cols, c + bw + rad)
            
            culture_map[r1:r2, c1:c2] += cult
        return culture_map

    culture_map = compute_culture_coverage(grid_work, placed)

    # Calcul boost pour chaque producteur
    for b in placed:
        if b['Type'] != 'Producteur': continue
        r, c, rot = b['row'], b['col'], b['rotation']
        h = b['Largeur'] if rot else b['Longueur']
        w = b['Longueur'] if rot else b['Largeur']
        
        # Moyenne sur tout le bâtiment (plus précis)
        cult_received = float(np.mean(culture_map[r:r+h, c:c+w]))
        
        b25 = b.get('Boost 25%')
        b50 = b.get('Boost 50%')
        b100 = b.get('Boost 100%')
        
        if pd.notna(b100) and cult_received >= b100:
            boost = 1.00
        elif pd.notna(b50) and cult_received >= b50:
            boost = 0.50
        elif pd.notna(b25) and cult_received >= b25:
            boost = 0.25
        else:
            boost = 0.00
        
        b['culture_recue'] = round(cult_received, 1)
        b['boost'] = boost
        b['prod_reelle'] = round(float(b['Quantite']) * (1 + boost), 1)

    # ====================== STATISTIQUES ======================
    from collections import defaultdict
    stats = defaultdict(float)
    for b in placed:
        if b['Type'] == 'Producteur' and b.get('Production') and b['Production'] != 'Rien':
            stats[b['Production']] += b['prod_reelle']

    non_places = [b['Nom'] for b in batiments if not b.get('placed', False)]
    cases_libres = int(np.sum(grid_work == 1))
    cases_non_place = sum(b['Longueur'] * b['Largeur'] for b in batiments if not b.get('placed', False))

    # ====================== EXPORT EXCEL ======================
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Bâtiments placés
        df_placed = pd.DataFrame([{
            'Nom': b['Nom'],
            'Type': b['Type'],
            'Production': b.get('Production', ''),
            'Row': b.get('row', ''),
            'Col': b.get('col', ''),
            'Rotation': 'Oui' if b.get('rotation') else 'Non',
            'Culture reçue': b.get('culture_recue', 0),
            'Boost': f"{int(b.get('boost',0)*100)}%",
            'Prod/heure réelle': b.get('prod_reelle', 0)
        } for b in placed])
        df_placed.to_excel(writer, sheet_name='Bâtiments placés', index=False)

        # 2. Non placés
        pd.DataFrame({'Bâtiment non placé': non_places}).to_excel(writer, sheet_name='Non placés', index=False)

        # 3. Stats production
        pd.DataFrame({
            'Ressource': list(stats.keys()),
            'Production totale / heure': [round(v, 1) for v in stats.values()]
        }).to_excel(writer, sheet_name='Production totale', index=False)

        # 4. Terrain visuel (texte)
        viz = grid_work.astype(object).copy()
        for b in placed:
            r, c, rot = b['row'], b['col'], b['rotation']
            h = b['Largeur'] if rot else b['Longueur']
            w = b['Longueur'] if rot else b['Largeur']
            label = b['Nom'][:10] + ("\n" + str(int(b.get('boost',0)*100)) + "%" if b['Type']=='Producteur' else "")
            for i in range(h):
                for j in range(w):
                    viz[r+i, c+j] = label
        pd.DataFrame(viz).to_excel(writer, sheet_name='Terrain', index=False, header=False)

        # 5. Résumé
        pd.DataFrame({
            'Indicateur': [
                'Cases libres',
                'Bâtiments non placés',
                'Cases des bâtiments non placés',
                'Production Guérison /h',
                'Production Nourriture /h',
                'Production Or /h'
            ],
            'Valeur': [
                cases_libres,
                len(non_places),
                cases_non_place,
                round(stats.get('Guerison', 0), 1),
                round(stats.get('Nourriture', 0), 1),
                round(stats.get('Or', 0), 1)
            ]
        }).to_excel(writer, sheet_name='Résumé', index=False)

    output.seek(0)

    st.success("✅ Placement terminé avec succès !")
    st.download_button(
        label="📥 Télécharger Ville_resultat.xlsx",
        data=output,
        file_name="Ville_resultat.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Aperçu
    st.subheader("📊 Production totale estimée")
    st.write(dict(stats))
    st.write(f"Cases libres restantes : **{cases_libres}**")
    st.write(f"Bâtiments non placés : **{len(non_places)}**")

except Exception as e:
    st.error("Erreur lors du traitement")
    st.exception(e)
