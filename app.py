# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Alignment, Font
from collections import defaultdict

st.set_page_config(page_title="Optimiseur de Ville - Phase 1 + Boost Opti", layout="wide")
st.title("Optimiseur de placement + optimisation des boosts")

uploaded_file = st.file_uploader("Choisir ton fichier Excel (Ville.xlsx)", type=["xlsx"])

if uploaded_file is None:
    st.info("Charge ton fichier pour commencer.")
    st.stop()

try:
    xls = pd.ExcelFile(uploaded_file)
    df_terrain = pd.read_excel(xls, "Terrain", header=None)
    df_bat     = pd.read_excel(xls, "Batiments")
    # df_actuel  = pd.read_excel(xls, "Actuel", header=None)   # non utilisé ici

    # ─── Chargement & nettoyage terrain ────────────────────────────────────────
    grid = df_terrain.fillna('').replace('X', -1).replace('', 0).astype(int).values
    while len(grid) > 0 and np.all(grid[-1] == 0): grid = grid[:-1]
    while grid.shape[1] > 0 and np.all(grid[:, -1] == 0): grid = np.delete(grid, -1, axis=1)
    rows, cols = grid.shape
    st.success(f"Terrain : {rows} × {cols}")

    # ─── Préparation bâtiments ─────────────────────────────────────────────────
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

    neutres     = [b for b in batiments if b.get('Type') == 'Neutre']
    culturels   = [b for b in batiments if b.get('Type') == 'Culturel']
    producteurs = [b for b in batiments if b.get('Type') == 'Producteur']

    for lst in (neutres, culturels, producteurs):
        lst.sort(key=lambda x: x['Longueur'] * x['Largeur'], reverse=True)

    prod_guerison = [p for p in producteurs if str(p.get('Production','')).strip() == 'Guerison']
    prod_autres   = [p for p in producteurs if str(p.get('Production','')).strip() != 'Guerison']
    producteurs = prod_guerison + prod_autres

    # ─── Fonctions de base ─────────────────────────────────────────────────────
    def can_place(g, r, c, h, w, rot=False):
        if rot: h, w = w, h
        if r + h > g.shape[0] or c + w > g.shape[1]: return False
        return np.all(g[r:r+h, c:c+w] == 1)

    def place(g, r, c, h, w, value, rot=False):
        if rot: h, w = w, h
        g[r:r+h, c:c+w] = value

    def find_positions(g, h, w, prefer_border=False):
        pos = []
        for ri in range(g.shape[0]):
            for ci in range(g.shape[1]):
                if can_place(g, ri, ci, h, w, False):
                    pos.append((ri, ci, False))
                if h != w and can_place(g, ri, ci, h, w, True):
                    pos.append((ri, ci, True))
        if prefer_border:
            pos.sort(key=lambda p: min(p[0], rows-1-p[0], p[1], cols-1-p[1]))
        return pos

    # ─── Phase 1 : Placement initial ───────────────────────────────────────────
    grid_work = grid.copy().astype(int)
    placed = []

    st.write("Phase 1 – Placement initial…")
    # Neutres → bords
    for b in neutres:
        for r,c,rot in find_positions(grid_work, b['Longueur'], b['Largeur'], True)[:60]:
            if can_place(grid_work, r, c, b['Longueur'], b['Largeur'], rot):
                place(grid_work, r, c, b['Longueur'], b['Largeur'], -2, rot)
                placed.append({**b, 'row':r, 'col':c, 'rotation':rot, 'placed':True})
                b['placed'] = True
                break

    # Culturels + Producteurs (alternance par taille)
    remaining = culturels + producteurs
    remaining.sort(key=lambda x: x['Longueur']*x['Largeur'], reverse=True)

    i = 0
    while i < len(remaining):
        b = remaining[i]
        ok = False
        for r,c,rot in find_positions(grid_work, b['Longueur'], b['Largeur'], False)[:140]:
            if can_place(grid_work, r, c, b['Longueur'], b['Largeur'], rot):
                val = -3 if b['Type']=='Culturel' else -4
                place(grid_work, r, c, b['Longueur'], b['Largeur'], val, rot)
                placed.append({**b, 'row':r, 'col':c, 'rotation':rot, 'placed':True})
                b['placed'] = True
                ok = True
                break
        if not ok:
            i += 1
        else:
            remaining = [x for x in remaining if not x.get('placed')]
            remaining.sort(key=lambda x: x['Longueur']*x['Largeur'], reverse=True)
            i = 0

    # ─── Phase 2 : Optimisation des boosts (déplacement autorisé sauf neutres) ──
    st.write("Phase 2 – Optimisation des boosts…")

    def compute_culture_map(g, placed_list):
        rmax, cmax = g.shape
        cmap = np.zeros((rmax, cmax), dtype=float)
        for b in placed_list:
            if b.get('Type') != 'Culturel' or not b.get('placed'): continue
            rr, cc, rrot = b['row'], b['col'], b['rotation']
            bh = b['Largeur'] if rrot else b['Longueur']
            bw = b['Longueur'] if rrot else b['Largeur']
            rad = int(b.get('Rayonnement', 0))
            cult = float(b.get('Culture', 0))
            r1 = max(0, rr - rad)
            r2 = min(rmax, rr + bh + rad)
            c1 = max(0, cc - rad)
            c2 = min(cmax, cc + bw + rad)
            cmap[r1:r2, c1:c2] += cult
        return cmap

    def get_boost_value(cult, b):
        t25 = b.get('Boost 25%')
        t50 = b.get('Boost 50%')
        t100 = b.get('Boost 100%')
        if pd.notna(t100) and cult >= float(t100): return 1.0
        if pd.notna(t50)  and cult >= float(t50):  return 0.5
        if pd.notna(t25)  and cult >= float(t25):  return 0.25
        return 0.0

    # Déplaçables = tout sauf neutres
    to_optimize = [b for b in placed if b.get('placed') and b['Type'] != 'Neutre']
    to_optimize.sort(key=lambda b: (
        0 if b['Type']=='Producteur' else 1,                          # prod avant cult
        -{'Guerison':100, 'Nourriture':60, 'Or':30}.get(b.get('Production',''), 10)
    ), reverse=True)

    improved = 0

    for building in to_optimize:
        orig_r, orig_c, orig_rot = building['row'], building['col'], building['rotation']
        orig_h = building['Largeur'] if orig_rot else building['Longueur']
        orig_w = building['Longueur'] if orig_rot else building['Largeur']

        # Libération temporaire
        place(grid_work, orig_r, orig_c, orig_h, orig_w, 1, orig_rot)

        best_score = -1e12
        best_r, best_c, best_rot = orig_r, orig_c, orig_rot

        for try_rot in [False, True]:
            hh = building['Largeur'] if try_rot else building['Longueur']
            ww = building['Longueur'] if try_rot else building['Largeur']

            cands = find_positions(grid_work, hh, ww, False)[:220]

            for nr, nc, _ in cands:
                if not can_place(grid_work, nr, nc, hh, ww, False): continue

                place(grid_work, nr, nc, hh, ww, -5, try_rot)
                temp_map = compute_culture_map(grid_work, placed)

                score = 0.0
                for p in [x for x in placed if x['Type']=='Producteur' and x.get('placed')]:
                    pr, pc, prot = p['row'], p['col'], p['rotation']
                    ph = p['Largeur'] if prot else p['Longueur']
                    pw = p['Longueur'] if prot else p['Largeur']
                    cult_p = float(np.mean(temp_map[pr:pr+ph, pc:pc+pw]))
                    boost_p = get_boost_value(cult_p, p)
                    w = {'Guerison':12, 'Nourriture':7, 'Or':4}.get(p.get('Production',''), 1.5)
                    score += boost_p * w * float(p.get('Quantite', 0))

                if score > best_score + 1.2:   # seuil d'amélioration
                    best_score = score
                    best_r, best_c, best_rot = nr, nc, try_rot

                place(grid_work, nr, nc, hh, ww, 1, try_rot)

        # Application meilleure position (ou retour origine)
        h_final = building['Largeur'] if best_rot else building['Longueur']
        w_final = building['Longueur'] if best_rot else building['Largeur']
        val = -3 if building['Type']=='Culturel' else -4
        place(grid_work, best_r, best_c, h_final, w_final, val, best_rot)

        building['row'] = best_r
        building['col'] = best_c
        building['rotation'] = best_rot

        if (best_r, best_c, best_rot) != (orig_r, orig_c, orig_rot):
            improved += 1

    st.write(f"Phase 2 terminée – {improved} bâtiments ont été repositionnés")

    # ─── Calcul final culture & stats ──────────────────────────────────────────
    culture_map = compute_culture_map(grid_work, placed)

    for b in placed:
        if b['Type'] != 'Producteur': continue
        r,c,rot = b['row'], b['col'], b['rotation']
        h = b['Largeur'] if rot else b['Longueur']
        w = b['Longueur'] if rot else b['Largeur']
        b['culture_recue'] = round(float(np.mean(culture_map[r:r+h, c:c+w])), 1)
        b['boost'] = get_boost_value(b['culture_recue'], b)
        b['prod_reelle'] = round(float(b.get('Quantite',0)) * (1 + b['boost']), 1)

    stats = defaultdict(float)
    for b in placed:
        p = b.get('Production','')
        if b['Type']=='Producteur' and p and p != 'Rien':
            stats[p] += b['prod_reelle']

    non_places = [b['Nom'] for b in batiments if not b.get('placed')]
    cases_libres = int(np.sum(grid_work == 1))
    cases_non_place = sum(b['Longueur']*b['Largeur'] for b in batiments if not b.get('placed'))

    # ─── Export Excel ──────────────────────────────────────────────────────────
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Bâtiments placés
        pd.DataFrame([{
            'Nom': b['Nom'],
            'Type': b.get('Type',''),
            'Production': b.get('Production',''),
            'Row': b.get('row',''),
            'Col': b.get('col',''),
            'Rotation': 'Oui' if b.get('rotation') else 'Non',
            'Culture reçue': b.get('culture_recue',0),
            'Boost': f"{int(b.get('boost',0)*100)}%",
            'Prod/heure réelle': b.get('prod_reelle',0)
        } for b in placed if b.get('placed')]).to_excel(writer, 'Bâtiments placés', index=False)

        # Non placés
        pd.DataFrame({'Non placés': non_places}).to_excel(writer, 'Non placés', index=False)

        # Production totale
        pd.DataFrame({
            'Ressource': list(stats),
            'Prod totale /h': [round(v,1) for v in stats.values()]
        }).to_excel(writer, 'Production totale', index=False)

        # Terrain fusionné + couleurs
        wb = writer.book
        ws = wb.create_sheet('Terrain')

        gray_wall = PatternFill("solid", start_color="666666")
        for ri in range(rows):
            for ci in range(cols):
                if grid_work[ri,ci] == -1:
                    ws.cell(ri+1, ci+1).value = 'X'
                    ws.cell(ri+1, ci+1).fill = gray_wall

        orange  = PatternFill("solid", start_color="FF9900")
        green   = PatternFill("solid", start_color="00CC44")
        gray    = PatternFill("solid", start_color="AAAAAA")
        center  = Alignment(horizontal="center", vertical="center", wrap_text=True)
        bold    = Font(bold=True, size=10)

        for b in placed:
            if not b.get('placed'): continue
            r,c,rot = b['row'], b['col'], b['rotation']
            h = b['Largeur'] if rot else b['Longueur']
            w = b['Longueur'] if rot else b['Largeur']

            fill = orange if b['Type']=='Culturel' else green if b['Type']=='Producteur' else gray
            txt = str(b['Nom'])[:15]
            if b['Type']=='Producteur' and 'boost' in b:
                txt += f"\n{int(b['boost']*100)}%"

            tl = f"{openpyxl.utils.get_column_letter(c+1)}{r+1}"
            br = f"{openpyxl.utils.get_column_letter(c+w)}{r+h}"
            ws.merge_cells(f"{tl}:{br}")

            cell = ws[tl]
            cell.value = txt
            cell.fill = fill
            cell.alignment = center
            cell.font = bold

            for ri in range(r, r+h):
                for ci in range(c, c+w):
                    ws.cell(ri+1, ci+1).fill = fill

        for col in range(1, cols+1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 5.8
        for rw in range(1, rows+10):
            ws.row_dimensions[rw].height = 32

        # Résumé
        pd.DataFrame({
            'Indicateur': ['Cases libres', 'Bât non placés', 'Cases non placées',
                           'Guérison /h', 'Nourriture /h', 'Or /h'],
            'Valeur': [cases_libres, len(non_places), cases_non_place,
                       round(stats.get('Guerison',0),1),
                       round(stats.get('Nourriture',0),1),
                       round(stats.get('Or',0),1)]
        }).to_excel(writer, 'Résumé', index=False)

    output.seek(0)
    st.success("Traitement terminé !")
    st.download_button("Télécharger résultat.xlsx", output, "Ville_optimisee.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("Production estimée")
    st.write(dict(stats))
    st.write(f"Cases libres : **{cases_libres}**   |   Non placés : **{len(non_places)}**")

except Exception as e:
    st.error("Erreur pendant le traitement")
    st.exception(e)
