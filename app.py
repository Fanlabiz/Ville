import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import io, copy, itertools, time

st.set_page_config(page_title="Placement de Batiments", layout="wide")
st.title("Optimiseur de placement de batiments")

# ═══════════════════════════════════════════════════════════════════════════════
# CHARGEMENT
# ═══════════════════════════════════════════════════════════════════════════════

def load_terrain(ws):
    grid = []
    for row in ws.iter_rows(values_only=True):
        last = max((i for i, v in enumerate(row) if v is not None), default=-1)
        if last == -1:
            continue
        grid.append(list(row[:last+1]))
    while grid and all(v is None for v in grid[-1]):
        grid.pop()
    return grid

def load_buildings(ws):
    rows = list(ws.iter_rows(values_only=True))
    buildings = []
    for row in rows[1:]:
        if not row or row[0] is None:
            continue
        qty_raw = row[11]
        if qty_raw is None:
            qty = 0.0
        else:
            s = str(qty_raw).strip().lstrip('=')
            try:
                qty = float(eval(s))
            except Exception:
                qty = 0.0
        try:
            b = {
                'nom':         str(row[0]).strip(),
                'longueur':    int(row[1]),
                'largeur':     int(row[2]),
                'nombre':      int(row[3]),
                'type':        str(row[4]).strip(),
                'culture':     float(row[5]) if row[5] else 0,
                'rayonnement': int(row[6]) if row[6] else 0,
                'boost25':     float(row[7]) if row[7] is not None else None,
                'boost50':     float(row[8]) if row[8] is not None else None,
                'boost100':    float(row[9]) if row[9] is not None else None,
                'production':  str(row[10]).strip() if row[10] else 'Rien',
                'quantite':    qty,
            }
            buildings.append(b)
        except Exception:
            continue
    return buildings

def grid_to_matrix(grid):
    if not grid:
        return np.zeros((0, 0), dtype=bool)
    cols = max(len(r) for r in grid)
    mat = []
    for row in grid:
        r2 = list(row) + [None] * (cols - len(row))
        mat.append([v == 1 for v in r2])
    return np.array(mat, dtype=bool)

# ═══════════════════════════════════════════════════════════════════════════════
# UTILITAIRES GRILLE
# ═══════════════════════════════════════════════════════════════════════════════

PROD_PRIORITY = ['Guerison', 'Nourriture', 'Or']

def expand_buildings(buildings):
    return [dict(b) for b in buildings for _ in range(b['nombre'])]

def can_place(grid, r, c, h, w):
    rn, cn = grid.shape
    if r + h > rn or c + w > cn:
        return False
    return bool(grid[r:r+h, c:c+w].all())

def do_place(grid, r, c, h, w):
    grid[r:r+h, c:c+w] = False

def rebuild_grid(grid_orig, placed):
    grid = grid_orig.copy()
    for p in placed:
        do_place(grid, p['row'], p['col'], p['h'], p['w'])
    return grid

def orientations(h, w):
    if h == w:
        return [(h, w)]
    return [(h, w), (w, h)]

# ═══════════════════════════════════════════════════════════════════════════════
# CULTURE & BOOST
# ═══════════════════════════════════════════════════════════════════════════════

def build_culture_map(placed, rows, cols):
    cmap = np.zeros((rows, cols), dtype=float)
    for p in placed:
        if p['type'] != 'Culturel' or p['culture'] <= 0:
            continue
        ray = p['rayonnement']
        r0, c0, h, w = p['row'], p['col'], p['h'], p['w']
        rmin = max(0, r0 - ray);  rmax = min(rows, r0 + h + ray)
        cmin = max(0, c0 - ray);  cmax = min(cols, c0 + w + ray)
        cmap[rmin:rmax, cmin:cmax] += p['culture']
        cmap[r0:r0+h, c0:c0+w]    -= p['culture']
    np.clip(cmap, 0, None, out=cmap)
    return cmap

def compute_culture_received(placed, rows, cols):
    cmap = build_culture_map(placed, rows, cols)
    for p in placed:
        if p['type'] == 'Producteur':
            r0, c0, h, w = p['row'], p['col'], p['h'], p['w']
            foot = cmap[r0:r0+h, c0:c0+w]
            p['culture_recue'] = float(foot.max()) if foot.size > 0 else 0.0
        else:
            p['culture_recue'] = 0.0

def get_boost(p):
    c = p.get('culture_recue', 0)
    b25, b50, b100 = p.get('boost25'), p.get('boost50'), p.get('boost100')
    if b100 is not None and c >= b100:
        return 100, 2.0
    if b50 is not None and c >= b50:
        return 50, 1.5
    if b25 is not None and c >= b25:
        return 25, 1.25
    return 0, 1.0

# ═══════════════════════════════════════════════════════════════════════════════
# STRATÉGIES DE POSITIONNEMENT
# ═══════════════════════════════════════════════════════════════════════════════

def find_border_pos(grid, rows, cols, h, w):
    best, best_dist = None, 9999
    for bh, bw in orientations(h, w):
        for r in range(rows):
            for c in range(cols):
                if can_place(grid, r, c, bh, bw):
                    dist = min(r, c, rows - r - bh, cols - c - bw)
                    if dist < best_dist:
                        best_dist = dist
                        best = (r, c, bh, bw)
    return best

def find_cultural_pos(grid, rows, cols, h, w, ray):
    best, best_score = None, -1
    for bh, bw in orientations(h, w):
        for r in range(rows):
            for c in range(cols):
                if not can_place(grid, r, c, bh, bw):
                    continue
                rmin = max(0, r - ray);  rmax = min(rows, r + bh + ray)
                cmin = max(0, c - ray);  cmax = min(cols, c + bw + ray)
                zone = grid[rmin:rmax, cmin:cmax].copy()
                dr0 = r - rmin;  dc0 = c - cmin
                zone[dr0:dr0+bh, dc0:dc0+bw] = False
                score = int(zone.sum())
                if score > best_score:
                    best_score = score
                    best = (r, c, bh, bw)
    return best

def find_producer_pos(grid, rows, cols, h, w, cmap):
    best, best_val = None, -1
    for bh, bw in orientations(h, w):
        for r in range(rows):
            for c in range(cols):
                if can_place(grid, r, c, bh, bw):
                    val = float(cmap[r:r+bh, c:c+bw].max())
                    if val > best_val or best is None:
                        best_val = val
                        best = (r, c, bh, bw)
    return best

def find_any_pos(grid, rows, cols, h, w):
    for bh, bw in orientations(h, w):
        for r in range(rows):
            for c in range(cols):
                if can_place(grid, r, c, bh, bw):
                    return (r, c, bh, bw)
    return None

def prod_sort_key(b):
    p = b['production']
    if p in PROD_PRIORITY:
        return (PROD_PRIORITY.index(p), -(b['longueur'] * b['largeur']))
    if p == 'Rien':
        return (999, 0)
    return (500, -(b['longueur'] * b['largeur']))

# ═══════════════════════════════════════════════════════════════════════════════
# PLACEMENT PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════

def run_placement(grid_orig, buildings_raw):
    grid      = grid_orig.copy()
    rows, cols = grid.shape
    instances  = expand_buildings(buildings_raw)
    placed, unplaced = [], []

    # 1. Neutres sur les bords
    neutres = sorted([b for b in instances if b['type'] == 'Neutre'],
                     key=lambda b: b['longueur'] * b['largeur'], reverse=True)
    for b in neutres:
        pos = find_border_pos(grid, rows, cols, b['longueur'], b['largeur'])
        if not pos:
            pos = find_any_pos(grid, rows, cols, b['longueur'], b['largeur'])
        if pos:
            r, c, h, w = pos
            placed.append({**b, 'row': r, 'col': c, 'h': h, 'w': w})
            do_place(grid, r, c, h, w)
        else:
            unplaced.append(b)

    # 2. Culturels & producteurs en alternance
    culturels   = sorted([b for b in instances if b['type'] == 'Culturel'],
                         key=lambda b: b['longueur'] * b['largeur'], reverse=True)
    producteurs = sorted([b for b in instances if b['type'] == 'Producteur'],
                         key=prod_sort_key)
    ci = pi = 0
    while ci < len(culturels) or pi < len(producteurs):
        if ci < len(culturels):
            b   = culturels[ci]
            pos = find_cultural_pos(grid, rows, cols, b['longueur'], b['largeur'],
                                    b['rayonnement'])
            if not pos:
                pos = find_any_pos(grid, rows, cols, b['longueur'], b['largeur'])
            if pos:
                r, c, h, w = pos
                placed.append({**b, 'row': r, 'col': c, 'h': h, 'w': w})
                do_place(grid, r, c, h, w)
            else:
                unplaced.append(b)
            ci += 1
        if pi < len(producteurs):
            b    = producteurs[pi]
            cmap = build_culture_map(placed, rows, cols)
            pos  = find_producer_pos(grid, rows, cols, b['longueur'], b['largeur'], cmap)
            if not pos:
                pos = find_any_pos(grid, rows, cols, b['longueur'], b['largeur'])
            if pos:
                r, c, h, w = pos
                placed.append({**b, 'row': r, 'col': c, 'h': h, 'w': w})
                do_place(grid, r, c, h, w)
            else:
                unplaced.append(b)
            pi += 1

    # 3. Passe de compaction orientée zones
    placed, unplaced = zone_compaction(grid_orig, placed, unplaced)

    # 4. Passe de décalage latéral : consolider l'espace pour les non-placés
    placed, unplaced = lateral_shift_pass(grid_orig, placed, unplaced)

    # 4b. Passe de décalages coordonnés par paires
    placed, unplaced = coordinated_shift_pass(grid_orig, placed, unplaced)

    # 5. Passe d'optimisation culturelle : repositionner les petits sites
    placed = cultural_reposition_pass(grid_orig, placed)

    compute_culture_received(placed, rows, cols)
    free_cells = int(rebuild_grid(grid_orig, placed).sum())
    return placed, unplaced, free_cells

# ═══════════════════════════════════════════════════════════════════════════════
# PASSE DE COMPACTION ORIENTÉE ZONES
# ═══════════════════════════════════════════════════════════════════════════════

def blockers_for_zone(placed, r, c, h, w):
    """Retourne l'ensemble des bâtiments qui empiètent sur la zone (r,c,h,w)."""
    result = []
    for p in placed:
        pr, pc, ph, pw = p['row'], p['col'], p['h'], p['w']
        if pr < r+h and pr+ph > r and pc < c+w and pc+pw > c:
            result.append(p)
    return result

def movability_score(p):
    """Score de mobilité : plus le score est bas, plus le bâtiment est facile à déplacer
    sans pénalité (neutres d'abord, puis culturels, puis producteurs faible priorité)."""
    t = p['type']
    if t == 'Neutre':
        return 0
    if t == 'Culturel':
        return 1
    prod = p.get('production', 'Rien')
    if prod == 'Rien':
        return 2
    if prod not in PROD_PRIORITY:
        return 3
    return 4 + PROD_PRIORITY.index(prod)

def try_relocate_group(grid_orig, placed, group):
    """
    Tente de déplacer tous les bâtiments du groupe vers de nouvelles positions
    libérées entre eux.  Retourne le nouveau placed si succès, None sinon.
    """
    rows, cols = grid_orig.shape
    group_ids  = set(id(p) for p in group)

    # Grille sans le groupe
    base_placed = [p for p in placed if id(p) not in group_ids]
    base_grid   = rebuild_grid(grid_orig, base_placed)

    # Replacer chaque membre du groupe séquentiellement
    new_positions = []
    temp_grid = base_grid.copy()
    for p in group:
        pos = find_any_pos(temp_grid, rows, cols, p['h'], p['w'])
        if pos is None:
            # Essayer l'orientation inverse
            pos = find_any_pos(temp_grid, rows, cols, p['w'], p['h'])
        if pos is None:
            return None   # impossible de recaser ce bâtiment
        r, c, h, w = pos
        new_positions.append((r, c, h, w))
        do_place(temp_grid, r, c, h, w)

    # Construire le nouveau placed
    new_placed = list(base_placed)
    for p, (r, c, h, w) in zip(group, new_positions):
        new_placed.append({**p, 'row': r, 'col': c, 'h': h, 'w': w})
    return new_placed

def zone_compaction(grid_orig, placed, unplaced):
    """
    Pour chaque bâtiment non placé (du plus grand au plus petit) :
    1. Cherche toutes les zones candidates dans le terrain où il pourrait tenir.
    2. Pour chaque zone, identifie les bâtiments qui la bloquent (blockers).
    3. Si les blockers sont tous "déplaçables" et peu nombreux (≤ MAX_BLOCKERS),
       tente de les relocaliser ailleurs.
    4. Si la relocalisation réussit et que la zone cible est maintenant libre,
       place le bâtiment cible.
    Les bâtiments sont triés par score de mobilité pour minimiser les perturbations.
    """
    MAX_BLOCKERS = 4   # max bâtiments à déplacer simultanément
    rows, cols   = grid_orig.shape

    still_unplaced = sorted(unplaced, key=lambda b: -(b['longueur'] * b['largeur']))
    current_placed = list(placed)

    for target in still_unplaced[:]:
        # Essai direct sans rien déplacer
        grid_cur = rebuild_grid(grid_orig, current_placed)
        pos = find_any_pos(grid_cur, rows, cols, target['longueur'], target['largeur'])
        if pos:
            r, c, h, w = pos
            current_placed.append({**target, 'row': r, 'col': c, 'h': h, 'w': w})
            still_unplaced.remove(target)
            # Cascade : tenter de caser d'autres non-placés dans la foulée
            for other in still_unplaced[:]:
                grid_cur2 = rebuild_grid(grid_orig, current_placed)
                pos2 = find_any_pos(grid_cur2, rows, cols,
                                    other['longueur'], other['largeur'])
                if pos2:
                    r2, c2, h2, w2 = pos2
                    current_placed.append({**other, 'row': r2, 'col': c2,
                                           'h': h2, 'w': w2})
                    still_unplaced.remove(other)
            continue

        # Construire et trier les zones candidates : peu de blockers + blockers mobiles
        candidates = []
        for bh, bw in orientations(target['longueur'], target['largeur']):
            for r in range(rows - bh + 1):
                for c in range(cols - bw + 1):
                    if not grid_orig[r:r+bh, c:c+bw].all():
                        continue
                    blk = blockers_for_zone(current_placed, r, c, bh, bw)
                    if len(blk) == 0:
                        current_placed.append({**target, 'row': r, 'col': c,
                                               'h': bh, 'w': bw})
                        still_unplaced.remove(target)
                        # Cascade
                        for other in still_unplaced[:]:
                            grid_cur2 = rebuild_grid(grid_orig, current_placed)
                            pos2 = find_any_pos(grid_cur2, rows, cols,
                                                other['longueur'], other['largeur'])
                            if pos2:
                                r2, c2, h2, w2 = pos2
                                current_placed.append({**other, 'row': r2, 'col': c2,
                                                       'h': h2, 'w': w2})
                                still_unplaced.remove(other)
                        break   # inner loop
                    if 0 < len(blk) <= MAX_BLOCKERS:
                        mob = sum(movability_score(p) for p in blk)
                        candidates.append((len(blk), mob, r, c, bh, bw, blk))
            else:
                continue
            break   # zone libre trouvée, quitter la boucle orientations aussi
        else:
            # Aucune zone libre directe : essayer les candidats triés
            candidates.sort(key=lambda x: (x[0], x[1]))
            placed_ok = False
            for nb_blk, mob, r, c, bh, bw, blk in candidates:
                if placed_ok:
                    break
                blk_sorted = sorted(blk, key=movability_score)
                new_placed = try_relocate_group(grid_orig, current_placed, blk_sorted)
                if new_placed is None:
                    continue
                test_grid = rebuild_grid(grid_orig, new_placed)
                if not test_grid[r:r+bh, c:c+bw].all():
                    continue
                # Succès
                current_placed = new_placed
                current_placed.append({**target, 'row': r, 'col': c,
                                       'h': bh, 'w': bw})
                still_unplaced.remove(target)
                placed_ok = True
                # Cascade dans la zone libérée
                for other in still_unplaced[:]:
                    grid_cur2 = rebuild_grid(grid_orig, current_placed)
                    pos2 = find_any_pos(grid_cur2, rows, cols,
                                        other['longueur'], other['largeur'])
                    if pos2:
                        r2, c2, h2, w2 = pos2
                        current_placed.append({**other, 'row': r2, 'col': c2,
                                               'h': h2, 'w': w2})
                        still_unplaced.remove(other)

    return current_placed, still_unplaced



# ═══════════════════════════════════════════════════════════════════════════════
# PASSE DE DÉCALAGES COORDONNÉS PAR PAIRES
# ═══════════════════════════════════════════════════════════════════════════════

def coordinated_shift_pass(grid_orig, placed, unplaced):
    """
    Teste toutes les paires de bâtiments (A, B) en décalant chacun dans des
    directions opposées ou complémentaires pour créer une zone contiguë libre
    suffisante pour un bâtiment non placé.

    Exemple : GTA décalé à gauche + CE décalé à droite libèrent une bande
    centrale que ni l'un ni l'autre ne pouvait créer seul.

    Pour chaque paire (A, B) candidates (triées par mobilité) :
      - Retirer A et B de la grille
      - Tester toutes combinaisons de directions/distances pour A et B
      - Si target peut se loger dans la zone libérée -> appliquer
    """
    MAX_SHIFT   = 6   # décalage max en cases
    rows, cols  = grid_orig.shape
    still_unplaced = sorted(unplaced, key=lambda b: -(b['longueur']*b['largeur']))
    current_placed = list(placed)

    for target in still_unplaced[:]:
        # Essai direct d'abord
        grid_cur = rebuild_grid(grid_orig, current_placed)
        pos = find_any_pos(grid_cur, rows, cols, target['longueur'], target['largeur'])
        if pos:
            r, c, h, w = pos
            current_placed.append({**target, 'row': r, 'col': c, 'h': h, 'w': w})
            still_unplaced.remove(target)
            continue

        # Candidats mobiles triés par score (neutres en premier)
        candidates = sorted(current_placed, key=movability_score)

        placed_ok = False
        # Paires de bâtiments
        for i, ca in enumerate(candidates):
            if placed_ok: break
            for cb in candidates[i+1:i+15]:   # limiter à 14 partenaires par candidat
                if placed_ok: break

                ca_r,ca_c,ca_h,ca_w = ca['row'],ca['col'],ca['h'],ca['w']
                cb_r,cb_c,cb_h,cb_w = cb['row'],cb['col'],cb['h'],cb['w']

                # Grille sans les deux
                base_placed = [p for p in current_placed if p is not ca and p is not cb]
                base_grid   = rebuild_grid(grid_orig, base_placed)

                # Directions à tester pour chaque bâtiment
                DIRS = [(0,d) for d in range(-MAX_SHIFT, MAX_SHIFT+1) if d != 0] +                        [(d,0) for d in range(-MAX_SHIFT, MAX_SHIFT+1) if d != 0]

                for (dra, dca) in DIRS:
                    if placed_ok: break
                    na_r, na_c = ca_r+dra, ca_c+dca
                    if na_r < 0 or na_c < 0 or na_r+ca_h > rows or na_c+ca_w > cols: continue
                    if not grid_orig[na_r:na_r+ca_h, na_c:na_c+ca_w].all(): continue
                    if not base_grid[na_r:na_r+ca_h, na_c:na_c+ca_w].all(): continue

                    # Poser A à sa nouvelle position
                    grid_a = base_grid.copy()
                    do_place(grid_a, na_r, na_c, ca_h, ca_w)

                    for (drb, dcb) in DIRS:
                        if placed_ok: break
                        nb_r, nb_c = cb_r+drb, cb_c+dcb
                        if nb_r < 0 or nb_c < 0 or nb_r+cb_h > rows or nb_c+cb_w > cols: continue
                        if not grid_orig[nb_r:nb_r+cb_h, nb_c:nb_c+cb_w].all(): continue
                        if not grid_a[nb_r:nb_r+cb_h, nb_c:nb_c+cb_w].all(): continue

                        # Poser B à sa nouvelle position
                        grid_ab = grid_a.copy()
                        do_place(grid_ab, nb_r, nb_c, cb_h, cb_w)

                        # Chercher position pour target
                        pos_t = find_any_pos(grid_ab, rows, cols,
                                             target['longueur'], target['largeur'])
                        if pos_t is None: continue

                        # Succès : appliquer les trois placements
                        rt, ct, ht, wt = pos_t
                        current_placed = base_placed
                        current_placed.append({**ca, 'row': na_r, 'col': na_c,
                                               'h': ca_h, 'w': ca_w})
                        current_placed.append({**cb, 'row': nb_r, 'col': nb_c,
                                               'h': cb_h, 'w': cb_w})
                        current_placed.append({**target, 'row': rt, 'col': ct,
                                               'h': ht, 'w': wt})
                        still_unplaced.remove(target)
                        placed_ok = True
                        # Cascade
                        for other in still_unplaced[:]:
                            g2 = rebuild_grid(grid_orig, current_placed)
                            p2 = find_any_pos(g2, rows, cols,
                                              other['longueur'], other['largeur'])
                            if p2:
                                r2,c2,h2,w2 = p2
                                current_placed.append({**other,'row':r2,'col':c2,
                                                       'h':h2,'w':w2})
                                still_unplaced.remove(other)
                        break

    return current_placed, still_unplaced

# ═══════════════════════════════════════════════════════════════════════════════
# PASSE DE DÉCALAGE LATÉRAL
# ═══════════════════════════════════════════════════════════════════════════════

def lateral_shift_pass(grid_orig, placed, unplaced):
    """
    Deux niveaux de décalage pour consolider les espaces fragmentés :

    Niveau 1 — décalage simple (1 bâtiment, 1..MAX_SHIFT cases).
    Niveau 2 — décalage par paire (2 bâtiments ADJACENTS à des zones libres,
                1..MAX_SHIFT_PAIR cases chacun).
                Cas typique : GTA gauche + Champs Élysées droite libèrent
                une zone centrale contiguë.

    Optimisations pour rester rapide :
    - Niveau 2 : MAX_SHIFT_PAIR réduit à 2
    - Niveau 2 : on ne teste que les bâtiments qui touchent une case libre
                 (filtrage par adjacence)
    """
    MAX_SHIFT      = 5
    MAX_SHIFT_PAIR = 2
    rows, cols = grid_orig.shape
    still_unplaced = sorted(unplaced, key=lambda b: -(b['longueur']*b['largeur']))
    current_placed = list(placed)

    def try_place_target(grid_state, target):
        return find_any_pos(grid_state, rows, cols,
                            target['longueur'], target['largeur'])

    def apply_cascade(current, still):
        for other in still[:]:
            g2 = rebuild_grid(grid_orig, current)
            p2 = find_any_pos(g2, rows, cols, other['longueur'], other['largeur'])
            if p2:
                r2,c2,h2,w2 = p2
                current.append({**other,'row':r2,'col':c2,'h':h2,'w':w2})
                still.remove(other)

    def touches_free(p, grid_cur):
        """Vrai si au moins une case voisine du bâtiment est libre dans grid_cur."""
        r0,c0,h,w = p['row'],p['col'],p['h'],p['w']
        for r in range(max(0,r0-1), min(rows,r0+h+1)):
            for c in range(max(0,c0-1), min(cols,c0+w+1)):
                if r0<=r<r0+h and c0<=c<c0+w: continue
                if grid_cur[r][c]: return True
        return False

    for target in still_unplaced[:]:
        # Essai direct
        grid_cur = rebuild_grid(grid_orig, current_placed)
        pos = try_place_target(grid_cur, target)
        if pos:
            r,c,h,w = pos
            current_placed.append({**target,'row':r,'col':c,'h':h,'w':w})
            still_unplaced.remove(target)
            apply_cascade(current_placed, still_unplaced)
            continue

        candidates = sorted(current_placed, key=movability_score)
        placed_ok = False

        # ── Niveau 1 : un seul bâtiment ───────────────────────────────────
        for shift_size in range(1, MAX_SHIFT + 1):
            if placed_ok: break
            for cand in candidates:
                if placed_ok: break
                cr,cc,ch,cw = cand['row'],cand['col'],cand['h'],cand['w']
                for dr,dc in [(-shift_size,0),(shift_size,0),
                               (0,-shift_size),(0,shift_size)]:
                    nr,nc = cr+dr, cc+dc
                    if nr<0 or nc<0 or nr+ch>rows or nc+cw>cols: continue
                    if not grid_orig[nr:nr+ch, nc:nc+cw].all(): continue
                    without = [p for p in current_placed if p is not cand]
                    tg = rebuild_grid(grid_orig, without)
                    if not tg[nr:nr+ch, nc:nc+cw].all(): continue
                    do_place(tg, nr, nc, ch, cw)
                    pos_t = try_place_target(tg, target)
                    if pos_t is None: continue
                    current_placed = without
                    current_placed.append({**cand,'row':nr,'col':nc,'h':ch,'w':cw})
                    rt,ct,ht,wt = pos_t
                    current_placed.append({**target,'row':rt,'col':ct,'h':ht,'w':wt})
                    still_unplaced.remove(target)
                    apply_cascade(current_placed, still_unplaced)
                    placed_ok = True
                    break

        if placed_ok: continue

        # ── Niveau 2 : deux bâtiments coordonnés ──────────────────────────
        # Ne tester que les bâtiments qui touchent des cases libres
        grid_cur = rebuild_grid(grid_orig, current_placed)
        adj_candidates = [p for p in candidates if touches_free(p, grid_cur)][:20]

        t_pair_start = time.time()
        for i, cand1 in enumerate(adj_candidates):
            if placed_ok: break
            if time.time() - t_pair_start > 3.0: break   # max 3s par bâtiment non-placé
            c1r,c1c,c1h,c1w = cand1['row'],cand1['col'],cand1['h'],cand1['w']
            for cand2 in adj_candidates[i+1:]:
                if placed_ok: break
                c2r,c2c,c2h,c2w = cand2['row'],cand2['col'],cand2['h'],cand2['w']
                # Filtrer par distance : seulement bâtiments proches
                dist = abs(c1r-c2r) + abs(c1c-c2c)
                if dist > (c1h+c1w+c2h+c2w) * 2: continue

                without = [p for p in current_placed
                           if p is not cand1 and p is not cand2]

                for sh1 in range(1, MAX_SHIFT_PAIR + 1):
                    if placed_ok: break
                    for dr1,dc1 in [(-sh1,0),(sh1,0),(0,-sh1),(0,sh1)]:
                        if placed_ok: break
                        nr1,nc1 = c1r+dr1, c1c+dc1
                        if nr1<0 or nc1<0 or nr1+c1h>rows or nc1+c1w>cols: continue
                        if not grid_orig[nr1:nr1+c1h, nc1:nc1+c1w].all(): continue

                        for sh2 in range(1, MAX_SHIFT_PAIR + 1):
                            if placed_ok: break
                            for dr2,dc2 in [(-sh2,0),(sh2,0),(0,-sh2),(0,sh2)]:
                                nr2,nc2 = c2r+dr2, c2c+dc2
                                if nr2<0 or nc2<0 or nr2+c2h>rows or nc2+c2w>cols: continue
                                if not grid_orig[nr2:nr2+c2h, nc2:nc2+c2w].all(): continue

                                tg = rebuild_grid(grid_orig, without)
                                if not tg[nr1:nr1+c1h, nc1:nc1+c1w].all(): continue
                                do_place(tg, nr1, nc1, c1h, c1w)
                                if not tg[nr2:nr2+c2h, nc2:nc2+c2w].all(): continue
                                do_place(tg, nr2, nc2, c2h, c2w)

                                pos_t = try_place_target(tg, target)
                                if pos_t is None: continue

                                # Succès
                                current_placed = without
                                current_placed.append({**cand1,'row':nr1,'col':nc1,'h':c1h,'w':c1w})
                                current_placed.append({**cand2,'row':nr2,'col':nc2,'h':c2h,'w':c2w})
                                rt,ct,ht,wt = pos_t
                                current_placed.append({**target,'row':rt,'col':ct,'h':ht,'w':wt})
                                still_unplaced.remove(target)
                                apply_cascade(current_placed, still_unplaced)
                                placed_ok = True
                                break

    return current_placed, still_unplaced

def cultural_reposition_pass(grid_orig, placed):
    """
    Repositionne les petits sites culturels (compacts, réduits, moyens) pour :
    1. Éliminer les sites ne couvrant aucun producteur (gaspillage)
    2. Maximiser le nombre de producteurs franchissant un nouveau seuil de boost

    Score d'une position = somme sur tous les producteurs couverts de :
      - 10000 * nb_nouveaux_seuils_débloqués  (franchissement de 25%/50%/100%)
      - 1000  si un seuil est franchi
      - culture_site / seuil_prochain          (fraction d'avancement)
    """
    rows, cols = grid_orig.shape
    current_placed = list(placed)

    # Recalculer la culture de chaque producteur sur la base du placed courant
    def get_producer_culture_from_map(cmap, p):
        r0,c0,h,w = p['row'],p['col'],p['h'],p['w']
        foot = cmap[r0:r0+h, c0:c0+w]
        return float(foot.max()) if foot.size > 0 else 0.0

    def score_site_at(site, nr, nc, nh, nw, current_placed_without_site):
        """Score d'un site culturel placé en (nr,nc,nh,nw)."""
        ray = site['rayonnement']
        rmin=max(0,nr-ray); rmax=min(rows,nr+nh+ray)
        cmin=max(0,nc-ray); cmax=min(cols,nc+nw+ray)

        # Culture map sans ce site + avec lui à la nouvelle position
        fake_site = {**site, 'row':nr, 'col':nc, 'h':nh, 'w':nw}
        cmap_new = build_culture_map(current_placed_without_site + [fake_site], rows, cols)
        cmap_old = build_culture_map(current_placed_without_site, rows, cols)

        score = 0
        for pp in current_placed_without_site:
            if pp['type'] != 'Producteur': continue
            pr,pc,ph,pw = pp['row'],pp['col'],pp['h'],pp['w']
            if not (pr<rmax and pr+ph>rmin and pc<cmax and pc+pw>cmin):
                continue   # hors portée, pas d'impact direct
            c_old = get_producer_culture_from_map(cmap_old, pp)
            c_new = get_producer_culture_from_map(cmap_new, pp)
            if c_new <= c_old: continue

            # Évaluer franchissements de seuils
            for seuil, weight in [(pp.get('boost25'), 1000),
                                  (pp.get('boost50'), 3000),
                                  (pp.get('boost100'), 8000)]:
                if seuil is None: continue
                if c_old < seuil <= c_new:
                    score += weight
                elif c_new < seuil:
                    # Pas encore franchi mais on s'en rapproche
                    score += weight * (c_new - c_old) / seuil
        return score

    # Sites candidats au repositionnement (petits : culture ≤ 1300)
    small_cultural = [p for p in current_placed
                      if p['type'] == 'Culturel' and p.get('culture', 0) <= 1300]

    for site in small_cultural:
        sh, sw = site['h'], site['w']
        placed_without = [p for p in current_placed if p is not site]

        # Score actuel
        score_cur = score_site_at(site, site['row'], site['col'], sh, sw, placed_without)

        best_score = score_cur
        best_pos   = None

        # Explorer toutes les positions libres du terrain
        # Grille sans ce site
        test_grid = rebuild_grid(grid_orig, placed_without)

        for bh, bw in orientations(sh, sw):
            for r in range(rows - bh + 1):
                for c in range(cols - bw + 1):
                    if not grid_orig[r:r+bh, c:c+bw].all(): continue
                    if not test_grid[r:r+bh, c:c+bw].all(): continue
                    if r == site['row'] and c == site['col']: continue  # même position

                    s = score_site_at(site, r, c, bh, bw, placed_without)
                    if s > best_score:
                        best_score = s
                        best_pos   = (r, c, bh, bw)

        if best_pos is not None:
            r, c, h, w = best_pos
            # Mettre à jour le site dans current_placed
            idx = current_placed.index(site)
            current_placed[idx] = {**site, 'row': r, 'col': c, 'h': h, 'w': w}

    return current_placed

# ═══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL
# ═══════════════════════════════════════════════════════════════════════════════

def generate_excel(placed, unplaced, free_cells, grid_matrix, terrain_grid):
    wb     = Workbook()
    rows_n, cols_n = grid_matrix.shape
    HDR_FILL = PatternFill("solid", fgColor="1F4E79")

    # ── Onglet 1 : Bâtiments placés ──────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Batiments places"
    hdr = ["Nom","Type","Production","Ligne","Colonne","Hauteur","Largeur",
           "Culture recue","Boost (%)","Quantite/h","Prod totale/h"]
    for i, h in enumerate(hdr, 1):
        cell = ws1.cell(row=1, column=i, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HDR_FILL
    for ri, p in enumerate(placed, 2):
        boost_pct, mult = get_boost(p)
        prod = p['quantite'] * mult if p['production'] != 'Rien' else 0
        vals = [p['nom'], p['type'], p['production'],
                p['row']+1, p['col']+1, p['h'], p['w'],
                round(p.get('culture_recue', 0), 1), boost_pct,
                p['quantite'], round(prod, 1)]
        for ci, v in enumerate(vals, 1):
            ws1.cell(row=ri, column=ci, value=v)
    for col in ws1.columns:
        ws1.column_dimensions[get_column_letter(col[0].column)].width = 22

    # ── Onglet 2 : Synthèse ───────────────────────────────────────────────────
    ws2 = wb.create_sheet("Synthese")
    ws2["A1"] = "Synthese par type de production"
    ws2["A1"].font = Font(bold=True, size=13)
    hdr2 = ["Production","Culture totale","Boost moyen (%)","Nb batiments","Production/h"]
    for i, h in enumerate(hdr2, 1):
        cell = ws2.cell(row=2, column=i, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HDR_FILL
    prod_groups = {}
    for p in placed:
        if p['type'] == 'Producteur' and p['production'] != 'Rien':
            key = p['production']
            boost_pct, mult = get_boost(p)
            g = prod_groups.setdefault(key, {'cultures':[],'boosts':[],'qtys':[]})
            g['cultures'].append(p.get('culture_recue', 0))
            g['boosts'].append(boost_pct)
            g['qtys'].append(p['quantite'] * mult)
    def pg_key(k):
        return PROD_PRIORITY.index(k) if k in PROD_PRIORITY else 999
    row = 3
    for key in sorted(prod_groups, key=pg_key):
        g = prod_groups[key]
        ws2.cell(row=row, column=1, value=key)
        ws2.cell(row=row, column=2, value=round(sum(g['cultures']), 1))
        ws2.cell(row=row, column=3, value=round(sum(g['boosts'])/len(g['boosts']), 1))
        ws2.cell(row=row, column=4, value=len(g['qtys']))
        ws2.cell(row=row, column=5, value=round(sum(g['qtys']), 1))
        row += 1
    for col in ws2.columns:
        ws2.column_dimensions[get_column_letter(col[0].column)].width = 25

    # ── Onglet 3 : Terrain visuel ─────────────────────────────────────────────
    ws3 = wb.create_sheet("Terrain")

    FILL_X = PatternFill("solid", fgColor="404040")
    FILL_C = PatternFill("solid", fgColor="FFA500")
    FILL_P = PatternFill("solid", fgColor="70AD47")
    FILL_N = PatternFill("solid", fgColor="BFBFBF")
    FILL_F = PatternFill("solid", fgColor="FFFFFF")

    x_set = set()
    for ri, row_data in enumerate(terrain_grid):
        for ci, v in enumerate(row_data):
            if v == 'X':
                x_set.add((ri, ci))

    cell_map = {}
    for p in placed:
        r0, c0, h, w = p['row'], p['col'], p['h'], p['w']
        for dr in range(h):
            for dc in range(w):
                cell_map.setdefault((r0+dr, c0+dc), p)

    for r in range(rows_n):
        ws3.row_dimensions[r+1].height = 20
    for c in range(cols_n):
        ws3.column_dimensions[get_column_letter(c+1)].width = 4

    for r in range(rows_n):
        for c in range(cols_n):
            cell = ws3.cell(row=r+1, column=c+1)
            cell.alignment = Alignment(wrap_text=True,
                                       horizontal='center', vertical='center')
            if (r, c) in x_set:
                cell.fill  = FILL_X
                cell.value = "X"
                cell.font  = Font(color="FFFFFF", size=6, bold=True)
            elif (r, c) in cell_map:
                p = cell_map[(r, c)]
                if p['type'] == 'Culturel':
                    cell.fill = FILL_C
                elif p['type'] == 'Producteur':
                    cell.fill = FILL_P
                else:
                    cell.fill = FILL_N
                if r == p['row'] and c == p['col']:
                    boost_pct, _ = get_boost(p)
                    label = p['nom']
                    if p['type'] == 'Producteur' and boost_pct > 0:
                        label += f"\n+{boost_pct}%"
                    cell.value = label
                    cell.font  = Font(bold=True, size=7)
            else:
                cell.fill = FILL_F

    # Fusionner les cellules de chaque bâtiment
    merged = set()
    for p in placed:
        r0, c0, h, w = p['row'], p['col'], p['h'], p['w']
        if h == 1 and w == 1:
            continue
        key = (r0, c0, h, w)
        if key in merged:
            continue
        merged.add(key)
        ws3.merge_cells(start_row=r0+1, start_column=c0+1,
                        end_row=r0+h,   end_column=c0+w)

    # ── Onglet 4 : Non placés ─────────────────────────────────────────────────
    ws4 = wb.create_sheet("Non places")
    ws4["A1"] = "Batiments non places"
    ws4["A1"].font = Font(bold=True, size=13)
    for i, h in enumerate(["Nom","Type","Taille","Production"], 1):
        cell = ws4.cell(row=2, column=i, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HDR_FILL
    total_cells = 0
    for ri, b in enumerate(unplaced, 3):
        ws4.cell(row=ri, column=1, value=b['nom'])
        ws4.cell(row=ri, column=2, value=b['type'])
        ws4.cell(row=ri, column=3, value=f"{b['longueur']}x{b['largeur']}")
        ws4.cell(row=ri, column=4, value=b['production'])
        total_cells += b['longueur'] * b['largeur']
    end = len(unplaced) + 4
    ws4.cell(row=end,   column=1, value="Cases libres restantes:").font   = Font(bold=True)
    ws4.cell(row=end,   column=2, value=free_cells)
    ws4.cell(row=end+1, column=1, value="Cases batiments non places:").font = Font(bold=True)
    ws4.cell(row=end+1, column=2, value=total_cells)
    for col in ws4.columns:
        ws4.column_dimensions[get_column_letter(col[0].column)].width = 30

    return wb

# ═══════════════════════════════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ═══════════════════════════════════════════════════════════════════════════════

uploaded = st.file_uploader("Charger le fichier Excel de la ville", type=["xlsx"])

if uploaded:
    wb_in = load_workbook(io.BytesIO(uploaded.read()), read_only=True)
    if 'Terrain' not in wb_in.sheetnames or 'Batiments' not in wb_in.sheetnames:
        st.error("Le fichier doit contenir les onglets 'Terrain' et 'Batiments'")
        st.stop()

    terrain_grid   = load_terrain(wb_in['Terrain'])
    buildings_raw  = load_buildings(wb_in['Batiments'])
    grid_matrix    = grid_to_matrix(terrain_grid)
    rows_n, cols_n = grid_matrix.shape

    st.success(f"Terrain charge : {rows_n} lignes x {cols_n} colonnes  |  "
               f"{int(grid_matrix.sum())} cases libres")

    types_count = {}
    for b in buildings_raw:
        types_count[b['type']] = types_count.get(b['type'], 0) + b['nombre']
    c1, c2, c3 = st.columns(3)
    c1.metric("Batiments neutres",     types_count.get('Neutre', 0))
    c2.metric("Batiments culturels",   types_count.get('Culturel', 0))
    c3.metric("Batiments producteurs", types_count.get('Producteur', 0))

    if st.button("Lancer le placement optimise", type="primary"):
        with st.spinner("Optimisation et compaction en cours..."):
            placed, unplaced, free_cells = run_placement(grid_matrix, buildings_raw)

        col_a, col_b, col_c = st.columns(3)
        col_a.metric("Bâtiments placés",     len(placed))
        col_b.metric("Bâtiments non placés", len(unplaced))
        col_c.metric("Cases libres",         free_cells)

        st.subheader("Synthese des productions")
        prod_data = {}
        for p in placed:
            if p['type'] == 'Producteur' and p['production'] != 'Rien':
                key = p['production']
                boost_pct, mult = get_boost(p)
                d = prod_data.setdefault(key, {'cultures':[],'boosts':[],'prod':0,'n':0})
                d['cultures'].append(p.get('culture_recue', 0))
                d['boosts'].append(boost_pct)
                d['prod'] += p['quantite'] * mult
                d['n']    += 1
        rows_list = []
        for k in sorted(prod_data,
                        key=lambda x: PROD_PRIORITY.index(x) if x in PROD_PRIORITY else 999):
            d = prod_data[k]
            rows_list.append({
                "Production":      k,
                "Culture totale":  round(sum(d['cultures']), 0),
                "Boost moyen (%)": round(sum(d['boosts'])/len(d['boosts']), 1),
                "Nb batiments":    d['n'],
                "Production/h":    round(d['prod'], 1),
            })
        if rows_list:
            st.dataframe(pd.DataFrame(rows_list), hide_index=True, use_container_width=True)

        if unplaced:
            st.subheader(f"{len(unplaced)} batiments non places")
            st.dataframe(pd.DataFrame([{
                'Nom':        b['nom'],
                'Type':       b['type'],
                'Taille':     f"{b['longueur']}x{b['largeur']}",
                'Production': b['production'],
            } for b in unplaced]), hide_index=True)

        wb_out = generate_excel(placed, unplaced, free_cells, grid_matrix, terrain_grid)
        buf = io.BytesIO()
        wb_out.save(buf)
        buf.seek(0)
        st.download_button(
            label="Telecharger les resultats Excel",
            data=buf,
            file_name="resultats_placement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
