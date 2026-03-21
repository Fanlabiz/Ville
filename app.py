import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import io, copy, itertools

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
        try:
            qty_raw = row[11]
            if qty_raw is None:
                qty = 0.0
            else:
                s = str(qty_raw).strip().lstrip('=')
                try:
                    qty = float(eval(s))
                except Exception:
                    qty = 0.0
            b = {
                'nom':        str(row[0]).strip(),
                'longueur':   int(row[1]),
                'largeur':    int(row[2]),
                'nombre':     int(row[3]),
                'type':       str(row[4]).strip(),
                'culture':    float(row[5]) if row[5] else 0,
                'rayonnement':int(row[6]) if row[6] else 0,
                'boost25':    float(row[7]) if row[7] is not None else None,
                'boost50':    float(row[8]) if row[8] is not None else None,
                'boost100':   float(row[9]) if row[9] is not None else None,
                'production': str(row[10]).strip() if row[10] else 'Rien',
                'quantite':   qty,
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

def do_remove(grid, r, c, h, w):
    grid[r:r+h, c:c+w] = True

def rebuild_grid(grid_orig, placed):
    """Reconstruit la grille libre à partir de la grille originale et des placements actuels."""
    grid = grid_orig.copy()
    for p in placed:
        do_place(grid, p['row'], p['col'], p['h'], p['w'])
    return grid

def orientations(b):
    """Retourne les orientations possibles (h, w) sans doublon."""
    return list(dict.fromkeys([(b['longueur'], b['largeur']),
                                (b['largeur'],  b['longueur'])]))

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
    for bh, bw in orientations({'longueur': h, 'largeur': w}):
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
    for bh, bw in orientations({'longueur': h, 'largeur': w}):
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
    for bh, bw in orientations({'longueur': h, 'largeur': w}):
        for r in range(rows):
            for c in range(cols):
                if can_place(grid, r, c, bh, bw):
                    val = float(cmap[r:r+bh, c:c+bw].max())
                    if val > best_val or best is None:
                        best_val = val
                        best = (r, c, bh, bw)
    return best

def find_any_pos(grid, rows, cols, h, w):
    for bh, bw in orientations({'longueur': h, 'largeur': w}):
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
    grid  = grid_orig.copy()
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
            b = culturels[ci]
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
            b = producteurs[pi]
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

    # 3. Passe de compaction : tente de recaser les non-placés
    placed, unplaced = compaction_pass(grid_orig, placed, unplaced)

    compute_culture_received(placed, rows, cols)
    free_cells = int(rebuild_grid(grid_orig, placed).sum())
    return placed, unplaced, free_cells

# ═══════════════════════════════════════════════════════════════════════════════
# PASSE DE COMPACTION
# ═══════════════════════════════════════════════════════════════════════════════

def compaction_pass(grid_orig, placed, unplaced):
    """
    Pour chaque bâtiment non placé, on cherche si en déplaçant un ou deux
    bâtiments déjà placés (de moindre priorité ou plus petits) on peut
    libérer assez de cases contiguës pour le caser.

    Stratégie :
      - Trier les non-placés du plus grand au plus petit.
      - Pour chacun, essayer de trouver une position en replaçant d'abord
        les bâtiments amovibles dans d'autres endroits libres.
      - On ne déplace que des bâtiments de même type ou de priorité inférieure.
    """
    rows, cols = grid_orig.shape
    still_unplaced = list(unplaced)
    still_unplaced.sort(key=lambda b: -(b['longueur'] * b['largeur']))

    newly_placed  = []

    for target in still_unplaced[:]:          # copie pour itérer pendant modification
        # Reconstruire la grille courante
        current_placed = placed + newly_placed
        grid = rebuild_grid(grid_orig, current_placed)

        # Essai direct d'abord (peut-être que des déplacements précédents ont libéré de la place)
        pos = find_any_pos(grid, rows, cols, target['longueur'], target['largeur'])
        if pos:
            r, c, h, w = pos
            newly_placed.append({**target, 'row': r, 'col': c, 'h': h, 'w': w})
            do_place(grid, r, c, h, w)
            still_unplaced.remove(target)
            continue

        # Tentative de déplacement de bâtiments candidats
        # Candidats : bâtiments de taille >= target qui pourraient libérer de l'espace
        # On préfère déplacer les neutres, puis les culturels, puis les producteurs de moindre priorité
        def movability_key(p):
            t = p['type']
            if t == 'Neutre':      return 0
            if t == 'Culturel':    return 1
            prod = p.get('production', 'Rien')
            if prod == 'Rien':     return 2
            if prod not in PROD_PRIORITY: return 3
            return 4 + PROD_PRIORITY.index(prod)

        candidates = sorted(current_placed, key=movability_key)

        moved = False
        for cand in candidates:
            # Simuler le retrait du candidat
            test_placed = [p for p in current_placed if p is not cand]
            test_grid   = rebuild_grid(grid_orig, test_placed)

            # Y a-t-il maintenant une place pour target ?
            pos_t = find_any_pos(test_grid, rows, cols, target['longueur'], target['largeur'])
            if pos_t is None:
                continue

            # Peut-on replaçer le candidat ailleurs ?
            do_place(test_grid, *pos_t)          # occupe provisoirement la place de target
            pos_c = find_any_pos(test_grid, rows, cols, cand['h'], cand['w'])
            if pos_c is None:
                # Essai avec l'autre orientation
                pos_c = find_any_pos(test_grid, rows, cols, cand['w'], cand['h'])
            if pos_c is None:
                continue                          # impossible de recaser le candidat

            # Appliquer les deux déplacements sur la liste principale
            rt, ct, ht, wt = pos_t
            rc, cc, hc, wc = pos_c

            # Retirer l'ancien candidat de placed / newly_placed
            if cand in placed:
                placed.remove(cand)
                placed.append({**cand, 'row': rc, 'col': cc, 'h': hc, 'w': wc})
            else:
                newly_placed.remove(cand)
                newly_placed.append({**cand, 'row': rc, 'col': cc, 'h': hc, 'w': wc})

            newly_placed.append({**target, 'row': rt, 'col': ct, 'h': ht, 'w': wt})
            still_unplaced.remove(target)
            moved = True
            break

        if not moved:
            # Tentative plus agressive : déplacer 2 bâtiments simultanément
            for cand1, cand2 in itertools.combinations(candidates[:12], 2):
                test_placed = [p for p in current_placed
                               if p is not cand1 and p is not cand2]
                test_grid   = rebuild_grid(grid_orig, test_placed)

                pos_t = find_any_pos(test_grid, rows, cols,
                                     target['longueur'], target['largeur'])
                if pos_t is None:
                    continue

                do_place(test_grid, *pos_t)

                pos_c1 = find_any_pos(test_grid, rows, cols, cand1['h'], cand1['w'])
                if pos_c1 is None:
                    pos_c1 = find_any_pos(test_grid, rows, cols, cand1['w'], cand1['h'])
                if pos_c1 is None:
                    continue
                do_place(test_grid, *pos_c1)

                pos_c2 = find_any_pos(test_grid, rows, cols, cand2['h'], cand2['w'])
                if pos_c2 is None:
                    pos_c2 = find_any_pos(test_grid, rows, cols, cand2['w'], cand2['h'])
                if pos_c2 is None:
                    continue

                rt, ct, ht, wt   = pos_t
                rc1,cc1,hc1,wc1  = pos_c1
                rc2,cc2,hc2,wc2  = pos_c2

                for lst in (placed, newly_placed):
                    if cand1 in lst:
                        lst.remove(cand1)
                        lst.append({**cand1,'row':rc1,'col':cc1,'h':hc1,'w':wc1})
                    if cand2 in lst:
                        lst.remove(cand2)
                        lst.append({**cand2,'row':rc2,'col':cc2,'h':hc2,'w':wc2})

                newly_placed.append({**target,'row':rt,'col':ct,'h':ht,'w':wt})
                still_unplaced.remove(target)
                moved = True
                break

    placed_final = placed + newly_placed
    return placed_final, still_unplaced

# ═══════════════════════════════════════════════════════════════════════════════
# EXPORT EXCEL
# ═══════════════════════════════════════════════════════════════════════════════

def generate_excel(placed, unplaced, free_cells, grid_matrix, terrain_grid):
    wb   = Workbook()
    rows_n, cols_n = grid_matrix.shape

    # ── Onglet 1 : Bâtiments placés ──────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Batiments places"
    hdr = ["Nom","Type","Production","Ligne","Colonne","Hauteur","Largeur",
           "Culture recue","Boost (%)","Quantite/h","Prod totale/h"]
    HDR_FILL = PatternFill("solid", fgColor="1F4E79")
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
    FILL_C = PatternFill("solid", fgColor="FFA500")   # orange
    FILL_P = PatternFill("solid", fgColor="70AD47")   # vert
    FILL_N = PatternFill("solid", fgColor="BFBFBF")   # gris
    FILL_F = PatternFill("solid", fgColor="FFFFFF")   # libre

    # X positions
    x_set = set()
    for ri, row_data in enumerate(terrain_grid):
        for ci, v in enumerate(row_data):
            if v == 'X':
                x_set.add((ri, ci))

    # Carte bâtiment par cellule
    cell_map = {}
    for p in placed:
        r0, c0, h, w = p['row'], p['col'], p['h'], p['w']
        for dr in range(h):
            for dc in range(w):
                cell_map.setdefault((r0+dr, c0+dc), p)

    # Dimensionner les cellules (colonne étroite, ligne haute)
    COL_W = 4   # largeur colonne en unités Excel
    ROW_H = 20  # hauteur ligne en points
    for r in range(rows_n):
        ws3.row_dimensions[r+1].height = ROW_H
    for c in range(cols_n):
        ws3.column_dimensions[get_column_letter(c+1)].width = COL_W

    # Remplir chaque cellule
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
                # Écrire le label uniquement dans la cellule en haut à gauche du bâtiment
                if r == p['row'] and c == p['col']:
                    boost_pct, _ = get_boost(p)
                    label = p['nom']
                    if p['type'] == 'Producteur' and boost_pct > 0:
                        label += f"\n+{boost_pct}%"
                    cell.value = label
                    cell.font  = Font(bold=True, size=7,
                                      color="000000" if p['type'] != 'Culturel' else "000000")
            else:
                cell.fill = FILL_F

    # Fusionner les cellules de chaque bâtiment pour un rendu propre
    merged = set()
    for p in placed:
        r0, c0, h, w = p['row'], p['col'], p['h'], p['w']
        if h == 1 and w == 1:
            continue    # pas besoin de fusionner
        key = (r0, c0, h, w)
        if key in merged:
            continue
        merged.add(key)
        r1_xl = r0 + 1;   c1_xl = c0 + 1
        r2_xl = r0 + h;   c2_xl = c0 + w
        ws3.merge_cells(start_row=r1_xl, start_column=c1_xl,
                        end_row=r2_xl,   end_column=c2_xl)

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
    ws4.cell(row=end,   column=1, value="Cases libres restantes:")  .font = Font(bold=True)
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
        col_c.metric("Cases libres",          free_cells)

        # Synthèse
        st.subheader("Synthese des productions")
        prod_data = {}
        for p in placed:
            if p['type'] == 'Producteur' and p['production'] != 'Rien':
                key = p['production']
                boost_pct, mult = get_boost(p)
                d = prod_data.setdefault(key, {'cultures':[], 'boosts':[], 'prod':0, 'n':0})
                d['cultures'].append(p.get('culture_recue', 0))
                d['boosts'].append(boost_pct)
                d['prod'] += p['quantite'] * mult
                d['n']    += 1
        rows_list = []
        for k in sorted(prod_data,
                        key=lambda x: PROD_PRIORITY.index(x) if x in PROD_PRIORITY else 999):
            d = prod_data[k]
            rows_list.append({
                "Production":       k,
                "Culture totale":   round(sum(d['cultures']), 0),
                "Boost moyen (%)":  round(sum(d['boosts'])/len(d['boosts']), 1),
                "Nb batiments":     d['n'],
                "Production/h":     round(d['prod'], 1),
            })
        if rows_list:
            st.dataframe(pd.DataFrame(rows_list), hide_index=True, use_container_width=True)

        if unplaced:
            st.subheader(f"{len(unplaced)} batiments non places")
            st.dataframe(pd.DataFrame([{
                'Nom':        b['nom'],
                'Type':       b['type'],
                'Taille':     f"{b['longueur']}x{b['largeur']}",
                'Production': b['production']
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
