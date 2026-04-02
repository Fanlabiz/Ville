import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import copy
import itertools
from collections import deque

st.set_page_config(page_title="Optimiseur de Ville", layout="wide")
st.title("🏙️ Optimiseur de placement de bâtiments")

# ─────────────────────────── PARSING ───────────────────────────

def parse_terrain(ws_or_df):
    """Parse terrain sheet -> grid of cell values, boundary mask.
    Accepts either an openpyxl worksheet (preferred) or a pandas DataFrame.
    Uses openpyxl to correctly expand merged cells."""
    from openpyxl.worksheet.worksheet import Worksheet as OpenpyxlWS
    if isinstance(ws_or_df, OpenpyxlWS):
        rows = ws_or_df.max_row
        cols = ws_or_df.max_column
        grid = [[None]*cols for _ in range(rows)]
        for r in range(1, rows+1):
            for c in range(1, cols+1):
                grid[r-1][c-1] = ws_or_df.cell(r, c).value
        # Expand merged cell values (openpyxl only stores value in top-left)
        for merged_range in ws_or_df.merged_cells.ranges:
            top_left_val = ws_or_df.cell(merged_range.min_row, merged_range.min_col).value
            for r in range(merged_range.min_row, merged_range.max_row+1):
                for c in range(merged_range.min_col, merged_range.max_col+1):
                    grid[r-1][c-1] = top_left_val
    else:
        grid = ws_or_df.values.tolist()
        rows = len(grid)
        cols = max(len(r) for r in grid)
        for r in grid:
            while len(r) < cols:
                r.append(None)

    outside = [[False]*cols for _ in range(rows)]
    q = deque()
    for r in range(rows):
        for c in range(cols):
            val = grid[r][c]
            is_x = isinstance(val, str) and val.strip().upper() == 'X'
            at_edge = (r == 0 or r == rows-1 or c == 0 or c == cols-1)
            if is_x or (at_edge and (val is None or (isinstance(val, float) and np.isnan(val)))):
                if not outside[r][c]:
                    outside[r][c] = True
                    q.append((r, c))
    while q:
        r, c = q.popleft()
        for dr, dc in [(-1,0),(1,0),(0,-1),(0,1)]:
            nr, nc = r+dr, c+dc
            if 0 <= nr < rows and 0 <= nc < cols and not outside[nr][nc]:
                val = grid[nr][nc]
                is_x = isinstance(val, str) and val.strip().upper() == 'X'
                is_empty = val is None or (isinstance(val, float) and np.isnan(val))
                if is_x or is_empty:
                    outside[nr][nc] = True
                    q.append((nr, nc))
    playable = [[False]*cols for _ in range(rows)]
    for r in range(rows):
        for c in range(cols):
            val = grid[r][c]
            is_x = isinstance(val, str) and val.strip().upper() == 'X'
            playable[r][c] = not (is_x or outside[r][c])
    return grid, playable, rows, cols


def parse_buildings_from_terrain(grid, playable, rows, cols, bldg_info, ws_openpyxl=None):
    """Extract placed buildings from terrain.
    If ws_openpyxl is provided (openpyxl worksheet), reads merged cell ranges directly
    for accurate per-instance bounding boxes. Falls back to flood-fill otherwise."""
    def make_bldg(name, r0, c0, h, w):
        info = bldg_info.get(name, {})
        return {
            'name': name, 'row': r0, 'col': c0, 'h': h, 'w': w,
            'type': info.get('type', 'Neutre'), 'culture': info.get('culture', 0),
            'rayonnement': info.get('rayonnement', 0),
            'boost25': info.get('boost25', None), 'boost50': info.get('boost50', None),
            'boost100': info.get('boost100', None),
            'production': info.get('production', 'Rien'), 'quantite': info.get('quantite', 0),
            'priorite': info.get('priorite', 0),
            'lon': info.get('lon', w), 'lar': info.get('lar', h),
        }

    if ws_openpyxl is not None:
        placed = []
        processed = set()
        for mr in ws_openpyxl.merged_cells.ranges:
            r0, c0 = mr.min_row - 1, mr.min_col - 1
            val = ws_openpyxl.cell(mr.min_row, mr.min_col).value
            if val is None or (isinstance(val, str) and val.strip().upper() == 'X'):
                continue
            if not (0 <= r0 < rows and 0 <= c0 < cols and playable[r0][c0]):
                continue
            name = str(val).strip()
            h = mr.max_row - mr.min_row + 1
            w = mr.max_col - mr.min_col + 1
            placed.append(make_bldg(name, r0, c0, h, w))
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    processed.add((r - 1, c - 1))
        # Handle un-merged single cells
        for r in range(rows):
            for c in range(cols):
                if (r, c) in processed or not playable[r][c]:
                    continue
                val = grid[r][c]
                if val is None or (isinstance(val, float) and np.isnan(val)):
                    continue
                if isinstance(val, str) and val.strip().upper() == 'X':
                    continue
                name = str(val).strip()
                placed.append(make_bldg(name, r, c, 1, 1))
                processed.add((r, c))
        return placed

    # Fallback: flood-fill (used when no openpyxl ws available)
    visited = [[False]*cols for _ in range(rows)]
    placed = []
    for r in range(rows):
        for c in range(cols):
            if not playable[r][c] or visited[r][c]:
                continue
            val = grid[r][c]
            if val is None or (isinstance(val, float) and np.isnan(val)):
                continue
            name = str(val).strip()
            cells = []
            q = deque([(r, c)])
            visited[r][c] = True
            while q:
                rr, cc = q.popleft()
                cells.append((rr, cc))
                for dr, dc in [(-1,0),(1,0),(0,-1),(0,1)]:
                    nr, nc = rr+dr, cc+dc
                    if 0 <= nr < rows and 0 <= nc < cols and not visited[nr][nc] and playable[nr][nc]:
                        nval = grid[nr][nc]
                        if nval is not None and not (isinstance(nval, float) and np.isnan(nval)) and str(nval).strip() == name:
                            visited[nr][nc] = True
                            q.append((nr, nc))
            min_r = min(cc[0] for cc in cells)
            max_r = max(cc[0] for cc in cells)
            min_c = min(cc[1] for cc in cells)
            max_c = max(cc[1] for cc in cells)
            placed.append(make_bldg(name, min_r, min_c, max_r-min_r+1, max_c-min_c+1))
    return placed


def parse_bldg_info(df_bat):
    """Parse buildings sheet -> dict name -> info"""
    df = df_bat.copy()
    df.columns = df.iloc[0]
    df = df.iloc[1:].reset_index(drop=True)
    info = {}
    for _, row in df.iterrows():
        name = str(row.get('Nom', '')).strip()
        if not name or name == 'nan':
            continue
        def safe_int(v, default=0):
            try: return int(float(v)) if pd.notna(v) else default
            except: return default
        def safe_float(v):
            try: return float(v) if pd.notna(v) else None
            except: return None
        info[name] = {
            'type': str(row.get('Type', 'Neutre')).strip(),
            'culture': safe_int(row.get('Culture', 0)),
            'rayonnement': safe_int(row.get('Rayonnement', 0)),
            'boost25': safe_float(row.get('Boost 25%')),
            'boost50': safe_float(row.get('Boost 50%')),
            'boost100': safe_float(row.get('Boost 100%')),
            'production': str(row.get('Production', 'Rien')).strip(),
            'quantite': safe_int(row.get('Quantite', 0)),
            'priorite': safe_int(row.get('Priorite', 0)),
            'lon': safe_int(row.get('Longueur', 1), 1),
            'lar': safe_int(row.get('Largeur', 1), 1),
            'nombre': safe_int(row.get('Nombre', 1), 1),
        }
    return info


# ─────────────────────────── CULTURE CALC ───────────────────────────

def compute_culture_received(placed):
    """For each Producteur, compute total culture received from Culturels in range"""
    culturels = [b for b in placed if b['type'] == 'Culturel']
    results = {}
    for b in placed:
        if b['type'] != 'Producteur':
            continue
        total_culture = 0
        contrib = []
        for cult in culturels:
            ray = cult['rayonnement']
            # Zone = band of 'ray' cells around cult building
            cult_r1, cult_c1 = cult['row'], cult['col']
            cult_r2, cult_c2 = cult['row'] + cult['h'] - 1, cult['col'] + cult['w'] - 1
            zone_r1 = cult_r1 - ray
            zone_c1 = cult_c1 - ray
            zone_r2 = cult_r2 + ray
            zone_c2 = cult_c2 + ray
            # Check if producer overlaps with zone (but not with the cultural itself)
            prod_r1, prod_c1 = b['row'], b['col']
            prod_r2, prod_c2 = b['row'] + b['h'] - 1, b['col'] + b['w'] - 1
            overlaps_zone = not (prod_r2 < zone_r1 or prod_r1 > zone_r2 or prod_c2 < zone_c1 or prod_c1 > zone_c2)
            if overlaps_zone:
                total_culture += cult['culture']
                contrib.append(cult['name'])
        results[id(b)] = {'culture': total_culture, 'contrib': contrib}
    return results


def get_boost(b, culture_received):
    """Return boost percentage for a producer"""
    if b['type'] != 'Producteur':
        return 0
    c = culture_received
    b100 = b['boost100']
    b50 = b['boost50']
    b25 = b['boost25']
    if b100 is not None and c >= b100:
        return 100
    if b50 is not None and c >= b50:
        return 50
    if b25 is not None and c >= b25:
        return 25
    return 0


def compute_total_production(placed, culture_map):
    """Compute total production per type"""
    prod_totals = {}
    for b in placed:
        if b['type'] != 'Producteur' or b['production'] == 'Rien' or b['production'] == 'nan':
            continue
        cr = culture_map.get(id(b), {}).get('culture', 0)
        boost = get_boost(b, cr)
        qty = b['quantite'] * (1 + boost / 100)
        prod = b['production']
        prod_totals[prod] = prod_totals.get(prod, 0) + qty
    return prod_totals


PRIORITY_ORDER = ['Guerison', 'Nourriture', 'Or']


def score_placement(placed):
    """Score: priority-weighted production total"""
    culture_map = compute_culture_received(placed)
    prod_totals = compute_total_production(placed, culture_map)
    score = 0
    weights = {'Guerison': 1e12, 'Nourriture': 1e8, 'Or': 1e4}
    for prod, total in prod_totals.items():
        w = weights.get(prod, 1)
        score += total * w
    return score


# ─────────────────────────── OPTIMIZER ───────────────────────────

def build_occupancy(placed, rows, cols):
    occ = [[None]*cols for _ in range(rows)]
    for b in placed:
        for dr in range(b['h']):
            for dc in range(b['w']):
                r, c = b['row']+dr, b['col']+dc
                if 0 <= r < rows and 0 <= c < cols:
                    occ[r][c] = b
    return occ


def can_place(b, row, col, h, w, playable, occ, rows, cols, exclude=None):
    if row < 0 or col < 0 or row+h > rows or col+w > cols:
        return False
    for dr in range(h):
        for dc in range(w):
            r, c = row+dr, col+dc
            if not playable[r][c]:
                return False
            if occ[r][c] is not None and occ[r][c] is not exclude:
                return False
    return True


def _free_cells_of(b, occ):
    """Remove b from occupancy grid temporarily, return cells freed."""
    cells = []
    for dr in range(b['h']):
        for dc in range(b['w']):
            occ[b['row']+dr][b['col']+dc] = None
            cells.append((b['row']+dr, b['col']+dc))
    return cells

def _occupy_cells(b, occ):
    for dr in range(b['h']):
        for dc in range(b['w']):
            occ[b['row']+dr][b['col']+dc] = b

def _move(b, r, c, h, w):
    b['row'], b['col'], b['h'], b['w'] = r, c, h, w

def try_single_move(placed, playable, rows, cols, occ, current_score):
    """Try moving one non-Neutre building to any free position. Return (improved, new_score)."""
    for b in placed:
        if b['type'] == 'Neutre':
            continue
        orig_r, orig_c, orig_h, orig_w = b['row'], b['col'], b['h'], b['w']
        _free_cells_of(b, occ)
        for (h, w) in set([(orig_h, orig_w), (orig_w, orig_h)]):
            for r in range(rows - h + 1):
                for c in range(cols - w + 1):
                    if r == orig_r and c == orig_c and h == orig_h and w == orig_w:
                        continue
                    if can_place(b, r, c, h, w, playable, occ, rows, cols):
                        _move(b, r, c, h, w)
                        s = score_placement(placed)
                        if s > current_score:
                            _occupy_cells(b, occ)
                            return True, s
                        _move(b, orig_r, orig_c, orig_h, orig_w)
        _occupy_cells(b, occ)
    return False, current_score


def try_swap_two(placed, playable, rows, cols, occ, current_score):
    """Try swapping positions of two non-Neutre buildings."""
    non_neutral = [b for b in placed if b['type'] != 'Neutre']
    for i, a in enumerate(non_neutral):
        for b in non_neutral[i+1:]:
            # Swap a and b positions (keep sizes, try both orientations)
            a_r, a_c, a_h, a_w = a['row'], a['col'], a['h'], a['w']
            b_r, b_c, b_h, b_w = b['row'], b['col'], b['h'], b['w']
            # Try all orientation combos
            for (ah, aw) in set([(a_h, a_w), (a_w, a_h)]):
                for (bh, bw) in set([(b_h, b_w), (b_w, b_h)]):
                    # Check if a fits at b's position and vice versa
                    _free_cells_of(a, occ)
                    _free_cells_of(b, occ)
                    a_fits = can_place(a, b_r, b_c, ah, aw, playable, occ, rows, cols)
                    b_fits = can_place(b, a_r, a_c, bh, bw, playable, occ, rows, cols)
                    if a_fits and b_fits:
                        _move(a, b_r, b_c, ah, aw)
                        _move(b, a_r, a_c, bh, bw)
                        s = score_placement(placed)
                        if s > current_score:
                            _occupy_cells(a, occ)
                            _occupy_cells(b, occ)
                            return True, s
                        _move(a, a_r, a_c, a_h, a_w)
                        _move(b, b_r, b_c, b_h, b_w)
                    # Restore
                    _occupy_cells(a, occ)
                    _occupy_cells(b, occ)
    return False, current_score


def try_cluster_move(placed, playable, rows, cols, occ, current_score):
    """For each producer, move it AND culturals towards each other to maximise culture.
    Tests a cluster: pick a producer position, then greedily relocate each cultural
    to the closest free spot within its rayonnement that covers the producer."""
    producers = [b for b in placed if b['type'] == 'Producteur'
                 and b['production'] not in ('Rien', 'nan', '')
                 and b['boost25'] is not None]
    culturels = [b for b in placed if b['type'] == 'Culturel' and b['culture'] > 0]

    best_score = current_score
    best_state = None

    for prod in producers:
        orig_prod_r, orig_prod_c = prod['row'], prod['col']
        # Try every playable position for the producer
        _free_cells_of(prod, occ)
        for pr in range(rows - prod['h'] + 1):
            for pc in range(cols - prod['w'] + 1):
                if not can_place(prod, pr, pc, prod['h'], prod['w'], playable, occ, rows, cols):
                    continue
                # Temporarily place producer here
                _move(prod, pr, pc, prod['h'], prod['w'])
                _occupy_cells(prod, occ)

                # For each cultural, find the best position covering this producer
                moved_culturels = []
                for cult in culturels:
                    orig_cr, orig_cc = cult['row'], cult['col']
                    ray = cult['rayonnement']
                    best_cult_pos = None
                    best_cult_dist = float('inf')
                    _free_cells_of(cult, occ)
                    # Search window: positions that could cover the producer
                    sr_min = max(0, pr - ray - cult['h'] + 1)
                    sr_max = min(rows - cult['h'], pr + prod['h'] + ray - 1)
                    sc_min = max(0, pc - ray - cult['w'] + 1)
                    sc_max = min(cols - cult['w'], pc + prod['w'] + ray - 1)
                    for sr in range(sr_min, sr_max + 1):
                        for sc in range(sc_min, sc_max + 1):
                            if not can_place(cult, sr, sc, cult['h'], cult['w'], playable, occ, rows, cols):
                                continue
                            # Check this position actually covers the producer
                            zr1 = sr - ray; zc1 = sc - ray
                            zr2 = sr + cult['h'] - 1 + ray; zc2 = sc + cult['w'] - 1 + ray
                            if (pr + prod['h'] - 1 < zr1 or pr > zr2 or
                                    pc + prod['w'] - 1 < zc1 or pc > zc2):
                                continue
                            dist = abs(sr - pr) + abs(sc - pc)
                            if dist < best_cult_dist:
                                best_cult_dist = dist
                                best_cult_pos = (sr, sc)
                    if best_cult_pos and best_cult_pos != (orig_cr, orig_cc):
                        _move(cult, best_cult_pos[0], best_cult_pos[1], cult['h'], cult['w'])
                        _occupy_cells(cult, occ)
                        moved_culturels.append((cult, orig_cr, orig_cc))
                    else:
                        _occupy_cells(cult, occ)

                s = score_placement(placed)
                if s > best_score:
                    best_score = s
                    best_state = copy.deepcopy(placed)

                # Restore all moved culturels
                for cult, orig_cr, orig_cc in moved_culturels:
                    _free_cells_of(cult, occ)
                    _move(cult, orig_cr, orig_cc, cult['h'], cult['w'])
                    _occupy_cells(cult, occ)
                # Restore producer
                _free_cells_of(prod, occ)
                _move(prod, orig_prod_r, orig_prod_c, prod['h'], prod['w'])
                _occupy_cells(prod, occ)

        # Final restore of prod
        _move(prod, orig_prod_r, orig_prod_c, prod['h'], prod['w'])
        _occupy_cells(prod, occ)

    if best_state is not None:
        # Apply the best state found
        for i, b in enumerate(placed):
            b.update({k: best_state[i][k] for k in ('row','col','h','w')})
        return True, best_score
    return False, current_score


def optimize(placed, playable, rows, cols, max_iter=30, progress_cb=None):
    """Multi-strategy iterative optimization:
    Phase 1 — cluster moves (move producer + culturals together)
    Phase 2 — single-building relocations
    Phase 3 — pairwise swaps
    Repeats until no improvement or max_iter reached."""
    best = copy.deepcopy(placed)
    best_score = score_placement(best)
    total_phases = 3

    for i in range(max_iter):
        if progress_cb:
            progress_cb(i / max_iter)
        occ = build_occupancy(best, rows, cols)
        improved = False

        # Phase 1: cluster moves (most powerful, run first)
        ok, s = try_cluster_move(best, playable, rows, cols, occ, best_score)
        if ok:
            best_score = s
            improved = True
            continue

        # Phase 2: single moves
        occ = build_occupancy(best, rows, cols)
        ok, s = try_single_move(best, playable, rows, cols, occ, best_score)
        if ok:
            best_score = s
            improved = True
            continue

        # Phase 3: pairwise swaps
        occ = build_occupancy(best, rows, cols)
        ok, s = try_swap_two(best, playable, rows, cols, occ, best_score)
        if ok:
            best_score = s
            improved = True
            continue

        if not improved:
            break

    if progress_cb:
        progress_cb(1.0)
    return best, best_score


# ─────────────────────────── EXCEL OUTPUT ───────────────────────────

ORANGE = "FFFFAA44"
GREEN  = "FF44BB44"
GRAY   = "FFAAAAAA"
LIGHT_ORANGE = "FFFFE0AA"
LIGHT_GREEN  = "FFAAFFAA"
LIGHT_GRAY   = "FFDDDDDD"
BLUE   = "FF4488FF"
YELLOW = "FFFFFF88"
WHITE  = "FFFFFFFF"


def col_color(btype):
    if btype == 'Culturel':   return LIGHT_ORANGE
    if btype == 'Producteur': return LIGHT_GREEN
    return LIGHT_GRAY


def make_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def make_border():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)


def write_sheet_liste(ws, placed, culture_map, original_placed):
    """Sheet 1: list of all placed buildings"""
    orig_by_name = {}
    for b in original_placed:
        orig_by_name.setdefault(b['name'], []).append(b)
    orig_used = {n: 0 for n in orig_by_name}
    orig_culture_map = compute_culture_received(original_placed)

    headers = ['Nom', 'Type', 'Production', 'Ligne', 'Colonne', 'Hauteur', 'Largeur',
               'Culture reçue', 'Boost avant', 'Boost après', 'Évolution boost',
               'Qté/h avant optim', 'Qté/h après optim', 'Gain/Perte']
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = make_fill("FF444444")
        cell.alignment = Alignment(horizontal='center')

    for ri, b in enumerate(placed, 2):
        cr = culture_map.get(id(b), {}).get('culture', 0)
        boost = get_boost(b, cr)
        qty_after = round(b['quantite'] * (1 + boost/100)) if b['type'] == 'Producteur' else 0

        # Find matching original building
        idx = orig_used.get(b['name'], 0)
        orig_list = orig_by_name.get(b['name'], [])
        if idx < len(orig_list):
            ob = orig_list[idx]
            orig_used[b['name']] = idx + 1
            orig_cr = orig_culture_map.get(id(ob), {}).get('culture', 0)
            orig_boost = get_boost(ob, orig_cr)
            qty_before = round(ob['quantite'] * (1 + orig_boost/100)) if ob['type'] == 'Producteur' else 0
        else:
            orig_boost = 0
            qty_before = 0

        gain = qty_after - qty_before
        boost_improved = b['type'] == 'Producteur' and boost > orig_boost
        boost_degraded = b['type'] == 'Producteur' and boost < orig_boost

        if b['type'] == 'Producteur':
            boost_evo = f"▲ +{boost - orig_boost}%" if boost_improved else (f"▼ {boost - orig_boost}%" if boost_degraded else "=")
        else:
            boost_evo = ''

        row_data = [b['name'], b['type'], b['production'],
                    b['row']+1, b['col']+1, b['h'], b['w'],
                    cr,
                    f"{orig_boost}%" if b['type'] == 'Producteur' else '',
                    f"{boost}%"      if b['type'] == 'Producteur' else '',
                    boost_evo,
                    qty_before if b['type'] == 'Producteur' else '',
                    qty_after  if b['type'] == 'Producteur' else '',
                    gain if (b['type'] == 'Producteur' and b['production'] not in ('Rien','nan','')) else '']

        # Row fill: yellow highlight for rows where boost improved
        if boost_improved:
            base_fill = make_fill(YELLOW)
        elif boost_degraded:
            base_fill = make_fill("FFFFE0E0")  # light red tint
        else:
            base_fill = make_fill(col_color(b['type']))

        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill = base_fill
            cell.border = make_border()
            # Evolution boost column: bold colored
            if ci == 11 and boost_evo:
                if boost_improved:
                    cell.font = Font(bold=True, color="FF006600")
                elif boost_degraded:
                    cell.font = Font(bold=True, color="FF880000")
            # Gain/Perte column
            if ci == len(headers) and isinstance(val, (int, float)) and val != '':
                cell.font = Font(bold=True, color="FF006600" if val >= 0 else "FF880000")

    for ci in range(1, len(headers)+1):
        ws.column_dimensions[get_column_letter(ci)].width = 16
    ws.column_dimensions[get_column_letter(1)].width = 28   # Nom
    ws.column_dimensions[get_column_letter(11)].width = 14  # Évolution


def write_sheet_synthese(ws, placed, culture_map, original_culture_map, original_placed):
    """Sheet 2 & 3: synthesis by production type"""
    prods = {}
    for b in placed:
        if b['type'] != 'Producteur' or b['production'] in ('Rien', 'nan'):
            continue
        cr = culture_map.get(id(b), {}).get('culture', 0)
        boost = get_boost(b, cr)
        prod = b['production']
        if prod not in prods:
            prods[prod] = {'bldgs': [], 'culture_total': 0, 'qty_brut': 0, 'qty_boost': 0,
                           'boosts': {0:0, 25:0, 50:0, 100:0}}
        prods[prod]['bldgs'].append(b)
        prods[prod]['culture_total'] += cr
        prods[prod]['qty_brut'] += b['quantite']
        prods[prod]['qty_boost'] += b['quantite'] * (1 + boost/100)
        prods[prod]['boosts'][boost] += 1

    # Compute original
    orig_prods = {}
    for b in original_placed:
        if b['type'] != 'Producteur' or b['production'] in ('Rien', 'nan'):
            continue
        cr = original_culture_map.get(id(b), {}).get('culture', 0)
        boost = get_boost(b, cr)
        prod = b['production']
        if prod not in orig_prods:
            orig_prods[prod] = {'qty_boost': 0}
        orig_prods[prod]['qty_boost'] += b['quantite'] * (1 + boost/100)

    headers = ['Type de production', 'Culture totale reçue',
               'Nb bâtiments à 0%', 'Nb à 25%', 'Nb à 50%', 'Nb à 100%',
               'Qté/h optimisée', 'Qté/h avant optim', 'Gain/Perte']
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = make_fill("FF444444")
        cell.alignment = Alignment(horizontal='center')

    priority_order = ['Guerison', 'Nourriture', 'Or']
    all_prods = list(prods.keys())
    sorted_prods = sorted(all_prods, key=lambda x: (priority_order.index(x) if x in priority_order else 99, x))

    for ri, prod in enumerate(sorted_prods, 2):
        p = prods[prod]
        orig_qty = orig_prods.get(prod, {}).get('qty_boost', 0)
        gain = p['qty_boost'] - orig_qty
        row_data = [prod, p['culture_total'],
                    p['boosts'][0], p['boosts'][25], p['boosts'][50], p['boosts'][100],
                    round(p['qty_boost']), round(orig_qty), round(gain)]
        fill_color = LIGHT_GREEN if gain >= 0 else "FFFFAAAA"
        fill = make_fill(fill_color)
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill = fill
            cell.border = make_border()
            if ci == len(headers):
                cell.font = Font(bold=True, color="FF006600" if gain >= 0 else "FF880000")
    for ci in range(1, len(headers)+1):
        ws.column_dimensions[get_column_letter(ci)].width = 20


def write_sheet_deplacements(ws, original_placed, optimized_placed):
    """Sheet 4: list of moved buildings + sequence"""
    orig_dict = {b['name']+'_'+str(i): b for i, b in enumerate(original_placed)}

    # Match by name (simplified)
    moves = []
    orig_by_name = {}
    for b in original_placed:
        orig_by_name.setdefault(b['name'], []).append(b)
    opt_by_name = {}
    for b in optimized_placed:
        opt_by_name.setdefault(b['name'], []).append(b)

    for name in set(list(orig_by_name.keys()) + list(opt_by_name.keys())):
        olist = orig_by_name.get(name, [])
        nlist = opt_by_name.get(name, [])
        for i, (ob, nb) in enumerate(zip(olist, nlist)):
            if ob['row'] != nb['row'] or ob['col'] != nb['col'] or ob['h'] != nb['h'] or ob['w'] != nb['w']:
                moves.append({
                    'name': name,
                    'from_r': ob['row']+1, 'from_c': ob['col']+1,
                    'from_h': ob['h'], 'from_w': ob['w'],
                    'to_r': nb['row']+1, 'to_c': nb['col']+1,
                    'to_h': nb['h'], 'to_w': nb['w'],
                })

    headers = ['Bâtiment', 'Ligne avant', 'Col avant', 'HxL avant',
               'Ligne après', 'Col après', 'HxL après', 'Séquence opération']
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = make_fill("FF444444")
        cell.alignment = Alignment(horizontal='center')

    ws.cell(row=2, column=1, value="SÉQUENCE DE DÉPLACEMENT RECOMMANDÉE").font = Font(bold=True, color="FFAA0000")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=8)

    # Sort by priority (move smaller buildings first to free space)
    moves_sorted = sorted(moves, key=lambda m: m['from_h']*m['from_w'])

    for step, m in enumerate(moves_sorted, 1):
        ri = step + 2
        seq = (f"Étape {step}: Déplacer '{m['name']}' "
               f"de ({m['from_r']},{m['from_c']}) "
               f"vers ({m['to_r']},{m['to_c']}). "
               f"Si occupé, mettre temporairement hors terrain.")
        row_data = [m['name'],
                    m['from_r'], m['from_c'], f"{m['from_h']}x{m['from_w']}",
                    m['to_r'], m['to_c'], f"{m['to_h']}x{m['to_w']}",
                    seq]
        fill = make_fill(YELLOW)
        for ci, val in enumerate(row_data, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill = fill
            cell.border = make_border()
    for ci in range(1, len(headers)+1):
        ws.column_dimensions[get_column_letter(ci)].width = 22
    ws.column_dimensions[get_column_letter(8)].width = 60


def write_sheet_terrain(ws, placed, playable, rows, cols, culture_map, label="Terrain optimisé"):
    """Sheet: visual terrain grid matching input file style (merged cells per building)"""
    ws.cell(row=1, column=1, value=label).font = Font(bold=True, size=14)

    # Column widths matching input: col A=25, others=13
    ws.column_dimensions['A'].width = 25
    for c in range(2, cols + 2):
        ws.column_dimensions[get_column_letter(c)].width = 13

    # Build lookup: (terrain_row, terrain_col) -> building
    cell_to_bldg = {}
    for b in placed:
        for dr in range(b['h']):
            for dc in range(b['w']):
                cell_to_bldg[(b['row']+dr, b['col']+dc)] = b

    # Colors for building types
    TYPE_FILL = {
        'Culturel':   PatternFill("solid", fgColor="FFFF8C00"),   # orange
        'Producteur': PatternFill("solid", fgColor="FF228B22"),   # green
        'Neutre':     PatternFill("solid", fgColor="FF808080"),   # gray
    }
    X_FILL   = PatternFill("solid", fgColor="FF000000")
    NO_FILL  = PatternFill(fill_type=None)

    # Build set of actual X border cells (vs empty cells outside terrain)
    x_cells = set()
    for r in range(rows):
        for c in range(cols):
            if not playable[r][c]:
                val = grid[r][c]
                if isinstance(val, str) and val.strip().upper() == 'X':
                    x_cells.add((r, c))

    # Pass 1: set X cells only; leave all empty cells completely untouched
    for r in range(rows):
        for c in range(cols):
            excel_row = r + 2
            excel_col = c + 1
            if (r, c) in x_cells:
                cell = ws.cell(row=excel_row, column=excel_col)
                cell.value = 'X'
                cell.fill = X_FILL
                cell.font = Font(color='FFFFFFFF', size=11)
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # Pass 2: handle buildings — merge first, then set label+fill on top-left
    done_bldgs = set()
    for b in placed:
        bid = id(b)
        if bid in done_bldgs:
            continue
        done_bldgs.add(bid)
        fill = TYPE_FILL.get(b['type'], TYPE_FILL['Neutre'])

        r1 = b['row'] + 2
        c1 = b['col'] + 1
        r2 = b['row'] + b['h'] - 1 + 2
        c2 = b['col'] + b['w'] - 1 + 1

        # Merge first (makes inner cells MergedCell read-only)
        if r1 != r2 or c1 != c2:
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

        # Set label + style on top-left cell only (always writable)
        cr = culture_map.get(bid, {}).get('culture', 0)
        boost = get_boost(b, cr)
        label_text = f"{b['name']}\n+{boost}%" if (b['type'] == 'Producteur' and boost > 0) else b['name']
        top_cell = ws.cell(row=r1, column=c1)
        top_cell.value = label_text
        top_cell.font = Font(bold=True, size=9, color='FFFFFFFF')
        top_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        top_cell.fill = fill


def generate_excel(original_placed, optimized_placed, playable, rows, cols):
    original_culture_map = compute_culture_received(original_placed)
    culture_map = compute_culture_received(optimized_placed)

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Liste bâtiments"
    write_sheet_liste(ws1, optimized_placed, culture_map, original_placed)

    ws2 = wb.create_sheet("Synthèse productions")
    write_sheet_synthese(ws2, optimized_placed, culture_map, original_culture_map, original_placed)

    ws3 = wb.create_sheet("Déplacements")
    write_sheet_deplacements(ws3, original_placed, optimized_placed)

    ws4 = wb.create_sheet("Terrain optimisé")
    write_sheet_terrain(ws4, optimized_placed, playable, rows, cols, culture_map, "Terrain optimisé")

    ws5 = wb.create_sheet("Terrain original")
    write_sheet_terrain(ws5, original_placed, playable, rows, cols, original_culture_map, "Terrain original")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─────────────────────────── STREAMLIT UI ───────────────────────────

uploaded = st.file_uploader("📂 Chargez votre fichier Excel (Ville.xlsx)", type=['xlsx'])

if uploaded:
    with st.spinner("Lecture du fichier..."):
        from openpyxl import load_workbook as _load_wb
        import io as _io
        _wb = _load_wb(uploaded)
        ws_terrain = _wb[_wb.sheetnames[0]]
        # Parse buildings sheet via pandas
        _buf2 = _io.BytesIO()
        _wb.save(_buf2); _buf2.seek(0)
        df_bat = pd.read_excel(_buf2, sheet_name=_wb.sheetnames[1], header=None)

        bldg_info = parse_bldg_info(df_bat)
        grid, playable, rows, cols = parse_terrain(ws_terrain)
        original_placed = parse_buildings_from_terrain(grid, playable, rows, cols, bldg_info, ws_openpyxl=ws_terrain)

    st.success(f"✅ Terrain {rows}×{cols} chargé — {len(original_placed)} bâtiments détectés")

    # Stats before
    orig_culture_map = compute_culture_received(original_placed)
    orig_prod = compute_total_production(original_placed, orig_culture_map)

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Productions actuelles (avant optimisation)")
        for prod, qty in sorted(orig_prod.items(), key=lambda x: (PRIORITY_ORDER.index(x[0]) if x[0] in PRIORITY_ORDER else 99)):
            st.metric(prod, f"{qty:,.0f}/h")

    n_iter = st.slider("Nombre d'itérations d'optimisation", 5, 100, 20, 5)

    if st.button("🚀 Lancer l'optimisation", type="primary"):
        progress = st.progress(0)
        status = st.empty()
        status.info("Optimisation en cours...")

        optimized = copy.deepcopy(original_placed)
        best, best_score = optimize(
            optimized, playable, rows, cols,
            max_iter=n_iter,
            progress_cb=lambda v: progress.progress(v)
        )

        new_culture_map = compute_culture_received(best)
        new_prod = compute_total_production(best, new_culture_map)

        status.success(f"✅ Optimisation terminée (score: {best_score:,.0f})")

        with col2:
            st.subheader("Productions après optimisation")
            for prod, qty in sorted(new_prod.items(), key=lambda x: (PRIORITY_ORDER.index(x[0]) if x[0] in PRIORITY_ORDER else 99)):
                orig = orig_prod.get(prod, 0)
                delta = qty - orig
                st.metric(prod, f"{qty:,.0f}/h", delta=f"{delta:+,.0f}/h")

        st.subheader("📥 Télécharger le fichier de résultats")
        excel_buf = generate_excel(original_placed, best, playable, rows, cols)
        st.download_button(
            label="⬇️ Télécharger résultats.xlsx",
            data=excel_buf,
            file_name="resultats_optimisation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Preview
        st.subheader("Aperçu — synthèse des boosts")
        rows_data = []
        for b in best:
            if b['type'] != 'Producteur' or b['production'] in ('Rien', 'nan'):
                continue
            cr = new_culture_map.get(id(b), {}).get('culture', 0)
            boost = get_boost(b, cr)
            rows_data.append({
                'Bâtiment': b['name'],
                'Production': b['production'],
                'Culture reçue': cr,
                'Boost': f"{boost}%",
                'Qté/h': round(b['quantite'] * (1+boost/100))
            })
        if rows_data:
            st.dataframe(pd.DataFrame(rows_data).sort_values('Production'), use_container_width=True)
