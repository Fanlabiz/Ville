# -*- coding: utf-8 -*-

# Optimiseur de placement de batiments – Streamlit App

# Depose un fichier Excel en input, telecharge le resultat optimise.

import streamlit as st
import pandas as pd
import numpy as np
from copy import deepcopy
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

# ———————————————

# PAGE CONFIG

# ———————————————

st.set_page_config(
page_title="Optimiseur de Ville",
page_icon="🏙️",
layout="wide",
)

st.markdown(’’’

<style>
@import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@600;800&family=Lato:wght@300;400;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Lato', sans-serif;
}
h1, h2, h3 {
    font-family: 'Cinzel', serif;
}
.main-title {
    font-family: 'Cinzel', serif;
    font-size: 2.2rem;
    font-weight: 800;
    color: #2C3E50;
    letter-spacing: 0.04em;
    margin-bottom: 0;
}
.sub-title {
    font-family: 'Lato', sans-serif;
    font-weight: 300;
    font-size: 1.0rem;
    color: #7F8C8D;
    margin-top: 0.2rem;
    margin-bottom: 1.5rem;
    letter-spacing: 0.06em;
    text-transform: uppercase;
}
.stat-box {
    background: linear-gradient(135deg, #f8f9fa, #e9ecef);
    border-left: 4px solid #2C3E50;
    padding: 0.9rem 1.2rem;
    border-radius: 6px;
    margin-bottom: 0.6rem;
}
.stat-label {
    font-size: 0.75rem;
    color: #6c757d;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    font-weight: 700;
}
.stat-value {
    font-size: 1.5rem;
    font-weight: 700;
    color: #2C3E50;
    font-family: 'Cinzel', serif;
}
.gain-positive { color: #27ae60; font-weight: 700; }
.gain-negative { color: #e74c3c; font-weight: 700; }
.step-card {
    background: #fff;
    border: 1px solid #dee2e6;
    border-radius: 8px;
    padding: 0.7rem 1rem;
    margin-bottom: 0.4rem;
    display: flex;
    align-items: center;
    gap: 1rem;
}
.badge-temp   { background:#FEF3C7; color:#92400E; border-radius:20px; padding:3px 10px; font-size:0.75rem; font-weight:700; }
.badge-place  { background:#D1FAE5; color:#065F46; border-radius:20px; padding:3px 10px; font-size:0.75rem; font-weight:700; }
.badge-restore{ background:#DBEAFE; color:#1E3A8A; border-radius:20px; padding:3px 10px; font-size:0.75rem; font-weight:700; }
.section-sep { border:none; border-top:2px solid #e9ecef; margin: 1.5rem 0; }
</style>

‘’’, unsafe_allow_html=True)

# ———————————————

# HEADER

# ———————————————

st.markdown(’<div class="main-title">🏙️ Optimiseur de Ville</div>’, unsafe_allow_html=True)
st.markdown(’<div class="sub-title">Maximisation de la production par boost culturel</div>’, unsafe_allow_html=True)
st.markdown(’<hr class="section-sep">’, unsafe_allow_html=True)

# ===============================================

# CORE FUNCTIONS

# ===============================================

def parse_terrain(terrain_df):
rows, cols = terrain_df.shape
grid = np.zeros((rows, cols), dtype=bool)
for r in range(rows):
for c in range(cols):
if str(terrain_df.iloc[r, c]).strip() == ‘1’:
grid[r, c] = True
return grid

def parse_buildings(bats_df):
buildings = []
for _, row in bats_df.iterrows():
nom   = str(row[‘Nom’]).strip()
L     = int(row[‘Longueur’])
W     = int(row[‘Largeur’])
N     = int(row[‘Nombre’])
typ   = str(row[‘Type’]).strip()
cult  = float(row[‘Culture’])   if pd.notna(row.get(‘Culture’))   else 0.0
ray   = int(row[‘Rayonnement’]) if pd.notna(row.get(‘Rayonnement’)) else 0
b25   = float(row[‘Boost 25%’])  if pd.notna(row.get(‘Boost 25%’))  else None
b50   = float(row[‘Boost 50%’])  if pd.notna(row.get(‘Boost 50%’))  else None
b100  = float(row[‘Boost 100%’]) if pd.notna(row.get(‘Boost 100%’)) else None
prod  = str(row[‘Production’]).strip() if pd.notna(row.get(‘Production’)) else ‘Rien’
qte   = float(row[‘Quantite’]) if pd.notna(row.get(‘Quantite’)) else 0.0
for i in range(N):
buildings.append(dict(nom=nom, L=L, W=W, type=typ, culture=cult, ray=ray,
b25=b25, b50=b50, b100=b100, prod=prod, qte=qte))
return buildings

def parse_actuel(actuel_df):
‘’‘Extract building instances from the painted grid.’’’
grid = {}
for r in range(actuel_df.shape[0]):
for c in range(actuel_df.shape[1]):
val = actuel_df.iloc[r, c]
if pd.notna(val) and str(val).strip() not in (‘X’, ‘1’, ‘0’, ‘’):
grid[(r, c)] = str(val).strip()

```
visited = set()
instances = []
for (r, c), name in sorted(grid.items()):
    if (r, c) in visited:
        continue
    cells, stack = [], [(r, c)]
    while stack:
        cur = stack.pop()
        if cur in visited:
            continue
        if grid.get(cur) == name:
            visited.add(cur)
            cells.append(cur)
            cr, cc = cur
            for nr, nc in [(cr-1,cc),(cr+1,cc),(cr,cc-1),(cr,cc+1)]:
                if (nr, nc) not in visited and grid.get((nr, nc)) == name:
                    stack.append((nr, nc))
    if cells:
        rmin = min(x[0] for x in cells)
        rmax = max(x[0] for x in cells)
        cmin = min(x[1] for x in cells)
        cmax = max(x[1] for x in cells)
        instances.append({'nom': name, 'r': rmin, 'c': cmin,
                           'h': rmax-rmin+1, 'w': cmax-cmin+1})
return instances
```

def build_placed(actuel_instances, buildings):
inst_by_name = {}
for inst in actuel_instances:
inst_by_name.setdefault(inst[‘nom’], []).append(inst)

```
placed = {}
for pid, bat in enumerate(buildings):
    nom = bat['nom']
    if nom in inst_by_name and inst_by_name[nom]:
        inst = inst_by_name[nom].pop(0)
        placed[pid] = {**bat, 'r': inst['r'], 'c': inst['c'],
                       'h': inst['h'], 'w': inst['w']}
return placed
```

def build_grid(placed, rows_t, cols_t):
grid = np.full((rows_t, cols_t), -1, dtype=int)
for pid, p in placed.items():
if ‘r’ not in p:
continue
for dr in range(p[‘h’]):
for dc in range(p[‘w’]):
grid[p[‘r’]+dr, p[‘c’]+dc] = pid
return grid

def can_place(grid_occ, r, c, h, w, terrain):
rows_t, cols_t = terrain.shape
if r < 0 or c < 0 or r+h > rows_t or c+w > cols_t:
return False
for dr in range(h):
for dc in range(w):
if not terrain[r+dr, c+dc]:
return False
if grid_occ[r+dr, c+dc] >= 0:
return False
return True

def compute_culture_received(placed):
cult_recv = {pid: 0.0 for pid, p in placed.items() if p[‘type’] == ‘Producteur’}
for cid, cp in placed.items():
if cp[‘type’] != ‘Culturel’ or cp[‘culture’] == 0:
continue
ray = cp[‘ray’]
cr, cc, ch, cw = cp[‘r’], cp[‘c’], cp[‘h’], cp[‘w’]
zr1, zr2 = cr - ray, cr + ch - 1 + ray
zc1, zc2 = cc - ray, cc + cw - 1 + ray
for pid, p in placed.items():
if p[‘type’] != ‘Producteur’:
continue
pr, pc, ph, pw = p[‘r’], p[‘c’], p[‘h’], p[‘w’]
if pr <= zr2 and pr+ph-1 >= zr1 and pc <= zc2 and pc+pw-1 >= zc1:
cult_recv[pid] += cp[‘culture’]
return cult_recv

def compute_boost(culture, b25, b50, b100):
if b100 is not None and culture >= b100: return 1.0
if b50  is not None and culture >= b50:  return 0.5
if b25  is not None and culture >= b25:  return 0.25
return 0.0

PROD_PRIORITY = [‘Guerison’, ‘Nourriture’, ‘Or’]

def compute_score(placed):
cult_recv = compute_culture_received(placed)
totals = {}
for pid, p in placed.items():
if p[‘type’] != ‘Producteur’ or p[‘prod’] == ‘Rien’:
continue
c = cult_recv.get(pid, 0)
boost = compute_boost(c, p[‘b25’], p[‘b50’], p[‘b100’])
prod = p[‘prod’]
totals[prod] = totals.get(prod, 0) + p[‘qte’] * (1 + boost)
score = 0
for i, pt in enumerate(PROD_PRIORITY):
score += (10 ** (len(PROD_PRIORITY) - i)) * totals.get(pt, 0)
return score

def compute_production_detail(placed):
cult_recv = compute_culture_received(placed)
by_prod = {}
per_bat = {}
for pid, p in placed.items():
if p[‘type’] != ‘Producteur’ or p[‘prod’] == ‘Rien’:
continue
c = cult_recv.get(pid, 0)
boost = compute_boost(c, p[‘b25’], p[‘b50’], p[‘b100’])
prod = p[‘prod’]
qty = p[‘qte’] * (1 + boost)
by_prod[prod] = by_prod.get(prod, 0) + qty
per_bat[pid] = {‘culture’: c, ‘boost’: boost, ‘qty_boost’: qty}
return by_prod, per_bat, cult_recv

# — OPTIMIZATION ———————————————————–

def try_swap(placed, pid1, pid2, terrain):
rows_t, cols_t = terrain.shape
p1, p2 = placed[pid1], placed[pid2]
r1, c1, h1, w1 = p1[‘r’], p1[‘c’], p1[‘h’], p1[‘w’]
r2, c2, h2, w2 = p2[‘r’], p2[‘c’], p2[‘h’], p2[‘w’]

```
base = {k: v for k, v in placed.items() if k not in (pid1, pid2)}
g = build_grid(base, rows_t, cols_t)

# Try p1->pos2, p2->pos1 (with optional rotation)
for nh1, nw1 in [(h1, w1), (w1, h1)]:
    if not can_place(g, r2, c2, nh1, nw1, terrain):
        continue
    for nh2, nw2 in [(h2, w2), (w2, h2)]:
        if not can_place(g, r1, c1, nh2, nw2, terrain):
            continue
        tmp = deepcopy(placed)
        tmp[pid1] = {**p1, 'r': r2, 'c': c2, 'h': nh1, 'w': nw1}
        tmp[pid2] = {**p2, 'r': r1, 'c': c1, 'h': nh2, 'w': nw2}
        return tmp
return None
```

def optimize(placed_orig, terrain, progress_cb=None):
rows_t, cols_t = terrain.shape
placed = deepcopy(placed_orig)
best_score = compute_score(placed)

```
cultural_ids = [pid for pid, p in placed.items() if p['type'] == 'Culturel']
neutral_ids  = [pid for pid, p in placed.items() if p['type'] == 'Neutre']
movable      = cultural_ids + neutral_ids

# -- Phase 1 : pairwise swaps ------------------------------------------
if progress_cb:
    progress_cb(0.05, "Phase 1 : échanges par paires…")

improved = True
iteration = 0
while improved:
    improved = False
    iteration += 1
    for i in range(len(movable)):
        for j in range(i+1, len(movable)):
            pid1, pid2 = movable[i], movable[j]
            if pid1 not in placed or pid2 not in placed:
                continue
            result = try_swap(placed, pid1, pid2, terrain)
            if result is not None:
                new_score = compute_score(result)
                if new_score > best_score:
                    placed = result
                    best_score = new_score
                    improved = True

# -- Phase 2 : relocalisation individuelle des bâtiments culturels -----
if progress_cb:
    progress_cb(0.45, "Phase 2 : relocalisation des bâtiments culturels…")

pass_num = 0
improved = True
while improved and pass_num < 30:
    improved = False
    pass_num += 1
    for cid in cultural_ids:
        if cid not in placed:
            continue
        cp = placed[cid]
        oh, ow = cp['h'], cp['w']
        base = {k: v for k, v in placed.items() if k != cid}
        g = build_grid(base, rows_t, cols_t)

        best_local = best_score
        best_pos   = None

        for nh, nw in ([(oh, ow)] if oh == ow else [(oh, ow), (ow, oh)]):
            for r in range(rows_t - nh + 1):
                for c in range(cols_t - nw + 1):
                    if not can_place(g, r, c, nh, nw, terrain):
                        continue
                    tmp = deepcopy(placed)
                    tmp[cid] = {**cp, 'r': r, 'c': c, 'h': nh, 'w': nw}
                    s = compute_score(tmp)
                    if s > best_local:
                        best_local = s
                        best_pos   = (r, c, nh, nw)

        if best_pos:
            r, c, nh, nw = best_pos
            placed[cid] = {**placed[cid], 'r': r, 'c': c, 'h': nh, 'w': nw}
            best_score = best_local
            improved = True

    if progress_cb:
        pct = 0.45 + 0.45 * (pass_num / 30)
        progress_cb(pct, f"Phase 2 : passe {pass_num}…")

if progress_cb:
    progress_cb(0.95, "Calcul de la séquence d'opérations…")

return placed, best_score
```

# — SEQUENCE BUILDER —————————————————––

def build_sequence(placed_orig, placed_best, terrain):
‘’‘Build the ordered step-by-step move sequence.’’’
rows_t, cols_t = terrain.shape

```
moved = []
for pid in placed_orig:
    if pid not in placed_best:
        continue
    po, pb = placed_orig[pid], placed_best[pid]
    if (po.get('r') != pb.get('r') or po.get('c') != pb.get('c')
            or po.get('h') != pb.get('h') or po.get('w') != pb.get('w')):
        moved.append({'pid': pid, 'nom': po['nom'], 'type': po['type'],
                      'fr': po['r'], 'fc': po['c'], 'fh': po['h'], 'fw': po['w'],
                      'tr': pb['r'], 'tc': pb['c'], 'th': pb['h'], 'tw': pb['w']})

moved_pids = {m['pid'] for m in moved}

# Build original cell map
orig_cell = {}
for pid, p in placed_orig.items():
    if 'r' not in p:
        continue
    for dr in range(p['h']):
        for dc in range(p['w']):
            orig_cell[(p['r']+dr, p['c']+dc)] = pid

# Find blockers (other moved buildings that occupy destination)
def get_blocker(m):
    for dr in range(m['th']):
        for dc in range(m['tw']):
            occ = orig_cell.get((m['tr']+dr, m['tc']+dc))
            if occ is not None and occ != m['pid'] and occ in moved_pids:
                return occ
    return None

# Topological sort -- detect cycles (direct swaps)
in_temp = set()   # PIDs currently held out of terrain
done    = set()
steps   = []
remaining = list(moved)
group_num = 0

while remaining:
    progress = False
    for m in list(remaining):
        blocker_pid = get_blocker(m)
        if blocker_pid is None or blocker_pid in done:
            # Destination is free (or blocker already moved away)
            group_num += 1
            if m['pid'] in in_temp:
                steps.append({'group': group_num, 'action': 'replacer',
                              'bat': m['nom'], 'type': m['type'],
                              'from': 'HORS TERRAIN',
                              'to': _coord(m['tr'], m['tc'], m['th'], m['tw'])})
                in_temp.discard(m['pid'])
            else:
                steps.append({'group': group_num, 'action': 'placer',
                              'bat': m['nom'], 'type': m['type'],
                              'from': _coord(m['fr'], m['fc'], m['fh'], m['fw']),
                              'to':   _coord(m['tr'], m['tc'], m['th'], m['tw'])})
            done.add(m['pid'])
            remaining.remove(m)
            progress = True
            break

    if not progress:
        # Cycle detected → put first remaining building out of terrain
        m = remaining[0]
        group_num += 1
        steps.append({'group': group_num, 'action': 'retirer',
                      'bat': m['nom'], 'type': m['type'],
                      'from': _coord(m['fr'], m['fc'], m['fh'], m['fw']),
                      'to':   'HORS TERRAIN (temporaire)'})
        in_temp.add(m['pid'])
        # Mark its cells as freed
        for dr in range(m['fh']):
            for dc in range(m['fw']):
                orig_cell.pop((m['fr']+dr, m['fc']+dc), None)
        moved_pids.discard(m['pid'])
        remaining.remove(m)
        # Re-add at end so it gets placed later
        remaining.append(m)

return steps, moved
```

def _coord(r, c, h, w):
return f”Ligne {r+1}, Col {c+1}  ({h}L × {w}C)”

# ===============================================

# EXCEL OUTPUT GENERATOR

# ===============================================

ORANGE_FILL  = PatternFill(‘solid’, start_color=‘FFFF9933’)
GREEN_FILL   = PatternFill(‘solid’, start_color=‘FF66BB44’)
GREY_FILL    = PatternFill(‘solid’, start_color=‘FFCCCCCC’)
LORANGE_FILL = PatternFill(‘solid’, start_color=‘FFFFCC99’)
LGREEN_FILL  = PatternFill(‘solid’, start_color=‘FFCCFFCC’)
LGREY_FILL   = PatternFill(‘solid’, start_color=‘FFEEEEEE’)
BLUE_HDR     = PatternFill(‘solid’, start_color=‘FF1F497D’)
YELLOW_FILL  = PatternFill(‘solid’, start_color=‘FFFEF3C7’)
LBLUE_FILL   = PatternFill(‘solid’, start_color=‘FFDBEAFE’)
LRED_FILL    = PatternFill(‘solid’, start_color=‘FFFEE2E2’)
DARK_FILL    = PatternFill(‘solid’, start_color=‘FF333333’)
WHITE_FILL   = PatternFill(‘solid’, start_color=‘FFFFFFFF’)

_thin = Side(style=‘thin’)
_bdr  = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

GRP_COLORS = [‘FFE8F4FD’,‘FFE8F5E9’,‘FFFFF8E1’,‘FFFFE8E8’,
‘FFF3E8FF’,‘FFE8FFF3’,‘FFFFE8F5’,‘FFF0F0FF’]

def _hdr(cell, text, bg=‘FF1F497D’, size=9):
cell.value = text
cell.font  = Font(bold=True, color=‘FFFFFFFF’, name=‘Calibri’, size=size)
cell.fill  = PatternFill(‘solid’, start_color=bg)
cell.alignment = Alignment(horizontal=‘center’, vertical=‘center’, wrap_text=True)
cell.border = _bdr

def _dc(cell, value, bold=False, bg=None, align=‘left’, num_fmt=None, size=9, color=‘FF000000’):
cell.value = value
cell.font  = Font(name=‘Calibri’, size=size, bold=bold, color=color)
cell.alignment = Alignment(horizontal=align, vertical=‘center’, wrap_text=False)
if bg:
cell.fill = PatternFill(‘solid’, start_color=bg)
if num_fmt:
cell.number_format = num_fmt
cell.border = _bdr

def generate_excel(placed_orig, placed_best, terrain, orig_prod, best_prod):
rows_t, cols_t = terrain.shape
wb = Workbook()

```
best_detail, per_bat_best, cult_best = compute_production_detail(placed_best)
_, per_bat_orig, cult_orig = compute_production_detail(placed_orig)
steps, moved_list = build_sequence(placed_orig, placed_best, terrain)

type_bg = {'Culturel': 'FFFFCC99', 'Producteur': 'FFCCFFCC', 'Neutre': 'FFEEEEEE'}

# -- Onglet 1 : Bâtiments optimisés --------------------------------------
ws1 = wb.active
ws1.title = "1-Batiments Optimises"
ws1.freeze_panes = 'A2'
ws1.row_dimensions[1].height = 28

hdrs1 = ['Nom','Type','Production','Ligne','Colonne','H','W',
         'Culture reçue','Boost atteint','Qte/h (boostée)']
for ci, h in enumerate(hdrs1, 1):
    _hdr(ws1.cell(1, ci), h)

col_w1 = [30,12,13,7,9,5,5,14,12,16]
for ci, w in enumerate(col_w1, 1):
    ws1.column_dimensions[get_column_letter(ci)].width = w

boost_lbl = {1.0:'100%', 0.5:'50%', 0.25:'25%', 0.0:'0%'}

for ri, (pid, p) in enumerate(sorted(placed_best.items(),
                                     key=lambda x: x[1]['nom']), 2):
    bg = type_bg.get(p['type'])
    info = per_bat_best.get(pid, {})
    c_recv = info.get('culture', 0)
    boost  = info.get('boost', 0)
    qb     = info.get('qty_boost', 0)
    row_vals = [p['nom'], p['type'], p['prod'],
                p.get('r',0)+1, p.get('c',0)+1, p.get('h',0), p.get('w',0),
                round(c_recv,1), boost_lbl.get(boost,'0%'), round(qb,1)]
    for ci, v in enumerate(row_vals, 1):
        nf = '#,##0' if ci in (8,10) else None
        al = 'right' if ci >= 4 else 'left'
        _dc(ws1.cell(ri, ci), v, bg=bg, align=al, num_fmt=nf)

# -- Onglet 2 : Résumé production ----------------------------------------
ws2 = wb.create_sheet("2-Resume Production")
ws2.freeze_panes = 'A2'
ws2.row_dimensions[1].height = 28

hdrs2 = ['Production','Nb bâtiments','Qte base/h','Qte optimisée/h',
         'Nb 0%','Nb 25%','Nb 50%','Nb 100%','Culture totale reçue']
for ci, h in enumerate(hdrs2, 1):
    _hdr(ws2.cell(1, ci), h)
for ci, w in enumerate([16,14,13,18,8,8,8,9,20], 1):
    ws2.column_dimensions[get_column_letter(ci)].width = w

# compute summary
def prod_summary(placed, cult_recv):
    d = {}
    for pid, p in placed.items():
        if p['type'] != 'Producteur' or p['prod'] == 'Rien':
            continue
        c    = cult_recv.get(pid, 0)
        boost = compute_boost(c, p['b25'], p['b50'], p['b100'])
        pr   = p['prod']
        if pr not in d:
            d[pr] = dict(count=0, qte_base=0, qte_boost=0,
                         n0=0, n25=0, n50=0, n100=0, cult_total=0)
        d[pr]['count']     += 1
        d[pr]['qte_base']  += p['qte']
        d[pr]['qte_boost'] += p['qte'] * (1 + boost)
        d[pr]['cult_total']+= c
        d[pr][{0.0:'n0',0.25:'n25',0.5:'n50',1.0:'n100'}[boost]] += 1
    return d

s_best = prod_summary(placed_best, cult_best)
PROD_ORDER = ['Guerison','Nourriture','Or','Pommades','Boiseries',
              'Bijoux','Epices','Scriberie','Cristal']

ri = 2
for pr in PROD_ORDER + [p for p in s_best if p not in PROD_ORDER]:
    if pr not in s_best:
        continue
    s = s_best[pr]
    row_vals = [pr, s['count'], round(s['qte_base'],0), round(s['qte_boost'],0),
                s['n0'], s['n25'], s['n50'], s['n100'], round(s['cult_total'],0)]
    for ci, v in enumerate(row_vals, 1):
        nf = '#,##0' if ci in (3,4,9) else None
        al = 'right' if ci > 1 else 'left'
        _dc(ws2.cell(ri, ci), v, align=al, num_fmt=nf)
    ri += 1

# -- Onglet 3 : Comparatif production ------------------------------------
ws3 = wb.create_sheet("3-Comparatif Production")
ws3.freeze_panes = 'A2'
ws3.row_dimensions[1].height = 28

hdrs3 = ['Production','Avant /h','Après /h','Gain /h','Gain %']
for ci, h in enumerate(hdrs3, 1):
    _hdr(ws3.cell(1, ci), h)
for ci, w in enumerate([16,16,16,12,10], 1):
    ws3.column_dimensions[get_column_letter(ci)].width = w

ri = 2
for pr in PROD_ORDER + [p for p in best_prod if p not in PROD_ORDER and p != 'Rien']:
    if pr not in best_prod and pr not in orig_prod:
        continue
    o  = orig_prod.get(pr, 0)
    b  = best_prod.get(pr, 0)
    dlt = b - o
    pct = (dlt / o * 100) if o > 0 else 0
    bg = 'FFE0FFE0' if dlt > 0 else ('FFFFE0E0' if dlt < 0 else None)
    row_vals = [pr, round(o,0), round(b,0), round(dlt,0), round(pct,1)]
    for ci, v in enumerate(row_vals, 1):
        nf = '#,##0' if ci in (2,3,4) else ('0.0' if ci == 5 else None)
        al = 'right' if ci > 1 else 'left'
        _dc(ws3.cell(ri, ci), v, bg=bg, align=al, num_fmt=nf)
    ri += 1

# -- Onglet 4 : Bâtiments déplacés ---------------------------------------
ws4 = wb.create_sheet("4-Batiments Deplaces")
ws4.freeze_panes = 'A2'
ws4.row_dimensions[1].height = 28

hdrs4 = ['Nom','Type','Avant Ligne','Avant Col','Avant Orient.',
         'Après Ligne','Après Col','Après Orient.']
for ci, h in enumerate(hdrs4, 1):
    _hdr(ws4.cell(1, ci), h)
for ci, w in enumerate([30,12,11,10,13,11,10,13], 1):
    ws4.column_dimensions[get_column_letter(ci)].width = w

for ri, m in enumerate(moved_list, 2):
    bg = type_bg.get(m['type'])
    row_vals = [m['nom'], m['type'],
                m['fr']+1, m['fc']+1, f"{m['fh']}L×{m['fw']}C",
                m['tr']+1, m['tc']+1, f"{m['th']}L×{m['tw']}C"]
    for ci, v in enumerate(row_vals, 1):
        al = 'right' if ci in (3,4,6,7) else 'left'
        _dc(ws4.cell(ri, ci), v, bg=bg, align=al)

ws4.cell(len(moved_list)+3, 1).value = f"Total : {len(moved_list)} bâtiment(s) déplacé(s)"
ws4.cell(len(moved_list)+3, 1).font  = Font(bold=True, name='Calibri', size=9)

# -- Onglet 5 : Séquence des opérations ----------------------------------
ws5 = wb.create_sheet("5-Sequence Operations")
ws5.row_dimensions[1].height = 22
ws5.row_dimensions[2].height = 15
ws5.row_dimensions[3].height = 28

ws5.merge_cells('A1:H1')
c = ws5['A1']
c.value = f"SÉQUENCE DES OPÉRATIONS -- {len(moved_list)} bâtiments déplacés, {len(steps)} étapes"
c.font  = Font(bold=True, name='Calibri', size=11, color='FFFFFFFF')
c.fill  = PatternFill('solid', start_color='FF1F497D')
c.alignment = Alignment(horizontal='center', vertical='center')

ws5.merge_cells('A2:H2')
c2 = ws5['A2']
c2.value = ("Pour chaque échange circulaire : (1) retirer A hors terrain, "
            "(2) déplacer B à sa destination, (3) replacer A à sa destination.")
c2.font  = Font(italic=True, name='Calibri', size=9)
c2.alignment = Alignment(horizontal='left', vertical='center')

hdrs5 = ['Étape','Groupe','Action','Bâtiment','Type',
         'Position actuelle','Position finale','Note']
for ci, h in enumerate(hdrs5, 1):
    _hdr(ws5.cell(3, ci), h)

col_w5 = [7,13,26,30,12,34,34,34]
for ci, w in enumerate(col_w5, 1):
    ws5.column_dimensions[get_column_letter(ci)].width = w

act_bg = {'retirer':  'FFFEF3C7',
          'placer':   'FFD1FAE5',
          'replacer': 'FFDBEAFE'}
act_lbl = {'retirer':  '1 -- Retirer hors terrain',
           'placer':   '2 -- Placer à destination',
           'replacer': '3 -- Replacer depuis hors terrain'}

grp_map = {}
grp_idx = -1
current_grp = None

for si, step in enumerate(steps, 1):
    ri = si + 3
    ws5.row_dimensions[ri].height = 18

    grp = step['group']
    if grp != current_grp:
        current_grp = grp
        grp_idx = (grp_idx + 1) % len(GRP_COLORS)
    gc = GRP_COLORS[grp_idx]
    ac = act_bg.get(step['action'], gc)

    note = ''
    if step['action'] == 'retirer':
        note = f"Libère la place pour l'échange {grp}"
    elif step['action'] == 'placer':
        note = f"Position définitive"
    elif step['action'] == 'replacer':
        note = f"Position définitive (depuis hors terrain)"

    row_vals = [si, f"Groupe {grp}", act_lbl[step['action']],
                step['bat'], step['type'],
                step['from'], step['to'], note]

    for ci, v in enumerate(row_vals, 1):
        cell = ws5.cell(ri, ci)
        if ci in (1, 2):
            _dc(cell, v, bold=(ci==1), bg=gc, align='center')
        elif ci == 3:
            _dc(cell, v, bold=True, bg=ac, align='center')
        elif 'HORS TERRAIN' in str(v):
            _dc(cell, v, bold=True, bg='FFFFF0AA', align='center')
        else:
            _dc(cell, v, bg=type_bg.get(step['type']) if ci == 5 else None, align='left')

# Legend
lr = len(steps) + 5
ws5.cell(lr, 1).value = "LÉGENDE :"
ws5.cell(lr, 1).font  = Font(bold=True, name='Calibri', size=9)
leg_items = [
    ('1 -- Retirer hors terrain',        'FFFEF3C7', "Mettre le bâtiment de côté temporairement"),
    ('2 -- Placer à destination',         'FFD1FAE5', "Déplacer à la position définitive optimisée"),
    ('3 -- Replacer depuis hors terrain', 'FFDBEAFE', "Reprendre le bâtiment mis de côté et le placer définitivement"),
]
for i, (lbl, bg, desc) in enumerate(leg_items):
    r = lr + 1 + i
    ws5.row_dimensions[r].height = 16
    c1 = ws5.cell(r, 1, value=lbl)
    c1.fill = PatternFill('solid', start_color=bg)
    c1.font = Font(bold=True, name='Calibri', size=9)
    c1.border = _bdr
    c2 = ws5.cell(r, 2, value=desc)
    c2.font = Font(name='Calibri', size=9)
    ws5.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)

# -- Onglet 6 : Carte du terrain ------------------------------------------
ws6 = wb.create_sheet("6-Carte Terrain")

# Column headers
ws6.column_dimensions['A'].width = 4
for c in range(cols_t):
    ws6.column_dimensions[get_column_letter(c+2)].width = 6.5

ws6.cell(1, 1).value = ""
for c in range(cols_t):
    cell = ws6.cell(1, c+2, value=c+1)
    cell.font = Font(bold=True, name='Calibri', size=7)
    cell.alignment = Alignment(horizontal='center')

# Build label grid
label_grid = {}
for pid, p in placed_best.items():
    if 'r' not in p:
        continue
    info  = per_bat_best.get(pid, {})
    boost = info.get('boost', 0)
    blbl  = boost_lbl.get(boost, '')
    for dr in range(p['h']):
        for dc in range(p['w']):
            is_topleft = (dr == 0 and dc == 0)
            label_grid[(p['r']+dr, p['c']+dc)] = {
                'name': p['nom'] if is_topleft else '',
                'type': p['type'],
                'boost': blbl if (is_topleft and p['type']=='Producteur' and p['prod']!='Rien') else '',
            }

for r in range(rows_t):
    ws6.row_dimensions[r+2].height = 11
    cell0 = ws6.cell(r+2, 1, value=r+1)
    cell0.font = Font(bold=True, name='Calibri', size=7)
    cell0.alignment = Alignment(horizontal='center')

    for c in range(cols_t):
        cell = ws6.cell(r+2, c+2)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.font = Font(name='Calibri', size=6)

        if not terrain[r, c]:
            cell.fill  = DARK_FILL
            cell.value = 'X'
            cell.font  = Font(name='Calibri', size=6, color='FFAAAAAA')
        elif (r, c) in label_grid:
            info = label_grid[(r, c)]
            typ  = info['type']
            cell.fill  = (ORANGE_FILL if typ == 'Culturel'
                           else GREEN_FILL if typ == 'Producteur'
                           else GREY_FILL)
            txt = info['name']
            if info['boost']:
                txt += f"\n{info['boost']}"
            cell.value = txt
            cell.font  = Font(name='Calibri', size=6,
                               color='FFFFFFFF' if typ == 'Culturel' else 'FF000000')
        else:
            cell.fill = WHITE_FILL

# Legend for map
map_lr = rows_t + 4
ws6.cell(map_lr, 1, value="LÉGENDE").font = Font(bold=True, name='Calibri', size=8)
for i, (lbl, fill) in enumerate([
    ('Culturel',   ORANGE_FILL),
    ('Producteur', GREEN_FILL),
    ('Neutre',     GREY_FILL),
    ('Hors terrain', DARK_FILL),
]):
    c = ws6.cell(map_lr+1+i, 1, value=lbl)
    c.fill = fill
    c.font = Font(name='Calibri', size=8,
                  color='FFFFFFFF' if lbl == 'Hors terrain' else 'FF000000')
    c.border = _bdr

# -- Save to buffer ------------------------------------------------------
buf = io.BytesIO()
wb.save(buf)
buf.seek(0)
return buf.getvalue()
```

# ===============================================

# STREAMLIT UI

# ===============================================

col_upload, col_info = st.columns([1, 1], gap=“large”)

with col_upload:
st.markdown(”### 📂 Fichier d’entrée”)
uploaded = st.file_uploader(
“Dépose ton fichier Excel ici”,
type=[“xlsx”],
help=“Le fichier doit contenir les onglets : Terrain, Batiments, Actuel”,
)

with col_info:
st.markdown(”### ℹ️ Format attendu”)
st.markdown(’’’

- **Onglet Terrain** : grille de `1` (libre) et `X` (occupé/bord)
- **Onglet Batiments** : colonnes Nom, Longueur, Largeur, Nombre, Type, Culture, Rayonnement, Boost 25%/50%/100%, Production, Quantite
- **Onglet Actuel** : terrain peint avec les noms des bâtiments placés
  ‘’’)

if uploaded is not None:
st.markdown(’<hr class="section-sep">’, unsafe_allow_html=True)

```
try:
    xl = pd.ExcelFile(uploaded)

    # Validate sheets
    required = {'Terrain', 'Batiments', 'Actuel'}
    missing = required - set(xl.sheet_names)
    if missing:
        st.error(f"Onglets manquants dans le fichier : {', '.join(missing)}")
        st.stop()

    terrain_df = pd.read_excel(xl, sheet_name='Terrain',  header=None)
    bats_df    = pd.read_excel(xl, sheet_name='Batiments', header=0)
    actuel_df  = pd.read_excel(xl, sheet_name='Actuel',    header=None)

    terrain  = parse_terrain(terrain_df)
    buildings = parse_buildings(bats_df)
    actuel_instances = parse_actuel(actuel_df)
    placed_orig = build_placed(actuel_instances, buildings)

    rows_t, cols_t = terrain.shape
    n_free  = int(terrain.sum())
    n_total = len(buildings)
    n_placed = len(placed_orig)
    n_cult   = sum(1 for p in placed_orig.values() if p['type'] == 'Culturel')
    n_prod   = sum(1 for p in placed_orig.values() if p['type'] == 'Producteur')
    n_neutre = sum(1 for p in placed_orig.values() if p['type'] == 'Neutre')

    # -- Stats du fichier ----------------------------------------------
    st.markdown("### 📊 Fichier chargé")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="stat-box"><div class="stat-label">Terrain</div>'
                    f'<div class="stat-value">{rows_t}×{cols_t}</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="stat-box"><div class="stat-label">Cases libres</div>'
                    f'<div class="stat-value">{n_free}</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="stat-box"><div class="stat-label">Bâtiments placés</div>'
                    f'<div class="stat-value">{n_placed}</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="stat-box"><div class="stat-label">🟠 Cult / 🟢 Prod / ⬜ Neutre</div>'
                    f'<div class="stat-value">{n_cult} / {n_prod} / {n_neutre}</div></div>',
                    unsafe_allow_html=True)

    # -- Production de référence ---------------------------------------
    orig_prod, _, _ = compute_production_detail(placed_orig)
    st.markdown("**Production initiale (référence) :**")
    prod_cols = st.columns(min(len(orig_prod), 5))
    for i, (pr, qty) in enumerate(sorted(orig_prod.items(),
                                          key=lambda x: -x[1])):
        with prod_cols[i % len(prod_cols)]:
            st.metric(pr, f"{qty:,.0f} /h")

    # -- Bouton lancement ---------------------------------------------
    st.markdown('<hr class="section-sep">', unsafe_allow_html=True)
    st.markdown("### 🚀 Lancer l'optimisation")

    if st.button("▶  Optimiser le placement", type="primary", use_container_width=True):
        prog_bar  = st.progress(0.0)
        prog_text = st.empty()

        def update_progress(pct, msg):
            prog_bar.progress(pct)
            prog_text.text(msg)

        with st.spinner("Optimisation en cours…"):
            placed_best, best_score = optimize(
                placed_orig, terrain, progress_cb=update_progress)

        prog_bar.progress(1.0)
        prog_text.text("✅ Optimisation terminée !")

        best_prod, _, _ = compute_production_detail(placed_best)

        # -- Résultats ---------------------------------------------
        st.markdown('<hr class="section-sep">', unsafe_allow_html=True)
        st.markdown("### 🏆 Résultats de l'optimisation")

        # Moved buildings
        steps, moved_list = build_sequence(placed_orig, placed_best, terrain)
        n_moved = len(moved_list)

        mc1, mc2, mc3 = st.columns(3)
        with mc1:
            st.markdown(f'<div class="stat-box"><div class="stat-label">Bâtiments déplacés</div>'
                        f'<div class="stat-value">{n_moved}</div></div>', unsafe_allow_html=True)
        with mc2:
            st.markdown(f'<div class="stat-box"><div class="stat-label">Étapes de manipulation</div>'
                        f'<div class="stat-value">{len(steps)}</div></div>', unsafe_allow_html=True)
        with mc3:
            score_gain = (best_score - compute_score(placed_orig)) / compute_score(placed_orig) * 100
            color = "gain-positive" if score_gain >= 0 else "gain-negative"
            st.markdown(f'<div class="stat-box"><div class="stat-label">Score global</div>'
                        f'<div class="stat-value"><span class="{color}">+{score_gain:.1f}%</span></div></div>',
                        unsafe_allow_html=True)

        # Production comparison table
        all_p = sorted(set(orig_prod) | set(best_prod),
                       key=lambda x: -(best_prod.get(x, 0)))
        PROD_PRIORITY_ALL = ['Guerison','Nourriture','Or']
        all_p = sorted(all_p, key=lambda x: (
            PROD_PRIORITY_ALL.index(x) if x in PROD_PRIORITY_ALL else 99, -best_prod.get(x, 0)))

        table_data = []
        for pr in all_p:
            if pr == 'Rien':
                continue
            o  = orig_prod.get(pr, 0)
            b  = best_prod.get(pr, 0)
            d  = b - o
            pt = f"+{d:,.0f}" if d >= 0 else f"{d:,.0f}"
            pp = f"+{d/o*100:.1f}%" if o > 0 else "--"
            table_data.append({
                'Production': pr,
                'Avant /h':   f"{o:,.0f}",
                'Après /h':   f"{b:,.0f}",
                'Gain /h':    pt,
                'Gain %':     pp,
            })
        st.dataframe(pd.DataFrame(table_data), use_container_width=True, hide_index=True)

        # Moved buildings summary
        if n_moved > 0:
            st.markdown(f"**{n_moved} bâtiment(s) à déplacer :**")
            mv_data = [{'Bâtiment': m['nom'], 'Type': m['type'],
                        'De': f"L{m['fr']+1}-C{m['fc']+1}",
                        'Vers': f"L{m['tr']+1}-C{m['tc']+1}"}
                       for m in moved_list]
            st.dataframe(pd.DataFrame(mv_data), use_container_width=True, hide_index=True)

        # -- Génération et téléchargement --------------------------
        st.markdown('<hr class="section-sep">', unsafe_allow_html=True)
        st.markdown("### 📥 Télécharger le résultat")

        with st.spinner("Génération du fichier Excel…"):
            excel_bytes = generate_excel(
                placed_orig, placed_best, terrain, orig_prod, best_prod)

        st.download_button(
            label="⬇️  Télécharger Ville_optimisée.xlsx",
            data=excel_bytes,
            file_name="Ville_optimisee.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )

        st.success("Le fichier contient 6 onglets : bâtiments optimisés, résumé production, "
                   "comparatif avant/après, bâtiments déplacés, séquence des opérations, "
                   "et carte du terrain.")

except Exception as e:
    st.error(f"Erreur lors du traitement du fichier : {e}")
    st.exception(e)
```

else:
st.info(“👆 Dépose ton fichier Excel pour commencer l’optimisation.”)
