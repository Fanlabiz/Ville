import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="Placement de Batiments", layout="wide")
st.title("Optimiseur de placement de batiments")

# ─── Chargement ───────────────────────────────────────────────────────────────

def load_terrain(ws):
    grid = []
    for row in ws.iter_rows(values_only=True):
        last = -1
        for i, v in enumerate(row):
            if v is not None:
                last = i
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
                qty = 0
            else:
                try:
                    s = str(qty_raw).strip().lstrip('=')
                    qty = float(eval(s))
                except Exception:
                    qty = 0.0
            b = {
                'nom': str(row[0]).strip(),
                'longueur': int(row[1]),
                'largeur': int(row[2]),
                'nombre': int(row[3]),
                'type': str(row[4]).strip(),
                'culture': float(row[5]) if row[5] else 0,
                'rayonnement': int(row[6]) if row[6] else 0,
                'boost25': float(row[7]) if row[7] is not None else None,
                'boost50': float(row[8]) if row[8] is not None else None,
                'boost100': float(row[9]) if row[9] is not None else None,
                'production': str(row[10]).strip() if row[10] else 'Rien',
                'quantite': qty,
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

# ─── Algorithme ───────────────────────────────────────────────────────────────

PROD_PRIORITY = ['Guerison', 'Nourriture', 'Or']

def expand_buildings(buildings):
    result = []
    for b in buildings:
        for _ in range(b['nombre']):
            result.append(dict(b))
    return result

def can_place(grid, r, c, h, w):
    rows, cols = grid.shape
    if r + h > rows or c + w > cols:
        return False
    return bool(grid[r:r+h, c:c+w].all())

def do_place(grid, r, c, h, w):
    grid[r:r+h, c:c+w] = False

def build_culture_map(placed, rows, cols):
    cmap = np.zeros((rows, cols), dtype=float)
    for p in placed:
        if p['type'] != 'Culturel' or p['culture'] <= 0:
            continue
        ray = p['rayonnement']
        r0, c0, h, w = p['row'], p['col'], p['h'], p['w']
        rmin = max(0, r0 - ray)
        rmax = min(rows, r0 + h + ray)
        cmin = max(0, c0 - ray)
        cmax = min(cols, c0 + w + ray)
        cmap[rmin:rmax, cmin:cmax] += p['culture']
        cmap[r0:r0+h, c0:c0+w] -= p['culture']
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

def find_border_pos(grid, rows, cols, h, w):
    best, best_dist = None, 9999
    for bh, bw in (set([(h, w), (w, h)])):
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
    for bh, bw in set([(h, w), (w, h)]):
        for r in range(rows):
            for c in range(cols):
                if not can_place(grid, r, c, bh, bw):
                    continue
                rmin = max(0, r - ray); rmax = min(rows, r + bh + ray)
                cmin = max(0, c - ray); cmax = min(cols, c + bw + ray)
                zone = grid[rmin:rmax, cmin:cmax].copy()
                dr0 = r - rmin; dc0 = c - cmin
                zone[dr0:dr0+bh, dc0:dc0+bw] = False
                score = int(zone.sum())
                if score > best_score:
                    best_score = score
                    best = (r, c, bh, bw)
    return best

def find_producer_pos(grid, rows, cols, h, w, cmap):
    best, best_val = None, -1
    for bh, bw in set([(h, w), (w, h)]):
        for r in range(rows):
            for c in range(cols):
                if can_place(grid, r, c, bh, bw):
                    val = float(cmap[r:r+bh, c:c+bw].max())
                    if val > best_val or best is None:
                        best_val = val
                        best = (r, c, bh, bw)
    return best

def find_any_pos(grid, rows, cols, h, w):
    for bh, bw in set([(h, w), (w, h)]):
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

def run_placement(grid_orig, buildings_raw):
    grid = grid_orig.copy()
    rows, cols = grid.shape
    instances = expand_buildings(buildings_raw)
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

    # 2. Culturels et producteurs en alternance
    culturels = sorted([b for b in instances if b['type'] == 'Culturel'],
                       key=lambda b: b['longueur'] * b['largeur'], reverse=True)
    producteurs = sorted([b for b in instances if b['type'] == 'Producteur'],
                         key=prod_sort_key)

    ci = pi = 0
    while ci < len(culturels) or pi < len(producteurs):
        if ci < len(culturels):
            b = culturels[ci]
            pos = find_cultural_pos(grid, rows, cols, b['longueur'], b['largeur'], b['rayonnement'])
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
            pos = find_producer_pos(grid, rows, cols, b['longueur'], b['largeur'], cmap)
            if not pos:
                pos = find_any_pos(grid, rows, cols, b['longueur'], b['largeur'])
            if pos:
                r, c, h, w = pos
                placed.append({**b, 'row': r, 'col': c, 'h': h, 'w': w})
                do_place(grid, r, c, h, w)
            else:
                unplaced.append(b)
            pi += 1

    compute_culture_received(placed, rows, cols)
    free_cells = int(grid.sum())
    return placed, unplaced, free_cells

# ─── Export Excel ─────────────────────────────────────────────────────────────

def generate_excel(placed, unplaced, free_cells, grid_matrix, terrain_grid):
    wb = Workbook()

    # Onglet 1 : liste des batiments places
    ws1 = wb.active
    ws1.title = "Batiments places"
    hdr = ["Nom", "Type", "Production", "Ligne", "Colonne", "Hauteur", "Largeur",
           "Culture recue", "Boost (%)", "Quantite/h", "Prod totale/h"]
    for i, h in enumerate(hdr, 1):
        cell = ws1.cell(row=1, column=i, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")
    for ri, p in enumerate(placed, 2):
        boost_pct, mult = get_boost(p)
        prod = p['quantite'] * mult if p['production'] != 'Rien' else 0
        vals = [p['nom'], p['type'], p['production'],
                p['row'] + 1, p['col'] + 1, p['h'], p['w'],
                round(p.get('culture_recue', 0), 1), boost_pct,
                p['quantite'], round(prod, 1)]
        for ci, v in enumerate(vals, 1):
            ws1.cell(row=ri, column=ci, value=v)
    for col in ws1.columns:
        ws1.column_dimensions[get_column_letter(col[0].column)].width = 22

    # Onglet 2 : synthese par production
    ws2 = wb.create_sheet("Synthese")
    ws2["A1"] = "Synthese par type de production"
    ws2["A1"].font = Font(bold=True, size=13)
    hdr2 = ["Production", "Culture totale", "Boost moyen (%)", "Nb batiments", "Production/h"]
    for i, h in enumerate(hdr2, 1):
        cell = ws2.cell(row=2, column=i, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E79")
    prod_groups = {}
    for p in placed:
        if p['type'] == 'Producteur' and p['production'] != 'Rien':
            key = p['production']
            boost_pct, mult = get_boost(p)
            if key not in prod_groups:
                prod_groups[key] = {'cultures': [], 'boosts': [], 'qtys': []}
            prod_groups[key]['cultures'].append(p.get('culture_recue', 0))
            prod_groups[key]['boosts'].append(boost_pct)
            prod_groups[key]['qtys'].append(p['quantite'] * mult)

    def pg_key(k):
        return PROD_PRIORITY.index(k) if k in PROD_PRIORITY else 999

    row = 3
    for key in sorted(prod_groups, key=pg_key):
        g = prod_groups[key]
        ws2.cell(row=row, column=1, value=key)
        ws2.cell(row=row, column=2, value=round(sum(g['cultures']), 1))
        ws2.cell(row=row, column=3, value=round(sum(g['boosts']) / len(g['boosts']), 1))
        ws2.cell(row=row, column=4, value=len(g['qtys']))
        ws2.cell(row=row, column=5, value=round(sum(g['qtys']), 1))
        row += 1
    for col in ws2.columns:
        ws2.column_dimensions[get_column_letter(col[0].column)].width = 25

    # Onglet 3 : terrain visuel
    ws3 = wb.create_sheet("Terrain")
    rows_n, cols_n = grid_matrix.shape
    cell_map = {}
    for p in placed:
        r0, c0, h, w = p['row'], p['col'], p['h'], p['w']
        for dr in range(h):
            for dc in range(w):
                if (r0 + dr, c0 + dc) not in cell_map:
                    cell_map[(r0 + dr, c0 + dc)] = p

    x_set = set()
    for ri, row_data in enumerate(terrain_grid):
        for ci, v in enumerate(row_data):
            if v == 'X':
                x_set.add((ri, ci))

    FILL_X = PatternFill("solid", fgColor="404040")
    FILL_C = PatternFill("solid", fgColor="FFA500")
    FILL_P = PatternFill("solid", fgColor="90EE90")
    FILL_N = PatternFill("solid", fgColor="D3D3D3")
    FILL_F = PatternFill("solid", fgColor="FFFFFF")

    for r in range(rows_n):
        ws3.row_dimensions[r + 1].height = 18
        for c in range(cols_n):
            ws3.column_dimensions[get_column_letter(c + 1)].width = 3
            cell = ws3.cell(row=r + 1, column=c + 1)
            cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
            if (r, c) in x_set:
                cell.fill = FILL_X
                cell.value = "X"
                cell.font = Font(color="FFFFFF", size=6)
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
                        label += f" +{boost_pct}%"
                    cell.value = label
                    cell.font = Font(bold=True, size=6)
            else:
                cell.fill = FILL_F

    # Onglet 4 : non places
    ws4 = wb.create_sheet("Non places")
    ws4["A1"] = "Batiments non places"
    ws4["A1"].font = Font(bold=True, size=13)
    for i, h in enumerate(["Nom", "Type", "Taille", "Production"], 1):
        c = ws4.cell(row=2, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E79")
    total_cells = 0
    for ri, b in enumerate(unplaced, 3):
        ws4.cell(row=ri, column=1, value=b['nom'])
        ws4.cell(row=ri, column=2, value=b['type'])
        ws4.cell(row=ri, column=3, value=f"{b['longueur']}x{b['largeur']}")
        ws4.cell(row=ri, column=4, value=b['production'])
        total_cells += b['longueur'] * b['largeur']
    end = len(unplaced) + 4
    ws4.cell(row=end, column=1, value="Cases libres restantes:").font = Font(bold=True)
    ws4.cell(row=end, column=2, value=free_cells)
    ws4.cell(row=end + 1, column=1, value="Cases des batiments non places:").font = Font(bold=True)
    ws4.cell(row=end + 1, column=2, value=total_cells)
    for col in ws4.columns:
        ws4.column_dimensions[get_column_letter(col[0].column)].width = 30

    return wb

# ─── Interface ────────────────────────────────────────────────────────────────

uploaded = st.file_uploader("Charger le fichier Excel de la ville", type=["xlsx"])

if uploaded:
    wb_in = load_workbook(io.BytesIO(uploaded.read()), read_only=True)
    if 'Terrain' not in wb_in.sheetnames or 'Batiments' not in wb_in.sheetnames:
        st.error("Le fichier doit contenir les onglets 'Terrain' et 'Batiments'")
        st.stop()

    terrain_grid = load_terrain(wb_in['Terrain'])
    buildings_raw = load_buildings(wb_in['Batiments'])
    grid_matrix = grid_to_matrix(terrain_grid)
    rows_n, cols_n = grid_matrix.shape

    st.success(f"Terrain charge : {rows_n} lignes x {cols_n} colonnes  |  {int(grid_matrix.sum())} cases libres")

    types_count = {}
    for b in buildings_raw:
        types_count[b['type']] = types_count.get(b['type'], 0) + b['nombre']
    c1, c2, c3 = st.columns(3)
    c1.metric("Batiments neutres", types_count.get('Neutre', 0))
    c2.metric("Batiments culturels", types_count.get('Culturel', 0))
    c3.metric("Batiments producteurs", types_count.get('Producteur', 0))

    if st.button("Lancer le placement optimise", type="primary"):
        with st.spinner("Optimisation en cours..."):
            placed, unplaced, free_cells = run_placement(grid_matrix, buildings_raw)

        st.success(f"Placement termine : {len(placed)} places, {len(unplaced)} non places, {free_cells} cases libres")

        # Synthese
        st.subheader("Synthese des productions")
        prod_data = {}
        for p in placed:
            if p['type'] == 'Producteur' and p['production'] != 'Rien':
                key = p['production']
                boost_pct, mult = get_boost(p)
                if key not in prod_data:
                    prod_data[key] = {'cultures': [], 'boosts': [], 'prod': 0, 'n': 0}
                prod_data[key]['cultures'].append(p.get('culture_recue', 0))
                prod_data[key]['boosts'].append(boost_pct)
                prod_data[key]['prod'] += p['quantite'] * mult
                prod_data[key]['n'] += 1

        rows_list = []
        for k in sorted(prod_data, key=lambda x: PROD_PRIORITY.index(x) if x in PROD_PRIORITY else 999):
            d = prod_data[k]
            rows_list.append({
                "Production": k,
                "Culture totale": round(sum(d['cultures']), 0),
                "Boost moyen (%)": round(sum(d['boosts']) / len(d['boosts']), 1),
                "Nb batiments": d['n'],
                "Production/h": round(d['prod'], 1),
            })
        if rows_list:
            st.dataframe(pd.DataFrame(rows_list), hide_index=True, use_container_width=True)

        if unplaced:
            st.subheader(f"{len(unplaced)} batiments non places")
            st.dataframe(pd.DataFrame([{
                'Nom': b['nom'], 'Type': b['type'],
                'Taille': f"{b['longueur']}x{b['largeur']}",
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
