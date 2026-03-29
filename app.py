import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import copy
from collections import defaultdict

st.set_page_config(page_title="Optimiseur de Ville", layout="wide")
st.title("🏙️ Optimiseur de placement des bâtiments")

# ─────────────────────────────────────────────
# LECTURE DU FICHIER EXCEL
# ─────────────────────────────────────────────

def load_excel(file):
    xl = pd.ExcelFile(file)
    # Terrain
    terrain_df = pd.read_excel(file, sheet_name='Terrain', header=None)
    terrain = terrain_df.fillna('').values.tolist()
    # Normalise: X=bord, 1=libre, ''=hors terrain
    grid = []
    for row in terrain:
        grid_row = []
        for cell in row:
            v = str(cell).strip()
            if v == 'X':
                grid_row.append('X')
            elif v == '1':
                grid_row.append(1)
            else:
                grid_row.append(None)  # hors terrain
        grid.append(grid_row)

    # Bâtiments
    bat_df = pd.read_excel(file, sheet_name='Batiments', header=0)
    bat_df.columns = [c.strip() for c in bat_df.columns]
    bat_df = bat_df.fillna(0)
    buildings = []
    for _, row in bat_df.iterrows():
        b = {
            'nom': str(row['Nom']).strip(),
            'longueur': int(row['Longueur']),
            'largeur': int(row['Largeur']),
            'nombre': int(row['Nombre']),
            'type': str(row['Type']).strip(),
            'culture': float(row['Culture']),
            'rayonnement': int(row['Rayonnement']),
            'boost25': float(row['Boost 25%']) if row['Boost 25%'] != 0 else None,
            'boost50': float(row['Boost 50%']) if row['Boost 50%'] != 0 else None,
            'boost100': float(row['Boost 100%']) if row['Boost 100%'] != 0 else None,
            'production': str(row['Production']).strip(),
            'quantite': float(row['Quantite']),
        }
        buildings.append(b)

    # Actuel (placement existant)
    actuel_df = pd.read_excel(file, sheet_name='Actuel', header=None)
    actuel = actuel_df.fillna('').values.tolist()

    return grid, buildings, actuel

# ─────────────────────────────────────────────
# PARSING DU PLACEMENT ACTUEL
# ─────────────────────────────────────────────

def parse_actuel(actuel, buildings):
    """Extrait les placements actuels depuis la feuille Actuel."""
    bname_to_info = {b['nom']: b for b in buildings}
    placed = {}  # nom_instance -> (row, col, h, w, orientation)
    # Trouve les rectangles de chaque bâtiment
    seen_cells = {}
    for r, row in enumerate(actuel):
        for c, cell in enumerate(row):
            v = str(cell).strip()
            if v and v not in ('X', '1', '0', ''):
                if v not in seen_cells:
                    seen_cells[v] = []
                seen_cells[v].append((r, c))

    # Groupe par nom
    name_groups = defaultdict(list)
    for name, cells in seen_cells.items():
        name_groups[name] = cells

    placed_list = []
    for name, cells in name_groups.items():
        if not cells:
            continue
        rows = [c[0] for c in cells]
        cols = [c[1] for c in cells]
        min_r, max_r = min(rows), max(rows)
        min_c, max_c = min(cols), max(cols)
        h = max_r - min_r + 1
        w = max_c - min_c + 1
        # Peut y avoir plusieurs occurrences: on décompose
        info = bname_to_info.get(name)
        if info is None:
            continue
        exp_h_h = info['largeur']  # orientation H: h=largeur
        exp_w_h = info['longueur']
        exp_h_v = info['longueur']
        exp_w_v = info['largeur']

        # Découpage en instances
        instances = split_into_instances(cells, info)
        for inst_cells in instances:
            r_cells = [c[0] for c in inst_cells]
            c_cells = [c[1] for c in inst_cells]
            ir, ic = min(r_cells), min(c_cells)
            ih = max(r_cells) - ir + 1
            iw = max(c_cells) - ic + 1
            placed_list.append({
                'nom': name,
                'row': ir, 'col': ic,
                'h': ih, 'w': iw,
            })

    return placed_list


def split_into_instances(cells, info):
    """Décompose un ensemble de cases en instances distinctes d'un bâtiment."""
    L, l = info['longueur'], info['largeur']
    cells_set = set(cells)
    remaining = set(cells)
    instances = []
    # Essaie de former des rectangles L×l ou l×L
    sorted_cells = sorted(remaining)
    while remaining:
        r, c = min(remaining)
        found = False
        for (h, w) in [(l, L), (L, l)]:
            rect = set()
            ok = True
            for dr in range(h):
                for dc in range(w):
                    if (r+dr, c+dc) not in remaining:
                        ok = False
                        break
                if not ok:
                    break
                for dc in range(w):
                    rect.add((r+dr, c+dc))
            if ok and len(rect) == h*w:
                instances.append(list(rect))
                remaining -= rect
                found = True
                break
        if not found:
            # Retire la cellule pour éviter boucle infinie
            remaining.discard((r, c))
    return instances if instances else [list(cells)]

# ─────────────────────────────────────────────
# GRILLE D'OCCUPATION
# ─────────────────────────────────────────────

def build_occupied_grid(grid, placed_list):
    rows = len(grid)
    cols = max(len(r) for r in grid)
    occ = {}  # (r,c) -> nom_instance
    for idx, p in enumerate(placed_list):
        key = f"{p['nom']}#{idx}"
        for dr in range(p['h']):
            for dc in range(p['w']):
                occ[(p['row']+dr, p['col']+dc)] = key
    return occ

def is_free(grid, occ, row, col, h, w):
    rows = len(grid)
    for dr in range(h):
        for dc in range(w):
            r, c = row+dr, col+dc
            if r >= len(grid) or c >= len(grid[r]):
                return False
            cell = grid[r][c]
            if cell == 'X' or cell is None:
                return False
            if (r, c) in occ:
                return False
    return True

def place_building(occ, row, col, h, w, key):
    for dr in range(h):
        for dc in range(w):
            occ[(row+dr, col+dc)] = key

def remove_building(occ, row, col, h, w):
    for dr in range(h):
        for dc in range(w):
            occ.pop((row+dr, col+dc), None)

# ─────────────────────────────────────────────
# CALCUL DE LA CULTURE REÇUE
# ─────────────────────────────────────────────

def compute_culture(placed_list, buildings_by_name):
    """Calcule la culture reçue par chaque bâtiment producteur."""
    culturels = [p for p in placed_list if buildings_by_name[p['nom']]['type'] == 'Culturel']
    culture_received = {}

    for idx, p in enumerate(placed_list):
        info = buildings_by_name[p['nom']]
        if info['type'] != 'Producteur':
            continue
        total_culture = 0
        for cp in culturels:
            ci = buildings_by_name[cp['nom']]
            ray = ci['rayonnement']
            if ray == 0:
                continue
            # Zone de rayonnement = bande de largeur ray autour du bâtiment culturel
            cr0 = cp['row'] - ray
            cr1 = cp['row'] + cp['h'] - 1 + ray
            cc0 = cp['col'] - ray
            cc1 = cp['col'] + cp['w'] - 1 + ray
            # Le producteur est-il partiellement dans cette zone?
            pr0, pr1 = p['row'], p['row'] + p['h'] - 1
            pc0, pc1 = p['col'], p['col'] + p['w'] - 1
            if pr0 <= cr1 and pr1 >= cr0 and pc0 <= cc1 and pc1 >= cc0:
                total_culture += ci['culture']
        culture_received[idx] = total_culture

    return culture_received

def get_boost(culture, info):
    b100 = info['boost100']
    b50  = info['boost50']
    b25  = info['boost25']
    if b100 and culture >= b100:
        return 1.0, '100%'
    elif b50 and culture >= b50:
        return 0.5, '50%'
    elif b25 and culture >= b25:
        return 0.25, '25%'
    else:
        return 0.0, '0%'

def compute_total_production(placed_list, buildings_by_name, culture_received):
    totals = defaultdict(float)
    for idx, p in enumerate(placed_list):
        info = buildings_by_name[p['nom']]
        if info['type'] != 'Producteur' or info['production'] == 'Rien':
            continue
        culture = culture_received.get(idx, 0)
        boost, _ = get_boost(culture, info)
        prod = info['quantite'] * (1 + boost)
        totals[info['production']] += prod
    return dict(totals)

# ─────────────────────────────────────────────
# OPTIMISATION
# ─────────────────────────────────────────────

PROD_PRIORITY = ['Guerison', 'Nourriture', 'Or']

def prod_score(totals):
    """Score pondéré par priorité de production."""
    weights = {'Guerison': 1e12, 'Nourriture': 1e8, 'Or': 1e4}
    score = 0
    for prod, qty in totals.items():
        w = weights.get(prod, 1)
        score += qty * w
    return score

def optimize(grid, placed_list, buildings_by_name, n_passes=3):
    """Optimise le placement par échanges de paires de bâtiments."""
    best = copy.deepcopy(placed_list)
    culture = compute_culture(best, buildings_by_name)
    best_totals = compute_total_production(best, buildings_by_name, culture)
    best_score = prod_score(best_totals)

    for pass_num in range(n_passes):
        improved = True
        while improved:
            improved = False
            n = len(best)
            for i in range(n):
                for j in range(i+1, n):
                    candidate = copy.deepcopy(best)
                    # Échange positions
                    pi, pj = candidate[i], candidate[j]
                    # Essaie de placer i à la position de j et vice versa
                    # (même orientation si dimensions compatibles)
                    swapped = try_swap(grid, candidate, i, j, buildings_by_name)
                    if swapped is None:
                        continue
                    c2 = compute_culture(swapped, buildings_by_name)
                    t2 = compute_total_production(swapped, buildings_by_name, c2)
                    s2 = prod_score(t2)
                    if s2 > best_score:
                        best = swapped
                        best_score = s2
                        improved = True

    return best

def try_swap(grid, placed_list, i, j, buildings_by_name):
    """Tente d'échanger les positions de deux bâtiments."""
    p = copy.deepcopy(placed_list)
    pi, pj = p[i], p[j]
    ri, ci_, hi, wi = pi['row'], pi['col'], pi['h'], pi['w']
    rj, cj, hj, wj = pj['row'], pj['col'], pj['h'], pj['w']

    # Essaie toutes les combinaisons d'orientations
    for (nhi, nwi) in [(hi, wi), (wi, hi)]:
        for (nhj, nwj) in [(hj, wj), (wj, hj)]:
            cand = copy.deepcopy(p)
            occ = build_occupied_grid(grid, cand)
            # Libère i et j
            remove_building(occ, ri, ci_, hi, wi)
            remove_building(occ, rj, cj, hj, wj)
            # Tente i à position j
            if not is_free(grid, occ, rj, cj, nhi, nwi):
                continue
            place_building(occ, rj, cj, nhi, nwi, f"{pi['nom']}#{i}")
            # Tente j à position i
            if not is_free(grid, occ, ri, ci_, nhj, nwj):
                remove_building(occ, rj, cj, nhi, nwi)
                continue
            # Succès
            cand[i] = {'nom': pi['nom'], 'row': rj, 'col': cj, 'h': nhi, 'w': nwi}
            cand[j] = {'nom': pj['nom'], 'row': ri, 'col': ci_, 'h': nhj, 'w': nwj}
            return cand
    return None

# ─────────────────────────────────────────────
# GÉNÉRATION DU FICHIER RÉSULTAT
# ─────────────────────────────────────────────

ORANGE = 'FFFFA500'
GREEN  = 'FF92D050'
GREY   = 'FFD9D9D9'
BLUE   = 'FF4472C4'
WHITE  = 'FFFFFFFF'
YELLOW = 'FFFFFF00'
LIGHT_BLUE = 'FFBDD7EE'

def col_for_type(btype):
    if btype == 'Culturel':
        return ORANGE
    elif btype == 'Producteur':
        return GREEN
    else:
        return GREY

def make_fill(hex_color):
    return PatternFill('solid', fgColor=hex_color[2:])

def make_border():
    thin = Side(style='thin')
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def write_result(grid, placed_list, original_placed, buildings_by_name):
    wb = openpyxl.Workbook()

    culture_received = compute_culture(placed_list, buildings_by_name)
    culture_orig = compute_culture(original_placed, buildings_by_name)
    totals_new = compute_total_production(placed_list, buildings_by_name, culture_received)
    totals_old = compute_total_production(original_placed, buildings_by_name, culture_orig)

    # ── Onglet 1 : Liste des bâtiments ─────────────────────
    ws1 = wb.active
    ws1.title = "Batiments placés"
    headers = ['Nom', 'Type', 'Production', 'Ligne', 'Colonne', 'H', 'L',
               'Orientation', 'Culture reçue', 'Boost atteint']
    hfill = make_fill(BLUE)
    hfont = Font(bold=True, color='FFFFFFFF')
    for ci_, h in enumerate(headers, 1):
        c = ws1.cell(1, ci_, h)
        c.fill = hfill
        c.font = hfont
        c.alignment = Alignment(horizontal='center')

    for idx, p in enumerate(placed_list):
        info = buildings_by_name[p['nom']]
        culture = culture_received.get(idx, 0)
        boost, boost_label = get_boost(culture, info)
        orient = 'H' if p['w'] >= p['h'] else 'V'
        row_data = [
            p['nom'], info['type'], info['production'],
            p['row']+1, p['col']+1, p['h'], p['w'],
            orient, culture, boost_label
        ]
        r = idx + 2
        for ci_, val in enumerate(row_data, 1):
            c = ws1.cell(r, ci_, val)
            c.fill = make_fill(col_for_type(info['type']))
            c.border = make_border()
            c.alignment = Alignment(horizontal='center')
    for col in ws1.columns:
        ws1.column_dimensions[get_column_letter(col[0].column)].width = 20

    # ── Onglet 2 : Synthèse culture & boosts ───────────────
    ws2 = wb.create_sheet("Synthèse culture")
    ws2.append(['Production', 'Culture totale reçue', 'Boost moyen', 'Nb bâtiments'])
    ws2['A1'].font = Font(bold=True)
    by_prod = defaultdict(lambda: {'culture': 0, 'boosts': [], 'count': 0})
    for idx, p in enumerate(placed_list):
        info = buildings_by_name[p['nom']]
        if info['type'] != 'Producteur' or info['production'] == 'Rien':
            continue
        culture = culture_received.get(idx, 0)
        boost, _ = get_boost(culture, info)
        by_prod[info['production']]['culture'] += culture
        by_prod[info['production']]['boosts'].append(boost)
        by_prod[info['production']]['count'] += 1
    for prod, data in sorted(by_prod.items()):
        avg_boost = sum(data['boosts']) / len(data['boosts']) if data['boosts'] else 0
        ws2.append([prod, data['culture'], f"{avg_boost*100:.1f}%", data['count']])
    for col in ws2.columns:
        ws2.column_dimensions[get_column_letter(col[0].column)].width = 22

    # ── Onglet 3 : Productions par heure ───────────────────
    ws3 = wb.create_sheet("Productions / heure")
    ws3.append(['Production', 'Quantité nouvelle', 'Quantité originale', 'Gain/Perte', '% variation'])
    ws3['A1'].font = Font(bold=True)
    all_prods = set(list(totals_new.keys()) + list(totals_old.keys()))
    prod_order = PROD_PRIORITY + [p for p in sorted(all_prods) if p not in PROD_PRIORITY]
    for prod in prod_order:
        if prod not in all_prods:
            continue
        new_v = totals_new.get(prod, 0)
        old_v = totals_old.get(prod, 0)
        gain = new_v - old_v
        pct = (gain / old_v * 100) if old_v else 0
        r = ws3.max_row + 1
        ws3.append([prod, round(new_v, 1), round(old_v, 1), round(gain, 1), f"{pct:+.1f}%"])
        fill_color = GREEN if gain >= 0 else 'FFFF0000'
        for ci_ in range(1, 6):
            ws3.cell(r, ci_).fill = make_fill(fill_color)
            ws3.cell(r, ci_).border = make_border()
    for col in ws3.columns:
        ws3.column_dimensions[get_column_letter(col[0].column)].width = 22

    # ── Onglet 4 : Bâtiments déplacés ─────────────────────
    ws4 = wb.create_sheet("Déplacements")
    ws4.append(['Nom', 'Ligne avant', 'Col avant', 'Ligne après', 'Col après', 'Orientation avant', 'Orientation après'])
    ws4['A1'].font = Font(bold=True)
    orig_by_nom = defaultdict(list)
    for p in original_placed:
        orig_by_nom[p['nom']].append(p)
    new_by_nom = defaultdict(list)
    for p in placed_list:
        new_by_nom[p['nom']].append(p)

    moves = []
    for nom in orig_by_nom:
        orig_instances = sorted(orig_by_nom[nom], key=lambda x: (x['row'], x['col']))
        new_instances  = sorted(new_by_nom.get(nom, []), key=lambda x: (x['row'], x['col']))
        for oi, ni in zip(orig_instances, new_instances):
            if oi['row'] != ni['row'] or oi['col'] != ni['col']:
                moves.append((nom, oi, ni))
                ws4.append([
                    nom,
                    oi['row']+1, oi['col']+1,
                    ni['row']+1, ni['col']+1,
                    'H' if oi['w'] >= oi['h'] else 'V',
                    'H' if ni['w'] >= ni['h'] else 'V',
                ])
    for col in ws4.columns:
        ws4.column_dimensions[get_column_letter(col[0].column)].width = 20

    # ── Onglet 5 : Séquence des opérations ─────────────────
    ws5 = wb.create_sheet("Séquence opérations")
    ws5.append(['Étape', 'Action', 'Bâtiment', 'De (ligne, col)', 'Vers (ligne, col)', 'Note'])
    ws5['A1'].font = Font(bold=True)
    step = 1
    for nom, oi, ni in moves:
        ws5.append([
            step,
            'DÉPLACER',
            nom,
            f"({oi['row']+1}, {oi['col']+1})",
            f"({ni['row']+1}, {ni['col']+1})",
            'Déplacer hors terrain si position occupée, puis replacer'
        ])
        step += 1
    for col in ws5.columns:
        ws5.column_dimensions[get_column_letter(col[0].column)].width = 25

    # ── Onglet 6 : Carte du terrain (nouveau placement) ────
    ws6 = wb.create_sheet("Carte optimisée")
    _draw_map(ws6, grid, placed_list, buildings_by_name, culture_received, "OPTIMISÉ")

    # ── Onglet 7 : Carte originale ──────────────────────────
    ws7 = wb.create_sheet("Carte originale")
    _draw_map(ws7, grid, original_placed, buildings_by_name, culture_orig, "ORIGINAL")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _draw_map(ws, grid, placed_list, buildings_by_name, culture_received, title):
    """Dessine la carte du terrain sur un onglet."""
    CELL_W = 12
    CELL_H = 25
    occ_map = {}  # (r,c) -> (nom, info, culture, boost_label)
    for idx, p in enumerate(placed_list):
        info = buildings_by_name[p['nom']]
        culture = culture_received.get(idx, 0)
        _, boost_label = get_boost(culture, info)
        for dr in range(p['h']):
            for dc in range(p['w']):
                occ_map[(p['row']+dr, p['col']+dc)] = (p['nom'], info, culture, boost_label)

    rows = len(grid)
    cols = max(len(r) for r in grid)

    # Titre
    ws.cell(1, 1, f"Carte du terrain – {title}")
    ws.cell(1, 1).font = Font(bold=True, size=14)

    for r in range(rows):
        for c in range(cols):
            cell_val = grid[r][c] if c < len(grid[r]) else None
            excel_cell = ws.cell(r+2, c+1)
            excel_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            excel_cell.border = make_border()
            if cell_val == 'X':
                excel_cell.fill = make_fill('FF404040')
                excel_cell.value = 'X'
                excel_cell.font = Font(color='FFFFFFFF', size=7)
            elif cell_val is None:
                excel_cell.fill = make_fill('FFF0F0F0')
            elif (r, c) in occ_map:
                nom, info, culture, boost_label = occ_map[(r, c)]
                col_hex = col_for_type(info['type'])
                excel_cell.fill = make_fill(col_hex)
                # Affiche le nom court + boost
                short = nom[:12]
                if info['type'] == 'Producteur' and boost_label != '0%':
                    short += f"\n{boost_label}"
                excel_cell.value = short
                excel_cell.font = Font(size=7, bold=True)
            else:
                excel_cell.fill = make_fill(WHITE)

    for c in range(cols):
        ws.column_dimensions[get_column_letter(c+1)].width = CELL_W
    for r in range(rows):
        ws.row_dimensions[r+2].height = CELL_H

    # Légende
    leg_col = cols + 2
    ws.cell(2, leg_col, "Légende").font = Font(bold=True)
    for i, (label, color) in enumerate([
        ("Culturel", ORANGE), ("Producteur", GREEN), ("Neutre", GREY), ("Bord", 'FF404040')
    ], 1):
        c = ws.cell(2+i, leg_col, label)
        c.fill = make_fill(color)
        c.border = make_border()

# ─────────────────────────────────────────────
# INTERFACE STREAMLIT
# ─────────────────────────────────────────────

uploaded = st.file_uploader("📂 Charger le fichier Excel (Ville.xlsx)", type=['xlsx'])

if uploaded:
    with st.spinner("Lecture du fichier…"):
        grid, buildings, actuel = load_excel(uploaded)
        buildings_by_name = {b['nom']: b for b in buildings}
        original_placed = parse_actuel(actuel, buildings)

    st.success(f"✅ {len(original_placed)} bâtiments détectés dans le placement actuel")

    # Affiche le résumé du placement original
    with st.expander("📊 Productions actuelles"):
        culture_orig = compute_culture(original_placed, buildings_by_name)
        totals_old = compute_total_production(original_placed, buildings_by_name, culture_orig)
        if totals_old:
            df_orig = pd.DataFrame([
                {'Production': k, 'Quantité/h': round(v, 1)}
                for k, v in sorted(totals_old.items())
            ])
            st.dataframe(df_orig, use_container_width=True)
        else:
            st.info("Aucune production calculée pour le placement actuel.")

    n_passes = st.slider("Nombre de passes d'optimisation", 1, 5, 2)

    if st.button("🚀 Lancer l'optimisation"):
        with st.spinner("Optimisation en cours (échange de paires)…"):
            optimized = optimize(grid, original_placed, buildings_by_name, n_passes)
            culture_new = compute_culture(optimized, buildings_by_name)
            totals_new = compute_total_production(optimized, buildings_by_name, culture_new)

        st.success("✅ Optimisation terminée!")

        # Résumé comparatif
        st.subheader("📈 Comparaison des productions")
        all_prods = set(list(totals_new.keys()) + list(totals_old.keys()))
        rows_cmp = []
        for prod in PROD_PRIORITY + sorted(all_prods - set(PROD_PRIORITY)):
            if prod not in all_prods:
                continue
            nv = totals_new.get(prod, 0)
            ov = totals_old.get(prod, 0)
            gain = nv - ov
            pct = (gain / ov * 100) if ov else 0
            rows_cmp.append({
                'Production': prod,
                'Avant': round(ov, 1),
                'Après': round(nv, 1),
                'Gain': round(gain, 1),
                '% variation': f"{pct:+.1f}%"
            })
        st.dataframe(pd.DataFrame(rows_cmp), use_container_width=True)

        # Déplacements
        orig_by_nom = defaultdict(list)
        for p in original_placed:
            orig_by_nom[p['nom']].append(p)
        new_by_nom = defaultdict(list)
        for p in optimized:
            new_by_nom[p['nom']].append(p)
        n_moved = sum(
            1 for nom in orig_by_nom
            for oi, ni in zip(
                sorted(orig_by_nom[nom], key=lambda x: (x['row'], x['col'])),
                sorted(new_by_nom.get(nom, []), key=lambda x: (x['row'], x['col']))
            )
            if oi['row'] != ni['row'] or oi['col'] != ni['col']
        )
        st.info(f"🔄 {n_moved} bâtiment(s) déplacé(s)")

        with st.spinner("Génération du fichier Excel de résultats…"):
            buf = write_result(grid, optimized, original_placed, buildings_by_name)

        st.download_button(
            label="⬇️ Télécharger le fichier de résultats",
            data=buf,
            file_name="Resultat_optimise.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
else:
    st.info("👆 Charge ton fichier Excel pour commencer.")
    st.markdown("""
**Structure attendue du fichier Excel :**
- Onglet `Terrain` : grille de `X` (bords) et `1` (cases libres)
- Onglet `Batiments` : liste des bâtiments avec leurs caractéristiques
- Onglet `Actuel` : placement actuel des bâtiments
""")
