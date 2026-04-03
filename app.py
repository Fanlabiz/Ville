"""
Optimiseur de placement de batiments - Application Streamlit
Compatible iPad Excel (francais) - Deploiement GitHub/Streamlit Cloud
"""

import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import copy

st.set_page_config(page_title="Optimiseur de Ville", layout="wide")
st.title("Optimiseur de placement de batiments")

# ══════════════════════════════════════════════════════
# CONSTANTES COULEURS (format ARGB 8 caracteres)
# ══════════════════════════════════════════════════════
C_ORANGE  = "FFFFA500"
C_GREEN   = "FF90EE90"
C_GRAY    = "FFD3D3D3"
C_BLUE    = "FF4472C4"
C_WHITE   = "FFFFFFFF"
C_BORDX   = "FF808080"
C_GAIN    = "FF006400"
C_LOSS    = "FFCC0000"

def mfill(hex8):
    return PatternFill("solid", fgColor=hex8)

def thin_border():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def style_header(cell, text):
    cell.value = text
    cell.font = Font(bold=True, color=C_WHITE)
    cell.fill = mfill(C_BLUE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = thin_border()

# ══════════════════════════════════════════════════════
# LECTURE DES DONNEES
# ══════════════════════════════════════════════════════

def read_terrain(ws):
    """Retourne grid[r][c] = 'X'|None, max_r, max_c (0-indexes)."""
    max_r, max_c = ws.max_row, ws.max_column
    grid = [[None] * max_c for _ in range(max_r)]
    for row in ws.iter_rows(min_row=1, max_row=max_r, max_col=max_c):
        for cell in row:
            if cell.value == "X":
                grid[cell.row - 1][cell.column - 1] = "X"
    return grid, max_r, max_c


def read_buildings_def(ws):
    """Lit l'onglet Batiments. Retourne liste de dicts."""
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(c).strip() if c else "" for c in rows[0]]
    result = []
    for row in rows[1:]:
        if not any(row):
            continue
        d = dict(zip(header, row))
        result.append({
            "nom":        str(d.get("Nom", "")).strip(),
            "longueur":   int(d.get("Longueur") or 1),
            "largeur":    int(d.get("Largeur") or 1),
            "nombre":     int(d.get("Nombre") or 1),
            "type":       str(d.get("Type", "")).strip(),
            "culture":    float(d.get("Culture") or 0),
            "rayonnement":int(d.get("Rayonnement") or 0),
            "boost25":    float(d.get("Boost 25%") or 0),
            "boost50":    float(d.get("Boost 50%") or 0),
            "boost100":   float(d.get("Boost 100%") or 0),
            "production": str(d.get("Production", "")).strip(),
            "quantite":   float(d.get("Quantite") or 0),
            "priorite":   float(d.get("Priorite") or 99),
        })
    return result


def read_placed_buildings(ws):
    """
    Lit les batiments places sur le terrain :
    - Cellules fusionnees  -> batiments multi-cases
    - Cellules simples non-X -> batiments 1x1
    Retourne liste de dicts {nom, r, c, rows, cols}.
    """
    placed = []
    merged_set = set()

    for mr in ws.merged_cells.ranges:
        top = ws.cell(mr.min_row, mr.min_col)
        name = str(top.value).strip() if top.value else ""
        if name and name != "X":
            placed.append({
                "nom":  name,
                "r":    mr.min_row - 1,
                "c":    mr.min_col - 1,
                "rows": mr.max_row - mr.min_row + 1,
                "cols": mr.max_col - mr.min_col + 1,
            })
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                merged_set.add((r, c))

    for row in ws.iter_rows():
        for cell in row:
            if (cell.row, cell.column) in merged_set:
                continue
            if cell.value and cell.value != "X":
                placed.append({
                    "nom":  str(cell.value).strip(),
                    "r":    cell.row - 1,
                    "c":    cell.column - 1,
                    "rows": 1,
                    "cols": 1,
                })
    return placed


def enrich(placed, buildings_def):
    """Ajoute les infos catalogue a chaque batiment place."""
    catalog = {b["nom"].strip(): b for b in buildings_def}
    result = []
    for p in placed:
        base = catalog.get(p["nom"].strip(), {
            "type": "Neutre", "culture": 0, "rayonnement": 0,
            "boost25": 0, "boost50": 0, "boost100": 0,
            "production": "Rien", "quantite": 0, "priorite": 99,
            "longueur": p["cols"], "largeur": p["rows"], "nombre": 1,
        })
        result.append({**base, **p})
    return result


# ══════════════════════════════════════════════════════
# MECANIQUE CULTURE / BOOST / SCORE
# ══════════════════════════════════════════════════════

def cells_of(b):
    """Cases occupees par le batiment b."""
    return {(b["r"] + dr, b["c"] + dc)
            for dr in range(b["rows"]) for dc in range(b["cols"])}


def radiation_zone(b):
    """Cases dans la zone de rayonnement du batiment culturel b."""
    ray = b.get("rayonnement", 0)
    r0, c0 = b["r"], b["c"]
    r1, c1 = r0 + b["rows"] - 1, c0 + b["cols"] - 1
    return {
        (r, c)
        for r in range(r0 - ray, r1 + ray + 1)
        for c in range(c0 - ray, c1 + ray + 1)
        if not (r0 <= r <= r1 and c0 <= c <= c1)
    }


def culture_received(producer, culturels):
    """Culture totale recue par un batiment producteur."""
    prod_cells = cells_of(producer)
    return sum(cult["culture"] for cult in culturels
               if prod_cells & radiation_zone(cult))


def boost_level(culture, b):
    """Boost obtenu (0, 25, 50 ou 100)."""
    if b["type"] != "Producteur":
        return 0
    if b["boost100"] and culture >= b["boost100"]:
        return 100
    if b["boost50"] and culture >= b["boost50"]:
        return 50
    if b["boost25"] and culture >= b["boost25"]:
        return 25
    return 0


def score_placement(placed):
    """Score total = somme(boost/priorite) pour les Producteurs."""
    culturels = [b for b in placed if b["type"] == "Culturel"]
    total = 0.0
    for b in placed:
        if b["type"] == "Producteur" and b["priorite"] > 0:
            cult = culture_received(b, culturels)
            boost = boost_level(cult, b)
            total += boost / b["priorite"]
    return total


# ══════════════════════════════════════════════════════
# OPTIMISEUR
# ══════════════════════════════════════════════════════

def make_x_grid(terrain_grid, max_r, max_c):
    return [[terrain_grid[r][c] == "X" for c in range(max_c)] for r in range(max_r)]


def make_occ_grid(placed, max_r, max_c, exclude_ids=None):
    occ = [[False] * max_c for _ in range(max_r)]
    excl = exclude_ids or set()
    for b in placed:
        if id(b) in excl:
            continue
        for r, c in cells_of(b):
            if 0 <= r < max_r and 0 <= c < max_c:
                occ[r][c] = True
    return occ


def can_place(r, c, rows, cols, x_grid, occ, max_r, max_c):
    if r < 0 or c < 0 or r + rows > max_r or c + cols > max_c:
        return False
    return not any(
        x_grid[r + dr][c + dc] or occ[r + dr][c + dc]
        for dr in range(rows) for dc in range(cols)
    )


def optimize(placed, terrain_grid, max_r, max_c, n_passes=2, progress_cb=None):
    """
    Optimise le placement par relocalisations successives.
    Retourne (placed_optimise, liste_des_deplacements).
    """
    x_grid = make_x_grid(terrain_grid, max_r, max_c)
    placed = [dict(b) for b in placed]
    all_moves = []

    total_iters = n_passes * len(placed) * 2
    iter_count = [0]

    for _ in range(n_passes):
        improved = True
        while improved:
            improved = False
            for b in placed:
                orig = (b["r"], b["c"], b["rows"], b["cols"])
                occ = make_occ_grid(placed, max_r, max_c, exclude_ids={id(b)})
                orig_score = score_placement(placed)
                best_s, best_pos = orig_score, None

                # Tester les deux orientations
                for rows, cols in {(b["rows"], b["cols"]), (b["cols"], b["rows"])}:
                    for r in range(max_r):
                        for c in range(max_c):
                            if (r, c, rows, cols) == orig:
                                continue
                            if can_place(r, c, rows, cols, x_grid, occ, max_r, max_c):
                                b["r"], b["c"], b["rows"], b["cols"] = r, c, rows, cols
                                s = score_placement(placed)
                                if s > best_s:
                                    best_s, best_pos = s, (r, c, rows, cols)
                                b["r"], b["c"], b["rows"], b["cols"] = orig

                if best_pos:
                    all_moves.append({
                        "nom":   b["nom"],
                        "old_r": b["r"], "old_c": b["c"],
                        "old_rows": b["rows"], "old_cols": b["cols"],
                        "new_r": best_pos[0], "new_c": best_pos[1],
                        "new_rows": best_pos[2], "new_cols": best_pos[3],
                    })
                    b["r"], b["c"], b["rows"], b["cols"] = best_pos
                    improved = True

                iter_count[0] += 1
                if progress_cb:
                    progress_cb(min(iter_count[0] / max(total_iters, 1), 0.98))

    return placed, all_moves


# ══════════════════════════════════════════════════════
# GENERATION DU FICHIER EXCEL DE SORTIE
# ══════════════════════════════════════════════════════

def build_excel_output(optimized, original_placed, terrain_grid, max_r, max_c, buildings_def):
    wb = openpyxl.Workbook()
    culturels = [b for b in optimized if b["type"] == "Culturel"]
    orig_culturels = [b for b in original_placed if b["type"] == "Culturel"]

    # ─────────────────────────────────────
    # ONGLET 1 : Liste batiments
    # ─────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Liste batiments"
    headers = ["Nom", "Type", "Production", "Coord (L,C)", "Orientation",
               "Culture recue", "Boost atteint", "Qte/h avec boost", "Score boost"]
    widths  = [30, 12, 22, 12, 12, 14, 13, 18, 12]
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        style_header(ws1.cell(1, ci), h)
        ws1.column_dimensions[get_column_letter(ci)].width = w

    row_i = 2
    for b in sorted(optimized, key=lambda x: (x["type"], x["nom"])):
        cult = culture_received(b, culturels) if b["type"] == "Producteur" else 0
        boost = boost_level(cult, b)
        prio = b["priorite"] if b["priorite"] > 0 else 99
        score = boost / prio if b["type"] == "Producteur" else ""
        qte_boost = b["quantite"] * (1 + boost / 100) if b["type"] == "Producteur" else ""
        orient = "H" if b["cols"] >= b["rows"] else "V"
        fill = mfill(C_ORANGE if b["type"] == "Culturel" else C_GREEN if b["type"] == "Producteur" else C_GRAY)
        vals = [b["nom"], b["type"], b["production"],
                f"L{b['r']+1} C{b['c']+1}", orient,
                round(cult, 1), f"{boost}%",
                round(qte_boost, 1) if qte_boost != "" else "",
                round(score, 3) if score != "" else ""]
        for ci, v in enumerate(vals, 1):
            cell = ws1.cell(row_i, ci, v)
            cell.fill = fill
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center", vertical="center")
        row_i += 1

    # ─────────────────────────────────────
    # ONGLET 2 : Synthese
    # ─────────────────────────────────────
    ws2 = wb.create_sheet("Synthese")
    hdrs2 = ["Production", "Culture totale recue", "Boost max atteint",
              "Qte/h initiale", "Qte/h optimisee", "Gain/perte Qte/h"]
    widths2 = [22, 22, 18, 16, 16, 18]
    for ci, (h, w) in enumerate(zip(hdrs2, widths2), 1):
        style_header(ws2.cell(1, ci), h)
        ws2.column_dimensions[get_column_letter(ci)].width = w

    # Aggreger par type de production
    prod_data = {}
    for b in optimized:
        if b["type"] != "Producteur" or b["production"] == "Rien":
            continue
        p = b["production"]
        cult = culture_received(b, culturels)
        boost = boost_level(cult, b)
        if p not in prod_data:
            prod_data[p] = {"cult": 0.0, "boost": 0, "qte_new": 0.0, "qte_orig": 0.0}
        prod_data[p]["cult"] += cult
        prod_data[p]["boost"] = max(prod_data[p]["boost"], boost)
        prod_data[p]["qte_new"] += b["quantite"] * (1 + boost / 100)

    for b in original_placed:
        if b["type"] != "Producteur" or b["production"] == "Rien":
            continue
        p = b["production"]
        cult = culture_received(b, orig_culturels)
        boost = boost_level(cult, b)
        if p in prod_data:
            prod_data[p]["qte_orig"] += b["quantite"] * (1 + boost / 100)

    row_i = 2
    for prod, data in sorted(prod_data.items()):
        gain = data["qte_new"] - data["qte_orig"]
        vals = [prod, round(data["cult"], 1), f"{data['boost']}%",
                round(data["qte_orig"], 1), round(data["qte_new"], 1), round(gain, 1)]
        for ci, v in enumerate(vals, 1):
            cell = ws2.cell(row_i, ci, v)
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center")
            if ci == 6:
                cell.font = Font(bold=True, color=C_GAIN if gain >= 0 else C_LOSS)
        row_i += 1

    # Score global
    row_i += 1
    ws2.cell(row_i, 1, "Score boost global").font = Font(bold=True)
    ws2.cell(row_i, 2, round(score_placement(optimized), 4)).font = Font(bold=True)
    ws2.cell(row_i, 3, "(Score initial)").font = Font(italic=True, color="FF888888")
    ws2.cell(row_i, 4, round(score_placement(original_placed), 4)).font = Font(italic=True, color="FF888888")

    # ─────────────────────────────────────
    # ONGLET 3 : Deplacements
    # ─────────────────────────────────────
    ws3 = wb.create_sheet("Deplacements")
    hdrs3 = ["#", "Batiment", "Position initiale", "Position finale", "Sequence d'operations"]
    widths3 = [4, 30, 16, 16, 70]
    for ci, (h, w) in enumerate(zip(hdrs3, widths3), 1):
        style_header(ws3.cell(1, ci), h)
        ws3.column_dimensions[get_column_letter(ci)].width = w

    # Identifier les vrais deplacements (position finale != initiale)
    orig_pos = {}
    for b in original_placed:
        orig_pos.setdefault(b["nom"], []).append((b["r"], b["c"], b["rows"], b["cols"]))

    real_moves = []
    used_orig = {nom: 0 for nom in orig_pos}
    for b in optimized:
        nom = b["nom"]
        if nom not in orig_pos:
            continue
        idx = used_orig.get(nom, 0)
        if idx < len(orig_pos[nom]):
            op = orig_pos[nom][idx]
            used_orig[nom] = idx + 1
            if op[0] != b["r"] or op[1] != b["c"]:
                real_moves.append({
                    "nom": nom,
                    "old_r": op[0], "old_c": op[1],
                    "new_r": b["r"], "new_c": b["c"],
                })

    if not real_moves:
        ws3.cell(2, 1, "Aucun deplacement effectue - placement deja optimal.")
    else:
        for step, mv in enumerate(real_moves, 1):
            old_str = f"L{mv['old_r']+1} C{mv['old_c']+1}"
            new_str = f"L{mv['new_r']+1} C{mv['new_c']+1}"

            # Verifier si la destination etait occupee initialement
            blockers = []
            for b2 in original_placed:
                if b2["nom"] == mv["nom"]:
                    continue
                for dr in range(b2["rows"]):
                    for dc in range(b2["cols"]):
                        if b2["r"]+dr == mv["new_r"] and b2["c"]+dc == mv["new_c"]:
                            if b2["nom"] not in blockers:
                                blockers.append(b2["nom"])

            if blockers:
                action = (
                    f"1) Sortir provisoirement du terrain : {', '.join(blockers)}. "
                    f"2) Deplacer '{mv['nom']}' de {old_str} vers {new_str}. "
                    f"3) Remettre {', '.join(blockers)} a leur position finale."
                )
            else:
                action = f"Deplacer '{mv['nom']}' de {old_str} vers {new_str}."

            ri = step + 1
            ws3.cell(ri, 1, step)
            ws3.cell(ri, 2, mv["nom"])
            ws3.cell(ri, 3, old_str)
            ws3.cell(ri, 4, new_str)
            ws3.cell(ri, 5, action)
            ws3.cell(ri, 5).alignment = Alignment(wrap_text=True, vertical="top")
            ws3.row_dimensions[ri].height = 48
            for ci in range(1, 6):
                ws3.cell(ri, ci).border = thin_border()
                if ci < 5:
                    ws3.cell(ri, ci).alignment = Alignment(horizontal="center", vertical="top")

    # ─────────────────────────────────────
    # ONGLET 4 : Terrain optimise (carte)
    # ─────────────────────────────────────
    ws4 = wb.create_sheet("Terrain optimise")

    # Construire la grille des batiments places
    placed_grid = {}
    for b in optimized:
        for dr in range(b["rows"]):
            for dc in range(b["cols"]):
                placed_grid[(b["r"]+dr, b["c"]+dc)] = b

    col_w = 14
    row_h = 20
    for r in range(max_r):
        ws4.row_dimensions[r+1].height = row_h
    for c in range(max_c):
        ws4.column_dimensions[get_column_letter(c+1)].width = col_w

    for r in range(max_r):
        for c in range(max_c):
            cell = ws4.cell(r+1, c+1)
            if terrain_grid[r][c] == "X":
                cell.value = "X"
                cell.fill = mfill(C_BORDX)
                cell.font = Font(bold=True, color=C_WHITE)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif (r, c) in placed_grid:
                b = placed_grid[(r, c)]
                fill_hex = C_ORANGE if b["type"] == "Culturel" else C_GREEN if b["type"] == "Producteur" else C_GRAY
                cell.fill = mfill(fill_hex)
                cell.border = thin_border()
                # Ecrire le nom uniquement dans la cellule en haut a gauche
                if b["r"] == r and b["c"] == c:
                    cult = culture_received(b, culturels) if b["type"] == "Producteur" else 0
                    boost = boost_level(cult, b)
                    label = b["nom"]
                    if b["type"] == "Producteur" and boost > 0:
                        label += f"\n+{boost}%"
                    cell.value = label
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    cell.font = Font(bold=(boost > 0))
                    # Fusionner les cases du batiment
                    if b["rows"] > 1 or b["cols"] > 1:
                        try:
                            ws4.merge_cells(
                                start_row=r+1, start_column=c+1,
                                end_row=r+b["rows"], end_column=c+b["cols"]
                            )
                        except Exception:
                            pass

    # Legende
    leg_r = max_r + 2
    ws4.cell(leg_r, 1, "Legende").font = Font(bold=True)
    for i, (label, color) in enumerate([("Culturel", C_ORANGE), ("Producteur", C_GREEN), ("Neutre", C_GRAY)], 1):
        cell = ws4.cell(leg_r+i, 1, label)
        cell.fill = mfill(color)
        cell.border = thin_border()
        cell.alignment = Alignment(horizontal="center")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════
# INTERFACE STREAMLIT
# ══════════════════════════════════════════════════════

uploaded = st.file_uploader(
    "Choisissez votre fichier Excel de ville (.xlsx)",
    type=["xlsx"],
    help="Le fichier doit contenir un onglet Terrain et un onglet Batiments."
)

if uploaded:
    try:
        wb_in = openpyxl.load_workbook(uploaded)
    except Exception as e:
        st.error(f"Impossible de lire le fichier : {e}")
        st.stop()

    sheet_names = wb_in.sheetnames
    if len(sheet_names) < 2:
        st.error("Le fichier doit contenir au moins 2 onglets (Terrain + Batiments).")
        st.stop()

    terrain_ws = wb_in[sheet_names[0]]
    bat_ws     = wb_in[sheet_names[1]]

    terrain_grid, max_r, max_c = read_terrain(terrain_ws)
    buildings_def = read_buildings_def(bat_ws)
    placed = enrich(read_placed_buildings(terrain_ws), buildings_def)
    original_placed = [dict(b) for b in placed]

    n_culturels  = sum(1 for b in placed if b["type"] == "Culturel")
    n_producteurs = sum(1 for b in placed if b["type"] == "Producteur")
    n_neutres    = sum(1 for b in placed if b["type"] == "Neutre")
    score_init   = score_placement(placed)

    st.success(f"Fichier charge avec succes : **{len(placed)} batiments** sur un terrain **{max_r} x {max_c}**")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Score initial", f"{score_init:.2f}")
    col2.metric("Batiments culturels", n_culturels)
    col3.metric("Batiments producteurs", n_producteurs)
    col4.metric("Batiments neutres", n_neutres)

    # Detail culture initiale
    with st.expander("Detail de la culture initiale par producteur"):
        culturels = [b for b in placed if b["type"] == "Culturel"]
        for b in sorted(placed, key=lambda x: x["nom"]):
            if b["type"] == "Producteur":
                cult = culture_received(b, culturels)
                boost = boost_level(cult, b)
                seuil = max(b["boost25"], b["boost50"], b["boost100"])
                st.write(
                    f"**{b['nom']}** - Culture recue : {cult:.0f} / "
                    f"Seuil 25% : {b['boost25']:.0f} | 50% : {b['boost50']:.0f} | 100% : {b['boost100']:.0f} "
                    f"→ **Boost : {boost}%**"
                )

    st.divider()
    n_passes = st.slider(
        "Nombre de passes d'optimisation",
        min_value=1, max_value=5, value=2,
        help="Plus de passes = meilleur resultat mais plus long. 2 passes est un bon compromis."
    )

    if st.button("Lancer l'optimisation", type="primary"):
        progress_bar = st.progress(0)
        status = st.empty()
        status.info("Optimisation en cours... Veuillez patienter.")

        def update_prog(v):
            progress_bar.progress(v)

        optimized, moves = optimize(
            placed, terrain_grid, max_r, max_c,
            n_passes=n_passes, progress_cb=update_prog
        )
        progress_bar.progress(1.0)

        score_opt = score_placement(optimized)
        delta = score_opt - score_init
        status.success("Optimisation terminee !")

        c1, c2, c3 = st.columns(3)
        c1.metric("Score initial", f"{score_init:.2f}")
        c2.metric("Score optimise", f"{score_opt:.2f}", delta=f"{delta:+.2f}")

        # Compter les deplacements reels
        moved = []
        orig_map = {}
        for b in original_placed:
            orig_map.setdefault(b["nom"], []).append((b["r"], b["c"]))
        used = {n: 0 for n in orig_map}
        for b in optimized:
            nom = b["nom"]
            if nom in orig_map:
                idx = used[nom]
                if idx < len(orig_map[nom]):
                    used[nom] += 1
                    op = orig_map[nom][idx]
                    if op[0] != b["r"] or op[1] != b["c"]:
                        moved.append(b)

        c3.metric("Batiments deplaces", len(moved))

        if moved:
            st.subheader("Batiments deplaces")
            used2 = {n: 0 for n in orig_map}
            for b in optimized:
                nom = b["nom"]
                if nom in orig_map:
                    idx = used2[nom]
                    if idx < len(orig_map[nom]):
                        used2[nom] += 1
                        op = orig_map[nom][idx]
                        if op[0] != b["r"] or op[1] != b["c"]:
                            cult = culture_received(b, [x for x in optimized if x["type"]=="Culturel"])
                            boost = boost_level(cult, b)
                            icon = "🟠" if b["type"]=="Culturel" else "🟢" if b["type"]=="Producteur" else "⬜"
                            st.write(
                                f"{icon} **{b['nom']}** : "
                                f"L{op[0]+1} C{op[1]+1} → L{b['r']+1} C{b['c']+1}"
                                + (f" | Boost apres : **{boost}%**" if b["type"]=="Producteur" else "")
                            )
        else:
            st.info("Le placement initial est deja optimal. Aucun deplacement ameliorant le score n'a ete trouve.")

        # Generer le fichier Excel
        st.divider()
        with st.spinner("Generation du fichier Excel..."):
            output_buf = build_excel_output(
                optimized, original_placed, terrain_grid, max_r, max_c, buildings_def
            )

        st.download_button(
            label="Telecharger le fichier resultat Excel",
            data=output_buf,
            file_name="ville_optimisee.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption(
            "Le fichier contient 4 onglets : "
            "**Liste batiments** (detail de chaque batiment), "
            "**Synthese** (production et gains par type), "
            "**Deplacements** (sequence d'operations), "
            "**Terrain optimise** (carte coloree)."
        )
