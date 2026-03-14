import streamlit as st
import pandas as pd
import random
import math
import copy
import time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="Détecteur & Optimiseur de bâtiments", page_icon="🏙️", layout="wide")
st.title("🏙️ Détecteur & Optimiseur de bâtiments")
st.markdown("Charge un fichier Excel avec 3 onglets : **Terrain**, **Batiments**, **Actuel**.")

uploaded_file = st.file_uploader("📂 Choisir le fichier Excel d'entrée", type=["xlsx"])
if not uploaded_file:
    st.stop()

# ─── Chargement ──────────────────────────────────────────────────────────────

@st.cache_data
def load_data(file_bytes):
    terrain_df   = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Terrain",   header=None)
    batiments_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Batiments", header=0)
    actuel_df    = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Actuel",    header=None)
    return terrain_df, batiments_df, actuel_df

file_bytes = uploaded_file.read()
terrain_df, batiments_df, actuel_df = load_data(file_bytes)
batiments_df.columns = [str(c).strip() for c in batiments_df.columns]

# ─── Helpers ─────────────────────────────────────────────────────────────────

def safe_int(v, d=0):
    try:
        f = float(v)
        return int(f) if not pd.isna(f) else d
    except (TypeError, ValueError):
        return d

def safe_float(v):
    try:
        f = float(v)
        return f if not pd.isna(f) else None
    except (TypeError, ValueError):
        return None

def df_to_grid(df):
    grid = []
    for _, row in df.iterrows():
        line = ['' if pd.isna(v) else str(v).strip() for v in row]
        grid.append(line)
    mc = max(len(r) for r in grid)
    for r in grid:
        while len(r) < mc:
            r.append('')
    return grid

# ─── Catalogue ───────────────────────────────────────────────────────────────

def build_catalog(df):
    catalog = {}
    for _, row in df.iterrows():
        nom = str(row.get("Nom", "")).strip()
        if not nom or nom == "nan":
            continue
        lon = safe_int(row.get("Longueur", 0))
        lar = safe_int(row.get("Largeur",  0))
        if lon == 0 or lar == 0:
            continue
        catalog[nom] = {
            "longueur":    lon,
            "largeur":     lar,
            "quantite":    safe_int(row.get("Quantite", 0)),
            "type":        str(row.get("Type", "")).strip(),
            "culture":     safe_int(row.get("Culture",  0)),
            "rayonnement": safe_int(row.get("Rayonnement", 0)),
            "boost25":     safe_float(row.get("Boost 25%")),
            "boost50":     safe_float(row.get("Boost 50%")),
            "boost100":    safe_float(row.get("Boost 100%")),
            "production":  str(row.get("Production", "")).strip(),
        }
    return catalog

catalog = build_catalog(batiments_df)

# ─── Terrain libre ────────────────────────────────────────────────────────────
# Extraire les cases intérieures libres (ni X ni bord) depuis l'onglet Terrain

def parse_terrain(df):
    """
    Retourne (ROWS, COLS, free_cells) où free_cells est un set de (r,c) 0-based
    correspondant aux cases intérieures libres (valeur 1).
    Les bords (X) définissent la frontière.
    """
    grid = df_to_grid(df)
    rows = len(grid)
    cols = max(len(r) for r in grid)
    free = set()
    for r in range(rows):
        for c in range(len(grid[r])):
            val = grid[r][c]
            if val == '1':
                free.add((r, c))
    return rows, cols, free

T_ROWS, T_COLS, terrain_free = parse_terrain(terrain_df)

# ─── Détection depuis l'onglet Actuel ────────────────────────────────────────

actuel_grid = df_to_grid(actuel_df)
A_ROWS = len(actuel_grid)
A_COLS = len(actuel_grid[0]) if A_ROWS else 0

def flood_fill(grid, rows, cols, r0, c0, name):
    cells, stack = set(), [(r0, c0)]
    while stack:
        r, c = stack.pop()
        if (r, c) in cells or not (0 <= r < rows and 0 <= c < cols):
            continue
        if grid[r][c] != name:
            continue
        cells.add((r, c))
        for dr, dc in [(-1,0),(1,0),(0,-1),(0,1)]:
            stack.append((r+dr, c+dc))
    return cells

def tile_box(cells, h, w):
    if len(cells) % (h * w) != 0:
        return None
    remaining = set(cells)
    placements = []
    while remaining:
        r0, c0 = min(remaining)
        tile = {(r0+dr, c0+dc) for dr in range(h) for dc in range(w)}
        if not tile.issubset(remaining):
            return None
        placements.append((r0, c0))
        remaining -= tile
    return placements

def detect_buildings(grid, rows, cols, catalog):
    visited = [[False]*cols for _ in range(rows)]
    found, warnings = [], []
    for r in range(rows):
        for c in range(cols):
            cell = grid[r][c]
            if not cell or cell in ("X","0","1","") or visited[r][c]:
                continue
            if cell not in catalog:
                continue
            blob = flood_fill(grid, rows, cols, r, c, cell)
            for br, bc in blob:
                visited[br][bc] = True
            bat = catalog[cell]
            lon, lar = bat["longueur"], bat["largeur"]
            placed = False
            for (h, w) in [(lar, lon), (lon, lar)]:
                pl = tile_box(blob, h, w)
                if pl is not None:
                    orient = "H" if w == lon else "V"
                    for (r0, c0) in pl:
                        tile_cells = {(r0+dr, c0+dc) for dr in range(h) for dc in range(w)}
                        found.append({
                            "nom":         cell,
                            "ligne":       r0 + 1,
                            "colonne":     c0 + 1,
                            "hauteur":     h,
                            "largeur":     w,
                            "orientation": orient,
                            "cases":       tile_cells,
                        })
                    placed = True
                    break
            if not placed:
                min_r = min(br for br, bc in blob)
                min_c = min(bc for br, bc in blob)
                max_r = max(br for br, bc in blob)
                max_c = max(bc for br, bc in blob)
                warnings.append(
                    f"⚠️ **{cell}** — bloc {len(blob)} cases "
                    f"[r{min_r+1}-{max_r+1}, c{min_c+1}-{max_c+1}] "
                    f"impossible à paver ({lar}×{lon})"
                )
    return found, warnings

# ─── Calcul culture & boost ───────────────────────────────────────────────────

def compute_culture_boost(found, catalog):
    """Calcule culture_recue et boost pour chaque Producteur. Modifie found in-place."""
    culturels   = [b for b in found if catalog.get(b["nom"],{}).get("type") == "Culturel"]
    producteurs = [b for b in found if catalog.get(b["nom"],{}).get("type") == "Producteur"]

    def ray_zone(bat, R):
        zone = set()
        r_min = bat["ligne"] - 1
        c_min = bat["colonne"] - 1
        r_max = r_min + bat["hauteur"] - 1
        c_max = c_min + bat["largeur"] - 1
        for r in range(r_min - R, r_max + R + 1):
            for c in range(c_min - R, c_max + R + 1):
                zone.add((r, c))
        return zone

    # Zones de rayonnement des Culturels
    cz_list = []
    for cult in culturels:
        bi = catalog[cult["nom"]]
        R  = bi["rayonnement"]
        cu = bi["culture"]
        if R > 0 and cu > 0:
            cz_list.append({"zone": ray_zone(cult, R), "culture": cu})

    for bat in found:
        bat["culture_recue"] = 0
        bat["boost"] = ""

    for prod in producteurs:
        total = sum(cz["culture"] for cz in cz_list if prod["cases"] & cz["zone"])
        prod["culture_recue"] = total
        bi = catalog[prod["nom"]]
        b25  = bi.get("boost25")
        b50  = bi.get("boost50")
        b100 = bi.get("boost100")
        if   b100 is not None and total >= b100: boost = "100%"
        elif b50  is not None and total >= b50:  boost = "50%"
        elif b25  is not None and total >= b25:  boost = "25%"
        elif b25  is not None:                   boost = "0%"
        else:                                    boost = ""
        prod["boost"] = boost

    return found

def boost_score(boost_str):
    """Convertit un boost en score numérique pour l'optimisation."""
    return {"100%": 4, "50%": 3, "25%": 2, "0%": 1, "": 0}.get(boost_str, 0)

def mean_boost_score(found, catalog):
    prods = [b for b in found if catalog.get(b["nom"],{}).get("type") == "Producteur"]
    if not prods:
        return 0.0
    return sum(boost_score(p["boost"]) for p in prods) / len(prods)

# ─── Optimisation par recuit simulé ──────────────────────────────────────────
#
# Représentation : liste de bâtiments avec (nom, r0, c0, h, w)
# À chaque itération, on choisit aléatoirement un mouvement parmi :
#   1. Déplacer un bâtiment vers une case libre aléatoire
#   2. Échanger les positions de deux bâtiments
#   3. Retourner l'orientation d'un bâtiment (si lon ≠ lar)
#
# Score = boost moyen des Producteurs (plus haut = meilleur)

def placement_to_found(placement):
    """Convertit une liste de dicts placement en liste found (avec cases)."""
    found = []
    for p in placement:
        r0, c0, h, w = p["r0"], p["c0"], p["h"], p["w"]
        tile_cells = {(r0+dr, c0+dc) for dr in range(h) for dc in range(w)}
        found.append({
            "nom":         p["nom"],
            "ligne":       r0 + 1,
            "colonne":     c0 + 1,
            "hauteur":     h,
            "largeur":     w,
            "orientation": "H" if w == catalog[p["nom"]]["longueur"] else "V",
            "cases":       tile_cells,
            "culture_recue": 0,
            "boost":       "",
        })
    return found

def found_to_placement(found):
    return [{
        "nom": b["nom"],
        "r0":  b["ligne"] - 1,
        "c0":  b["colonne"] - 1,
        "h":   b["hauteur"],
        "w":   b["largeur"],
    } for b in found]

def occupied_cells(placement):
    occ = set()
    for p in placement:
        for dr in range(p["h"]):
            for dc in range(p["w"]):
                occ.add((p["r0"]+dr, p["c0"]+dc))
    return occ

def is_valid_placement(p, free_cells, occ_without_p):
    """Vérifie que le bâtiment p tient dans les cases libres sans chevauchement."""
    for dr in range(p["h"]):
        for dc in range(p["w"]):
            rc = (p["r0"]+dr, p["c0"]+dc)
            if rc not in free_cells or rc in occ_without_p:
                return False
    return True

def score_placement(placement, catalog):
    found = placement_to_found(placement)
    found = compute_culture_boost(found, catalog)
    return mean_boost_score(found, catalog), found

def simulated_annealing(initial_placement, catalog, free_cells,
                        T_start=2.0, T_end=0.05, max_iter=8000,
                        progress_cb=None):
    """
    Recuit simulé sur le placement des bâtiments.
    Retourne (best_placement, best_score, best_found).
    """
    current = copy.deepcopy(initial_placement)
    cur_score, cur_found = score_placement(current, catalog)
    best = copy.deepcopy(current)
    best_score = cur_score
    best_found = copy.deepcopy(cur_found)

    free_list = sorted(free_cells)  # liste pour le tirage aléatoire

    n = len(current)
    T = T_start
    alpha = (T_end / T_start) ** (1.0 / max(max_iter, 1))

    for it in range(max_iter):
        T *= alpha
        if progress_cb and it % 500 == 0:
            progress_cb(it / max_iter, cur_score, best_score)

        # Construire l'occupation actuelle
        occ = occupied_cells(current)

        # Choisir un mouvement aléatoire
        move = random.randint(0, 2)
        idx  = random.randint(0, n - 1)
        p    = current[idx]
        bi   = catalog[p["nom"]]

        new_placement = copy.deepcopy(current)
        np_ = new_placement[idx]

        if move == 0:
            # Déplacer vers une position libre aléatoire
            # Essayer plusieurs candidats
            occ_without = occ - {(p["r0"]+dr, p["c0"]+dc)
                                  for dr in range(p["h"]) for dc in range(p["w"])}
            candidates = random.sample(free_list, min(30, len(free_list)))
            placed = False
            for (nr, nc) in candidates:
                np_["r0"] = nr; np_["c0"] = nc
                if is_valid_placement(np_, free_cells, occ_without):
                    placed = True
                    break
            if not placed:
                continue

        elif move == 1:
            # Échanger deux bâtiments de même taille (ou essayer quand même)
            idx2 = random.randint(0, n - 1)
            if idx2 == idx:
                continue
            p2   = current[idx2]
            np2_ = new_placement[idx2]
            # Essayer de placer p à la position de p2 et vice-versa
            occ_without = (occ
                - {(p["r0"]+dr,  p["c0"]+dc)  for dr in range(p["h"])  for dc in range(p["w"])}
                - {(p2["r0"]+dr, p2["c0"]+dc) for dr in range(p2["h"]) for dc in range(p2["w"])})
            # Placer idx à position de idx2
            np_["r0"], np_["c0"] = p2["r0"], p2["c0"]
            # Adapter la taille si différente — sinon abandonner
            if p["h"] != p2["h"] or p["w"] != p2["w"]:
                # Essayer les orientations possibles
                fits = False
                for (h2, w2) in [(p["h"], p["w"]), (p["w"], p["h"])]:
                    np_["h"] = h2; np_["w"] = w2
                    if is_valid_placement(np_, free_cells, occ_without):
                        fits = True; break
                if not fits:
                    continue
                occ_without2 = occ_without | {(np_["r0"]+dr, np_["c0"]+dc)
                                               for dr in range(np_["h"]) for dc in range(np_["w"])}
                np2_["r0"], np2_["c0"] = p["r0"], p["c0"]
                for (h2, w2) in [(p2["h"], p2["w"]), (p2["w"], p2["h"])]:
                    np2_["h"] = h2; np2_["w"] = w2
                    if is_valid_placement(np2_, free_cells, occ_without2):
                        break
                else:
                    continue
            else:
                np2_["r0"], np2_["c0"] = p["r0"], p["c0"]
                occ_without2 = occ_without | {(np_["r0"]+dr, np_["c0"]+dc)
                                               for dr in range(np_["h"]) for dc in range(np_["w"])}
                if not is_valid_placement(np2_, free_cells, occ_without2):
                    continue

        else:
            # Retourner l'orientation si lon ≠ lar
            if bi["longueur"] == bi["largeur"]:
                continue
            new_h = p["w"]; new_w = p["h"]
            np_["h"] = new_h; np_["w"] = new_w
            occ_without = occ - {(p["r0"]+dr, p["c0"]+dc)
                                  for dr in range(p["h"]) for dc in range(p["w"])}
            if not is_valid_placement(np_, free_cells, occ_without):
                continue

        # Évaluer le nouveau placement
        new_score, new_found = score_placement(new_placement, catalog)
        delta = new_score - cur_score

        if delta > 0 or random.random() < math.exp(delta / max(T, 1e-10)):
            current = new_placement
            cur_score = new_score
            cur_found = new_found
            if cur_score > best_score:
                best = copy.deepcopy(current)
                best_score = cur_score
                best_found = copy.deepcopy(cur_found)

    if progress_cb:
        progress_cb(1.0, cur_score, best_score)

    return best, best_score, best_found

# ─── Résumé boost par production ─────────────────────────────────────────────

def boost_summary(found, catalog):
    """Retourne un dict {production: {0%: n, 25%: n, 50%: n, 100%: n}}."""
    summary = {}
    for bat in found:
        bi = catalog.get(bat["nom"], {})
        if bi.get("type") != "Producteur":
            continue
        prod = bi.get("production", "") or "—"
        boost = bat.get("boost", "") or "0%"
        if prod not in summary:
            summary[prod] = {"0%": 0, "25%": 0, "50%": 0, "100%": 0}
        if boost in summary[prod]:
            summary[prod][boost] += 1
    return summary

# ─── Lancement ───────────────────────────────────────────────────────────────

with st.spinner("Détection des bâtiments…"):
    initial_found, detect_warnings = detect_buildings(actuel_grid, A_ROWS, A_COLS, catalog)
    initial_found = compute_culture_boost(initial_found, catalog)

initial_score = mean_boost_score(initial_found, catalog)
nb_found = len(initial_found)

# ─── Affichage initial ───────────────────────────────────────────────────────

tab1, tab2 = st.tabs(["📊 Situation initiale", "🔧 Optimisation"])

with tab1:
    st.subheader(f"{nb_found} bâtiment(s) détectés  —  Score boost moyen : {initial_score:.3f}")
    if detect_warnings:
        for w in detect_warnings:
            st.markdown(w)

    col1, col2 = st.columns([3, 2])
    with col1:
        st.markdown("### Journal")
        for bat in initial_found:
            bi = catalog.get(bat["nom"], {})
            line = f"✅ **{bat['nom']}** — l{bat['ligne']} c{bat['colonne']} ({bat['hauteur']}×{bat['largeur']}, {bat['orientation']})"
            if bi.get("type") == "Producteur":
                line += f" | Culture : **{bat['culture_recue']}** | Boost : **{bat['boost']}**"
            st.markdown(line)

    with col2:
        st.markdown("### Résumé par production")
        summ = boost_summary(initial_found, catalog)
        if summ:
            rows = []
            for prod, d in sorted(summ.items()):
                rows.append({"Production": prod, "0%": d["0%"], "25%": d["25%"],
                              "50%": d["50%"], "100%": d["100%"]})
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ─── Onglet optimisation ─────────────────────────────────────────────────────

with tab2:
    st.markdown("### Paramètres du recuit simulé")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        max_iter = st.slider("Nombre d'itérations", 2000, 20000, 8000, step=1000)
    with col_b:
        T_start = st.slider("Température initiale", 0.5, 5.0, 2.0, step=0.5)
    with col_c:
        T_end = st.slider("Température finale", 0.01, 0.5, 0.05, step=0.01)

    if st.button("🚀 Lancer l'optimisation", type="primary"):
        # Vérifier que le terrain a bien des cases libres
        if not terrain_free:
            st.error("Aucune case libre trouvée dans l'onglet Terrain.")
            st.stop()

        initial_placement = found_to_placement(initial_found)

        prog_bar  = st.progress(0.0)
        prog_text = st.empty()
        t0 = time.time()

        def progress_cb(frac, cur_sc, best_sc):
            prog_bar.progress(frac)
            elapsed = time.time() - t0
            prog_text.markdown(
                f"Itération {int(frac*max_iter)}/{max_iter} — "
                f"Score courant : **{cur_sc:.3f}** — "
                f"Meilleur : **{best_sc:.3f}** — "
                f"⏱ {elapsed:.1f}s"
            )

        with st.spinner("Optimisation en cours…"):
            best_placement, best_score, best_found = simulated_annealing(
                initial_placement, catalog, terrain_free,
                T_start=T_start, T_end=T_end, max_iter=max_iter,
                progress_cb=progress_cb,
            )

        elapsed = time.time() - t0
        prog_bar.progress(1.0)
        prog_text.markdown(
            f"✅ Terminé en {elapsed:.1f}s — "
            f"Score initial : **{initial_score:.3f}** → "
            f"Score optimisé : **{best_score:.3f}** "
            f"({'+'if best_score>=initial_score else ''}{best_score-initial_score:.3f})"
        )

        st.subheader("Résultat de l'optimisation")
        col3, col4 = st.columns([3, 2])
        with col3:
            st.markdown("### Journal optimisé")
            for bat in best_found:
                bi = catalog.get(bat["nom"], {})
                line = f"✅ **{bat['nom']}** — l{bat['ligne']} c{bat['colonne']} ({bat['hauteur']}×{bat['largeur']}, {bat['orientation']})"
                if bi.get("type") == "Producteur":
                    line += f" | Culture : **{bat['culture_recue']}** | Boost : **{bat['boost']}**"
                st.markdown(line)

        with col4:
            st.markdown("### Résumé par production (optimisé)")
            summ_opt = boost_summary(best_found, catalog)
            if summ_opt:
                rows_opt = []
                for prod, d in sorted(summ_opt.items()):
                    rows_opt.append({"Production": prod, "0%": d["0%"], "25%": d["25%"],
                                     "50%": d["50%"], "100%": d["100%"]})
                st.dataframe(pd.DataFrame(rows_opt), use_container_width=True, hide_index=True)

        # Stocker pour le téléchargement
        st.session_state["best_found"]   = best_found
        st.session_state["best_score"]   = best_score
        st.session_state["initial_found"] = initial_found
        st.session_state["initial_score"] = initial_score

# ─── Génération Excel ─────────────────────────────────────────────────────────

def make_excel(initial_found, catalog, best_found=None, best_score=None, initial_score=None):
    wb = Workbook()

    H_FILL   = PatternFill("solid", start_color="1F4E79")
    H_FONT   = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    D_FONT   = Font(name="Arial", size=10)
    CTR      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LFT      = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    BDR      = Border(left=Side(style="thin"), right=Side(style="thin"),
                      top=Side(style="thin"),  bottom=Side(style="thin"))
    OK_F     = PatternFill("solid", start_color="C6EFCE")
    WN_F     = PatternFill("solid", start_color="FFEB9C")
    ALT_F    = PatternFill("solid", start_color="DEEAF1")
    PROD_F   = PatternFill("solid", start_color="FFF2CC")
    B100_F   = PatternFill("solid", start_color="92D050")
    B50_F    = PatternFill("solid", start_color="FFEB9C")
    B25_F    = PatternFill("solid", start_color="FCE4D6")

    def hdr(ws, headers):
        for ci, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=ci, value=h)
            c.fill = H_FILL; c.font = H_FONT
            c.alignment = CTR; c.border = BDR

    def dcell(ws, r, ci, val, align=CTR, fill=None):
        c = ws.cell(row=r, column=ci, value=val)
        c.font = D_FONT; c.border = BDR; c.alignment = align
        if fill: c.fill = fill
        return c

    def write_journal(ws, found):
        hdr(ws, ["#","Bâtiment","Type","Ligne","Colonne","Orient.",
                 "H","L","Culture reçue","Boost","Production"])
        for ri, bat in enumerate(found, 2):
            bi = catalog.get(bat["nom"], {})
            bt = bi.get("type","")
            is_p = bt == "Producteur"
            rf = PROD_F if is_p else (ALT_F if ri%2==0 else PatternFill("solid", start_color="FFFFFF"))
            bf = {"100%":B100_F,"50%":B50_F,"25%":B25_F}.get(bat.get("boost",""))
            vals = [ri-1, bat["nom"], bt, bat["ligne"], bat["colonne"],
                    bat["orientation"], bat["hauteur"], bat["largeur"],
                    bat.get("culture_recue","") if is_p else "",
                    bat.get("boost","")         if is_p else "",
                    bi.get("production","")     if is_p else ""]
            for ci, v in enumerate(vals, 1):
                cell = dcell(ws, ri, ci, v, align=LFT if ci==2 else CTR, fill=rf)
                if ci == 10 and bf: cell.fill = bf
        for ci, w in enumerate([5,30,12,7,7,8,5,5,14,8,14], 1):
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = "A2"

    def write_resume(ws, found, label=""):
        hdr(ws, ["Bâtiment","Trouvés","Attendus","Statut","Type","Production"])
        counts = {}
        for f in found: counts[f["nom"]] = counts.get(f["nom"],0)+1
        for ri,(nom,cnt) in enumerate(sorted(counts.items()),2):
            bi = catalog.get(nom,{})
            exp = bi.get("quantite","?")
            try: ok = int(exp)==cnt
            except: ok = False
            sf = OK_F if ok else WN_F
            for ci,v in enumerate([nom,cnt,exp,"✅" if ok else "⚠️",
                                    bi.get("type",""),bi.get("production","")],1):
                c = dcell(ws,ri,ci,v,align=LFT if ci==1 else CTR)
                if ci in(2,3): c.fill=sf
        for ci,w in enumerate([30,10,10,8,12,14],1):
            ws.column_dimensions[get_column_letter(ci)].width=w
        ws.freeze_panes="A2"

    def write_boost_summary(ws, found, label=""):
        hdr(ws, ["Production","0%","25%","50%","100%","Total","% à 100%"])
        summ = boost_summary(found, catalog)
        for ri,(prod,d) in enumerate(sorted(summ.items()),2):
            total = sum(d.values())
            pct   = round(d["100%"]/total*100,1) if total else 0
            for ci,v in enumerate([prod,d["0%"],d["25%"],d["50%"],d["100%"],total,f"{pct}%"],1):
                c = dcell(ws,ri,ci,v,align=LFT if ci==1 else CTR)
                if ci==5 and d["100%"]>0: c.fill=B100_F
                if ci==3 and d["25%"]>0:  c.fill=B25_F
                if ci==4 and d["50%"]>0:  c.fill=B50_F
        for ci,w in enumerate([20,8,8,8,8,8,10],1):
            ws.column_dimensions[get_column_letter(ci)].width=w
        ws.freeze_panes="A2"

    # Onglets situation initiale
    ws1 = wb.active; ws1.title = "Journal initial"
    write_journal(ws1, initial_found)

    ws2 = wb.create_sheet("Résumé initial")
    write_resume(ws2, initial_found)

    ws3 = wb.create_sheet("Boosts initiaux")
    write_boost_summary(ws3, initial_found)

    # Onglets optimisés
    if best_found:
        ws4 = wb.create_sheet("Journal optimisé")
        write_journal(ws4, best_found)

        ws5 = wb.create_sheet("Résumé optimisé")
        write_resume(ws5, best_found)

        ws6 = wb.create_sheet("Boosts optimisés")
        write_boost_summary(ws6, best_found)

        # Onglet comparaison
        ws7 = wb.create_sheet("Comparaison")
        hdr(ws7, ["Métrique","Avant","Après","Δ"])
        summ_i = boost_summary(initial_found, catalog)
        summ_b = boost_summary(best_found, catalog)
        tot_i = {k:sum(d[k] for d in summ_i.values()) for k in["0%","25%","50%","100%"]}
        tot_b = {k:sum(d[k] for d in summ_b.values()) for k in["0%","25%","50%","100%"]}
        rows_cmp = [
            ("Score moyen boost", round(initial_score,3), round(best_score,3),
             round(best_score-initial_score,3)),
            ("Producteurs à 0%",   tot_i["0%"],   tot_b["0%"],   tot_b["0%"]-tot_i["0%"]),
            ("Producteurs à 25%",  tot_i["25%"],  tot_b["25%"],  tot_b["25%"]-tot_i["25%"]),
            ("Producteurs à 50%",  tot_i["50%"],  tot_b["50%"],  tot_b["50%"]-tot_i["50%"]),
            ("Producteurs à 100%", tot_i["100%"], tot_b["100%"], tot_b["100%"]-tot_i["100%"]),
        ]
        for ri,(label,avant,apres,delta) in enumerate(rows_cmp,2):
            col_d_fill = B100_F if isinstance(delta,(int,float)) and delta>0 else (
                         WN_F   if isinstance(delta,(int,float)) and delta<0 else None)
            for ci,v in enumerate([label,avant,apres,
                                   ("+"+str(delta) if isinstance(delta,(int,float)) and delta>0
                                    else str(delta))],1):
                c = dcell(ws7,ri,ci,v,align=LFT if ci==1 else CTR)
                if ci==4 and col_d_fill: c.fill=col_d_fill
        for ci,w in enumerate([28,12,12,10],1):
            ws7.column_dimensions[get_column_letter(ci)].width=w

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ─── Bouton de téléchargement ─────────────────────────────────────────────────

st.divider()
best_found_dl   = st.session_state.get("best_found")
best_score_dl   = st.session_state.get("best_score")
init_score_dl   = st.session_state.get("initial_score", initial_score)
init_found_dl   = st.session_state.get("initial_found", initial_found)

output_buf = make_excel(init_found_dl, catalog, best_found_dl, best_score_dl, init_score_dl)

label = "📥 Télécharger resultats.xlsx"
if best_found_dl:
    label += f"  (initial {init_score_dl:.3f} → optimisé {best_score_dl:.3f})"

st.download_button(
    label=label,
    data=output_buf,
    file_name="resultats_batiments.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
