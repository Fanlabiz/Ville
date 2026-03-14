import streamlit as st
import pandas as pd
import random, math, copy, time
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
    batiments_df.columns = [str(c).strip() for c in batiments_df.columns]
    return terrain_df, batiments_df, actuel_df

file_bytes = uploaded_file.read()
terrain_df, batiments_df, actuel_df = load_data(file_bytes)

# ─── Utilitaires ─────────────────────────────────────────────────────────────

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
        grid.append(['' if pd.isna(v) else str(v).strip() for v in row])
    mc = max(len(r) for r in grid)
    for r in grid:
        while len(r) < mc:
            r.append('')
    return grid

def build_free_mask(df):
    """Cases libres du terrain (valeur=1), coordonnées 0-based."""
    free = set()
    for r, row in df.iterrows():
        for c, val in enumerate(row):
            if str(val).strip() == '1':
                free.add((r, c))
    return free

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
            "longueur":    lon, "largeur": lar,
            "quantite":    safe_int(row.get("Quantite",    0)),
            "type":        str(row.get("Type",    "")).strip(),
            "culture":     safe_int(row.get("Culture",     0)),
            "rayonnement": safe_int(row.get("Rayonnement", 0)),
            "boost25":     safe_float(row.get("Boost 25%")),
            "boost50":     safe_float(row.get("Boost 50%")),
            "boost100":    safe_float(row.get("Boost 100%")),
            "production":  str(row.get("Production", "")).strip(),
        }
    return catalog

# ─── Détection ───────────────────────────────────────────────────────────────

def flood_fill(grid, r0, c0, name, rows, cols):
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

def detect_buildings(grid, catalog):
    rows = len(grid); cols = len(grid[0]) if rows else 0
    visited = [[False]*cols for _ in range(rows)]
    found, warnings = [], []
    for r in range(rows):
        for c in range(cols):
            cell = grid[r][c]
            if not cell or cell in ("X","0","1","") or visited[r][c]:
                continue
            if cell not in catalog:
                continue
            blob = flood_fill(grid, r, c, cell, rows, cols)
            for br, bc in blob:
                visited[br][bc] = True
            bat = catalog[cell]
            lon, lar = bat["longueur"], bat["largeur"]
            placed = False
            for h, w in [(lar, lon), (lon, lar)]:
                pl = tile_box(blob, h, w)
                if pl is not None:
                    orient = "H" if w == lon else "V"
                    for r0, c0 in pl:
                        cases = {(r0+dr, c0+dc) for dr in range(h) for dc in range(w)}
                        found.append({
                            "nom": cell, "ligne": r0+1, "colonne": c0+1,
                            "hauteur": h, "largeur": w,
                            "orientation": orient, "cases": cases,
                        })
                    placed = True; break
            if not placed:
                mr = min(br for br,bc in blob); mc2 = min(bc for br,bc in blob)
                xr = max(br for br,bc in blob); xc  = max(bc for br,bc in blob)
                warnings.append(
                    f"⚠️ **{cell}** — {len(blob)} cases "
                    f"[r{mr+1}-{xr+1}, c{mc2+1}-{xc+1}] impossible à paver {lar}×{lon}"
                )
    return found, warnings

# ─── Culture & Boost ─────────────────────────────────────────────────────────

def rayonnement_zone(bat, R):
    r0 = bat["ligne"] - 1; c0 = bat["colonne"] - 1
    r1 = r0 + bat["hauteur"] - 1; c1 = c0 + bat["largeur"] - 1
    return {(r, c) for r in range(r0-R, r1+R+1) for c in range(c0-R, c1+R+1)}

def compute_boost_label(culture, bi):
    b25 = bi.get("boost25"); b50 = bi.get("boost50"); b100 = bi.get("boost100")
    if b100 is not None and culture >= b100: return "100%"
    if b50  is not None and culture >= b50:  return "50%"
    if b25  is not None and culture >= b25:  return "25%"
    if b25  is not None:                     return "0%"
    return ""

def boost_score_val(culture, bi):
    return {"100%": 3, "50%": 2, "25%": 1, "0%": 0, "": 0}[
        compute_boost_label(culture, bi)]

def compute_culture_boost(found, catalog):
    """Calcule culture_recue et boost pour chaque bâtiment. Modifie found in-place."""
    czones = []
    for b in found:
        if catalog.get(b["nom"],{}).get("type") != "Culturel":
            continue
        bi = catalog[b["nom"]]; R = bi["rayonnement"]
        if R > 0 and bi["culture"] > 0:
            czones.append({"zone": rayonnement_zone(b, R), "culture": bi["culture"]})
    for bat in found:
        bat["culture_recue"] = 0; bat["boost"] = ""
        if catalog.get(bat["nom"],{}).get("type") == "Producteur":
            total = sum(cz["culture"] for cz in czones if bat["cases"] & cz["zone"])
            bat["culture_recue"] = total
            bat["boost"] = compute_boost_label(total, catalog[bat["nom"]])
    return found

def mean_boost_score(found, catalog):
    scores = [boost_score_val(b["culture_recue"], catalog[b["nom"]])
              for b in found if catalog.get(b["nom"],{}).get("type") == "Producteur"]
    return sum(scores) / len(scores) if scores else 0.0

def boost_distribution(found, catalog):
    d = {"0%": 0, "25%": 0, "50%": 0, "100%": 0, "(aucun)": 0}
    for b in found:
        if catalog.get(b["nom"],{}).get("type") != "Producteur":
            continue
        k = b["boost"] if b["boost"] else "(aucun)"
        d[k] = d.get(k, 0) + 1
    return d

def prod_summary(found, catalog):
    """Résumé par type de Production."""
    prod_map = {}
    for p in found:
        if catalog.get(p["nom"],{}).get("type") != "Producteur":
            continue
        pl = catalog[p["nom"]].get("production","") or ""
        k  = pl if pl and pl not in ("","Rien","nan") else p["nom"]
        if k not in prod_map:
            prod_map[k] = {"0%":0,"25%":0,"50%":0,"100%":0,"(aucun)":0}
        bk = p["boost"] if p["boost"] else "(aucun)"
        prod_map[k][bk] = prod_map[k].get(bk, 0) + 1
    return prod_map

# ─── Optimisation par recuit simulé (swaps) ──────────────────────────────────

def make_bat(original, r0, c0, h, w, catalog):
    new = dict(original)
    new["ligne"] = r0+1; new["colonne"] = c0+1
    new["hauteur"] = h;  new["largeur"] = w
    lon = catalog[original["nom"]]["longueur"]
    new["orientation"] = "H" if w == lon else "V"
    new["cases"] = {(r0+dr, c0+dc) for dr in range(h) for dc in range(w)}
    return new

def try_swap(found, i, j, free_mask, catalog):
    """
    Tente d'échanger les positions de found[i] et found[j].
    Teste toutes les orientations compatibles.
    Retourne (new_i, new_j) ou (None, None).
    """
    a = found[i]; b = found[j]
    a_r0, a_c0 = a["ligne"]-1, a["colonne"]-1
    b_r0, b_c0 = b["ligne"]-1, b["colonne"]-1

    # Cases de tous les autres bâtiments (hors i et j)
    other = set()
    for k, bat in enumerate(found):
        if k != i and k != j:
            other |= bat["cases"]

    a_lon = catalog[a["nom"]]["longueur"]; a_lar = catalog[a["nom"]]["largeur"]
    b_lon = catalog[b["nom"]]["longueur"]; b_lar = catalog[b["nom"]]["largeur"]
    a_orients = [(a_lar, a_lon)] if a_lar == a_lon else [(a_lar, a_lon), (a_lon, a_lar)]
    b_orients = [(b_lar, b_lon)] if b_lar == b_lon else [(b_lar, b_lon), (b_lon, b_lar)]

    for new_ah, new_aw in a_orients:
        new_a = {(b_r0+dr, b_c0+dc) for dr in range(new_ah) for dc in range(new_aw)}
        if not new_a.issubset(free_mask) or (new_a & other):
            continue
        for new_bh, new_bw in b_orients:
            new_b = {(a_r0+dr, a_c0+dc) for dr in range(new_bh) for dc in range(new_bw)}
            if not new_b.issubset(free_mask) or (new_b & other) or (new_a & new_b):
                continue
            return (make_bat(a, b_r0, b_c0, new_ah, new_aw, catalog),
                    make_bat(b, a_r0, a_c0, new_bh, new_bw, catalog))
    return None, None

def simulated_annealing(found_init, catalog, free_mask,
                        T_start=2.0, T_end=0.01, n_iter=10000,
                        progress_cb=None):
    """
    Recuit simulé par échanges (swaps) de positions entre paires de bâtiments.
    Adapté aux terrains très denses. Prend une copie indépendante de found_init.
    """
    # Copie totalement indépendante de l'état initial
    current = copy.deepcopy(found_init)
    current = compute_culture_boost(current, catalog)

    best       = copy.deepcopy(current)
    best_score = mean_boost_score(best, catalog)
    cur_score  = best_score

    T     = T_start
    alpha = (T_end / T_start) ** (1.0 / n_iter)
    n     = len(current)

    for it in range(n_iter):
        T *= alpha

        i = random.randrange(n)
        j = random.randrange(n)
        if i == j:
            continue

        new_a, new_b = try_swap(current, i, j, free_mask, catalog)
        if new_a is None:
            continue

        # Construire nouvel état par copie partielle (plus rapide que deepcopy total)
        new_found = list(current)            # copie superficielle de la liste
        new_found[i] = new_a
        new_found[j] = new_b
        # Recalculer uniquement la culture/boost (in-place sur new_found)
        new_found = compute_culture_boost(new_found, catalog)
        new_score = mean_boost_score(new_found, catalog)

        delta = new_score - cur_score
        if delta > 0 or random.random() < math.exp(delta / T):
            current   = new_found
            cur_score = new_score
            if cur_score > best_score:
                best       = copy.deepcopy(current)
                best_score = cur_score

        if progress_cb and it % 200 == 0:
            progress_cb(it / n_iter, cur_score, best_score)

    return best, best_score

# ─── Initialisation ──────────────────────────────────────────────────────────

free_mask   = build_free_mask(terrain_df)
catalog     = build_catalog(batiments_df)
actuel_grid = df_to_grid(actuel_df)

with st.spinner("Détection des bâtiments…"):
    found_init, detect_warnings = detect_buildings(actuel_grid, catalog)
    found_init = compute_culture_boost(found_init, catalog)

initial_score = mean_boost_score(found_init, catalog)

# ─── Interface ───────────────────────────────────────────────────────────────

tab1, tab2 = st.tabs(["📍 État actuel", "🔧 Optimisation"])

with tab1:
    nb = len(found_init); nw = len(detect_warnings)
    st.subheader(f"{nb} bâtiment(s) détecté(s)"
                 + (f"  ·  ⚠️ {nw} avertissement(s)" if nw else ""))
    st.metric("Score boost moyen", f"{initial_score:.3f} / 3.000")

    col1, col2 = st.columns([3, 2])
    with col1:
        st.markdown("### Journal")
        for bat in found_init:
            bi = catalog.get(bat["nom"], {}); t = bi.get("type","")
            line = (f"✅ **{bat['nom']}** — l{bat['ligne']}, c{bat['colonne']} "
                    f"({bat['hauteur']}×{bat['largeur']}, {bat['orientation']})")
            if t == "Producteur":
                line += f" | Culture : **{bat['culture_recue']}** | Boost : **{bat['boost']}**"
            st.markdown(line)
        for w in detect_warnings:
            st.markdown(w)

    with col2:
        st.markdown("### Résumé par Production")
        pm = prod_summary(found_init, catalog)
        if pm:
            rows = [{"Production": k, "0%": v["0%"], "25%": v["25%"],
                     "50%": v["50%"], "100%": v["100%"]}
                    for k, v in sorted(pm.items())]
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

with tab2:
    st.markdown("""
    ### Recuit simulé par échanges de positions
    L'optimiseur **échange** les positions de paires de bâtiments pour maximiser
    le boost moyen des Producteurs.
    > ℹ️ Le terrain est occupé à ~98% — les déplacements libres sont impossibles,
    > seuls les échanges (swaps) sont utilisés.
    """)

    cola, colb, colc = st.columns(3)
    with cola:
        n_iter  = st.slider("Itérations", 2000, 30000, 10000, step=1000)
    with colb:
        t_start = st.slider("Température initiale", 0.5, 10.0, 2.0, step=0.5)
    with colc:
        seed = st.number_input("Graine aléatoire", 0, 9999, 42)

    if st.button("🚀 Lancer l'optimisation", type="primary"):
        random.seed(int(seed))

        prog_bar  = st.progress(0.0)
        prog_text = st.empty()

        def progress_cb(frac, cur_s, best_s):
            prog_bar.progress(min(frac, 1.0))
            prog_text.markdown(
                f"Itération {int(frac*n_iter)}/{n_iter} — "
                f"T en cours — Score courant : `{cur_s:.3f}` — **Meilleur : `{best_s:.3f}`**"
            )

        t0 = time.time()
        with st.spinner("Optimisation en cours…"):
            # IMPORTANT : on passe found_init (état initial non modifié)
            # simulated_annealing en fait une copie profonde en interne
            found_opt, opt_score = simulated_annealing(
                found_init, catalog, free_mask,
                T_start=float(t_start), T_end=0.01, n_iter=int(n_iter),
                progress_cb=progress_cb
            )
        elapsed = time.time() - t0
        prog_bar.progress(1.0)
        prog_text.success(f"✅ Terminé en {elapsed:.1f}s")

        gain = opt_score - initial_score
        c1, c2, c3 = st.columns(3)
        c1.metric("Score initial",   f"{initial_score:.3f}")
        c2.metric("Score optimisé",  f"{opt_score:.3f}")
        c3.metric("Gain",            f"{gain:+.3f}", delta_color="normal")

        # Tableau comparatif des boosts
        st.markdown("### Comparaison des boosts")
        bd_i = boost_distribution(found_init, catalog)
        bd_o = boost_distribution(found_opt,  catalog)
        comp = pd.DataFrame({
            "Boost":  ["0%","25%","50%","100%","(aucun)"],
            "Avant":  [bd_i[k] for k in ["0%","25%","50%","100%","(aucun)"]],
            "Après":  [bd_o[k] for k in ["0%","25%","50%","100%","(aucun)"]],
        })
        comp["Δ"] = comp["Après"] - comp["Avant"]
        st.dataframe(comp, use_container_width=True, hide_index=True)

        # Résumé par Production après optimisation
        st.markdown("### Résumé par Production — après optimisation")
        pm_opt = prod_summary(found_opt, catalog)
        if pm_opt:
            rows_opt = [{"Production": k, "0%": v["0%"], "25%": v["25%"],
                         "50%": v["50%"], "100%": v["100%"]}
                        for k, v in sorted(pm_opt.items())]
            st.dataframe(pd.DataFrame(rows_opt), use_container_width=True, hide_index=True)

        # ── Génération Excel ─────────────────────────────────────────────────

        def make_excel(f_init, f_opt, catalog, warnings):
            wb = Workbook()
            HF = PatternFill("solid", start_color="1F4E79")
            HN = Font(bold=True, color="FFFFFF", name="Arial", size=11)
            DN = Font(name="Arial", size=10)
            CE = Alignment(horizontal="center", vertical="center", wrap_text=True)
            LE = Alignment(horizontal="left",   vertical="center", wrap_text=True)
            BR = Border(left=Side(style="thin"), right=Side(style="thin"),
                        top=Side(style="thin"),  bottom=Side(style="thin"))
            OK   = PatternFill("solid", start_color="C6EFCE")
            WN2  = PatternFill("solid", start_color="FFEB9C")
            AL   = PatternFill("solid", start_color="DEEAF1")
            PR   = PatternFill("solid", start_color="FFF2CC")
            B100 = PatternFill("solid", start_color="92D050")
            B50  = PatternFill("solid", start_color="FFEB9C")
            B25  = PatternFill("solid", start_color="FCE4D6")

            def hdr(ws, cols):
                for ci, h in enumerate(cols, 1):
                    cell = ws.cell(row=1, column=ci, value=h)
                    cell.fill=HF; cell.font=HN; cell.alignment=CE; cell.border=BR

            def dc(ws, r, c, val, align=CE, fill=None):
                cell = ws.cell(row=r, column=c, value=val)
                cell.font=DN; cell.border=BR; cell.alignment=align
                if fill: cell.fill=fill
                return cell

            def write_journal(ws, found):
                hdr(ws, ["#","Bâtiment","Type","Ligne","Colonne","Orient.",
                         "Haut.","Larg.","Culture reçue","Boost","Production"])
                for ri, bat in enumerate(found, 2):
                    bi   = catalog.get(bat["nom"], {})
                    btyp = bi.get("type","")
                    is_p = btyp == "Producteur"
                    boost = bat.get("boost","")
                    rf = PR if is_p else (AL if ri%2==0 else PatternFill("solid",start_color="FFFFFF"))
                    row = [ri-1, bat["nom"], btyp, bat["ligne"], bat["colonne"],
                           bat["orientation"], bat["hauteur"], bat["largeur"],
                           bat.get("culture_recue","") if is_p else "",
                           boost                       if is_p else "",
                           bi.get("production","")     if is_p else ""]
                    for ci, v in enumerate(row, 1):
                        cell = dc(ws, ri, ci, v, align=LE if ci==2 else CE, fill=rf)
                        # Colorier la colonne Boost
                        if ci == 10 and is_p:
                            if boost == "100%": cell.fill = B100
                            elif boost == "50%": cell.fill = B50
                            elif boost == "25%": cell.fill = B25
                for ci, w in enumerate([5,30,12,7,7,7,6,6,13,8,13], 1):
                    ws.column_dimensions[get_column_letter(ci)].width = w
                ws.freeze_panes = "A2"

            def write_resume(ws, found):
                hdr(ws, ["Production","0%","25%","50%","100%","(aucun)","Total"])
                pm = prod_summary(found, catalog)
                for ri, (k, cnts) in enumerate(sorted(pm.items()), 2):
                    total = sum(cnts.values())
                    row = [k, cnts["0%"], cnts["25%"], cnts["50%"],
                           cnts["100%"], cnts["(aucun)"], total]
                    for ci, v in enumerate(row, 1):
                        dc(ws, ri, ci, v, align=LE if ci==1 else CE)
                for ci, w in enumerate([30,8,8,8,8,8,7], 1):
                    ws.column_dimensions[get_column_letter(ci)].width = w
                ws.freeze_panes = "A2"

            def write_producteurs(ws, found):
                hdr(ws, ["Bâtiment","Ligne","Colonne","Production",
                         "Culture reçue","Seuil 25%","Seuil 50%","Seuil 100%","Boost"])
                prods = [b for b in found if catalog.get(b["nom"],{}).get("type")=="Producteur"]
                for ri, bat in enumerate(prods, 2):
                    bi = catalog.get(bat["nom"], {})
                    boost = bat.get("boost","")
                    bf = B100 if boost=="100%" else B50 if boost=="50%" else B25 if boost=="25%" else None
                    row = [bat["nom"], bat["ligne"], bat["colonne"],
                           bi.get("production",""), bat.get("culture_recue",0),
                           bi.get("boost25","") if bi.get("boost25") is not None else "",
                           bi.get("boost50","") if bi.get("boost50") is not None else "",
                           bi.get("boost100","") if bi.get("boost100") is not None else "",
                           boost]
                    for ci, v in enumerate(row, 1):
                        cell = dc(ws, ri, ci, v, align=LE if ci==1 else CE)
                        if ci == 9 and bf: cell.fill = bf
                for ci, w in enumerate([30,7,7,14,14,10,10,10,10], 1):
                    ws.column_dimensions[get_column_letter(ci)].width = w
                ws.freeze_panes = "A2"

            # ── Onglets état initial ─────────────────────────────────────────
            ws1 = wb.active; ws1.title = "Journal initial"
            write_journal(ws1, f_init)

            ws2 = wb.create_sheet("Résumé initial")
            write_resume(ws2, f_init)

            ws3 = wb.create_sheet("Producteurs initial")
            write_producteurs(ws3, f_init)

            # ── Onglets état optimisé ────────────────────────────────────────
            ws4 = wb.create_sheet("Journal optimisé")
            write_journal(ws4, f_opt)

            ws5 = wb.create_sheet("Résumé optimisé")
            write_resume(ws5, f_opt)

            ws6 = wb.create_sheet("Producteurs optimisé")
            write_producteurs(ws6, f_opt)

            # ── Comparaison bâtiment par bâtiment ───────────────────────────
            ws7 = wb.create_sheet("Comparaison")
            hdr(ws7, ["Bâtiment","Type",
                      "Ligne init","Col. init","Boost init",
                      "Ligne opt","Col. opt","Boost opt","Δ"])
            for ri, (bi_bat, bo_bat) in enumerate(zip(f_init, f_opt), 2):
                btyp   = catalog.get(bi_bat["nom"],{}).get("type","")
                moved  = (bi_bat["ligne"] != bo_bat["ligne"] or
                          bi_bat["colonne"] != bo_bat["colonne"] or
                          bi_bat["orientation"] != bo_bat["orientation"])
                b_init = bi_bat.get("boost",""); b_opt = bo_bat.get("boost","")
                improved = (btyp == "Producteur" and b_init != b_opt)
                flag   = ("📦" if moved else "") + ("⬆️" if improved else "")
                rfill  = OK if improved else (AL if moved else None)
                row = [bi_bat["nom"], btyp,
                       bi_bat["ligne"], bi_bat["colonne"], b_init,
                       bo_bat["ligne"], bo_bat["colonne"], b_opt, flag]
                for ci, v in enumerate(row, 1):
                    dc(ws7, ri, ci, v, align=LE if ci==1 else CE, fill=rfill)
            for ci, w in enumerate([30,12,8,8,8,8,8,8,6], 1):
                ws7.column_dimensions[get_column_letter(ci)].width = w
            ws7.freeze_panes = "A2"

            if warnings:
                ws8 = wb.create_sheet("Avertissements")
                ws8.cell(row=1,column=1,value="Avertissements").font=Font(bold=True)
                for ri, w in enumerate(warnings, 2):
                    ws8.cell(row=ri,column=1,value=w.replace("⚠️ ","").replace("**",""))
                ws8.column_dimensions["A"].width = 70

            buf = io.BytesIO()
            wb.save(buf); buf.seek(0)
            return buf

        output_buf = make_excel(found_init, found_opt, catalog, detect_warnings)

        st.download_button(
            label="📥 Télécharger resultats_optimises.xlsx",
            data=output_buf,
            file_name="resultats_optimises.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
