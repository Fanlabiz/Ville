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


def _to_float(val, default=0.0):
    """
    Convertit une valeur en float de facon robuste.
    Gere les formules Excel non evaluees (ex: '=49980/2').
    """
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if not s:
        return default
    if s.startswith("="):
        import re
        expr = s[1:]
        if re.fullmatch(r"[\d\s\+\-\*\/\(\)\.]+", expr):
            try:
                return float(eval(expr))
            except Exception:
                pass
        return default
    try:
        return float(s)
    except (ValueError, TypeError):
        return default


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
            "longueur":   int(_to_float(d.get("Longueur"), 1)),
            "largeur":    int(_to_float(d.get("Largeur"), 1)),
            "nombre":     int(_to_float(d.get("Nombre"), 1)),
            "type":       str(d.get("Type", "")).strip(),
            "culture":    _to_float(d.get("Culture")),
            "rayonnement":int(_to_float(d.get("Rayonnement"))),
            "boost25":    _to_float(d.get("Boost 25%")),
            "boost50":    _to_float(d.get("Boost 50%")),
            "boost100":   _to_float(d.get("Boost 100%")),
            "production": str(d.get("Production", "")).strip(),
            "quantite":   _to_float(d.get("Quantite")),
            "priorite":   _to_float(d.get("Priorite"), 0.0),
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
            "production": "Rien", "quantite": 0, "priorite": 0,
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
    """Score total = somme(boost*priorite) pour les Producteurs (grande priorite = plus important)."""
    culturels = [b for b in placed if b["type"] == "Culturel"]
    total = 0.0
    for b in placed:
        if b["type"] == "Producteur" and b["priorite"] > 0:
            cult = culture_received(b, culturels)
            boost = boost_level(cult, b)
            total += boost * b["priorite"]
    return total


# ══════════════════════════════════════════════════════
# PLACEMENT INITIAL DES BATIMENTS MANQUANTS
# ══════════════════════════════════════════════════════

def place_missing_buildings(placed, buildings_def, terrain_grid, max_r, max_c,
                            n_trials=10, time_budget=20.0):
    """
    Place les batiments manquants en combinant FFD multi-start et
    un scoring guide par des regles metier :

    Regle 1 (Neutres en bordure) :
      Les batiments Neutres preferent les cases adjacentes aux X,
      pour laisser le coeur du terrain aux batiments qui impactent le score.

    Regle 2 (Culturels orientes vers les producteurs) :
      Un culturel est place a la position ou sa zone de rayonnement contient
      le plus de cases libres ou deja occupees par des producteurs.
      Cases libres = futurs producteurs potentiels.
      Cases neutre/culturel dans la zone = inutiles pour ce calcul.

    Regle 3 (Eviter la fragmentation) :
      Apres placement hypothetique, penaliser les positions qui creent
      des espaces residuels trop petits ou trop etroits (incapables
      d'accueillir un batiment de taille minimale significative).

    Ces regles sont des preferences, pas des obligations.
    Si aucune position ideale n'existe, on prend la moins mauvaise
    pour garantir le placement de tous les batiments.

    Ordre de placement : Producteurs -> Culturels -> Neutres
    (les culturels voient les producteurs deja en place pour la regle 2).

    Retourne (new_placed, n_placed, n_failed).
    """
    import random, time
    from collections import Counter
    from itertools import groupby

    placed_counts = Counter(b["nom"].strip() for b in placed)
    to_place_base = []
    for b_def in buildings_def:
        nom = b_def["nom"].strip()
        needed = b_def["nombre"] - placed_counts.get(nom, 0)
        for _ in range(needed):
            to_place_base.append(dict(b_def, nom=nom,
                                      rows=b_def["largeur"],
                                      cols=b_def["longueur"]))

    if not to_place_base:
        return [dict(b) for b in placed], 0, 0

    x_grid = make_x_grid(terrain_grid, max_r, max_c)

    # Pre-calculer la distance au X le plus proche pour chaque case interieure
    # via BFS multi-source depuis toutes les cases X simultanement -> O(n²)
    from collections import deque
    dist_to_x = [[999] * max_c for _ in range(max_r)]
    bfs_queue = deque()
    for r in range(max_r):
        for c in range(max_c):
            if terrain_grid[r][c] == "X":
                dist_to_x[r][c] = 0
                bfs_queue.append((r, c))
    while bfs_queue:
        r, c = bfs_queue.popleft()
        for dr, dc in ((-1,0),(1,0),(0,-1),(0,1)):
            nr, nc = r+dr, c+dc
            if 0 <= nr < max_r and 0 <= nc < max_c and dist_to_x[nr][nc] == 999:
                dist_to_x[nr][nc] = dist_to_x[r][c] + 1
                bfs_queue.append((nr, nc))


    def build_occ(placed_list):
        occ = [[False] * max_c for _ in range(max_r)]
        for b in placed_list:
            for dr in range(b["rows"]):
                for dc in range(b["cols"]):
                    rr, cc = b["r"] + dr, b["c"] + dc
                    if 0 <= rr < max_r and 0 <= cc < max_c:
                        occ[rr][cc] = True
        return occ

    def cells_rect(r, c, rows, cols):
        return {(r + dr, c + dc) for dr in range(rows) for dc in range(cols)}

    # ── Regle 2 : score culturel ──
    def score_culturel(r, c, rows, cols, ray, occ, prod_cells_set):
        """
        Compte les cases utiles dans la zone de rayonnement du culturel
        si place en (r,c) avec taille rows×cols et rayonnement ray.
        Utile = libre ET dans le terrain, OU deja occupee par un producteur.
        """
        r0, c0 = r, c
        r1, c1 = r + rows - 1, c + cols - 1
        count = 0
        for rz in range(r0 - ray, r1 + ray + 1):
            for cz in range(c0 - ray, c1 + ray + 1):
                if r0 <= rz <= r1 and c0 <= cz <= c1:
                    continue  # case du batiment lui-meme
                if not (0 <= rz < max_r and 0 <= cz < max_c):
                    continue
                if x_grid[rz][cz]:
                    continue
                if (rz, cz) in prod_cells_set:
                    count += 2   # producteur deja en place : bonus
                elif not occ[rz][cz]:
                    count += 1   # case libre : futur producteur possible
        return count

    # ── Regle 3 : penalite fragmentation ──
    MIN_USEFUL = 3  # taille minimale d'un bloc utile (3 cases dans au moins une dimension)

    def fragmentation_penalty(r, c, rows, cols, occ):
        """
        Apres placement hypothetique en (r,c), verifie le voisinage immediat.
        Compte les zones libres adjacentes qui deviendraient trop petites ou
        trop etroites (< MIN_USEFUL cases dans chaque dimension d'un rectangle
        minimal). On ne scanne que le voisinage elargi pour rester rapide.
        """
        # Marquer temporairement les cases du batiment comme occupees
        temp_occ_cells = cells_rect(r, c, rows, cols)

        penalty = 0
        # Regarder les espaces libres contigus dans le voisinage elargi (rayon 2)
        margin = 2
        visited = set()
        for rn in range(r - margin, r + rows + margin):
            for cn in range(c - margin, c + cols + margin):
                if (rn, cn) in visited or (rn, cn) in temp_occ_cells:
                    continue
                if not (0 <= rn < max_r and 0 <= cn < max_c):
                    continue
                if x_grid[rn][cn] or occ[rn][cn]:
                    continue
                # BFS pour mesurer la taille et les dimensions de la zone libre adjacente
                zone = set()
                queue = [(rn, cn)]
                visited.add((rn, cn))
                while queue:
                    qr, qc = queue.pop()
                    zone.add((qr, qc))
                    for dr, dc in ((-1,0),(1,0),(0,-1),(0,1)):
                        nr, nc = qr+dr, qc+dc
                        if (nr, nc) not in visited and (nr, nc) not in temp_occ_cells:
                            if 0 <= nr < max_r and 0 <= nc < max_c:
                                if not x_grid[nr][nc] and not occ[nr][nc]:
                                    visited.add((nr, nc))
                                    queue.append((nr, nc))
                if not zone:
                    continue
                # Calculer les dimensions du bounding box de la zone
                rs = [zr for zr, zc in zone]
                cs = [zc for zr, zc in zone]
                height = max(rs) - min(rs) + 1
                width  = max(cs) - min(cs) + 1
                # Penaliser si trop petit ou trop etroit
                if len(zone) < MIN_USEFUL or height < 2 or width < 2:
                    penalty += 1
        return penalty

    # ── Fonction de score rapide (regles 1 et 2 seulement) ──
    def fast_score(b, r, c, rows, cols, occ, prod_cells_set):
        """Score rapide sans fragmentation, pour le screening initial."""
        s = 0.0
        if b["type"] == "Neutre":
            avg_dist = sum(dist_to_x[r+dr][c+dc]
                           for dr in range(rows) for dc in range(cols)) / (rows * cols)
            s += 10.0 / (1.0 + avg_dist)
        elif b["type"] == "Culturel":
            ray = b.get("rayonnement", 1)
            s += score_culturel(r, c, rows, cols, ray, occ, prod_cells_set)
        return s

    TOP_N = 10  # nombre de candidats evalues pour la regle 3

    def smart_place_one(ordered, seed):
        """
        Place les batiments dans l'ordre donne.

        Strategie par type :
        - Producteurs : first-fit pur (raster scan). Les regles metier
                        s'appliquent aux Culturels et Neutres, pas aux
                        Producteurs dont la position optimale sera trouvee
                        par l'optimiseur greedy apres le placement initial.
        - Culturels   : scoring regle 2 (cases prod/libres dans rayonnement)
                        + regle 3 (fragmentation, seuil strict <=1).
                        Fallback first-fit si fragmentation trop elevee.
        - Neutres     : scoring regle 1 (proximite des bords X)
                        + regle 3 (fragmentation, seuil genereux <=2).
                        Fallback first-fit si fragmentation trop elevee.
        """
        random.seed(seed)
        result = [dict(b) for b in placed]
        n_ok = n_fail = 0

        for b in ordered:
            occ = build_occ(result)

            # Trouver la premiere position raster valide (fallback absolu)
            first_fit = None
            all_candidates = []
            prod_cells_set = None

            for rows, cols in [(b["rows"], b["cols"]), (b["cols"], b["rows"])]:
                for r in range(max_r):
                    for c in range(max_c):
                        if can_place(r, c, rows, cols, x_grid, occ, max_r, max_c):
                            if first_fit is None:
                                first_fit = (r, c, rows, cols)
                            # Scorer seulement Culturels et Neutres
                            if b["type"] != "Producteur":
                                if b["type"] == "Culturel" and prod_cells_set is None:
                                    prod_cells_set = {
                                        (pb["r"] + dr, pb["c"] + dc)
                                        for pb in result if pb["type"] == "Producteur"
                                        for dr in range(pb["rows"])
                                        for dc in range(pb["cols"])
                                    }
                                s = fast_score(b, r, c, rows, cols, occ,
                                               prod_cells_set or set())
                                all_candidates.append((s, r, c, rows, cols))

            if first_fit is None:
                n_fail += 1
                continue

            if b["type"] == "Producteur" or not all_candidates:
                # Producteur : toujours first-fit
                r, c, rows, cols = first_fit
            else:
                all_candidates.sort(key=lambda x: -x[0])

                # Seuil de fragmentation selon le type
                max_penalty = 1 if b["type"] == "Culturel" else 2

                best_pos   = None
                best_final = None
                for s0, r, c, rows, cols in all_candidates[:TOP_N]:
                    penalty = fragmentation_penalty(r, c, rows, cols, occ)
                    final_s = s0 - 2.0 * penalty
                    if best_final is None or final_s > best_final:
                        best_final = final_s
                        best_pos   = (r, c, rows, cols)

                r, c, rows, cols = best_pos
                # Si fragmentation trop elevee -> first_fit
                if fragmentation_penalty(r, c, rows, cols, occ) > max_penalty:
                    r, c, rows, cols = first_fit

            result.append({**b, "r": r, "c": c, "rows": rows, "cols": cols})
            n_ok += 1

        return result, n_ok, n_fail

    def quick_score(placed_list, max_inner=2):
        """Screening rapide par quelques passes greedy."""
        p = [dict(b) for b in placed_list]
        moves = []
        for _ in range(max_inner):
            improved = False
            for b in p:
                if b["type"] == "Neutre":
                    continue
                best_s, best_pos = _best_position_for(b, p, x_grid, max_r, max_c)
                if best_pos:
                    _apply_move(b, best_pos, moves)
                    improved = True
            if not improved:
                break
        return score_placement(p)

    # Tri des bâtiments pour ffd_pure :
    # - Non-Neutres (Producteurs + Culturels) : taille décroissante ENSEMBLE.
    #   En les mélangeant dans le même flux, un Culturel de 3x4 se place
    #   naturellement à côté d'un Producteur de 3x4 → meilleure couverture.
    # - Neutres : taille décroissante, placés en bordure (regle 1).
    priority = {"Producteur": 0, "Neutre": 1, "Culturel": 2}
    sorted_base = sorted(to_place_base,
                         key=lambda b: (priority.get(b["type"], 9), -(b["rows"] * b["cols"])))

    sorted_ffd = sorted(to_place_base, key=lambda b: -(b["rows"] * b["cols"]))
    # BFS multi-source depuis toutes les cases X
    from collections import deque as _deque
    _dist = [[9999] * max_c for _ in range(max_r)]
    _bfs = _deque()
    for _r in range(max_r):
        for _c in range(max_c):
            if terrain_grid[_r][_c] == "X":
                _dist[_r][_c] = 0; _bfs.append((_r, _c))
    while _bfs:
        _r, _c = _bfs.popleft()
        for _dr, _dc in ((-1,0),(1,0),(0,-1),(0,1)):
            _nr, _nc = _r+_dr, _c+_dc
            if 0<=_nr<max_r and 0<=_nc<max_c and _dist[_nr][_nc]==9999:
                _dist[_nr][_nc] = _dist[_r][_c]+1; _bfs.append((_nr,_nc))
    # Cases internes triees par distance croissante au bord (bordure d'abord)
    cells_border_first = sorted(
        [(_r, _c) for _r in range(max_r) for _c in range(max_c) if not x_grid[_r][_c]],
        key=lambda rc: _dist[rc[0]][rc[1]]
    )
    # Cases internes triees par distance croissante au CENTRE (scan centripète)
    # Utilisé pour les casernes haute priorité: elles se placent au centre du terrain
    # et peuvent ainsi être entourées de culturels de tous côtés.
    _ctr_r, _ctr_c = max_r / 2.0, max_c / 2.0
    cells_center_first = sorted(
        [(_r, _c) for _r in range(max_r) for _c in range(max_c) if not x_grid[_r][_c]],
        key=lambda rc: abs(rc[0] - _ctr_r) + abs(rc[1] - _ctr_c)
    )

    # Listes séparées pour ffd_pure :
    # - non_neutres_ffd : Producteurs + Culturels avec tri prioritaire :
    #   1. Producteurs à haute priorité (casernes prio>=10) en tête, par taille décroissante
    #   2. Culturels forts (culture*rayonnement élevé) immédiatement après pour couvrir les casernes
    #   3. Reste des bâtiments par taille décroissante
    #   Ainsi les casernes se groupent en haut-gauche et les culturels forts les entourent.
    # - neutres_ffd     : Neutres triés par taille décroissante, placés en bordure
    max_prio = max((b["priorite"] for b in to_place_base if b["type"] == "Producteur"), default=1)
    prio_threshold = max_prio * 0.8  # considère "haute priorité" les 20% supérieurs

    # non_neutres_ffd : Producteurs + Culturels mélangés par taille décroissante.
    # Ce mélange naturel garantit 0 échecs et une bonne intrication spatiale.
    # Le biais dans ffd_pure oriente chaque Culturel vers les casernes en priorité.
    # La greedy_pass triée par priorité concentre ensuite les culturels sur les casernes.
    def _nn_sort_key(b):
        # Utilisé pour le groupby dans les trials suivants
        return -(b["rows"] * b["cols"])

    non_neutres_ffd = sorted(
        [b for b in to_place_base if b["type"] != "Neutre"],
        key=_nn_sort_key
    )
    neutres_ffd = sorted(
        [b for b in to_place_base if b["type"] == "Neutre"],
        key=lambda b: -(b["rows"] * b["cols"])
    )

    def ffd_pure(nn_ordered, n_ordered):
        """
        FFD en 3 phases :
        1. Neutres en bordure (garantit 0 echecs pour les Neutres).
        2. Casernes au centre (scan centripete pour les grouper).
        3. Culturels avec biais casernes + autres producteurs en first-fit.
        """
        result = [dict(b) for b in placed]
        n_ok = n_fail = 0
        placed_prods = []
        placed_top_prods = []

        # Phase 1 : Neutres en bordure
        for b in n_ordered:
            occ = build_occ(result)
            ok = False
            for r, c in cells_border_first:
                for rows, cols in [(b["rows"], b["cols"]), (b["cols"], b["rows"])]:
                    if can_place(r, c, rows, cols, x_grid, occ, max_r, max_c):
                        result.append({**b, "r": r, "c": c, "rows": rows, "cols": cols})
                        n_ok += 1; ok = True; break
                if ok: break
            if not ok:
                n_fail += 1

        # Phase 2 : Casernes au centre + Culturels avec biais + autres producteurs
        for b in nn_ordered:
            occ = build_occ(result)
            is_top = b["type"] == "Producteur" and b["priorite"] >= prio_threshold

            if is_top:
                chosen = None
                for r, c in cells_center_first:
                    for rows, cols in [(b["rows"], b["cols"]), (b["cols"], b["rows"])]:
                        if can_place(r, c, rows, cols, x_grid, occ, max_r, max_c):
                            chosen = (r, c, rows, cols); break
                    if chosen: break

            elif b["type"] == "Culturel":
                first_fit = None; chosen_top = None; chosen_any = None
                for rows, cols in [(b["rows"], b["cols"]), (b["cols"], b["rows"])]:
                    if chosen_top: break
                    for r in range(max_r):
                        if chosen_top: break
                        for c in range(max_c):
                            if can_place(r, c, rows, cols, x_grid, occ, max_r, max_c):
                                if first_fit is None: first_fit = (r, c, rows, cols)
                                ray = b.get("rayonnement", 1)
                                r0, c0, r1, c1 = r, c, r+rows-1, c+cols-1
                                zone = {(rz,cz) for rz in range(r0-ray,r1+ray+1)
                                        for cz in range(c0-ray,c1+ray+1)
                                        if not(r0<=rz<=r1 and c0<=cz<=c1)
                                        and 0<=rz<max_r and 0<=cz<max_c}
                                if placed_top_prods and any(cells_of(p)&zone for p in placed_top_prods):
                                    chosen_top=(r,c,rows,cols); break
                                elif chosen_any is None and placed_prods and any(cells_of(p)&zone for p in placed_prods):
                                    chosen_any=(r,c,rows,cols)
                chosen = chosen_top or chosen_any or first_fit

            else:
                chosen = None
                for rows, cols in [(b["rows"], b["cols"]), (b["cols"], b["rows"])]:
                    if chosen: break
                    for r in range(max_r):
                        if chosen: break
                        for c in range(max_c):
                            if can_place(r, c, rows, cols, x_grid, occ, max_r, max_c):
                                chosen = (r, c, rows, cols); break

            if chosen:
                r, c, rows, cols = chosen
                result.append({**b, "r": r, "c": c, "rows": rows, "cols": cols})
                if b["type"] == "Producteur":
                    placed_prods.append(result[-1])
                    if is_top: placed_top_prods.append(result[-1])
                n_ok += 1
            else:
                n_fail += 1

        return result, n_ok, n_fail

    best_placed      = None
    best_screen_score = -1
    best_n_placed    = 0
    best_n_failed    = len(to_place_base)
    t_start = time.time()

    # ── Trials : FFD avec mélange aléatoire des non-Neutres de même taille ──
    # Le trial 0 utilise l'ordre pur (taille décroissante).
    # Les trials suivants mélangent à l'intérieur de chaque groupe de taille,
    # ce qui peut produire de meilleurs mélanges Producteurs/Culturels.
    for trial in range(n_trials + 1):
        if trial > 0 and time.time() - t_start > time_budget:
            break

        random.seed(trial * 17 + 3)
        if trial == 0:
            nn_ordered = non_neutres_ffd  # ordre pur
            n_ordered  = neutres_ffd
        else:
            groups = []
            for _, g in groupby(non_neutres_ffd,
                                key=lambda b: -(b["rows"] * b["cols"])):
                grp = list(g); random.shuffle(grp); groups.append(grp)
            nn_ordered = [b for g in groups for b in g]
            # Mélanger aussi les Neutres (même taille) pour varier
            n_groups = []
            for _, g in groupby(neutres_ffd, key=lambda b: b["rows"] * b["cols"]):
                grp = list(g); random.shuffle(grp); n_groups.append(grp)
            n_ordered = [b for g in n_groups for b in g]

        result, n_ok, n_fail = ffd_pure(nn_ordered, n_ordered)

        # Compter les échecs non-Neutres (plus graves que les Neutres)
        _placed_noms = Counter(b["nom"].strip() for b in result)
        n_fail_non_neutre = sum(
            max(0, b_def["nombre"] - _placed_noms.get(b_def["nom"].strip(), 0))
            for b_def in buildings_def if b_def["type"] != "Neutre"
        )

        # Priorité absolue : 0 échecs non-Neutres. Les Neutres manquants
        # sont récupérés par le post-processing FFD garanti.
        if n_fail_non_neutre > 0:
            continue  # jamais acceptable

        s = quick_score(result)

        if (n_fail < best_n_failed or
                (n_fail == best_n_failed and s > best_screen_score)):
            best_placed       = result
            best_screen_score = s
            best_n_placed     = n_ok
            best_n_failed     = n_fail

    if best_placed is None:
        # Fallback : placement simple sans scoring
        result = [dict(b) for b in placed]
        n_ok = n_fail = 0
        for b in sorted_base:
            occ = build_occ(result)
            ok = False
            for rows, cols in [(b["rows"], b["cols"]), (b["cols"], b["rows"])]:
                if ok: break
                for r in range(max_r):
                    if ok: break
                    for c in range(max_c):
                        if can_place(r, c, rows, cols, x_grid, occ, max_r, max_c):
                            result.append({**b, "r": r, "c": c, "rows": rows, "cols": cols})
                            n_ok += 1; ok = True; break
            if not ok:
                n_fail += 1
        best_placed, best_n_placed, best_n_failed = result, n_ok, n_fail

    # ── Post-processing : placer les batiments manquants en FFD pur ──
    # Apres le meilleur trial (qui peut avoir des echecs dus au scoring),
    # on recupere les batiments non places en FFD raster scan.
    # Cela garantit que tous les batiments sont places si le terrain le permet,
    # sans degrader la qualite du placement des Producteurs et Culturels.
    if best_n_failed > 0:
        from collections import Counter as _Counter
        placed_counts_final = _Counter(b["nom"].strip() for b in best_placed)
        n_recovered = n_still_fail = 0

        # Trier les manquants par taille decroissante pour eviter la fragmentation
        missing_to_place = []
        for b_def in buildings_def:
            nom = b_def["nom"].strip()
            missing = b_def["nombre"] - placed_counts_final.get(nom, 0)
            for _ in range(missing):
                missing_to_place.append(dict(b_def, nom=nom,
                                             rows=b_def["largeur"],
                                             cols=b_def["longueur"]))
        missing_to_place.sort(key=lambda b: -(b["rows"] * b["cols"]))

        for b in missing_to_place:
            occ = build_occ(best_placed)
            ok = False
            for rows, cols in [(b["rows"], b["cols"]), (b["cols"], b["rows"])]:
                if ok: break
                for r in range(max_r):
                    if ok: break
                    for c in range(max_c):
                        if can_place(r, c, rows, cols, x_grid, occ, max_r, max_c):
                            best_placed.append({**b, "r": r, "c": c,
                                                "rows": rows, "cols": cols})
                            n_recovered += 1; ok = True; break
            if not ok:
                n_still_fail += 1

        best_n_placed += n_recovered
        best_n_failed  = n_still_fail

    return best_placed, best_n_placed, best_n_failed




# ══════════════════════════════════════════════════════
# OPTIMISEUR
# ══════════════════════════════════════════════════════

def make_x_grid(terrain_grid, max_r, max_c):
    """
    Retourne une grille booléenne où True = case INVALIDE pour le placement.
    Une case est invalide si :
      - elle contient un 'X' (bord du terrain), OU
      - elle est None ET se trouve à l'EXTÉRIEUR du périmètre X
        (terrain non rectangulaire : zone hors de l'enceinte).
    L'extérieur est identifié par flood-fill depuis les bords de la grille.
    """
    from collections import deque

    # Flood-fill depuis tous les bords None pour identifier l'extérieur
    exterior = set()
    queue = deque()

    def _seed(r, c):
        if 0 <= r < max_r and 0 <= c < max_c and terrain_grid[r][c] != "X" and (r, c) not in exterior:
            exterior.add((r, c))
            queue.append((r, c))

    for r in range(max_r):
        _seed(r, 0)
        _seed(r, max_c - 1)
    for c in range(max_c):
        _seed(0, c)
        _seed(max_r - 1, c)

    while queue:
        r, c = queue.popleft()
        for dr, dc in ((-1, 0), (1, 0), (0, -1), (0, 1)):
            _seed(r + dr, c + dc)

    # Une case est invalide si X ou extérieure
    return [
        [terrain_grid[r][c] == "X" or (r, c) in exterior
         for c in range(max_c)]
        for r in range(max_r)
    ]


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


def _score_of(b, culturels):
    """Contribution d'un seul batiment producteur au score global."""
    if b["type"] != "Producteur" or b["priorite"] <= 0:
        return 0.0
    cult = culture_received(b, culturels)
    return boost_level(cult, b) * b["priorite"]


def _score_of_cached(b, cult_total_map):
    """
    Contribution d'un producteur au score, en utilisant un dict pre-calcule
    cult_total_map : {id(prod): culture_recue} pour eviter de rappeler culture_received.
    """
    if b["type"] != "Producteur" or b["priorite"] <= 0:
        return 0.0
    cult = cult_total_map.get(id(b), 0.0)
    return boost_level(cult, b) * b["priorite"]


def _build_cult_map(producteurs, culturels):
    """
    Pre-calcule la culture recue par chaque producteur depuis les culturels donnes.
    Retourne {id(prod): culture_recue}.
    """
    # Pre-calculer les zones de rayonnement de chaque culturel (cache)
    zones = {id(c): radiation_zone(c) for c in culturels}
    result = {}
    for p in producteurs:
        pc = cells_of(p)
        total = sum(c["culture"] for c in culturels if pc & zones[id(c)])
        result[id(p)] = total
    return result


def _best_position_for(b, placed, x_grid, max_r, max_c):
    """
    Cherche la meilleure position pour b via un score incremental rapide.
    Toutes les zones de rayonnement sont pre-calculees une seule fois.
    Les Neutres sont ignores immediatement.
    """
    if b["type"] == "Neutre":
        return score_placement(placed), None

    orig = (b["r"], b["c"], b["rows"], b["cols"])
    occ  = make_occ_grid(placed, max_r, max_c, exclude_ids={id(b)})
    culturels   = [p for p in placed if p["type"] == "Culturel"]
    producteurs = [p for p in placed if p["type"] == "Producteur"]

    # Culture actuelle de chaque producteur (pre-calculee avec cache de zones)
    cult_map = _build_cult_map(producteurs, culturels)
    orig_score = sum(_score_of_cached(p, cult_map) for p in producteurs)
    best_s, best_pos = orig_score, None

    if b["type"] == "Producteur":
        # Seule la contribution de b change quand on le deplace.
        # Les culturels ne bougent pas -> leurs zones restent identiques.
        # Pour la nouvelle position de b, recalculer sa culture recue.
        contrib_orig = _score_of_cached(b, cult_map)
        base = orig_score - contrib_orig
        # Zones des culturels pre-calculees
        cult_zones = {id(c): radiation_zone(c) for c in culturels}

        for rows, cols in {(b["rows"], b["cols"]), (b["cols"], b["rows"])}:
            for r in range(max_r):
                for c in range(max_c):
                    if (r, c, rows, cols) == orig:
                        continue
                    if can_place(r, c, rows, cols, x_grid, occ, max_r, max_c):
                        b["r"], b["c"], b["rows"], b["cols"] = r, c, rows, cols
                        pc_new = cells_of(b)
                        cult_new = sum(c2["culture"] for c2 in culturels
                                       if pc_new & cult_zones[id(c2)])
                        contrib_new = boost_level(cult_new, b) * b["priorite"] if b["priorite"] > 0 else 0
                        s = base + contrib_new
                        if s > best_s:
                            best_s, best_pos = s, (r, c, rows, cols)
                        b["r"], b["c"], b["rows"], b["cols"] = orig

    else:  # Culturel
        # b est un culturel. Son deplacement change la culture recue
        # uniquement par les producteurs dans sa zone actuelle ou nouvelle.
        culturels_autres = [p for p in culturels if p is not b]
        # Zones des autres culturels (fixes)
        autres_zones = {id(c): radiation_zone(c) for c in culturels_autres}

        # Cases de chaque producteur (fixes)
        prod_cells = {id(p): cells_of(p) for p in producteurs}

        # Culture de chaque producteur via les AUTRES culturels seulement
        cult_sans_b = {}
        for p in producteurs:
            pc = prod_cells[id(p)]
            cult_sans_b[id(p)] = sum(c["culture"] for c in culturels_autres
                                     if pc & autres_zones[id(c)])

        # Zone actuelle de b et producteurs qu'il couvre
        zone_orig = radiation_zone(b)
        affected_orig = [p for p in producteurs if prod_cells[id(p)] & zone_orig]
        affected_orig_ids = {id(p) for p in affected_orig}

        # Score de base = score sans b (producteurs recalcules avec cult_sans_b)
        base = sum(
            boost_level(cult_sans_b[id(p)], p) * p["priorite"]
            for p in producteurs
        )

        for rows, cols in {(b["rows"], b["cols"]), (b["cols"], b["rows"])}:
            for r in range(max_r):
                for c in range(max_c):
                    if (r, c, rows, cols) == orig:
                        continue
                    if not can_place(r, c, rows, cols, x_grid, occ, max_r, max_c):
                        continue

                    b["r"], b["c"], b["rows"], b["cols"] = r, c, rows, cols
                    zone_new = radiation_zone(b)

                    # Producteurs touches par la nouvelle zone (anciens + nouveaux)
                    affected_new_ids = {id(p) for p in producteurs
                                        if prod_cells[id(p)] & zone_new}
                    all_affected_ids = affected_orig_ids | affected_new_ids
                    all_affected = [p for p in producteurs if id(p) in all_affected_ids]

                    # Pour ces producteurs, recalculer leur score avec b en nouvelle pos
                    delta = 0.0
                    b_cells = cells_of(b)  # deja mis a jour
                    # La zone de b en nouvelle pos = zone_new
                    for p in all_affected:
                        pc = prod_cells[id(p)]
                        # Culture = autres culturels + b si dans zone_new
                        extra = b["culture"] if pc & zone_new else 0.0
                        cult_new_p = cult_sans_b[id(p)] + extra
                        score_new_p = boost_level(cult_new_p, p) * p["priorite"]
                        # Score actuel sans b
                        score_old_p = boost_level(cult_sans_b[id(p)], p) * p["priorite"]
                        delta += score_new_p - score_old_p

                    s = base + delta
                    if s > best_s:
                        best_s, best_pos = s, (r, c, rows, cols)
                    b["r"], b["c"], b["rows"], b["cols"] = orig

    return best_s, best_pos


def _apply_move(b, pos, all_moves):
    all_moves.append({
        "nom":      b["nom"],
        "old_r":    b["r"],    "old_c":    b["c"],
        "old_rows": b["rows"], "old_cols": b["cols"],
        "new_r":    pos[0],    "new_c":    pos[1],
        "new_rows": pos[2],    "new_cols": pos[3],
    })
    b["r"], b["c"], b["rows"], b["cols"] = pos


def _culture_coverage(culturel, placed):
    """Nombre de producteurs couverts par ce culturel."""
    zone = radiation_zone(culturel)
    return sum(1 for b in placed if b["type"] == "Producteur" and cells_of(b) & zone)


def _dist(b1, b2):
    """Distance Manhattan entre centres de deux batiments."""
    r1 = b1["r"] + b1["rows"] / 2
    c1 = b1["c"] + b1["cols"] / 2
    r2 = b2["r"] + b2["rows"] / 2
    c2 = b2["c"] + b2["cols"] / 2
    return abs(r1 - r2) + abs(c1 - c2)


def optimize(placed, terrain_grid, max_r, max_c, n_passes=2, progress_cb=None):
    """
    Optimisation en 3 etapes :

    Etape 1 (greedy convergence) : chaque batiment cherche sa meilleure position
      globale en score direct. Ordre de la liste = culturels puis producteurs,
      ce qui permet aux culturels de s'installer et aux producteurs de se grouper.
      Repete n_passes fois jusqu'a convergence.

    Etape 2 (rapprochement agressif des culturels inutiles) : les culturels qui
      ne couvrent aucun producteur sont deplaces AU PLUS PRES des producteurs
      les moins bien alimentes, meme si le gain de score est nul.
      L'idee : rompre l'optimum local pour permettre a l'etape 3 d'aller plus loin.

    Etape 3 (reconvergence finale) : une nouvelle passe greedy complete apres
      les deplacements forcees de l'etape 2.

    Repete les etapes 2+3 jusqu'a ce qu'il n'y ait plus de culturels inutiles
    ou qu'aucun progres ne soit possible.
    """
    x_grid = make_x_grid(terrain_grid, max_r, max_c)
    placed = [dict(b) for b in placed]
    all_moves = []

    n = len(placed)
    total_ops = n_passes * n * 2 + 5 * n
    op = [0]

    def tick(k=1):
        op[0] += k
        if progress_cb:
            progress_cb(min(op[0] / max(total_ops, 1), 0.98))

    def greedy_pass(max_inner=10):
        """
        Passe greedy : chaque batiment Culturel ou Producteur cherche sa meilleure
        position. Les Neutres sont ignores (ils ne contribuent pas au score).
        max_inner limite le nombre d'iterations internes pour eviter les boucles longues.
        """
        for _ in range(max_inner):
            improved = False
            # Trier par priorité décroissante : les casernes (prio=10) bougent en premier.
            # Ainsi les culturels se déplacent ensuite autour des casernes bien placées.
            sorted_placed = sorted(
                [b for b in placed if b["type"] != "Neutre"],
                key=lambda b: -b.get("priorite", 0)
            )
            for b in sorted_placed:
                best_s, best_pos = _best_position_for(b, placed, x_grid, max_r, max_c)
                if best_pos:
                    _apply_move(b, best_pos, all_moves)
                    improved = True
                tick()
            if not improved:
                break

    # ── Etape 1 ──
    for _ in range(n_passes):
        greedy_pass()

    # ── Etapes 2+3 : boucle de deblocage ──
    for _outer in range(n_passes + 1):
        culturels = [b for b in placed if b["type"] == "Culturel"]
        producteurs = [b for b in placed if b["type"] == "Producteur"]
        inutiles = [c for c in culturels if _culture_coverage(c, placed) == 0]

        if not inutiles:
            break

        # Pour chaque culturel inutile, trouver la position la plus proche
        # d'un producteur peu couvert qui soit libre, et l'y deplacer.
        prod_by_cult = sorted(producteurs, key=lambda p: culture_received(p, culturels))
        any_forced = False

        for cult in sorted(inutiles, key=lambda c: c["culture"], reverse=True):
            orig_cult = (cult["r"], cult["c"], cult["rows"], cult["cols"])
            occ = make_occ_grid(placed, max_r, max_c, exclude_ids={id(cult)})

            # Chercher d'abord s'il existe une position ameliorant le score
            # (sans contrainte de proximite)
            best_improve_s, best_improve_pos = score_placement(placed), None
            for rows, cols in {(cult["rows"], cult["cols"]), (cult["cols"], cult["rows"])}:
                for r in range(max_r):
                    for c in range(max_c):
                        if (r, c, rows, cols) == orig_cult: continue
                        if not can_place(r, c, rows, cols, x_grid, occ, max_r, max_c): continue
                        cult["r"], cult["c"], cult["rows"], cult["cols"] = r, c, rows, cols
                        s = score_placement(placed)
                        if s > best_improve_s:
                            best_improve_s, best_improve_pos = s, (r, c, rows, cols)
                        cult["r"], cult["c"], cult["rows"], cult["cols"] = orig_cult

            if best_improve_pos:
                # Il existe une position qui ameliore le score : on l'applique
                _apply_move(cult, best_improve_pos, all_moves)
                any_forced = True
                continue

            # Sinon : deplacement force vers le producteur le moins bien alimente
            # On choisit la position libre la plus proche de ce producteur
            best_dist = float("inf")
            best_forced_pos = None
            target_prod = prod_by_cult[0]  # producteur avec le moins de culture

            for rows, cols in {(cult["rows"], cult["cols"]), (cult["cols"], cult["rows"])}:
                for r in range(max_r):
                    for c in range(max_c):
                        if (r, c, rows, cols) == orig_cult: continue
                        if not can_place(r, c, rows, cols, x_grid, occ, max_r, max_c): continue
                        # Distance entre le culturel (si place en r,c) et le producteur cible
                        cult["r"], cult["c"], cult["rows"], cult["cols"] = r, c, rows, cols
                        # Verifier que le rayonnement couvrirait le producteur
                        zone = radiation_zone(cult)
                        prod_cells = cells_of(target_prod)
                        if zone & prod_cells:
                            # Position qui couvre directement -> priorite absolue
                            d = -1
                        else:
                            d = _dist(cult, target_prod)
                        if d < best_dist:
                            best_dist, best_forced_pos = d, (r, c, rows, cols)
                        cult["r"], cult["c"], cult["rows"], cult["cols"] = orig_cult

            if best_forced_pos and best_forced_pos != orig_cult:
                _apply_move(cult, best_forced_pos, all_moves)
                any_forced = True
            tick()

        if not any_forced:
            break

        # ── Etape 3 : reconvergence apres deplacements forces ──
        greedy_pass()

    # ── Etape finale : swaps de paires de Culturels ──
    # L'optimiseur greedy peut etre bloque dans un optimum local ou
    # aucun deplacement individuel n'ameliore le score.
    # Echanger simultanement deux Culturels peut debloquer la situation :
    # chacun occupe la position de l'autre, ce qui peut mieux couvrir
    # les producteurs dans les deux zones.
    culturels_list = [b for b in placed if b["type"] == "Culturel"]
    producteurs_list = [b for b in placed if b["type"] == "Producteur"]
    prod_cells_cache = {id(p): cells_of(p) for p in producteurs_list}

    swap_improved = True
    n_swap_passes = 0
    max_swap_passes = 1  # 1 passe suffit pour un bon gain sans exploser le temps
    while swap_improved and n_swap_passes < max_swap_passes:
        swap_improved = False
        for i, c1 in enumerate(culturels_list):
            for c2 in culturels_list[i+1:]:
                zone1 = radiation_zone(c1)
                zone2 = radiation_zone(c2)
                # Producteurs dont la couverture changera
                affected = [p for p in producteurs_list
                            if prod_cells_cache[id(p)] & zone1
                            or prod_cells_cache[id(p)] & zone2]
                if not affected:
                    continue

                score_before = sum(
                    boost_level(culture_received(p, culturels_list), p) * p["priorite"]
                    for p in affected
                )

                # Tester le swap : c1 prend la position de c2 et vice versa.
                # Chaque bâtiment CONSERVE ses propres dimensions (rows, cols) :
                # les dimensions sont des propriétés fixes du bâtiment dans le jeu.
                # On teste aussi les pivots individuels si le bâtiment n'est pas carré.
                old1 = (c1["r"], c1["c"], c1["rows"], c1["cols"])
                old2 = (c2["r"], c2["c"], c2["rows"], c2["cols"])
                best_delta = 0.0
                best_config = None

                # Générer les variantes: chaque bâtiment peut rester dans son orientation
                # ou pivoter (seulement si rows != cols, i.e. pas carré)
                c1_orientations = [(c1["rows"], c1["cols"])]
                if c1["rows"] != c1["cols"]:
                    c1_orientations.append((c1["cols"], c1["rows"]))
                c2_orientations = [(c2["rows"], c2["cols"])]
                if c2["rows"] != c2["cols"]:
                    c2_orientations.append((c2["cols"], c2["rows"]))

                occ_without = make_occ_grid(placed, max_r, max_c,
                                            exclude_ids={id(c1), id(c2)})

                for r1_new, co1_new in c1_orientations:
                    for r2_new, co2_new in c2_orientations:
                        # c1 va à la position de c2 (avec son orientation r1_new x co1_new)
                        # c2 va à la position de c1 (avec son orientation r2_new x co2_new)
                        new_r1, new_c1 = old2[0], old2[1]  # position de c2
                        new_r2, new_c2 = old1[0], old1[1]  # position de c1

                        # Vérifier que les dimensions tiennent dans les cases disponibles
                        if not can_place(new_r1, new_c1, r1_new, co1_new,
                                         x_grid, occ_without, max_r, max_c):
                            continue
                        if not can_place(new_r2, new_c2, r2_new, co2_new,
                                         x_grid, occ_without, max_r, max_c):
                            continue

                        new1 = (new_r1, new_c1, r1_new, co1_new)
                        new2 = (new_r2, new_c2, r2_new, co2_new)

                        c1["r"], c1["c"], c1["rows"], c1["cols"] = new1
                        c2["r"], c2["c"], c2["rows"], c2["cols"] = new2

                        score_after = sum(
                            boost_level(culture_received(p, culturels_list), p) * p["priorite"]
                            for p in affected
                        )
                        delta = score_after - score_before
                        if delta > best_delta:
                            best_delta = delta
                            best_config = (new1, new2)

                        c1["r"], c1["c"], c1["rows"], c1["cols"] = old1
                        c2["r"], c2["c"], c2["rows"], c2["cols"] = old2

                if best_config:
                    c1["r"], c1["c"], c1["rows"], c1["cols"] = best_config[0]
                    c2["r"], c2["c"], c2["rows"], c2["cols"] = best_config[1]
                    all_moves.append({
                        "nom": c1["nom"], "old_r": old1[0], "old_c": old1[1],
                        "old_rows": old1[2], "old_cols": old1[3],
                        "new_r": best_config[0][0], "new_c": best_config[0][1],
                        "new_rows": best_config[0][2], "new_cols": best_config[0][3],
                    })
                    all_moves.append({
                        "nom": c2["nom"], "old_r": old2[0], "old_c": old2[1],
                        "old_rows": old2[2], "old_cols": old2[3],
                        "new_r": best_config[1][0], "new_c": best_config[1][1],
                        "new_rows": best_config[1][2], "new_cols": best_config[1][3],
                    })
                    swap_improved = True
        # Apres une passe de swaps, relancer la greedy pour consolider
        if swap_improved:
            greedy_pass()
        n_swap_passes += 1

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
        prio = b["priorite"]
        score = boost * prio if b["type"] == "Producteur" else ""
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

    # ── Section 1 : Score global et boosts par type de batiment ──
    score_avant = score_placement(original_placed)
    score_apres = score_placement(optimized)

    # En-tete section 1
    titre1 = ws2.cell(1, 1, "Boosts par type de batiment producteur")
    titre1.font = Font(bold=True, size=12)
    ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    titre1.alignment = Alignment(horizontal="center")
    titre1.fill = mfill("FF1F4E79")
    titre1.font = Font(bold=True, size=12, color=C_WHITE)

    hdrs_boost = ["Type de batiment", "Priorite",
                  "Avant : 0%", "Avant : 25%", "Avant : 50%", "Avant : 100%",
                  "Apres : 0%", "Apres : 25%", "Apres : 50%", "Apres : 100%"]
    widths_boost = [32, 10, 11, 11, 11, 12, 11, 11, 11, 12]
    for ci, (h, w) in enumerate(zip(hdrs_boost, widths_boost), 1):
        style_header(ws2.cell(2, ci), h)
        ws2.column_dimensions[get_column_letter(ci)].width = w

    # Calculer boosts avant/apres par nom de batiment unique
    def boost_counts(placed_list, cult_list):
        """Retourne dict {nom: {0:n, 25:n, 50:n, 100:n, priorite:p}}"""
        counts = {}
        for b in placed_list:
            if b["type"] != "Producteur":
                continue
            nom = b["nom"]
            cult = culture_received(b, cult_list)
            boost = boost_level(cult, b)
            if nom not in counts:
                counts[nom] = {0: 0, 25: 0, 50: 0, 100: 0, "priorite": b["priorite"]}
            counts[nom][boost] += 1
        return counts

    counts_avant = boost_counts(original_placed, orig_culturels)
    counts_apres = boost_counts(optimized, culturels)

    # Union de tous les noms de batiments producteurs
    all_prod_names = sorted(set(list(counts_avant.keys()) + list(counts_apres.keys())))

    C_BOOST0   = "FFFFD7D7"  # rouge pale  = pas de boost
    C_BOOST25  = "FFFFF2CC"  # jaune pale  = 25%
    C_BOOST50  = "FFD9EAD3"  # vert pale   = 50%
    C_BOOST100 = "FF93C47D"  # vert vif    = 100%
    boost_colors = {0: C_BOOST0, 25: C_BOOST25, 50: C_BOOST50, 100: C_BOOST100}

    row_i = 3
    for nom in all_prod_names:
        av = counts_avant.get(nom, {0: 0, 25: 0, 50: 0, 100: 0, "priorite": 99})
        ap = counts_apres.get(nom, {0: 0, 25: 0, 50: 0, 100: 0, "priorite": 99})
        prio = av.get("priorite") or ap.get("priorite") or 99

        vals = [nom, prio,
                av[0], av[25], av[50], av[100],
                ap[0], ap[25], ap[50], ap[100]]

        for ci, v in enumerate(vals, 1):
            cell = ws2.cell(row_i, ci, v)
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center", vertical="center")
            # Colorier les colonnes de boost
            if ci == 3:  cell.fill = mfill(C_BOOST0)
            elif ci == 4: cell.fill = mfill(C_BOOST25)
            elif ci == 5: cell.fill = mfill(C_BOOST50)
            elif ci == 6: cell.fill = mfill(C_BOOST100)
            elif ci == 7: cell.fill = mfill(C_BOOST0)
            elif ci == 8: cell.fill = mfill(C_BOOST25)
            elif ci == 9: cell.fill = mfill(C_BOOST50)
            elif ci == 10: cell.fill = mfill(C_BOOST100)
            # Mettre en gras les valeurs ameliorees
            if ci in (7, 8, 9, 10):
                boost_val = [0, 25, 50, 100][ci - 7]
                avant_val = av[boost_val]
                apres_val = ap[boost_val]
                if boost_val > 0 and apres_val > avant_val:
                    cell.font = Font(bold=True, color="FF006400")
                elif boost_val > 0 and apres_val < avant_val:
                    cell.font = Font(bold=True, color="FFCC0000")
        row_i += 1

    # Ligne de total
    total_row = row_i
    ws2.cell(total_row, 1, "TOTAL").font = Font(bold=True)
    ws2.cell(total_row, 1).fill = mfill("FFD9D9D9")
    ws2.cell(total_row, 1).border = thin_border()
    ws2.cell(total_row, 2).border = thin_border()
    for ci, col_boost in enumerate([0, 25, 50, 100], 3):
        # Avant
        total_av = sum(counts_avant.get(n, {col_boost: 0})[col_boost] for n in all_prod_names)
        cell = ws2.cell(total_row, ci, total_av)
        cell.font = Font(bold=True)
        cell.border = thin_border()
        cell.alignment = Alignment(horizontal="center")
        cell.fill = mfill(boost_colors[col_boost])
        # Apres
        total_ap = sum(counts_apres.get(n, {col_boost: 0})[col_boost] for n in all_prod_names)
        cell2 = ws2.cell(total_row, ci + 4, total_ap)
        cell2.font = Font(bold=True)
        cell2.border = thin_border()
        cell2.alignment = Alignment(horizontal="center")
        cell2.fill = mfill(boost_colors[col_boost])
    row_i = total_row + 2

    # ── Section 2 : Score global ──
    titre2 = ws2.cell(row_i, 1, "Score de boost de production")
    titre2.font = Font(bold=True, size=12, color=C_WHITE)
    titre2.fill = mfill("FF1F4E79")
    titre2.alignment = Alignment(horizontal="center")
    ws2.merge_cells(start_row=row_i, start_column=1, end_row=row_i, end_column=4)
    row_i += 1

    for ci, h in enumerate(["", "Score avant", "Score apres", "Gain"], 1):
        cell = ws2.cell(row_i, ci, h)
        if h:
            style_header(cell, h)
        ws2.column_dimensions[get_column_letter(ci)].width = max(
            ws2.column_dimensions[get_column_letter(ci)].width, 18)
    row_i += 1

    delta_score = score_apres - score_avant
    ws2.cell(row_i, 1, "Score boost global").font = Font(bold=True)
    ws2.cell(row_i, 1).border = thin_border()
    ws2.cell(row_i, 1).fill = mfill("FFD9D9D9")

    cell_av = ws2.cell(row_i, 2, round(score_avant, 2))
    cell_av.font = Font(bold=True)
    cell_av.border = thin_border()
    cell_av.alignment = Alignment(horizontal="center")

    cell_ap = ws2.cell(row_i, 3, round(score_apres, 2))
    cell_ap.font = Font(bold=True)
    cell_ap.border = thin_border()
    cell_ap.alignment = Alignment(horizontal="center")

    cell_gain = ws2.cell(row_i, 4, round(delta_score, 2))
    cell_gain.font = Font(bold=True,
                          color=C_GAIN if delta_score >= 0 else C_LOSS)
    cell_gain.border = thin_border()
    cell_gain.alignment = Alignment(horizontal="center")
    row_i += 2

    # ── Section 3 : Production par type ──
    titre3 = ws2.cell(row_i, 1, "Production par type")
    titre3.font = Font(bold=True, size=12, color=C_WHITE)
    titre3.fill = mfill("FF1F4E79")
    titre3.alignment = Alignment(horizontal="center")
    ws2.merge_cells(start_row=row_i, start_column=1, end_row=row_i, end_column=6)
    row_i += 1

    hdrs3b = ["Production", "Culture totale recue", "Boost max",
              "Qte/h initiale", "Qte/h optimisee", "Gain/perte Qte/h"]
    widths3b = [22, 22, 12, 16, 16, 18]
    for ci, (h, w) in enumerate(zip(hdrs3b, widths3b), 1):
        style_header(ws2.cell(row_i, ci), h)
        ws2.column_dimensions[get_column_letter(ci)].width = max(
            ws2.column_dimensions[get_column_letter(ci)].width, w)
    row_i += 1

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
        # Construire final_state depuis optimized
        final_state = {}
        for b in optimized:
            final_state.setdefault(b["nom"], []).append(
                [b["r"], b["c"], b["rows"], b["cols"]])

        # Construire current_state depuis original_placed
        current_state = {}
        for b in original_placed:
            current_state.setdefault(b["nom"], []).append(
                [b["r"], b["c"], b["rows"], b["cols"]])

        def cells_of_pos(r, c, rows, cols):
            return {(r+dr, c+dc) for dr in range(rows) for dc in range(cols)}

        def build_grid(state, exclude_nom=None, exclude_idx=None):
            g = {}
            for nom2, positions in state.items():
                for i2, pos2 in enumerate(positions):
                    if nom2 == exclude_nom and i2 == exclude_idx:
                        continue
                    r2, c2, rows2, cols2 = pos2
                    for dr in range(rows2):
                        for dc in range(cols2):
                            g[(r2+dr, c2+dc)] = (nom2, i2)
            return g

        # Preparer les infos de chaque move avec dims finales.
        # L'appariement entre real_moves et final_state se fait par POSITION FINALE
        # (new_r, new_c), pas par index, pour éviter les échanges de dimensions
        # entre bâtiments qui se déplacent vers/depuis les mêmes cases.
        move_info = []
        for mv in real_moves:
            nom = mv["nom"]
            # Trouver l'index dans current_state qui correspond à la position initiale
            mv_idx = 0
            for i, pos in enumerate(current_state.get(nom, [])):
                if pos[0] == mv["old_r"] and pos[1] == mv["old_c"]:
                    mv_idx = i
                    break
            # Trouver les dimensions finales en cherchant par position finale (new_r, new_c)
            # plutôt que par index mv_idx (qui peut pointer vers le mauvais bâtiment)
            fin_dims_list = final_state.get(nom, [])
            fin_rows, fin_cols = None, None
            for fd in fin_dims_list:
                if fd[0] == mv["new_r"] and fd[1] == mv["new_c"]:
                    fin_rows, fin_cols = fd[2], fd[3]
                    break
            if fin_rows is None:
                # Fallback: garder les dimensions originales (pas de pivot)
                cs = current_state.get(nom, [[0, 0, 1, 1]])
                fin_rows = cs[mv_idx][2] if mv_idx < len(cs) else 1
                fin_cols = cs[mv_idx][3] if mv_idx < len(cs) else 1
            move_info.append({
                "nom": nom, "mv_idx": mv_idx,
                "old_r": mv["old_r"], "old_c": mv["old_c"],
                "new_r": mv["new_r"], "new_c": mv["new_c"],
                "fin_rows": fin_rows, "fin_cols": fin_cols,
            })

        # Tri topologique : un move j doit preceder le move i si le batiment j
        # (dans son etat initial) occupe une case que le batiment i veut occuper.
        n_mv = len(move_info)
        init_grid = build_grid(current_state)

        predecesseurs = [set() for _ in range(n_mv)]
        for i, mi in enumerate(move_info):
            dest = cells_of_pos(mi["new_r"], mi["new_c"], mi["fin_rows"], mi["fin_cols"])
            for cell in dest:
                if cell in init_grid:
                    bnom, bidx = init_grid[cell]
                    if bnom == mi["nom"] and bidx == mi["mv_idx"]:
                        continue
                    for j, mj in enumerate(move_info):
                        if j != i and mj["nom"] == bnom and mj["mv_idx"] == bidx:
                            predecesseurs[i].add(j)
                            break

        successeurs = [[] for _ in range(n_mv)]
        for i in range(n_mv):
            for j in predecesseurs[i]:
                successeurs[j].append(i)

        in_degree = [len(predecesseurs[i]) for i in range(n_mv)]
        from collections import deque
        queue = deque(i for i in range(n_mv) if in_degree[i] == 0)
        ordered = []
        while queue:
            i = queue.popleft()
            ordered.append(i)
            for s in successeurs[i]:
                in_degree[s] -= 1
                if in_degree[s] == 0:
                    queue.append(s)
        # Ajouter les restants (cycles eventuels) tels quels
        ordered += [i for i in range(n_mv) if i not in ordered]

        # Ecrire les steps dans l'ordre topologique
        step = 1
        for i in ordered:
            mi = move_info[i]
            nom = mi["nom"]
            new_r, new_c = mi["new_r"], mi["new_c"]
            fin_rows, fin_cols = mi["fin_rows"], mi["fin_cols"]

            pos = current_state[nom][mi["mv_idx"]]
            cur_r, cur_c, cur_rows, cur_cols = pos[0], pos[1], pos[2], pos[3]
            cur_str = f"L{cur_r+1} C{cur_c+1}"
            new_str = f"L{new_r+1} C{new_c+1}"
            old_str = f"L{mi['old_r']+1} C{mi['old_c']+1}"

            # Detecter un changement d'orientation
            pivot = (cur_rows != fin_rows or cur_cols != fin_cols)
            if pivot:
                orient_avant = "horizontal" if cur_cols >= cur_rows else "vertical"
                orient_apres = "horizontal" if fin_cols >= fin_rows else "vertical"
                pivot_note = f" (pivoter de {orient_avant} vers {orient_apres} : {cur_rows}x{cur_cols} -> {fin_rows}x{fin_cols})"
            else:
                pivot_note = ""

            # Verifier bloqueurs residuels (cycles)
            grid_now = build_grid(current_state, exclude_nom=nom, exclude_idx=mi["mv_idx"])
            dest_cells = cells_of_pos(new_r, new_c, fin_rows, fin_cols)
            blockers_now = {}
            for cell in dest_cells:
                if cell in grid_now:
                    bnom2, bidx2 = grid_now[cell]
                    key = (bnom2, bidx2)
                    if key not in blockers_now:
                        br2, bc2 = current_state[bnom2][bidx2][0], current_state[bnom2][bidx2][1]
                        blockers_now[key] = (br2, bc2)

            if blockers_now:
                blocker_lines = []
                for (bnom2, bidx2), (br2, bc2) in blockers_now.items():
                    fin_b = final_state.get(bnom2, [])
                    if bidx2 < len(fin_b):
                        fr2, fc2 = fin_b[bidx2][0], fin_b[bidx2][1]
                        final_str = f"L{fr2+1} C{fc2+1}"
                    else:
                        final_str = "inconnue"
                    blocker_lines.append(f"{bnom2} (L{br2+1} C{bc2+1} -> {final_str})")
                action = (
                    f"1) Deplacer d'abord : {'; '.join(blocker_lines)}. "
                    f"2) Deplacer '{nom}' de {cur_str} vers {new_str}{pivot_note}."
                )
            else:
                action = f"Deplacer '{nom}' de {cur_str} vers {new_str}{pivot_note}."

            current_state[nom][mi["mv_idx"]] = [new_r, new_c, fin_rows, fin_cols]

            ri = step + 1
            ws3.cell(ri, 1, step)
            ws3.cell(ri, 2, nom)
            ws3.cell(ri, 3, old_str)
            ws3.cell(ri, 4, new_str)
            ws3.cell(ri, 5, action)
            ws3.cell(ri, 5).alignment = Alignment(wrap_text=True, vertical="top")
            ws3.row_dimensions[ri].height = 60
            for ci in range(1, 6):
                ws3.cell(ri, ci).border = thin_border()
                if ci < 5:
                    ws3.cell(ri, ci).alignment = Alignment(horizontal="center", vertical="top")
            step += 1


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
            # Ignorer les cellules secondaires d'une fusion (MergedCell = lecture seule)
            from openpyxl.cell.cell import MergedCell
            if isinstance(cell, MergedCell):
                continue
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

# Initialisation du session_state : les resultats persistent entre les reruns
if "result_excel" not in st.session_state:
    st.session_state.result_excel    = None   # bytes du fichier Excel
    st.session_state.score_init      = None
    st.session_state.score_opt       = None
    st.session_state.moved_summary   = []     # liste de chaines a afficher
    st.session_state.last_filename   = None   # nom du fichier charge

uploaded = st.file_uploader(
    "Choisissez votre fichier Excel de ville (.xlsx)",
    type=["xlsx"],
    help="Le fichier doit contenir un onglet Terrain et un onglet Batiments."
)

# Si un nouveau fichier est charge, on efface les resultats precedents
if uploaded:
    if uploaded.name != st.session_state.last_filename:
        st.session_state.result_excel  = None
        st.session_state.score_init    = None
        st.session_state.score_opt     = None
        st.session_state.moved_summary = []
        st.session_state.last_filename = uploaded.name

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

    n_culturels   = sum(1 for b in placed if b["type"] == "Culturel")
    n_producteurs = sum(1 for b in placed if b["type"] == "Producteur")
    n_neutres     = sum(1 for b in placed if b["type"] == "Neutre")
    score_init    = score_placement(placed)

    st.success(f"Fichier charge : **{len(placed)} batiments** sur un terrain **{max_r} x {max_c}**")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Score initial", f"{score_init:.2f}")
    col2.metric("Batiments culturels", n_culturels)
    col3.metric("Batiments producteurs", n_producteurs)
    col4.metric("Batiments neutres", n_neutres)

    with st.expander("Detail de la culture initiale par producteur"):
        culturels_init = [b for b in placed if b["type"] == "Culturel"]
        for b in sorted(placed, key=lambda x: x["nom"]):
            if b["type"] == "Producteur":
                cult = culture_received(b, culturels_init)
                boost = boost_level(cult, b)
                st.write(
                    f"**{b['nom']}** - Culture recue : {cult:.0f} / "
                    f"Seuil 25% : {b['boost25']:.0f} | 50% : {b['boost50']:.0f} | 100% : {b['boost100']:.0f} "
                    f"→ **Boost : {boost}%**"
                )

    st.divider()

    # ── Batiments manquants ──
    placed_counts = {}
    for b in placed:
        placed_counts[b["nom"].strip()] = placed_counts.get(b["nom"].strip(), 0) + 1
    missing_list = []
    for b_def in buildings_def:
        nom = b_def["nom"].strip()
        needed = b_def["nombre"] - placed_counts.get(nom, 0)
        if needed > 0:
            missing_list.append(f"**{nom}** : {needed} a placer")

    if missing_list:
        with st.expander(f"⚠️ {len(missing_list)} type(s) de batiments non encore places sur le terrain"):
            for m in missing_list:
                st.write(m)
        do_place_missing = st.checkbox(
            "Placer automatiquement les batiments manquants avant d'optimiser",
            value=True,
            help="Les batiments manquants seront places au mieux sur les cases libres, puis l'optimisation s'executera sur l'ensemble."
        )
    else:
        do_place_missing = False

    n_passes = st.slider(
        "Nombre de passes d'optimisation",
        min_value=1, max_value=5, value=2,
        help="Plus de passes = meilleur resultat mais plus long. 2 passes est un bon compromis."
    )

    if st.button("Lancer l'optimisation", type="primary"):
        # Effacer les resultats precedents avant de relancer
        st.session_state.result_excel  = None
        st.session_state.score_opt     = None
        st.session_state.moved_summary = []

        progress_bar = st.progress(0)
        status = st.empty()

        # Placer les batiments manquants si demande
        placed_for_optim = placed
        n_placed_new = 0
        n_failed_new = 0
        if do_place_missing and missing_list:
            status.info("Placement des batiments manquants (plusieurs essais)...")
            n_trials = min(3, n_passes)
            placed_for_optim, n_placed_new, n_failed_new = place_missing_buildings(
                placed, buildings_def, terrain_grid, max_r, max_c, n_trials=n_trials
            )
            if n_failed_new > 0:
                # Identifier si les echecs sont des Neutres uniquement
                from collections import Counter as _Cnt
                placed_names = _Cnt(b["nom"].strip() for b in placed_for_optim)
                fail_types = _Cnt()
                fail_names = []
                for b_def in buildings_def:
                    nom = b_def["nom"].strip()
                    short = b_def["nombre"] - placed_names.get(nom, 0)
                    if short > 0:
                        fail_types[b_def["type"]] += short
                        fail_names.append(f"{nom} ({short}×)")
                if fail_types.get("Neutre", 0) == n_failed_new:
                    # Tous les echecs sont des Neutres -> explication adaptee
                    st.warning(
                        f"⚠️ **{n_failed_new} bâtiment(s) Neutre(s) n'ont pas pu être placés** "
                        f"(terrain trop fragmenté après placement des Producteurs et Culturels). "
                        f"**Cela n'affecte pas le score** : les Neutres ne produisent ni culture "
                        f"ni ressources. Bâtiments concernés : {', '.join(fail_names)}. "
                        f"Vous pouvez les placer manuellement dans les espaces libres restants."
                    )
                else:
                    st.warning(
                        f"⚠️ {n_failed_new} bâtiment(s) n'ont pas pu être placés "
                        f"(terrain plein). Types : {dict(fail_types)}."
                    )
            if n_placed_new > 0:
                st.success(f"✅ {n_placed_new} bâtiment(s) placés sur le terrain.")

        status.info("Optimisation en cours... Veuillez patienter.")

        def update_prog(v):
            progress_bar.progress(v)

        optimized, moves = optimize(
            placed_for_optim, terrain_grid, max_r, max_c,
            n_passes=n_passes, progress_cb=update_prog
        )
        progress_bar.progress(1.0)
        status.success("Optimisation terminee !")

        score_opt = score_placement(optimized)

        # Calculer la liste des deplacements reels
        # original_placed = etat avant placement des manquants + avant optimisation
        orig_map = {}
        for b in original_placed:
            orig_map.setdefault(b["nom"], []).append((b["r"], b["c"]))
        used = {n: 0 for n in orig_map}
        summary_lines = []
        for b in optimized:
            nom = b["nom"]
            if nom in orig_map:
                idx = used[nom]
                if idx < len(orig_map[nom]):
                    used[nom] += 1
                    op = orig_map[nom][idx]
                    if op[0] != b["r"] or op[1] != b["c"]:
                        cult_val = culture_received(b, [x for x in optimized if x["type"] == "Culturel"])
                        boost    = boost_level(cult_val, b)
                        icon = "🟠" if b["type"] == "Culturel" else "🟢" if b["type"] == "Producteur" else "⬜"
                        line = (
                            f"{icon} **{nom}** : "
                            f"L{op[0]+1} C{op[1]+1} → L{b['r']+1} C{b['c']+1}"
                            + (f" | Boost apres : **{boost}%**" if b["type"] == "Producteur" else "")
                        )
                        summary_lines.append(line)

        # Generer le fichier Excel et stocker dans session_state
        # Pour l'onglet Deplacements, on compare toujours l'etat du fichier INPUT
        # (original_placed) avec l'etat optimise final.
        # Si le terrain etait (partiellement) vide, les nouveaux batiments places
        # n'ont pas de "position initiale" -> ils n'apparaissent pas dans Deplacements.
        with st.spinner("Generation du fichier Excel..."):
            output_buf = build_excel_output(
                optimized, original_placed, terrain_grid, max_r, max_c, buildings_def
            )
            st.session_state.result_excel  = output_buf.getvalue()
            st.session_state.score_init    = score_placement(placed_for_optim)
            st.session_state.score_opt     = score_opt
            st.session_state.moved_summary = summary_lines

# ── Affichage des resultats (hors du bloc if uploaded pour persister) ──
if st.session_state.result_excel is not None:
    st.divider()
    delta = st.session_state.score_opt - st.session_state.score_init
    c1, c2, c3 = st.columns(3)
    c1.metric("Score initial",  f"{st.session_state.score_init:.2f}")
    c2.metric("Score optimise", f"{st.session_state.score_opt:.2f}", delta=f"{delta:+.2f}")
    c3.metric("Batiments deplaces", len(st.session_state.moved_summary))

    if st.session_state.moved_summary:
        st.subheader("Batiments deplaces")
        for line in st.session_state.moved_summary:
            st.write(line)
    else:
        st.info("Le placement initial est deja optimal.")

    st.divider()
    st.download_button(
        label="⬇️  Telecharger le fichier resultat Excel",
        data=st.session_state.result_excel,
        file_name="ville_optimisee.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption(
        "Le fichier contient 4 onglets : "
        "**Liste batiments**, **Synthese**, **Deplacements**, **Terrain optimise**."
    )
