import streamlit as st
import numpy as np
import random
import io
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Optimiseur de batiments",
    page_icon="🏛️",
    layout="wide"
)

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@400;700&family=Lato:wght@300;400;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Lato', sans-serif;
}

.main { background: #0f1117; }

h1, h2, h3 {
    font-family: 'Cinzel', serif !important;
    color: #c9a84c !important;
    letter-spacing: 2px;
}

.hero {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    border: 1px solid #c9a84c44;
    border-radius: 12px;
    padding: 2.5rem;
    margin-bottom: 2rem;
    text-align: center;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: '';
    position: absolute;
    top: -50%; left: -50%;
    width: 200%; height: 200%;
    background: radial-gradient(ellipse at center, #c9a84c11 0%, transparent 60%);
    pointer-events: none;
}
.hero h1 { font-size: 2.2rem; margin: 0; text-shadow: 0 0 30px #c9a84c66; }
.hero p  { color: #8899aa; margin: 0.5rem 0 0; font-size: 1.05rem; }

.stat-card {
    background: #1a1a2e;
    border: 1px solid #c9a84c33;
    border-radius: 8px;
    padding: 1.2rem;
    text-align: center;
}
.stat-card .value {
    font-family: 'Cinzel', serif;
    font-size: 2rem;
    color: #c9a84c;
    line-height: 1;
}
.stat-card .label {
    color: #8899aa;
    font-size: 0.8rem;
    margin-top: 0.3rem;
    text-transform: uppercase;
    letter-spacing: 1px;
}

.boost-bar-bg {
    background: #1a1a2e;
    border-radius: 99px;
    height: 22px;
    width: 100%;
    overflow: hidden;
    border: 1px solid #ffffff11;
}
.boost-bar-fill {
    height: 100%;
    border-radius: 99px;
    display: flex;
    align-items: center;
    justify-content: flex-end;
    padding-right: 8px;
    font-size: 0.75rem;
    font-weight: 700;
    color: #0f1117;
    transition: width 0.8s ease;
}

.step-badge {
    background: #c9a84c;
    color: #0f1117;
    border-radius: 50%;
    width: 28px; height: 28px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    font-weight: 700;
    font-family: 'Cinzel', serif;
    margin-right: 8px;
    font-size: 0.9rem;
}

.warning-box {
    background: #2d1b00;
    border: 1px solid #c9a84c66;
    border-radius: 8px;
    padding: 1rem;
    color: #c9a84c;
}
.success-box {
    background: #001a0d;
    border: 1px solid #2ecc7166;
    border-radius: 8px;
    padding: 1rem;
    color: #2ecc71;
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Fonctions core (identiques au script CLI)
# ---------------------------------------------------------------------------

def load_terrain(wb):
    ws = wb["Terrain"]
    rows = []
    for row in ws.iter_rows(values_only=True):
        line = [int(c) if c is not None else 0 for c in row if c is not None]
        if line:
            rows.append(line)
    max_len = max(len(r) for r in rows)
    rows = [r + [0] * (max_len - len(r)) for r in rows]
    terrain = np.array(rows, dtype=int)
    # Nouvelle convention : 1=libre, 0=occupe -> inversion pour usage interne (0=libre, 1=occupe)
    return 1 - terrain


def load_buildings(wb):
    ws = wb["Batiments"]
    headers = None
    buildings = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            headers = [str(c).strip() if c else f"col{j}" for j, c in enumerate(row)]
            continue
        if not any(row):
            continue
        d = dict(zip(headers, row))
        b = {
            "nom":         str(d.get("Nom", f"B{i}")).strip(),
            "longueur":    int(d.get("Longueur") or 1),
            "largeur":     int(d.get("Largeur") or 1),
            "quantite":    int(next((d[k] for k in d if k.lower().replace("é","e").replace("è","e") == "quantite" and d[k]), 1)),
            "culture":     float(d.get("Culture") or 0),
            "rayonnement": int(d.get("Rayonnement") or 0),
            "boost_25":    float(d.get("Boost 25%") or 0),
            "boost_50":    float(d.get("Boost 50%") or 0),
            "boost_100":   float(d.get("Boost 100%") or 0),
        }
        buildings.append(b)
    return buildings


def get_orientations(template):
    h, w = template["largeur"], template["longueur"]
    orients = [("H", h, w)]
    if h != w:
        orients.append(("V", w, h))
    return orients


def corner_score(grid, r, c, h, w):
    rows, cols = grid.shape
    neighbors = set()
    for dr in range(h):
        for dc in range(w):
            for nr, nc in [(r+dr-1,c+dc),(r+dr+1,c+dc),(r+dr,c+dc-1),(r+dr,c+dc+1)]:
                if 0 <= nr < rows and 0 <= nc < cols and grid[nr,nc] == 0:
                    neighbors.add((nr, nc))
    free_adj = len(neighbors)
    border_bonus = 0
    if r == 0 or r + h == rows: border_bonus += h
    if c == 0 or c + w == cols: border_bonus += w
    return free_adj - border_bonus * 2


class FastPlacement:
    def __init__(self, terrain):
        self.terrain     = terrain.copy()
        self.shape       = terrain.shape
        self.grid        = terrain.copy()
        self.instances   = []
        self.culture_map = np.zeros(self.shape, dtype=float)

    def place(self, template, row, col, orientation, h, w):
        self.grid[row:row+h, col:col+w] = 2
        self.instances.append({
            "template": template, "row": row, "col": col,
            "orientation": orientation, "h": h, "w": w,
        })
        if template["culture"] > 0 and template["rayonnement"] > 0:
            R  = template["rayonnement"]
            r0 = max(0, row - R);  r1 = min(self.shape[0], row + h + R)
            c0 = max(0, col - R);  c1 = min(self.shape[1], col + w + R)
            self.culture_map[r0:r1, c0:c1] += template["culture"]
            self.culture_map[row:row+h, col:col+w] -= template["culture"]

    def culture_received(self, row, col, h, w, boostable=True):
        # Seuls les batiments boostables (boost_25 > 0) peuvent recevoir de la culture
        if not boostable:
            return 0.0
        # Culture recue = valeur maximale parmi toutes les cases du batiment
        # (un 2x2 partiellement dans un rayon recoit autant qu'un 1x1 bien centré)
        return float(self.culture_map[row:row+h, col:col+w].max())

    def boost_for(self, template, row, col, h, w):
        if template["boost_25"] <= 0 and template["boost_50"] <= 0 and template["boost_100"] <= 0:
            return 0
        received = self.culture_received(row, col, h, w, boostable=True)
        if template["boost_100"] > 0 and received >= template["boost_100"]: return 100
        if template["boost_50"]  > 0 and received >= template["boost_50"]:  return 50
        if template["boost_25"]  > 0 and received >= template["boost_25"]:  return 25
        return 0

    def total_score(self):
        score = 0.0
        for inst in self.instances:
            t = inst["template"]
            if t["boost_25"] <= 0 and t["boost_50"] <= 0 and t["boost_100"] <= 0:
                continue
            score += self.boost_for(t, inst["row"], inst["col"], inst["h"], inst["w"])
        return score

    def free_positions(self, h, w):
        positions = []
        for r in range(self.shape[0] - h + 1):
            for c in range(self.shape[1] - w + 1):
                if not np.any(self.grid[r:r+h, c:c+w]):
                    positions.append((r, c))
        return positions

    def best_emitter_position(self, template, sample=150, compact_weight=1.0):
        """
        Place un emetteur en maximisant le nombre de bâtiments boostables
        qui franchissent un nouveau seuil de boost grâce à cet emetteur.
        Priorité : débloquer un nouveau palier (0->25, 25->50, 50->100)
        plutôt que simplement couvrir des cases boostables.
        """
        orientations = get_orientations(template)
        best = None  # (score, r, c, label, h, w)
        culture_apportee = template["culture"]

        # Pré-calculer la culture ACTUELLE de chaque boostable (depuis culture_map à jour)
        # Cela tient compte des emetteurs déjà placés avant celui-ci
        boostable_instances = []
        for inst in self.instances:
            t = inst["template"]
            if t["boost_25"] > 0 or t["boost_50"] > 0 or t["boost_100"] > 0:
                r0, c0 = inst["row"], inst["col"]
                h0, w0 = inst["h"], inst["w"]
                current = float(self.culture_map[r0:r0+h0, c0:c0+w0].max())
                boostable_instances.append({
                    "row": r0, "col": c0, "h": h0, "w": w0,
                    "current": current,
                    "boost_25":  t["boost_25"],
                    "boost_50":  t["boost_50"],
                    "boost_100": t["boost_100"],
                })

        for (label, h, w) in orientations:
            positions = self.free_positions(h, w)
            if not positions:
                continue

            candidates = random.sample(positions, min(sample, len(positions)))

            for (r, c) in candidates:
                R = template["rayonnement"]
                r0 = max(0, r - R);  r1 = min(self.shape[0], r + h + R)
                c0 = max(0, c - R);  c1 = min(self.shape[1], c + w + R)

                score = 0.0
                for bi in boostable_instances:
                    # Ce boostable est-il dans le rayon de cet emetteur ?
                    # Vérifier si au moins une case du boostable est dans le rayon
                    in_range = False
                    for dr in range(bi["h"]):
                        for dc in range(bi["w"]):
                            if r0 <= bi["row"]+dr < r1 and c0 <= bi["col"]+dc < c1:
                                in_range = True
                                break
                        if in_range:
                            break
                    if not in_range:
                        continue

                    culture_avant = bi["current"]
                    culture_apres = culture_avant + culture_apportee

                    # Calculer le gain de palier
                    def palier(culture, b25, b50, b100):
                        if b100 > 0 and culture >= b100: return 100
                        if b50  > 0 and culture >= b50:  return 50
                        if b25  > 0 and culture >= b25:  return 25
                        return 0

                    boost_avant = palier(culture_avant, bi["boost_25"], bi["boost_50"], bi["boost_100"])
                    boost_apres = palier(culture_apres, bi["boost_25"], bi["boost_50"], bi["boost_100"])
                    gain = boost_apres - boost_avant

                    # Bonus fort si on débloque un nouveau palier
                    if gain > 0:
                        score += gain * 1000
                    else:
                        # Même sans débloquer un palier, rapprocher du prochain seuil a de la valeur
                        prochain_seuil = None
                        for seuil in sorted([bi["boost_25"], bi["boost_50"], bi["boost_100"]]):
                            if seuil > 0 and culture_avant < seuil:
                                prochain_seuil = seuil
                                break
                        if prochain_seuil:
                            # Plus on est proche du seuil, plus c'est utile
                            manque_avant = prochain_seuil - culture_avant
                            manque_apres = max(0, prochain_seuil - culture_apres)
                            score += (manque_avant - manque_apres) * 10

                compact = corner_score(self.grid, r, c, h, w)
                score -= compact * compact_weight

                if best is None or score > best[0]:
                    best = (score, r, c, label, h, w)

        if best is None:
            return None
        _, r, c, label, h, w = best
        return r, c, label, h, w

    def best_position_and_orientation(self, template, sample=150, compact_weight=1.0, boost_weight=10.0):
        needs_boost = (template["boost_25"] > 0 or template["boost_50"] > 0 or template["boost_100"] > 0)
        orientations = get_orientations(template)
        best = None
        for (label, h, w) in orientations:
            positions = self.free_positions(h, w)
            if not positions:
                continue
            candidates = positions if needs_boost else random.sample(positions, min(sample, len(positions)))
            for (r, c) in candidates:
                compact = corner_score(self.grid, r, c, h, w)
                culture = self.culture_received(r, c, h, w, boostable=needs_boost) if needs_boost else 0
                score   = culture * boost_weight - compact * compact_weight
                if best is None or score > best[0]:
                    best = (score, r, c, label, h, w)
        if best is None:
            return None
        _, r, c, label, h, w = best
        return r, c, label, h, w


def expand_templates(templates):
    result = []
    for t in templates:
        for _ in range(t["quantite"]):
            result.append(t)
    return result


def run_placement(terrain, templates, shuffle=False, compact_weight=1.0, boost_weight=10.0):
    instances = expand_templates(templates)

    # Séparer en 3 groupes :
    # 1. Boostables (ont besoin de culture) -> placés EN PREMIER pour choisir la meilleure position
    # 2. Emetteurs (génèrent de la culture)  -> placés AUTOUR des boostables déjà posés
    # 3. Neutres                             -> placés en dernier
    boostables = [b for b in instances if b["boost_25"] > 0 or b["boost_50"] > 0 or b["boost_100"] > 0]
    emetteurs  = [b for b in instances if b["culture"] > 0 and not (b["boost_25"] > 0 or b["boost_50"] > 0 or b["boost_100"] > 0)]
    neutres    = [b for b in instances if b["culture"] == 0 and b["boost_25"] == 0 and b["boost_50"] == 0 and b["boost_100"] == 0]

    if shuffle:
        grands_em = [e for e in emetteurs if e["rayonnement"] >= 3]
        petits_em = [e for e in emetteurs if e["rayonnement"] < 3]
        random.shuffle(grands_em); random.shuffle(petits_em)
        emetteurs = grands_em + petits_em
        random.shuffle(boostables)
        random.shuffle(neutres)
    else:
        boostables.sort(key=lambda b: -(b["largeur"] * b["longueur"]))
        emetteurs.sort(key=lambda b: (-b["rayonnement"], -(b["largeur"] * b["longueur"])))
        neutres.sort(key=lambda b: -(b["largeur"] * b["longueur"]))

    p = FastPlacement(terrain)
    unplaced = []

    # Étape 1 : placer les boostables (sans culture reçue pour l'instant, juste compacité)
    for tmpl in boostables:
        result = p.best_position_and_orientation(tmpl, sample=150,
                                                  compact_weight=compact_weight,
                                                  boost_weight=0.0)  # pas encore de culture
        if result is None:
            unplaced.append(tmpl["nom"])
        else:
            r, c, label, h, w = result
            p.place(tmpl, r, c, label, h, w)

    # Étape 2 : placer les émetteurs en maximisant la culture sur les boostables déjà posés
    for tmpl in emetteurs:
        result = p.best_emitter_position(tmpl, sample=150, compact_weight=compact_weight)
        if result is None:
            unplaced.append(tmpl["nom"])
        else:
            r, c, label, h, w = result
            p.place(tmpl, r, c, label, h, w)

    # Étape 3 : neutres
    for tmpl in neutres:
        result = p.best_position_and_orientation(tmpl, sample=150,
                                                  compact_weight=compact_weight,
                                                  boost_weight=0.0)
        if result is None:
            unplaced.append(tmpl["nom"])
        else:
            r, c, label, h, w = result
            p.place(tmpl, r, c, label, h, w)

    return p, unplaced


def local_improve(placement, terrain, n_iter=300):
    best_score = placement.total_score()
    for _ in range(n_iter):
        candidates = []
        for idx, inst in enumerate(placement.instances):
            t = inst["template"]
            if t["boost_25"] <= 0 and t["boost_50"] <= 0 and t["boost_100"] <= 0:
                continue
            b = placement.boost_for(t, inst["row"], inst["col"], inst["h"], inst["w"])
            if b < 100:
                candidates.append((idx, b))
        if not candidates:
            break
        candidates.sort(key=lambda x: x[1])
        pool = candidates[:max(1, len(candidates)//2)]
        idx, _ = random.choice(pool)
        inst = placement.instances[idx]
        t = inst["template"]
        new_p = FastPlacement(terrain)
        for j, other in enumerate(placement.instances):
            if j != idx:
                ot = other["template"]
                new_p.place(ot, other["row"], other["col"], other["orientation"], other["h"], other["w"])
        result = new_p.best_position_and_orientation(t, sample=999, compact_weight=0.5, boost_weight=15.0)
        if result is None:
            continue
        r, c, label, h, w = result
        new_p.place(t, r, c, label, h, w)
        new_score = new_p.total_score()
        if len(new_p.instances) == len(placement.instances) and new_score >= best_score:
            placement = new_p
            best_score = new_score
    return placement


def optimize(terrain, templates, n_restarts, local_iter, progress_cb=None):
    total = sum(t["quantite"] for t in templates)
    best_p        = None
    best_score    = -1.0
    best_placed   = -1
    best_unplaced = []
    for i in range(n_restarts):
        cw = random.uniform(0.5, 2.5) if i > 0 else 1.5
        bw = random.uniform(8.0, 20.0) if i > 0 else 12.0
        p, unplaced = run_placement(terrain, templates, shuffle=(i > 0),
                                    compact_weight=cw, boost_weight=bw)
        placed = len(p.instances)
        if placed == total:
            p = local_improve(p, terrain, n_iter=local_iter)
        s = p.total_score()
        if placed > best_placed or (placed == best_placed and s > best_score):
            best_score    = s
            best_placed   = placed
            best_p        = p
            best_unplaced = unplaced
        if progress_cb:
            progress_cb(i + 1, n_restarts, placed, total, s)
    return best_p, best_unplaced


# ---------------------------------------------------------------------------
# Export Excel
# ---------------------------------------------------------------------------

COLORS = [
    "FF6B6B","4ECDC4","45B7D1","96CEB4","FFEAA7","DDA0DD","98D8C8",
    "F7DC6F","BB8FCE","85C1E9","82E0AA","F0B27A","AED6F1","F9E79F",
    "D2B4DE","A9CCE3","A9DFBF","FAD7A0","F5CBA7","D5DBDB",
    "C0392B","1ABC9C","2980B9","27AE60","E67E22","8E44AD","2C3E50",
    "E74C3C","3498DB","F39C12","16A085","9B59B6","1F618D","117A65",
    "784212","1B4F72","4A235A","78281F","0B5345","6E2F1A","4D5656",
    "7D6608","7B241C","154360","0E6655","B03A2E","1A5276","1E8449",
    "B7770D","145A32","922B21","7F8C8D","2E4057","48C9B0","F8C471",
    "C39BD3","7DCEA0","F1948A","A3E4D7","D2B4DE","A9CCE3","F9E79F",
]


def build_excel(placement, templates):
    wb = Workbook()
    type_color = {t["nom"]: COLORS[i % len(COLORS)] for i, t in enumerate(templates)}
    shape = placement.shape

    # Carte
    ws_map = wb.active
    ws_map.title = "Carte"
    grid_labels = [[""] * shape[1] for _ in range(shape[0])]
    grid_colors = [[None]          * shape[1] for _ in range(shape[0])]
    for r in range(shape[0]):
        for c in range(shape[1]):
            if placement.terrain[r, c] == 1:
                grid_labels[r][c] = "X"
                grid_colors[r][c] = "555555"
    for inst in placement.instances:
        t      = inst["template"]
        r0, c0 = inst["row"], inst["col"]
        h, w   = inst["h"], inst["w"]
        color  = type_color.get(t["nom"], "CCCCCC")
        arrow  = "V" if inst["orientation"] == "V" else "H"
        label  = t["nom"][:5] + arrow
        for dr in range(h):
            for dc in range(w):
                grid_labels[r0+dr][c0+dc] = label if dr == 0 and dc == 0 else ""
                grid_colors[r0+dr][c0+dc] = color
    thin   = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for r in range(shape[0]):
        for c in range(shape[1]):
            cell = ws_map.cell(row=r+1, column=c+1)
            cell.value     = grid_labels[r][c]
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = border
            if grid_colors[r][c]:
                cell.fill = PatternFill("solid", fgColor=grid_colors[r][c])
                cell.font = Font(size=6, bold=True,
                                 color="FFFFFF" if grid_colors[r][c] == "555555" else "000000")
        ws_map.row_dimensions[r+1].height = 16
    for c in range(shape[1]):
        ws_map.column_dimensions[get_column_letter(c+1)].width = 7

    # Résumé
    ws_sum = wb.create_sheet("Resume")
    hdrs = ["Nom", "Orientation", "Ligne", "Colonne", "Taille", "Culture recue", "Boost (%)"]
    for ci, h in enumerate(hdrs, 1):
        cell = ws_sum.cell(row=1, column=ci, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="2C3E50")
        ws_sum.column_dimensions[get_column_letter(ci)].width = 18
    for ri, inst in enumerate(placement.instances, 2):
        t        = inst["template"]
        r0, c0   = inst["row"], inst["col"]
        h, w     = inst["h"], inst["w"]
        boostable = t["boost_25"] > 0 or t["boost_50"] > 0 or t["boost_100"] > 0
        received = placement.culture_received(r0, c0, h, w, boostable=boostable)
        boost    = placement.boost_for(t, r0, c0, h, w)
        orient   = "Vertical" if inst["orientation"] == "V" else "Horizontal"
        for ci, val in enumerate([t["nom"], orient, r0+1, c0+1, f"{h}x{w}", round(received,1), boost], 1):
            ws_sum.cell(row=ri, column=ci, value=val)
        fill_c = {"100": "A9DFBF", "50": "FAD7A0", "25": "FDEBD0"}.get(str(boost))
        if fill_c:
            for ci in range(1, 8):
                ws_sum.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor=fill_c)
    last = len(placement.instances) + 3
    ws_sum.cell(row=last, column=1, value="Score total").font = Font(bold=True)
    ws_sum.cell(row=last, column=2, value=placement.total_score()).font = Font(bold=True)

    # Analyse boosts
    ws_ana = wb.create_sheet("Analyse Boosts")
    boost_insts = [inst for inst in placement.instances
                   if inst["template"]["boost_25"] > 0 or inst["template"]["boost_50"] > 0 or inst["template"]["boost_100"] > 0]
    total_b = len(boost_insts)
    at = {0:[], 25:[], 50:[], 100:[]}
    for inst in boost_insts:
        t = inst["template"]
        b = placement.boost_for(t, inst["row"], inst["col"], inst["h"], inst["w"])
        at[b].append(inst)
    score_obtenu = sum(placement.boost_for(i["template"], i["row"], i["col"], i["h"], i["w"]) for i in boost_insts)
    score_max    = total_b * 100
    ws_ana.cell(row=1, column=1, value="ANALYSE DES BOOSTS").font = Font(bold=True, size=12)
    data = [
        ("Batiments boost-sensibles", total_b),
        ("Score obtenu", score_obtenu),
        ("Score maximum possible", score_max),
        ("Efficacite", f"{100*score_obtenu//(score_max or 1)}%"),
        ("", ""),
        ("100% boost", f"{len(at[100])} batiments ({100*len(at[100])//(total_b or 1)}%)"),
        (" 50% boost", f"{len(at[50])} batiments ({100*len(at[50])//(total_b or 1)}%)"),
        (" 25% boost", f"{len(at[25])} batiments ({100*len(at[25])//(total_b or 1)}%)"),
        ("  0% boost", f"{len(at[0])} batiments ({100*len(at[0])//(total_b or 1)}%)"),
    ]
    fills = {"100% boost":"A9DFBF"," 50% boost":"FAD7A0"," 25% boost":"FDEBD0","  0% boost":"FADBD8"}
    for ri, (k, v) in enumerate(data, 3):
        ws_ana.cell(row=ri, column=1, value=k).font = Font(bold=True)
        ws_ana.cell(row=ri, column=2, value=v)
        if k in fills:
            for ci in [1,2]:
                ws_ana.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor=fills[k])
    if at[0]:
        ws_ana.cell(row=14, column=1, value="Batiments a 0% :").font = Font(bold=True, color="C0392B")
        for ri, inst in enumerate(at[0], 15):
            t = inst["template"]
            received = placement.culture_received(inst["row"], inst["col"], inst["h"], inst["w"])
            manque = t["boost_25"] - received
            ws_ana.cell(row=ri, column=1, value=t["nom"])
            ws_ana.cell(row=ri, column=2, value=f"recu={received:.0f}  seuil={t['boost_25']:.0f}  manque={manque:.0f}")
            for ci in [1,2]:
                ws_ana.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor="FADBD8")
    ws_ana.column_dimensions["A"].width = 35
    ws_ana.column_dimensions["B"].width = 35

    # Non placés
    placed_counts = {}
    for inst in placement.instances:
        n = inst["template"]["nom"]
        placed_counts[n] = placed_counts.get(n, 0) + 1
    ws_unp = wb.create_sheet("Non Places")
    red = PatternFill("solid", fgColor="C0392B")
    for ci, h in enumerate(["Batiment","Demandes","Places","Manquants"], 1):
        cell = ws_unp.cell(row=1, column=ci, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = red
        ws_unp.column_dimensions[get_column_letter(ci)].width = 20
    ri = 2
    for t in templates:
        placed  = placed_counts.get(t["nom"], 0)
        missing = t["quantite"] - placed
        if missing > 0:
            for ci, val in enumerate([t["nom"], t["quantite"], placed, missing], 1):
                ws_unp.cell(row=ri, column=ci, value=val).fill = PatternFill("solid", fgColor="FADBD8")
            ri += 1
    if ri == 2:
        ws_unp.cell(row=2, column=1, value="✓ Tous les batiments ont été placés !")

    # Légende
    ws_leg = wb.create_sheet("Legende")
    for ci, h in enumerate(["Couleur","Batiment","Taille orig.","Culture","Rayonnement"], 1):
        ws_leg.cell(row=1, column=ci, value=h).font = Font(bold=True)
        ws_leg.column_dimensions[get_column_letter(ci)].width = 20
    for ri2, t in enumerate(templates, 2):
        color = type_color.get(t["nom"], "CCCCCC")
        ws_leg.cell(row=ri2, column=1).fill = PatternFill("solid", fgColor=color)
        ws_leg.cell(row=ri2, column=2, value=t["nom"])
        ws_leg.cell(row=ri2, column=3, value=f"{t['largeur']}x{t['longueur']}")
        ws_leg.cell(row=ri2, column=4, value=t["culture"])
        ws_leg.cell(row=ri2, column=5, value=t["rayonnement"])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# UI Streamlit
# ---------------------------------------------------------------------------

st.markdown("""
<div class="hero">
    <h1>🏛️ Optimiseur de Bâtiments</h1>
    <p>Placement automatique et optimisation des boosts culturels</p>
</div>
""", unsafe_allow_html=True)

# Étape 1 — Upload
st.markdown('<span class="step-badge">1</span> **Chargez votre fichier Excel**', unsafe_allow_html=True)
uploaded = st.file_uploader("", type=["xlsx"], label_visibility="collapsed")

if uploaded:
    try:
        wb = load_workbook(uploaded, data_only=True)
        terrain   = load_terrain(wb)
        templates = load_buildings(wb)
        total     = sum(t["quantite"] for t in templates)
        surface   = sum(t["quantite"] * t["largeur"] * t["longueur"] for t in templates)
        free      = int((terrain == 0).sum())

        # Stats terrain
        st.markdown("---")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(f'<div class="stat-card"><div class="value">{terrain.shape[0]}x{terrain.shape[1]}</div><div class="label">Taille du terrain</div></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="stat-card"><div class="value">{free}</div><div class="label">Cases libres</div></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="stat-card"><div class="value">{total}</div><div class="label">Bâtiments</div></div>', unsafe_allow_html=True)
        with c4:
            marge = free - surface
            color = "#2ecc71" if marge >= 0 else "#e74c3c"
            st.markdown(f'<div class="stat-card"><div class="value" style="color:{color}">{marge:+d}</div><div class="label">Marge (cases)</div></div>', unsafe_allow_html=True)

        if surface > free:
            st.markdown(f'<div class="warning-box">⚠️ Surface des batiments ({surface}) supérieure aux cases libres ({free}). Certains batiments ne pourront pas être placés.</div>', unsafe_allow_html=True)

        # Étape 2 — Paramètres
        st.markdown("---")
        st.markdown('<span class="step-badge">2</span> **Paramètres d\'optimisation**', unsafe_allow_html=True)
        col_a, col_b = st.columns(2)
        with col_a:
            n_restarts = st.slider("Nombre d'essais", min_value=5, max_value=30, value=20,
                                   help="Plus d'essais = meilleur résultat, mais plus long")
        with col_b:
            local_iter = st.slider("Itérations d'amélioration", min_value=50, max_value=500, value=300,
                                   help="Affinement des boosts après chaque essai réussi")

        # Étape 3 — Lancer
        st.markdown("---")
        st.markdown('<span class="step-badge">3</span> **Lancer l\'optimisation**', unsafe_allow_html=True)

        if st.button("🚀 Optimiser le placement", use_container_width=True, type="primary"):
            progress_bar = st.progress(0)
            status_text  = st.empty()
            log_area     = st.empty()
            logs = []

            def progress_cb(i, total_r, placed, total_b, score):
                pct = i / total_r
                progress_bar.progress(pct)
                status_text.markdown(f"Essai **{i}/{total_r}** — placés : **{placed}/{total_b}** — score : **{score:.0f}**")
                logs.append(f"Essai {i:2d} : placés={placed}/{total_b}, score={score:.0f}")
                log_area.code("\n".join(logs[-8:]))

            with st.spinner("Optimisation en cours..."):
                placement, unplaced = optimize(
                    terrain, templates,
                    n_restarts=n_restarts,
                    local_iter=local_iter,
                    progress_cb=progress_cb
                )

            progress_bar.progress(1.0)
            status_text.empty()
            log_area.empty()

            placed_count = len(placement.instances)
            score        = placement.total_score()

            # Boost-sensibles
            boost_insts = [inst for inst in placement.instances
                           if inst["template"]["boost_25"] > 0 or inst["template"]["boost_50"] > 0 or inst["template"]["boost_100"] > 0]
            total_b   = len(boost_insts)
            score_max = total_b * 100
            efficacite = int(100 * score // (score_max or 1))

            at = {0:0, 25:0, 50:0, 100:0}
            for inst in boost_insts:
                t = inst["template"]
                b = placement.boost_for(t, inst["row"], inst["col"], inst["h"], inst["w"])
                at[b] += 1

            # Résultats
            st.markdown("---")
            if placed_count == total:
                st.markdown('<div class="success-box">✅ Tous les batiments ont été placés avec succès !</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="warning-box">⚠️ {placed_count}/{total} batiments placés — {len(unplaced)} non placés</div>', unsafe_allow_html=True)

            st.markdown("### 📊 Résultats des boosts")
            r1, r2, r3, r4, r5 = st.columns(5)
            with r1:
                st.markdown(f'<div class="stat-card"><div class="value" style="color:#c9a84c">{efficacite}%</div><div class="label">Efficacité</div></div>', unsafe_allow_html=True)
            with r2:
                st.markdown(f'<div class="stat-card"><div class="value" style="color:#2ecc71">{at[100]}</div><div class="label">À 100%</div></div>', unsafe_allow_html=True)
            with r3:
                st.markdown(f'<div class="stat-card"><div class="value" style="color:#f39c12">{at[50]}</div><div class="label">À 50%</div></div>', unsafe_allow_html=True)
            with r4:
                st.markdown(f'<div class="stat-card"><div class="value" style="color:#e67e22">{at[25]}</div><div class="label">À 25%</div></div>', unsafe_allow_html=True)
            with r5:
                st.markdown(f'<div class="stat-card"><div class="value" style="color:#e74c3c">{at[0]}</div><div class="label">À 0%</div></div>', unsafe_allow_html=True)

            # Barre de progression visuelle
            st.markdown("<br>", unsafe_allow_html=True)
            bar_color = "#2ecc71" if efficacite >= 90 else ("#f39c12" if efficacite >= 70 else "#e74c3c")
            st.markdown(f"""
            <div class="boost-bar-bg">
                <div class="boost-bar-fill" style="width:{efficacite}%; background:{bar_color}">
                    {efficacite}%
                </div>
            </div>
            <p style="color:#8899aa; font-size:0.8rem; margin-top:4px">
                Score : {int(score)} / {score_max} points possibles
            </p>
            """, unsafe_allow_html=True)

            # Téléchargement
            st.markdown("---")
            st.markdown('<span class="step-badge">4</span> **Téléchargez le résultat**', unsafe_allow_html=True)
            excel_buf = build_excel(placement, templates)
            st.download_button(
                label="⬇️ Télécharger le fichier résultat (.xlsx)",
                data=excel_buf,
                file_name="placement_resultat.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary"
            )

    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier : {e}")
        st.info("Vérifiez que votre fichier contient bien les onglets 'Terrain' et 'Batiments'.")

else:
    st.markdown("""
    <div style="background:#1a1a2e; border:1px dashed #c9a84c44; border-radius:8px; padding:2rem; text-align:center; color:#8899aa; margin-top:1rem">
        <div style="font-size:2.5rem">📂</div>
        <div style="margin-top:0.5rem">Glissez-déposez votre fichier Excel ici</div>
        <div style="font-size:0.8rem; margin-top:0.3rem">Format attendu : onglets <b>Terrain</b> et <b>Batiments</b></div>
    </div>
    """, unsafe_allow_html=True)
