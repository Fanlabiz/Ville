import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import copy
from dataclasses import dataclass, field
from typing import Optional

st.set_page_config(page_title=“Placement de Bâtiments”, layout=“wide”)
st.title(“🏙️ Placement Automatique de Bâtiments”)

# ─────────────────────────────────────────────

# Structures de données

# ─────────────────────────────────────────────

@dataclass
class Batiment:
nom: str
longueur: int
largeur: int
nombre: int
type_bat: str        # Neutre / Culturel / Producteur
culture: float
rayonnement: int
boost_25: float
boost_50: float
boost_100: float
production: str
quantite: float

@dataclass
class BatimentPlace:
batiment: Batiment
row: int
col: int
rotation: bool       # True = pivoté (lon<->lar)
culture_recue: float = 0.0

```
@property
def h(self):
    return self.batiment.largeur if not self.rotation else self.batiment.longueur

@property
def w(self):
    return self.batiment.longueur if not self.rotation else self.batiment.largeur

def cases(self):
    return [(self.row + dr, self.col + dc) for dr in range(self.h) for dc in range(self.w)]

def boost(self):
    c = self.culture_recue
    b25, b50, b100 = self.batiment.boost_25, self.batiment.boost_50, self.batiment.boost_100
    if pd.notna(b100) and c >= b100:
        return 100
    if pd.notna(b50) and c >= b50:
        return 50
    if pd.notna(b25) and c >= b25:
        return 25
    return 0

def prod_reelle(self):
    b = self.boost()
    q = self.batiment.quantite
    if b == 100:
        return q * 2
    if b == 50:
        return q * 1.5
    if b == 25:
        return q * 1.25
    return q
```

# ─────────────────────────────────────────────

# Lecture du fichier input

# ─────────────────────────────────────────────

def lire_terrain(df_raw):
terrain = []
for _, row in df_raw.iterrows():
ligne = []
for v in row:
if str(v).strip().upper() == ‘X’:
ligne.append(‘X’)
elif str(v).strip() == ‘1’:
ligne.append(1)
elif str(v).strip() == ‘0’:
ligne.append(0)
else:
continue
if ligne:
terrain.append(ligne)
# Normaliser la largeur
W = max(len(r) for r in terrain)
for r in terrain:
while len(r) < W:
r.append(‘X’)
return terrain

def lire_batiments(df):
df = df.copy()
df.columns = [‘Nom’,‘Longueur’,‘Largeur’,‘Nombre’,‘Type’,‘Culture’,‘Rayonnement’,
‘Boost 25%’,‘Boost 50%’,‘Boost 100%’,‘Production’,‘Quantite’]
bats = []
for _, r in df.iterrows():
try:
b = Batiment(
nom=str(r[‘Nom’]).strip(),
longueur=int(r[‘Longueur’]),
largeur=int(r[‘Largeur’]),
nombre=int(r[‘Nombre’]),
type_bat=str(r[‘Type’]).strip(),
culture=float(r[‘Culture’]) if pd.notna(r[‘Culture’]) else 0,
rayonnement=int(r[‘Rayonnement’]) if pd.notna(r[‘Rayonnement’]) else 0,
boost_25=float(r[‘Boost 25%’]) if pd.notna(r[‘Boost 25%’]) else float(‘nan’),
boost_50=float(r[‘Boost 50%’]) if pd.notna(r[‘Boost 50%’]) else float(‘nan’),
boost_100=float(r[‘Boost 100%’]) if pd.notna(r[‘Boost 100%’]) else float(‘nan’),
production=str(r[‘Production’]).strip(),
quantite=float(r[‘Quantite’]) if pd.notna(r[‘Quantite’]) else 0,
)
bats.append(b)
except Exception:
pass
return bats

# ─────────────────────────────────────────────

# Calcul de la culture

# ─────────────────────────────────────────────

def calculer_zone_rayonnement(bp: BatimentPlace):
“”“Retourne l’ensemble des cases dans la zone de rayonnement d’un bâtiment culturel.”””
r0, c0, h, w = bp.row, bp.col, bp.h, bp.w
ray = bp.batiment.rayonnement
cases = set()
for dr in range(-ray, h + ray):
for dc in range(-ray, w + ray):
# exclure les cases occupées par le bâtiment lui-même
if 0 <= dr < h and 0 <= dc < w:
continue
cases.add((r0 + dr, c0 + dc))
return cases

def recalculer_culture(places: list[BatimentPlace]):
“”“Recalcule la culture reçue pour tous les producteurs.”””
culturels = [bp for bp in places if bp.batiment.type_bat == ‘Culturel’]
# Construire map cases_rayonnement -> culture
culture_map = {}  # (r,c) -> culture totale
for cp in culturels:
zone = calculer_zone_rayonnement(cp)
for case in zone:
culture_map[case] = culture_map.get(case, 0) + cp.batiment.culture

```
for bp in places:
    if bp.batiment.type_bat == 'Producteur':
        # Un producteur reçoit la culture si au moins une de ses cases est dans la zone
        total_culture = 0
        culturels_touches = set()
        for cp in culturels:
            zone = calculer_zone_rayonnement(cp)
            # Si une case du producteur est dans la zone
            for case in bp.cases():
                if case in zone:
                    culturels_touches.add(id(cp))
                    break
        for cp in culturels:
            if id(cp) in culturels_touches:
                total_culture += cp.batiment.culture
        bp.culture_recue = total_culture
```

# ─────────────────────────────────────────────

# Algorithme de placement

# ─────────────────────────────────────────────

class Placer:
def **init**(self, terrain_base, batiments: list[Batiment]):
self.terrain_base = terrain_base
self.H = len(terrain_base)
self.W = len(terrain_base[0])
self.batiments = batiments
self.places: list[BatimentPlace] = []
self.non_places: list[Batiment] = []
# Grille d’occupation : None = libre (si terrain==1), ‘X’ = bord, ref BatimentPlace
self.grille = [[None]*self.W for _ in range(self.H)]
for r in range(self.H):
for c in range(len(terrain_base[r])):
if terrain_base[r][c] == ‘X’:
self.grille[r][c] = ‘X’

```
def est_libre(self, r, c, h, w):
    if r < 0 or c < 0 or r+h > self.H or c+w > self.W:
        return False
    for dr in range(h):
        for dc in range(w):
            if self.grille[r+dr][c+dc] is not None:
                return False
    return True

def occuper(self, bp: BatimentPlace):
    for r, c in bp.cases():
        self.grille[r][c] = bp

def liberer(self, bp: BatimentPlace):
    for r, c in bp.cases():
        if self.grille[r][c] is bp:
            self.grille[r][c] = None

def cases_bord_libres(self, h, w):
    """Cases adjacentes aux X (bords), triées."""
    candidates = []
    for r in range(self.H):
        for c in range(self.W):
            if not self.est_libre(r, c, h, w):
                continue
            # est-on adjacent à un bord X ?
            adj = False
            for dr in range(h):
                for dc in range(w):
                    nr, nc = r+dr, c+dc
                    for d in [(-1,0),(1,0),(0,-1),(0,1)]:
                        rr, cc = nr+d[0], nc+d[1]
                        if 0 <= rr < self.H and 0 <= cc < self.W and self.grille[rr][cc] == 'X':
                            adj = True
            if adj:
                candidates.append((r, c))
    return candidates

def toutes_cases_libres(self, h, w):
    candidates = []
    for r in range(self.H):
        for c in range(self.W):
            if self.est_libre(r, c, h, w):
                candidates.append((r, c))
    return candidates

def score_cultural(self, r, c, h, w):
    """Score d'une position pour un bâtiment culturel : maximise les cases de rayonnement
    exploitables (pas sur X, pas en dehors du terrain, pas déjà utilisé par autre culturel)."""
    score = 0
    ray = 1  # sera recalculé par le vrai rayonnement
    # Compter les cases de rayonnement qui sont sur le terrain intérieur
    for dr in range(-ray, h+ray):
        for dc in range(-ray, w+ray):
            if 0 <= dr < h and 0 <= dc < w:
                continue
            rr, cc = r+dr, c+dc
            if 0 <= rr < self.H and 0 <= cc < self.W and self.grille[rr][cc] is None:
                score += 1
    return score

def placer_un(self, b: Batiment, forcer_bord=False, priorite_culture=False):
    """Tente de placer un bâtiment. Retourne True si succès."""
    orientations = [(b.largeur, b.longueur, False), (b.longueur, b.largeur, True)]
    if b.longueur == b.largeur:
        orientations = [(b.largeur, b.longueur, False)]

    best = None
    best_score = -1

    for h, w, rot in orientations:
        if forcer_bord:
            candidates = self.cases_bord_libres(h, w)
        else:
            candidates = self.toutes_cases_libres(h, w)

        for (r, c) in candidates:
            if priorite_culture:
                score = self.score_cultural(r, c, h, w)
            else:
                # Pour les producteurs : maximiser la culture reçue potentielle
                score = self._score_producteur(r, c, h, w)
            if score > best_score:
                best_score = score
                best = (r, c, h, w, rot)

    if best:
        r, c, h, w, rot = best
        bp = BatimentPlace(batiment=b, row=r, col=c, rotation=rot)
        self.occuper(bp)
        self.places.append(bp)
        return True
    return False

def _score_producteur(self, r, c, h, w):
    """Score d'emplacement d'un producteur : culture potentielle reçue."""
    total = 0
    for bp in self.places:
        if bp.batiment.type_bat != 'Culturel':
            continue
        zone = calculer_zone_rayonnement(bp)
        for dr in range(h):
            for dc in range(w):
                if (r+dr, c+dc) in zone:
                    total += bp.batiment.culture
                    break
    return total

def run(self, progress_cb=None):
    neutres = [b for b in self.batiments if b.type_bat == 'Neutre']
    culturels = [b for b in self.batiments if b.type_bat == 'Culturel']
    producteurs = [b for b in self.batiments if b.type_bat == 'Producteur']

    # Ordre de priorité pour la production
    prio_prod = {'Guerison': 0, 'Nourriture': 1, 'Or': 2}

    def taille(b): return b.longueur * b.largeur

    neutres.sort(key=taille, reverse=True)
    culturels.sort(key=taille, reverse=True)
    producteurs.sort(key=lambda b: (prio_prod.get(b.production, 99), -taille(b)))

    total = sum(b.nombre for b in self.batiments)
    placed = 0

    # 1. Neutres sur les bords
    for b in neutres:
        for _ in range(b.nombre):
            if not self.placer_un(b, forcer_bord=True):
                if not self.placer_un(b, forcer_bord=False):
                    self.non_places.append(b)
            placed += 1
            if progress_cb:
                progress_cb(placed / total)

    # 2. Alternance culturels / producteurs
    # On interleave : 1 culturel puis N producteurs, etc.
    ci = 0
    pi = 0
    culturel_queue = []
    for b in culturels:
        culturel_queue.extend([b]*b.nombre)
    producteur_queue = []
    for b in producteurs:
        producteur_queue.extend([b]*b.nombre)

    while ci < len(culturel_queue) or pi < len(producteur_queue):
        # Placer un culturel
        if ci < len(culturel_queue):
            b = culturel_queue[ci]; ci += 1
            if not self.placer_un(b, forcer_bord=False, priorite_culture=True):
                if not self.placer_un(b, forcer_bord=False):
                    self.non_places.append(b)
            placed += 1
            if progress_cb: progress_cb(placed / total)
        # Placer un producteur
        if pi < len(producteur_queue):
            b = producteur_queue[pi]; pi += 1
            if not self.placer_un(b, forcer_bord=False, priorite_culture=False):
                self.non_places.append(b)
            placed += 1
            if progress_cb: progress_cb(placed / total)

    recalculer_culture(self.places)
```

# ─────────────────────────────────────────────

# Génération du fichier résultat

# ─────────────────────────────────────────────

ORANGE = “FFAA00”
VERT   = “00BB44”
GRIS   = “AAAAAA”
BLEU   = “4488FF”
BLANC  = “FFFFFF”
JAUNE  = “FFFF99”

def couleur_type(type_bat):
if type_bat == ‘Culturel’:   return ORANGE
if type_bat == ‘Producteur’: return VERT
if type_bat == ‘Neutre’:     return GRIS
return BLANC

def creer_excel(places: list[BatimentPlace], non_places: list[Batiment],
terrain_base, batiments_input: list[Batiment]) -> bytes:
wb = Workbook()

```
# ── Onglet 1 : Bâtiments placés ────────────────────────────────────────
ws1 = wb.active
ws1.title = "Bâtiments placés"
hdrs = ["Nom","Type","Production","Ligne","Col","Rotation","Culture reçue","Boost","Prod/heure réelle"]
ws1.append(hdrs)
for cell in ws1[1]:
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", start_color="CCCCCC")

for bp in places:
    rot_str = "Oui" if bp.rotation else "Non"
    boost_str = f"{bp.boost()}%"
    ws1.append([
        bp.batiment.nom, bp.batiment.type_bat, bp.batiment.production,
        bp.row+1, bp.col+1, rot_str,
        round(bp.culture_recue, 1), boost_str,
        round(bp.prod_reelle(), 2)
    ])
    row_idx = ws1.max_row
    color = couleur_type(bp.batiment.type_bat)
    for cell in ws1[row_idx]:
        cell.fill = PatternFill("solid", start_color=color)

for col in ws1.columns:
    max_len = max(len(str(cell.value or "")) for cell in col)
    ws1.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 3, 40)

# ── Onglet 2 : Non placés ───────────────────────────────────────────────
ws2 = wb.create_sheet("Non placés")
ws2.append(["Non placés"])
ws2["A1"].font = Font(bold=True)
if non_places:
    for b in non_places:
        ws2.append([b.nom])
else:
    ws2.append(["(aucun)"])

# ── Onglet 3 : Production totale ────────────────────────────────────────
ws3 = wb.create_sheet("Production totale")
ws3.append(["Ressource", "Boost 0%", "Boost 25%", "Boost 50%", "Boost 100%", "Prod totale /h"])
for cell in ws3[1]:
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", start_color="CCCCCC")

# Regrouper par production
from collections import defaultdict
prod_data = defaultdict(lambda: {'b0':0,'b25':0,'b50':0,'b100':0,'total':0})
for bp in places:
    if bp.batiment.type_bat == 'Producteur' and bp.batiment.production != 'Rien':
        key = bp.batiment.production
        b = bp.boost()
        q = bp.prod_reelle()
        prod_data[key]['total'] += q
        prod_data[key][f'b{b}'] += 1

# Ordre prioritaire
ordre = ['Guerison','Nourriture','Or']
keys_ordered = ordre + [k for k in prod_data if k not in ordre]
for key in keys_ordered:
    if key in prod_data:
        d = prod_data[key]
        ws3.append([key, d['b0'], d['b25'], d['b50'], d['b100'], round(d['total'], 2)])

for col in ws3.columns:
    ws3.column_dimensions[get_column_letter(col[0].column)].width = 18

# ── Onglet 4 : Résumé culture ───────────────────────────────────────────
ws4 = wb.create_sheet("Résumé culture")
ws4.append(["Production","Culture totale reçue","Seuil 25%","Seuil 50%","Seuil 100%","Boost atteint"])
for cell in ws4[1]:
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", start_color="CCCCCC")

prod_culture = defaultdict(list)
for bp in places:
    if bp.batiment.type_bat == 'Producteur' and bp.batiment.production != 'Rien':
        prod_culture[bp.batiment.production].append(bp)

for prod, bps in prod_culture.items():
    cult_total = sum(b.culture_recue for b in bps)
    b25  = bps[0].batiment.boost_25  if pd.notna(bps[0].batiment.boost_25)  else '-'
    b50  = bps[0].batiment.boost_50  if pd.notna(bps[0].batiment.boost_50)  else '-'
    b100 = bps[0].batiment.boost_100 if pd.notna(bps[0].batiment.boost_100) else '-'
    boosts = [bp.boost() for bp in bps]
    boost_max = max(boosts) if boosts else 0
    ws4.append([prod, round(cult_total,1), b25, b50, b100, f"{boost_max}%"])

for col in ws4.columns:
    ws4.column_dimensions[get_column_letter(col[0].column)].width = 20

# ── Onglet 5 : Terrain visuel ───────────────────────────────────────────
ws5 = wb.create_sheet("Terrain")
H = len(terrain_base)
W = len(terrain_base[0])

# Construire la map des cases -> bâtiment placé
case_map = {}
for bp in places:
    for (r, c) in bp.cases():
        case_map[(r, c)] = bp

thin = Side(style='thin', color="888888")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# Écrire les cellules
for r in range(H):
    for c in range(W):
        cell = ws5.cell(row=r+1, column=c+1)
        v = terrain_base[r][c] if c < len(terrain_base[r]) else 'X'
        if v == 'X':
            cell.fill = PatternFill("solid", start_color="444444")
            cell.value = "X"
            cell.font = Font(color="FFFFFF", bold=True, size=7)
        elif (r, c) in case_map:
            bp = case_map[(r, c)]
            color = couleur_type(bp.batiment.type_bat)
            cell.fill = PatternFill("solid", start_color=color)
            # N'écrire le nom que dans la case en haut à gauche du bâtiment
            if (r, c) == (bp.row, bp.col):
                boost_str = f" [{bp.boost()}%]" if bp.batiment.type_bat == 'Producteur' else ""
                cell.value = bp.batiment.nom + boost_str
                cell.font = Font(size=7, bold=True)
            cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        else:
            cell.fill = PatternFill("solid", start_color="E8F5E9")
        cell.border = border

# Largeur des colonnes uniforme
for c in range(1, W+1):
    ws5.column_dimensions[get_column_letter(c)].width = 12
for r in range(1, H+1):
    ws5.row_dimensions[r].height = 30

# ── Onglet 6 : Statistiques ─────────────────────────────────────────────
ws6 = wb.create_sheet("Statistiques")
ws6.append(["Statistique", "Valeur"])
ws6["A1"].font = Font(bold=True)
ws6["B1"].font = Font(bold=True)

total_cases = sum(1 for r in range(H) for c in range(W)
                  if (c < len(terrain_base[r]) and terrain_base[r][c] == 1))
used_cases  = sum(bp.h * bp.w for bp in places)
free_cases  = total_cases - used_cases
np_cases    = sum(b.longueur * b.largeur for b in non_places)

ws6.append(["Bâtiments placés",  len(places)])
ws6.append(["Bâtiments non placés", len(non_places)])
ws6.append(["Cases totales disponibles", total_cases])
ws6.append(["Cases utilisées", used_cases])
ws6.append(["Cases libres", free_cases])
ws6.append(["Cases représentées par les non-placés", np_cases])

for col in ws6.columns:
    ws6.column_dimensions[get_column_letter(col[0].column)].width = 40

# Légende couleurs
ws6.append([])
ws6.append(["Légende"])
ws6[f"A{ws6.max_row}"].font = Font(bold=True)
for label, color in [("Culturel", ORANGE), ("Producteur", VERT), ("Neutre", GRIS)]:
    ws6.append([label])
    cell = ws6.cell(row=ws6.max_row, column=1)
    cell.fill = PatternFill("solid", start_color=color)
    cell.font = Font(bold=True)

buf = io.BytesIO()
wb.save(buf)
buf.seek(0)
return buf.read()
```

# ─────────────────────────────────────────────

# Interface Streamlit

# ─────────────────────────────────────────────

uploaded = st.file_uploader(“📂 Charger le fichier Excel (Ville.xlsx)”, type=[“xlsx”])

if uploaded:
try:
xl = pd.ExcelFile(uploaded)
sheets = xl.sheet_names
st.success(f”Fichier chargé : {len(sheets)} onglets détectés ({’, ’.join(sheets)})”)

```
    df_terrain_raw = pd.read_excel(uploaded, sheet_name=sheets[0], header=None)
    df_bats_raw    = pd.read_excel(uploaded, sheet_name=sheets[1], header=0)

    terrain = lire_terrain(df_terrain_raw)
    batiments = lire_batiments(df_bats_raw)

    H = len(terrain)
    W = len(terrain[0])
    total_bats = sum(b.nombre for b in batiments)
    free_cells = sum(1 for r in range(H) for c in range(W) if terrain[r][c] == 1)

    col1, col2, col3 = st.columns(3)
    col1.metric("Taille du terrain", f"{H} × {W}")
    col2.metric("Cases libres", free_cells)
    col3.metric("Bâtiments à placer", total_bats)

    with st.expander("📋 Liste des bâtiments", expanded=False):
        df_show = df_bats_raw.copy()
        st.dataframe(df_show, use_container_width=True)

    if st.button("🚀 Lancer le placement", type="primary"):
        progress_bar = st.progress(0, text="Placement en cours…")

        placer = Placer(terrain, batiments)

        def update_progress(v):
            progress_bar.progress(min(v, 1.0), text=f"Placement en cours… {int(v*100)}%")

        placer.run(progress_cb=update_progress)
        recalculer_culture(placer.places)

        progress_bar.progress(1.0, text="Placement terminé !")

        placed = placer.places
        non_placed = placer.non_places

        st.markdown("---")
        c1, c2, c3 = st.columns(3)
        c1.metric("✅ Bâtiments placés", len(placed))
        c2.metric("❌ Non placés", len(non_placed))
        boosted = sum(1 for bp in placed if bp.boost() > 0 and bp.batiment.type_bat == 'Producteur')
        c3.metric("⚡ Producteurs boostés", boosted)

        # Aperçu production
        from collections import defaultdict
        prod_data = defaultdict(float)
        for bp in placed:
            if bp.batiment.type_bat == 'Producteur' and bp.batiment.production != 'Rien':
                prod_data[bp.batiment.production] += bp.prod_reelle()

        if prod_data:
            st.markdown("**Production totale par ressource :**")
            df_prod = pd.DataFrame([
                {"Ressource": k, "Prod/heure": round(v, 2)}
                for k, v in sorted(prod_data.items(), key=lambda x: (
                    {'Guerison':0,'Nourriture':1,'Or':2}.get(x[0], 99), x[0]))
            ])
            st.dataframe(df_prod, use_container_width=True, hide_index=True)

        if non_placed:
            with st.expander(f"❌ Bâtiments non placés ({len(non_placed)})", expanded=False):
                for b in non_placed:
                    st.write(f"- {b.nom} ({b.longueur}×{b.largeur})")

        # Génération du fichier résultat
        excel_bytes = creer_excel(placed, non_placed, terrain, batiments)

        st.download_button(
            label="📥 Télécharger le résultat Excel",
            data=excel_bytes,
            file_name="Resultat_placement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

except Exception as e:
    st.error(f"Erreur lors du traitement : {e}")
    import traceback
    st.code(traceback.format_exc())
```

else:
st.info(“👆 Veuillez charger votre fichier Excel pour commencer.”)
with st.expander(“📖 Instructions”, expanded=True):
st.markdown(”””
**Structure du fichier Excel attendue :**

- **Onglet 1 – Terrain** : Grille de cases. `X` = bord/mur, `1` = case libre, `0` = case occupée.
- **Onglet 2 – Bâtiments** : Liste des bâtiments avec les colonnes :
  `Nom, Longueur, Largeur, Nombre, Type, Culture, Rayonnement, Boost 25%, Boost 50%, Boost 100%, Production, Quantite`
- **Onglet 3 – Actuel** *(optionnel)* : Terrain avec bâtiments déjà placés.

**Types de bâtiments :**

- 🟠 **Culturel** : Génère de la culture dans sa zone de rayonnement.
- 🟢 **Producteur** : Reçoit de la culture et bénéficie de boosts de production.
- ⚫ **Neutre** : Placé en priorité sur les bords.

**Ordre de placement :**

1. Neutres → placés sur les bords
1. Alternance Culturels / Producteurs (Guérison en priorité)
   “””)
