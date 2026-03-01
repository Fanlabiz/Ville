import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import copy
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Dict
import itertools

st.set_page_config(page_title=“Placement de Bâtiments”, layout=“wide”)

# ─────────────────────────────────────────────

# Data classes

# ─────────────────────────────────────────────

@dataclass
class Building:
name: str
length: int
width: int
quantity: int
type: str          # ‘culturel’ or ‘producteur’
culture: float = 0
rayonnement: int = 0
boost_25: float = 0
boost_50: float = 0
boost_100: float = 0
production: str = “”

@dataclass
class PlacedBuilding:
name: str
row: int
col: int
length: int   # actual length (may be rotated)
width: int    # actual width
type: str
culture: float = 0
rayonnement: int = 0
boost_25: float = 0
boost_50: float = 0
boost_100: float = 0
production: str = “”
received_culture: float = 0
boost_level: int = 0  # 0, 25, 50, 100

# ─────────────────────────────────────────────

# Parsing

# ─────────────────────────────────────────────

def parse_input(file) -> Tuple[np.ndarray, List[Building]]:
xls = pd.ExcelFile(file)

```
# Terrain (first sheet)
terrain_df = pd.read_excel(xls, sheet_name=0, header=None)
terrain = terrain_df.fillna(0).astype(int).values

# Buildings (second sheet)
buildings_df = pd.read_excel(xls, sheet_name=1, header=0)
buildings_df.columns = [str(c).strip() for c in buildings_df.columns]

buildings = []
for _, row in buildings_df.iterrows():
    cols = list(buildings_df.columns)
    def get(idx, default=0):
        v = row[cols[idx]] if idx < len(cols) else default
        if pd.isna(v): return default
        return v
    
    b = Building(
        name=str(get(0, "?")),
        length=int(get(1, 1)),
        width=int(get(2, 1)),
        quantity=int(get(3, 1)),
        type=str(get(4, "producteur")).strip().lower(),
        culture=float(get(5, 0)),
        rayonnement=int(get(6, 0)),
        boost_25=float(get(7, 0)),
        boost_50=float(get(8, 0)),
        boost_100=float(get(9, 0)),
        production=str(get(10, "")).strip() if not pd.isna(get(10, "")) else "",
    )
    buildings.append(b)

return terrain, buildings
```

# ─────────────────────────────────────────────

# Placement engine

# ─────────────────────────────────────────────

def can_place(grid: np.ndarray, row: int, col: int, l: int, w: int) -> bool:
rows, cols = grid.shape
if row + l > rows or col + w > cols:
return False
return np.all(grid[row:row+l, col:col+w] == 1)

def place(grid: np.ndarray, row: int, col: int, l: int, w: int, val: int = 0):
grid[row:row+l, col:col+w] = val

def get_cells(pb: PlacedBuilding):
return [(pb.row + dr, pb.col + dc) for dr in range(pb.length) for dc in range(pb.width)]

def cells_in_rayonnement(pb: PlacedBuilding, terrain_shape):
“”“All cells within rayonnement distance of a cultural building (including building cells).”””
r = pb.rayonnement
row_min = max(0, pb.row - r)
row_max = min(terrain_shape[0], pb.row + pb.length + r)
col_min = max(0, pb.col - r)
col_max = min(terrain_shape[1], pb.col + pb.width + r)
return set((rr, cc) for rr in range(row_min, row_max) for cc in range(col_min, col_max))

def compute_culture(placed: List[PlacedBuilding], terrain_shape):
“”“Compute received culture for each producer building.”””
cultural = [p for p in placed if p.type == ‘culturel’]
for p in placed:
if p.type != ‘producteur’:
continue
prod_cells = set(get_cells(p))
total = 0.0
for c in cultural:
ray_cells = cells_in_rayonnement(c, terrain_shape)
if prod_cells & ray_cells:  # overlap
total += c.culture
p.received_culture = total
# Boost level
if p.boost_100 > 0 and total >= p.boost_100:
p.boost_level = 100
elif p.boost_50 > 0 and total >= p.boost_50:
p.boost_level = 50
elif p.boost_25 > 0 and total >= p.boost_25:
p.boost_level = 25
else:
p.boost_level = 0

def score_placement(placed: List[PlacedBuilding], terrain_shape) -> Tuple[int, int, int]:
“”“Returns (guerison_score, nourriture_score, or_score) - each is sum of boost levels.”””
compute_culture(placed, terrain_shape)
g = n = o = 0
for p in placed:
if p.type != ‘producteur’: continue
bl = p.boost_level
prod = p.production.lower() if p.production else “”
if “guérison” in prod or “guerison” in prod or “guéri” in prod:
g += bl
elif “nourriture” in prod or “nourrit” in prod:
n += bl
elif “or” in prod:
o += bl
return g, n, o

PRODUCTION_PRIORITY = [“guerison”, “guérison”, “nourriture”, “or”]

def prod_priority(prod: str) -> int:
prod = prod.lower()
for i, p in enumerate(PRODUCTION_PRIORITY):
if p in prod:
return i
return 99

def optimize_placement(terrain_orig: np.ndarray, buildings: List[Building]):
“””
Greedy placement with priority:
1. Place all buildings
2. Maximize boost for Guerison
3. Maximize boost for Nourriture
4. Maximize boost for Or

```
Strategy:
- First place cultural buildings to maximize coverage
- Then place producers, prioritizing by production type
"""
terrain = terrain_orig.copy()
placed: List[PlacedBuilding] = []

# Expand buildings list
all_buildings = []
for b in buildings:
    for _ in range(b.quantity):
        all_buildings.append(copy.copy(b))

# Sort: cultural first, then producers by priority
def sort_key(b):
    if b.type == 'culturel':
        return (0, 0)
    return (1, prod_priority(b.production))

all_buildings.sort(key=sort_key)

# For each building, find best position
failed = []

for b in all_buildings:
    best_pos = None
    best_score = (-1, -1, -1, -1)
    
    orientations = [(b.length, b.width)]
    if b.length != b.width:
        orientations.append((b.width, b.length))
    
    rows, cols = terrain.shape
    
    for (l, w) in orientations:
        for r in range(rows):
            for c in range(cols):
                if can_place(terrain, r, c, l, w):
                    # Simulate placement
                    pb = PlacedBuilding(
                        name=b.name, row=r, col=c, length=l, width=w,
                        type=b.type, culture=b.culture, rayonnement=b.rayonnement,
                        boost_25=b.boost_25, boost_50=b.boost_50, boost_100=b.boost_100,
                        production=b.production
                    )
                    trial = placed + [pb]
                    compute_culture(trial, terrain.shape)
                    g, n, o = score_placement(trial, terrain.shape)
                    
                    # Tiebreaker: prefer positions that maximize future coverage (center of free area)
                    score = (g, n, o, 0)
                    if score > best_score:
                        best_score = score
                        best_pos = (r, c, l, w)
    
    if best_pos:
        r, c, l, w = best_pos
        pb = PlacedBuilding(
            name=b.name, row=r, col=c, length=l, width=w,
            type=b.type, culture=b.culture, rayonnement=b.rayonnement,
            boost_25=b.boost_25, boost_50=b.boost_50, boost_100=b.boost_100,
            production=b.production
        )
        placed.append(pb)
        place(terrain, r, c, l, w, 0)
    else:
        failed.append(b.name)

compute_culture(placed, terrain_orig.shape)
return placed, failed
```

# ─────────────────────────────────────────────

# Output Excel

# ─────────────────────────────────────────────

COLORS = [
“4472C4”, “ED7D31”, “A9D18E”, “FF0000”, “FFC000”,
“5B9BD5”, “70AD47”, “255E91”, “843C0C”, “375623”,
“7030A0”, “00B050”, “FF00FF”, “00B0F0”, “92D050”,
]

def build_output_excel(terrain_orig: np.ndarray, placed: List[PlacedBuilding], failed: List[str]) -> bytes:
wb = Workbook()

```
# ── Sheet 1: Terrain visuel ──
ws = wb.active
ws.title = "Terrain"

rows, cols = terrain_orig.shape

# Assign color per building name
names = list(dict.fromkeys(p.name for p in placed))
color_map = {n: COLORS[i % len(COLORS)] for i, n in enumerate(names)}

# Build cell→building map
cell_map: Dict[Tuple[int,int], PlacedBuilding] = {}
for pb in placed:
    for (r, c) in get_cells(pb):
        cell_map[(r, c)] = pb

thin = Side(style='thin', color="888888")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

for r in range(rows):
    ws.row_dimensions[r+2].height = 20
    for c in range(cols):
        cell = ws.cell(row=r+2, col=c+2)
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
        cell.font = Font(size=7, bold=True)
        ws.column_dimensions[get_column_letter(c+2)].width = 5
        
        if (r, c) in cell_map:
            pb = cell_map[(r, c)]
            color = color_map[pb.name]
            cell.fill = PatternFill("solid", fgColor=color)
            # Show name only in top-left cell
            if r == pb.row and c == pb.col:
                short = pb.name[:8]
                cell.value = short
                cell.font = Font(size=6, bold=True, color="FFFFFF")
        elif terrain_orig[r, c] == 1:
            cell.fill = PatternFill("solid", fgColor="D9E1F2")
        else:
            cell.fill = PatternFill("solid", fgColor="404040")

# Row/col labels
for c in range(cols):
    ws.cell(row=1, col=c+2).value = c
    ws.cell(row=1, col=c+2).font = Font(size=7, bold=True)
    ws.cell(row=1, col=c+2).alignment = Alignment(horizontal='center')
for r in range(rows):
    ws.cell(row=r+2, col=1).value = r
    ws.cell(row=r+2, col=1).font = Font(size=7, bold=True)

ws.column_dimensions['A'].width = 3
ws.freeze_panes = "B2"

# ── Sheet 2: Légende / Résultats ──
ws2 = wb.create_sheet("Résultats")
ws2.column_dimensions['A'].width = 25
ws2.column_dimensions['B'].width = 12
ws2.column_dimensions['C'].width = 12
ws2.column_dimensions['D'].width = 12
ws2.column_dimensions['E'].width = 20
ws2.column_dimensions['F'].width = 15
ws2.column_dimensions['G'].width = 15

header_fill = PatternFill("solid", fgColor="1F3864")
header_font = Font(color="FFFFFF", bold=True)

def hrow(ws, row, values):
    for i, v in enumerate(values):
        c = ws.cell(row=row, column=i+1, value=v)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal='center')

row = 1
ws2.cell(row=row, column=1, value="RÉSULTATS DE PLACEMENT").font = Font(bold=True, size=14)
row += 2

if failed:
    ws2.cell(row=row, column=1, value="⚠️ Bâtiments non placés:").font = Font(bold=True, color="FF0000")
    row += 1
    for f in failed:
        ws2.cell(row=row, column=1, value=f)
        row += 1
    row += 1
else:
    ws2.cell(row=row, column=1, value="✅ Tous les bâtiments ont été placés").font = Font(bold=True, color="00B050")
    row += 2

# Producer summary
hrow(ws2, row, ["Bâtiment", "Ligne", "Col", "Culture reçue", "Boost atteint", "Production", "Couleur"])
row += 1

for pb in placed:
    if pb.type != 'producteur':
        continue
    color = color_map.get(pb.name, "FFFFFF")
    cells_row = [
        pb.name,
        pb.row, pb.col,
        pb.received_culture,
        f"{pb.boost_level}%",
        pb.production,
        ""
    ]
    for i, v in enumerate(cells_row):
        c = ws2.cell(row=row, column=i+1, value=v)
        c.alignment = Alignment(horizontal='center')
    ws2.cell(row=row, column=7).fill = PatternFill("solid", fgColor=color)
    row += 1

row += 1
hrow(ws2, row, ["Bâtiment culturel", "Ligne", "Col", "Culture produite", "Rayonnement", "", "Couleur"])
row += 1

for pb in placed:
    if pb.type != 'culturel':
        continue
    color = color_map.get(pb.name, "FFFFFF")
    for i, v in enumerate([pb.name, pb.row, pb.col, pb.culture, pb.rayonnement, ""]):
        ws2.cell(row=row, column=i+1, value=v).alignment = Alignment(horizontal='center')
    ws2.cell(row=row, column=7).fill = PatternFill("solid", fgColor=color)
    row += 1

row += 2
# Summary by production type
ws2.cell(row=row, column=1, value="RÉCAPITULATIF PAR TYPE").font = Font(bold=True, size=12)
row += 1
hrow(ws2, row, ["Type de production", "Nb bâtiments", "Boost 0%", "Boost 25%", "Boost 50%", "Boost 100%"])
row += 1

prod_types = {}
for pb in placed:
    if pb.type != 'producteur' or not pb.production:
        continue
    pt = pb.production
    if pt not in prod_types:
        prod_types[pt] = {0:0, 25:0, 50:0, 100:0}
    prod_types[pt][pb.boost_level] += 1

for pt, counts in prod_types.items():
    total = sum(counts.values())
    for i, v in enumerate([pt, total, counts[0], counts[25], counts[50], counts[100]]):
        ws2.cell(row=row, column=i+1, value=v).alignment = Alignment(horizontal='center')
    row += 1

out = io.BytesIO()
wb.save(out)
out.seek(0)
return out.read()
```

# ─────────────────────────────────────────────

# Streamlit UI

# ─────────────────────────────────────────────

st.title(“🏗️ Placement Optimisé de Bâtiments”)
st.markdown(”””
Importez votre fichier Excel avec :

- **Onglet 1** : La grille du terrain (1 = case libre, 0 = case occupée)
- **Onglet 2** : La liste des bâtiments avec leurs caractéristiques
  “””)

with st.expander(“📋 Format attendu du fichier Excel”):
st.markdown(”””
**Onglet Terrain** : Grille de 0 et 1 (sans en-tête)

```
**Onglet Bâtiments** (avec en-têtes) :
| Nom | Longueur | Largeur | Quantité | Type | Culture | Rayonnement | Boost 25% | Boost 50% | Boost 100% | Production |
|-----|----------|---------|----------|------|---------|-------------|-----------|-----------|------------|------------|
| Marché | 2 | 2 | 3 | culturel | 50 | 3 | | | | |
| Ferme | 1 | 2 | 5 | producteur | | | 100 | 200 | 400 | Nourriture |

**Types** : `culturel` ou `producteur`
""")
```

uploaded = st.file_uploader(“📁 Choisissez votre fichier Excel (.xlsx)”, type=[“xlsx”])

if uploaded:
try:
terrain, buildings = parse_input(uploaded)

```
    col1, col2 = st.columns(2)
    with col1:
        st.success(f"✅ Terrain chargé : {terrain.shape[0]} lignes × {terrain.shape[1]} colonnes")
        free = int(terrain.sum())
        st.info(f"Cases libres : {free} / {terrain.size}")
    with col2:
        total_b = sum(b.quantity for b in buildings)
        st.success(f"✅ {len(buildings)} types de bâtiments ({total_b} bâtiments au total)")
    
    # Show terrain preview
    with st.expander("👁️ Aperçu du terrain"):
        terrain_display = pd.DataFrame(terrain)
        st.dataframe(terrain_display, use_container_width=True)
    
    with st.expander("👁️ Liste des bâtiments"):
        b_data = [{
            "Nom": b.name, "L": b.length, "W": b.width, "Qté": b.quantity,
            "Type": b.type, "Culture": b.culture, "Rayon": b.rayonnement,
            "B25": b.boost_25, "B50": b.boost_50, "B100": b.boost_100,
            "Production": b.production
        } for b in buildings]
        st.dataframe(pd.DataFrame(b_data), use_container_width=True)
    
    if st.button("🚀 Lancer l'optimisation", type="primary"):
        with st.spinner("Optimisation en cours... (peut prendre quelques instants)"):
            placed, failed = optimize_placement(terrain, buildings)
        
        if failed:
            st.error(f"⚠️ {len(failed)} bâtiment(s) n'ont pas pu être placés : {', '.join(failed)}")
        else:
            st.success("✅ Tous les bâtiments ont été placés !")
        
        # Show boost summary
        g, n, o = score_placement(placed, terrain.shape)
        
        col1, col2, col3 = st.columns(3)
        
        prod_summary = {}
        for pb in placed:
            if pb.type != 'producteur' or not pb.production: continue
            pt = pb.production
            if pt not in prod_summary:
                prod_summary[pt] = {0:0, 25:0, 50:0, 100:0}
            prod_summary[pt][pb.boost_level] += 1
        
        if prod_summary:
            st.subheader("📊 Récapitulatif des boosts")
            cols = st.columns(len(prod_summary))
            for idx, (pt, counts) in enumerate(prod_summary.items()):
                with cols[idx]:
                    st.metric(f"**{pt}**", f"{sum(counts.values())} bâtiments")
                    for level, cnt in sorted(counts.items()):
                        if cnt > 0:
                            emoji = "⭐" * (level // 25) if level > 0 else "◯"
                            st.write(f"{emoji} Boost {level}%: {cnt}")
        
        # Visual terrain
        st.subheader("🗺️ Terrain avec bâtiments placés")
        
        rows_t, cols_t = terrain.shape
        names = list(dict.fromkeys(p.name for p in placed))
        color_map = {n: COLORS[i % len(COLORS)] for i, n in enumerate(names)}
        cell_map = {}
        for pb in placed:
            for (r, c) in get_cells(pb):
                cell_map[(r, c)] = pb
        
        # Build HTML grid
        cell_size = max(20, min(40, 600 // cols_t))
        html = f'<div style="overflow-x:auto;font-family:monospace;font-size:{max(8,cell_size//3)}px">'
        html += '<table style="border-collapse:collapse">'
        for r in range(rows_t):
            html += "<tr>"
            for c in range(cols_t):
                if (r, c) in cell_map:
                    pb = cell_map[(r, c)]
                    col_hex = "#" + color_map[pb.name]
                    label = pb.name[:4] if r == pb.row and c == pb.col else ""
                    html += f'<td title="{pb.name}" style="width:{cell_size}px;height:{cell_size}px;background:{col_hex};border:1px solid #555;color:white;text-align:center;font-size:{max(6,cell_size//5)}px;font-weight:bold">{label}</td>'
                elif terrain[r, c] == 1:
                    html += f'<td style="width:{cell_size}px;height:{cell_size}px;background:#D9E1F2;border:1px solid #aaa"></td>'
                else:
                    html += f'<td style="width:{cell_size}px;height:{cell_size}px;background:#444;border:1px solid #333"></td>'
            html += "</tr>"
        html += "</table></div>"
        st.markdown(html, unsafe_allow_html=True)
        
        # Legend
        st.subheader("🔑 Légende")
        leg_cols = st.columns(min(5, len(names)))
        for i, name in enumerate(names):
            with leg_cols[i % 5]:
                col_hex = "#" + color_map[name]
                st.markdown(f'<span style="background:{col_hex};color:white;padding:2px 8px;border-radius:4px">{name}</span>', unsafe_allow_html=True)
        
        # Generate output Excel
        excel_bytes = build_output_excel(terrain, placed, failed)
        
        st.download_button(
            label="⬇️ Télécharger le fichier résultat (.xlsx)",
            data=excel_bytes,
            file_name="resultat_placement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

except Exception as e:
    st.error(f"❌ Erreur lors de la lecture du fichier : {e}")
    import traceback
    st.code(traceback.format_exc())
```

else:
st.info(“👆 Veuillez importer votre fichier Excel pour commencer.”)

```
# Sample file generator
if st.button("📥 Télécharger un fichier exemple"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Terrain"
    terrain_ex = [
        [1,1,1,1,1,1,1,1,1,1],
        [1,1,1,1,1,1,1,1,1,1],
        [1,1,0,0,1,1,1,1,1,1],
        [1,1,0,0,1,1,1,1,1,1],
        [1,1,1,1,1,1,1,1,1,1],
        [1,1,1,1,1,1,1,1,1,1],
        [1,1,1,1,1,1,0,0,0,1],
        [1,1,1,1,1,1,0,0,0,1],
    ]
    for row in terrain_ex:
        ws.append(row)
    
    ws2 = wb.create_sheet("Bâtiments")
    ws2.append(["Nom","Longueur","Largeur","Quantité","Type","Culture","Rayonnement","Boost 25%","Boost 50%","Boost 100%","Production"])
    ws2.append(["Marché",2,2,2,"culturel",80,3,"","","",""])
    ws2.append(["Cathédrale",1,1,1,"culturel",50,2,"","","",""])
    ws2.append(["Hôpital",2,1,3,"producteur","","",100,200,400,"Guérison"])
    ws2.append(["Ferme",1,2,4,"producteur","","",80,160,320,"Nourriture"])
    ws2.append(["Mine d'or",1,1,5,"producteur","","",60,120,240,"Or"])
    
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    st.download_button(
        "⬇️ Fichier exemple.xlsx",
        data=out.read(),
        file_name="exemple_terrain.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
```