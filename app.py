import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="Détecteur de bâtiments", page_icon="🏙️", layout="wide")

st.title("🏙️ Détecteur de bâtiments sur terrain")
st.markdown("Charge un fichier Excel contenant le terrain et les bâtiments pour identifier les bâtiments placés.")

uploaded_file = st.file_uploader("📂 Choisir le fichier Excel d'entrée", type=["xlsx"])

if not uploaded_file:
    st.info("Charge un fichier Excel avec 3 onglets : **Terrain**, **Batiments**, **Actuel**.")
    st.stop()

@st.cache_data
def load_data(file_bytes):
    terrain_df   = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Terrain",  header=None)
    batiments_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Batiments", header=0)
    actuel_df    = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Actuel",   header=None)
    return terrain_df, batiments_df, actuel_df

file_bytes = uploaded_file.read()
terrain_df, batiments_df, actuel_df = load_data(file_bytes)
batiments_df.columns = [str(c).strip() for c in batiments_df.columns]

def df_to_grid(df):
    grid = []
    for _, row in df.iterrows():
        line = ['' if pd.isna(v) else str(v).strip() for v in row]
        grid.append(line)
    max_cols = max(len(r) for r in grid)
    for r in grid:
        while len(r) < max_cols:
            r.append('')
    return grid

actuel_grid = df_to_grid(actuel_df)
ROWS = len(actuel_grid)
COLS = len(actuel_grid[0])

def build_catalog(df):
    catalog = {}
    for _, row in df.iterrows():
        nom = str(row.get("Nom", "")).strip()
        if not nom or nom == "nan":
            continue
        try:
            longueur = int(row.get("Longueur", 0))
            largeur  = int(row.get("Largeur",  0))
        except (ValueError, TypeError):
            continue
        catalog[nom] = {
            "longueur":    longueur,
            "largeur":     largeur,
            "quantite":    row.get("Quantite",    0),
            "type":        str(row.get("Type",    "")).strip(),
            "culture":     row.get("Culture",     0),
            "rayonnement": row.get("Rayonnement", 0),
            "boost25":     row.get("Boost 25%",   ""),
            "boost50":     row.get("Boost 50%",   ""),
            "boost100":    row.get("Boost 100%",  ""),
            "production":  str(row.get("Production", "")).strip(),
        }
    return catalog

catalog = build_catalog(batiments_df)

def flood_fill(grid, r0, c0, name):
    """Retourne l'ensemble des (r,c) connexes (4-connexité) portant 'name'."""
    cells = set()
    stack = [(r0, c0)]
    while stack:
        r, c = stack.pop()
        if (r, c) in cells:
            continue
        if r < 0 or r >= ROWS or c < 0 or c >= COLS:
            continue
        if grid[r][c] != name:
            continue
        cells.add((r, c))
        for dr, dc in [(-1,0),(1,0),(0,-1),(0,1)]:
            stack.append((r+dr, c+dc))
    return cells

def tile_box(cells, h, w):
    """
    Tente de paver l'ensemble 'cells' par des tuiles h×w.
    Parcourt en ordre ligne/colonne et place chaque tuile depuis le coin sup-gauche libre.
    Retourne une liste de (r0,c0) si succès, None sinon.
    """
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
    visited = [[False]*COLS for _ in range(ROWS)]
    found   = []
    journal = []

    for r in range(ROWS):
        for c in range(COLS):
            cell = grid[r][c]
            if not cell or cell in ("X", "0", "1", ""):
                continue
            if cell not in catalog:
                continue
            if visited[r][c]:
                continue

            blob = flood_fill(grid, r, c, cell)
            for br, bc in blob:
                visited[br][bc] = True

            bat = catalog[cell]
            lon = bat["longueur"]
            lar = bat["largeur"]

            placed = False
            for (h, w) in [(lar, lon), (lon, lar)]:
                placements = tile_box(blob, h, w)
                if placements is not None:
                    orientation = "Horizontal" if w == lon else "Vertical"
                    for (r0, c0) in placements:
                        found.append({
                            "Nom":         cell,
                            "Ligne":       r0 + 1,
                            "Colonne":     c0 + 1,
                            "Orientation": orientation,
                            "Hauteur":     h,
                            "Largeur":     w,
                        })
                        journal.append(
                            f"✅ **{cell}** — ligne {r0+1}, colonne {c0+1} "
                            f"({h}×{w}, {orientation})"
                        )
                    placed = True
                    break

            if not placed:
                min_r = min(br for br, bc in blob)
                min_c = min(bc for br, bc in blob)
                max_r = max(br for br, bc in blob)
                max_c = max(bc for br, bc in blob)
                journal.append(
                    f"⚠️ **{cell}** — bloc de {len(blob)} cases "
                    f"[r{min_r+1}-{max_r+1}, c{min_c+1}-{max_c+1}] "
                    f"impossible à paver avec une tuile {lar}×{lon}"
                )

    return found, journal

with st.spinner("Analyse du terrain en cours…"):
    found, journal = detect_buildings(actuel_grid, catalog)

warnings = [j for j in journal if j.startswith("⚠️")]
ok_count = len([j for j in journal if j.startswith("✅")])

st.subheader(f"📋 Journal — {ok_count} bâtiment(s) identifié(s)"
             + (f"  ·  ⚠️ {len(warnings)} avertissement(s)" if warnings else ""))

col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("### Journal détaillé")
    for entry in journal:
        st.markdown(entry)

with col2:
    st.markdown("### Résumé par bâtiment")
    if found:
        summary_df = (
            pd.DataFrame(found)
            .groupby("Nom")
            .size()
            .reset_index(name="Trouvés")
        )
        summary_df["Attendus"] = summary_df["Nom"].map(
            lambda n: catalog.get(n, {}).get("quantite", "?")
        )
        summary_df["Statut"] = summary_df.apply(
            lambda row: "✅" if str(row["Attendus"]) == str(row["Trouvés"]) else "⚠️",
            axis=1
        )
        st.dataframe(summary_df[["Nom","Trouvés","Attendus","Statut"]],
                     use_container_width=True)

def generate_output_excel(found, journal, catalog):
    wb = Workbook()
    header_fill = PatternFill("solid", start_color="1F4E79")
    header_font = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    data_font   = Font(name="Arial", size=10)
    center      = Alignment(horizontal="center", vertical="center")
    left_align  = Alignment(horizontal="left",   vertical="center")
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"),  bottom=Side(style="thin")
    )
    ok_fill   = PatternFill("solid", start_color="C6EFCE")
    warn_fill = PatternFill("solid", start_color="FFEB9C")
    alt_fill  = PatternFill("solid", start_color="DEEAF1")

    ws = wb.active
    ws.title = "Journal"
    for ci, h in enumerate(["#","Bâtiment","Ligne","Colonne","Orientation","Hauteur","Largeur"], 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = center; cell.border = thin_border

    for ri, entry in enumerate(found, 2):
        row_fill = alt_fill if ri % 2 == 0 else PatternFill("solid", start_color="FFFFFF")
        for ci, val in enumerate(
            [ri-1, entry["Nom"], entry["Ligne"], entry["Colonne"],
             entry["Orientation"], entry["Hauteur"], entry["Largeur"]], 1
        ):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font = data_font; cell.fill = row_fill; cell.border = thin_border
            cell.alignment = left_align if ci == 2 else center

    for ci, w in enumerate([6,34,8,8,12,8,8], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("Résumé")
    for ci, h in enumerate(["Bâtiment","Trouvés","Attendus","Statut","Type","Production"], 1):
        cell = ws2.cell(row=1, column=ci, value=h)
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = center; cell.border = thin_border

    counts = {}
    for f in found:
        counts[f["Nom"]] = counts.get(f["Nom"], 0) + 1

    for ri, (nom, occ) in enumerate(sorted(counts.items()), 2):
        bat  = catalog.get(nom, {})
        exp  = bat.get("quantite", "?")
        try:
            status = "✅" if int(exp) == occ else "⚠️"
            sfill  = ok_fill if int(exp) == occ else warn_fill
        except (ValueError, TypeError):
            status = "?"; sfill = warn_fill
        for ci, val in enumerate(
            [nom, occ, exp, status, bat.get("type",""), bat.get("production","")], 1
        ):
            cell = ws2.cell(row=ri, column=ci, value=val)
            cell.font = data_font; cell.border = thin_border
            cell.alignment = left_align if ci == 1 else center
            if ci in (2, 3):
                cell.fill = sfill

    for ci, w in enumerate([34,10,10,8,12,14], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

output_buf = generate_output_excel(found, journal, catalog)

st.divider()
st.subheader("⬇️ Télécharger le fichier de résultats")
st.download_button(
    label="📥 Télécharger resultats_batiments.xlsx",
    data=output_buf,
    file_name="resultats_batiments.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
