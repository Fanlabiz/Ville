import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

st.set_page_config(page_title="Détecteur de bâtiments", page_icon="🏙️", layout="wide")
st.title("🏙️ Détecteur de bâtiments sur terrain")
st.markdown("Charge un fichier Excel avec 3 onglets : **Terrain**, **Batiments**, **Actuel**.")

uploaded_file = st.file_uploader("📂 Choisir le fichier Excel d'entrée", type=["xlsx"])
if not uploaded_file:
    st.stop()

@st.cache_data
def load_data(file_bytes):
    terrain_df   = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Terrain",   header=None)
    batiments_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Batiments", header=0)
    actuel_df    = pd.read_excel(io.BytesIO(file_bytes), sheet_name="Actuel",    header=None)
    return terrain_df, batiments_df, actuel_df

file_bytes = uploaded_file.read()
terrain_df, batiments_df, actuel_df = load_data(file_bytes)
batiments_df.columns = [str(c).strip() for c in batiments_df.columns]

# ─── Grille 2D ───────────────────────────────────────────────────────────────

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
COLS = len(actuel_grid[0]) if ROWS else 0

# ─── Catalogue ───────────────────────────────────────────────────────────────

def safe_int(v, default=0):
    try:
        f = float(v)
        return int(f) if not pd.isna(f) else default
    except (TypeError, ValueError):
        return default

def safe_float(v):
    try:
        f = float(v)
        return f if not pd.isna(f) else None
    except (TypeError, ValueError):
        return None

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

catalog = build_catalog(batiments_df)

# ─── Détection des bâtiments ─────────────────────────────────────────────────

def flood_fill(grid, r0, c0, name):
    cells, stack = set(), [(r0, c0)]
    while stack:
        r, c = stack.pop()
        if (r, c) in cells or not (0 <= r < ROWS and 0 <= c < COLS):
            continue
        if grid[r][c] != name:
            continue
        cells.add((r, c))
        for dr, dc in [(-1,0),(1,0),(0,-1),(0,1)]:
            stack.append((r+dr, c+dc))
    return cells

def tile_box(cells, h, w):
    """Pave l'ensemble de cases par des tuiles h×w (ordre ligne/col). None si impossible."""
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
    """
    Retourne une liste de bâtiments placés :
    {nom, ligne, colonne, hauteur, largeur, orientation, cases: set of (r,c)}
    """
    visited = [[False]*COLS for _ in range(ROWS)]
    found, warnings = [], []

    for r in range(ROWS):
        for c in range(COLS):
            cell = grid[r][c]
            if not cell or cell in ("X", "0", "1", "") or visited[r][c]:
                continue
            if cell not in catalog:
                continue
            blob = flood_fill(grid, r, c, cell)
            for br, bc in blob:
                visited[br][bc] = True

            bat = catalog[cell]
            lon, lar = bat["longueur"], bat["largeur"]
            placed = False
            for (h, w) in [(lar, lon), (lon, lar)]:
                placements = tile_box(blob, h, w)
                if placements is not None:
                    orientation = "Horizontal" if w == lon else "Vertical"
                    for (r0, c0) in placements:
                        tile_cells = {(r0+dr, c0+dc) for dr in range(h) for dc in range(w)}
                        found.append({
                            "nom":         cell,
                            "ligne":       r0 + 1,
                            "colonne":     c0 + 1,
                            "hauteur":     h,
                            "largeur":     w,
                            "orientation": orientation,
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
                    f"⚠️ **{cell}** — blob de {len(blob)} cases "
                    f"[r{min_r+1}-{max_r+1}, c{min_c+1}-{max_c+1}] "
                    f"impossible à paver avec une tuile {lar}×{lon}"
                )
    return found, warnings

# ─── Calcul de la culture et du boost ────────────────────────────────────────

def compute_culture_and_boost(found, catalog):
    """
    Pour chaque bâtiment Producteur, calcule la culture reçue et le boost.

    Zone de rayonnement d'un bâtiment Culturel de rayonnement R :
      toutes les cases dont la distance de Chebyshev (max |Δr|,|Δc|) au bâtiment
      est ≤ R, c.-à-d. le rectangle agrandi de R cases de chaque côté.

    Un Producteur reçoit la culture d'un Culturel si au moins une de ses cases
    est dans la zone de rayonnement du Culturel.
    """
    # Séparer Culturels et Producteurs
    culturels   = [b for b in found if catalog.get(b["nom"], {}).get("type") == "Culturel"]
    producteurs = [b for b in found if catalog.get(b["nom"], {}).get("type") == "Producteur"]

    # Pour chaque Culturel, construire l'ensemble des cases de sa zone de rayonnement
    def rayonnement_cases(bat, R):
        """Toutes les cases dans le rectangle étendu de R cases autour du bâtiment."""
        zone = set()
        r_min = bat["ligne"] - 1       # 0-based, coin sup-gauche
        c_min = bat["colonne"] - 1
        r_max = r_min + bat["hauteur"] - 1
        c_max = c_min + bat["largeur"] - 1
        for r in range(r_min - R, r_max + R + 1):
            for c in range(c_min - R, c_max + R + 1):
                zone.add((r, c))
        return zone

    # Pré-calculer les zones de rayonnement de tous les Culturels
    cultural_zones = []
    for cult in culturels:
        bat_info = catalog[cult["nom"]]
        R        = bat_info["rayonnement"]
        culture  = bat_info["culture"]
        if R > 0 and culture > 0:
            zone = rayonnement_cases(cult, R)
            cultural_zones.append({"zone": zone, "culture": culture, "nom": cult["nom"]})

    # Pour chaque bâtiment, initialiser culture_recue et boost
    for bat in found:
        bat["culture_recue"] = 0
        bat["boost"]         = ""

    # Calculer la culture reçue par chaque Producteur
    for prod in producteurs:
        total_culture = 0
        sources = []
        for cz in cultural_zones:
            # Le Producteur reçoit la culture si au moins une de ses cases
            # est dans la zone de rayonnement du Culturel
            if prod["cases"] & cz["zone"]:
                total_culture += cz["culture"]
                sources.append(cz["nom"])
        prod["culture_recue"] = total_culture

        # Calculer le boost
        bat_info = catalog[prod["nom"]]
        b25  = bat_info.get("boost25")
        b50  = bat_info.get("boost50")
        b100 = bat_info.get("boost100")

        boost = ""
        if b100 is not None and total_culture >= b100:
            boost = "100%"
        elif b50 is not None and total_culture >= b50:
            boost = "50%"
        elif b25 is not None and total_culture >= b25:
            boost = "25%"
        elif b25 is not None:
            boost = "0%"

        prod["boost"] = boost

    return found

# ─── Lancement ───────────────────────────────────────────────────────────────

with st.spinner("Analyse du terrain en cours…"):
    found, detect_warnings = detect_buildings(actuel_grid, catalog)
    found = compute_culture_and_boost(found, catalog)

# ─── Affichage ───────────────────────────────────────────────────────────────

nb_found = len(found)
nb_warn  = len(detect_warnings)

st.subheader(f"📋 Journal — {nb_found} bâtiment(s) identifié(s)"
             + (f"  ·  ⚠️ {nb_warn} avertissement(s)" if nb_warn else ""))

col1, col2 = st.columns([3, 2])

with col1:
    st.markdown("### Journal détaillé")
    for bat in found:
        bat_info = catalog.get(bat["nom"], {})
        bat_type = bat_info.get("type", "")
        line = f"✅ **{bat['nom']}** — ligne {bat['ligne']}, colonne {bat['colonne']} ({bat['hauteur']}×{bat['largeur']}, {bat['orientation']})"
        if bat_type == "Producteur":
            line += f" | Culture reçue : **{bat['culture_recue']}**"
            if bat["boost"]:
                line += f" | Boost : **{bat['boost']}**"
        st.markdown(line)
    for w in detect_warnings:
        st.markdown(w)

with col2:
    st.markdown("### Résumé par bâtiment")
    if found:
        counts = {}
        for f in found:
            counts[f["nom"]] = counts.get(f["nom"], 0) + 1
        summary_rows = []
        for nom, cnt in sorted(counts.items()):
            exp    = catalog.get(nom, {}).get("quantite", "?")
            status = "✅" if str(exp) == str(cnt) else "⚠️"
            summary_rows.append({"Bâtiment": nom, "Trouvés": cnt, "Attendus": exp, "": status})
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True, hide_index=True)

# ─── Génération Excel de sortie ──────────────────────────────────────────────

def make_excel(found, detect_warnings, catalog):
    wb = Workbook()

    H_FILL   = PatternFill("solid", start_color="1F4E79")
    H_FONT   = Font(bold=True, color="FFFFFF", name="Arial", size=11)
    D_FONT   = Font(name="Arial", size=10)
    CENTER   = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT     = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    BORDER   = Border(left=Side(style="thin"), right=Side(style="thin"),
                      top=Side(style="thin"),  bottom=Side(style="thin"))
    OK_FILL  = PatternFill("solid", start_color="C6EFCE")
    WN_FILL  = PatternFill("solid", start_color="FFEB9C")
    ALT_FILL = PatternFill("solid", start_color="DEEAF1")
    PROD_FILL= PatternFill("solid", start_color="FFF2CC")

    def hdr(ws, row, cols):
        for ci, h in enumerate(cols, 1):
            cell = ws.cell(row=row, column=ci, value=h)
            cell.fill = H_FILL; cell.font = H_FONT
            cell.alignment = CENTER; cell.border = BORDER

    def data_cell(ws, r, c, val, align=CENTER, fill=None):
        cell = ws.cell(row=r, column=c, value=val)
        cell.font = D_FONT; cell.border = BORDER; cell.alignment = align
        if fill:
            cell.fill = fill
        return cell

    # ── Onglet Journal ──────────────────────────────────────────────────────
    ws = wb.active
    ws.title = "Journal"
    journal_cols = ["#", "Bâtiment", "Type", "Ligne", "Colonne",
                    "Orientation", "Hauteur", "Largeur",
                    "Culture reçue", "Boost", "Production"]
    hdr(ws, 1, journal_cols)

    for ri, bat in enumerate(found, 2):
        bat_info = catalog.get(bat["nom"], {})
        bat_type = bat_info.get("type", "")
        is_prod  = bat_type == "Producteur"
        row_fill = PROD_FILL if is_prod else (ALT_FILL if ri % 2 == 0 else PatternFill("solid", start_color="FFFFFF"))

        boost_val = bat.get("boost", "") if is_prod else ""
        cult_val  = bat.get("culture_recue", "") if is_prod else ""
        prod_val  = bat_info.get("production", "") if is_prod else ""

        row_data = [
            ri - 1,
            bat["nom"],
            bat_type,
            bat["ligne"],
            bat["colonne"],
            bat["orientation"],
            bat["hauteur"],
            bat["largeur"],
            cult_val,
            boost_val,
            prod_val,
        ]
        for ci, val in enumerate(row_data, 1):
            data_cell(ws, ri, ci, val,
                      align=LEFT if ci == 2 else CENTER,
                      fill=row_fill)

    col_widths = [5, 30, 12, 7, 7, 12, 8, 8, 14, 8, 14]
    for ci, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A2"

    # ── Onglet Résumé ───────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Résumé")
    hdr(ws2, 1, ["Bâtiment", "Trouvés", "Attendus", "Statut", "Type", "Production"])

    counts = {}
    for f in found:
        counts[f["nom"]] = counts.get(f["nom"], 0) + 1

    for ri, (nom, cnt) in enumerate(sorted(counts.items()), 2):
        bat_info = catalog.get(nom, {})
        exp  = bat_info.get("quantite", "?")
        try:
            ok = int(exp) == cnt
        except (ValueError, TypeError):
            ok = False
        sfill = OK_FILL if ok else WN_FILL
        status = "✅" if ok else "⚠️"
        row_data = [nom, cnt, exp, status, bat_info.get("type",""), bat_info.get("production","")]
        for ci, val in enumerate(row_data, 1):
            cell = data_cell(ws2, ri, ci, val, align=LEFT if ci==1 else CENTER)
            if ci in (2, 3):
                cell.fill = sfill

    for ci, w in enumerate([30,10,10,8,12,14], 1):
        ws2.column_dimensions[get_column_letter(ci)].width = w
    ws2.freeze_panes = "A2"

    # ── Onglet Producteurs (détail culture+boost) ───────────────────────────
    ws3 = wb.create_sheet("Producteurs")
    hdr(ws3, 1, ["Bâtiment", "Ligne", "Colonne", "Production",
                 "Culture reçue", "Seuil 25%", "Seuil 50%", "Seuil 100%", "Boost atteint"])

    producteurs = [b for b in found if catalog.get(b["nom"], {}).get("type") == "Producteur"]
    for ri, bat in enumerate(producteurs, 2):
        bat_info = catalog.get(bat["nom"], {})
        b25  = bat_info.get("boost25")
        b50  = bat_info.get("boost50")
        b100 = bat_info.get("boost100")
        boost = bat.get("boost", "")

        boost_fill = None
        if boost == "100%":   boost_fill = PatternFill("solid", start_color="92D050")
        elif boost == "50%":  boost_fill = PatternFill("solid", start_color="FFEB9C")
        elif boost == "25%":  boost_fill = PatternFill("solid", start_color="FCE4D6")

        row_data = [
            bat["nom"], bat["ligne"], bat["colonne"],
            bat_info.get("production", ""),
            bat.get("culture_recue", 0),
            b25 if b25 is not None else "",
            b50 if b50 is not None else "",
            b100 if b100 is not None else "",
            boost,
        ]
        for ci, val in enumerate(row_data, 1):
            cell = data_cell(ws3, ri, ci, val, align=LEFT if ci==1 else CENTER)
            if ci == 9 and boost_fill:
                cell.fill = boost_fill

    for ci, w in enumerate([30,7,7,14,14,10,10,10,12], 1):
        ws3.column_dimensions[get_column_letter(ci)].width = w
    ws3.freeze_panes = "A2"

    # Avertissements
    if detect_warnings:
        ws4 = wb.create_sheet("Avertissements")
        ws4.cell(row=1, column=1, value="Avertissements").font = Font(bold=True, name="Arial")
        for ri, w in enumerate(detect_warnings, 2):
            ws4.cell(row=ri, column=1, value=w.replace("⚠️ ", "").replace("**",""))
        ws4.column_dimensions["A"].width = 70

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

output_buf = make_excel(found, detect_warnings, catalog)

st.divider()
st.download_button(
    label="📥 Télécharger resultats_batiments.xlsx",
    data=output_buf,
    file_name="resultats_batiments.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
