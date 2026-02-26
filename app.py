import streamlit as st
import pandas as pd
import numpy as np
import io
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches

st.set_page_config(
page_title=“Placement de Bâtiments”,
page_icon=“🏗️”,
layout=“wide”
)

st.markdown(”””

<style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700&display=swap');
    
    html, body, [class*="css"] { font-family: 'Sora', sans-serif; }
    
    .main { background-color: #f0f4f8; }
    
    h1 { color: #1a2e44; font-weight: 700; }
    
    .stButton > button {
        background-color: #2563eb;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1.5rem;
        font-family: 'Sora', sans-serif;
        font-weight: 600;
        transition: background 0.2s;
    }
    .stButton > button:hover { background-color: #1d4ed8; }
    
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 1rem 1.5rem;
        box-shadow: 0 2px 8px rgba(0,0,0,0.07);
        text-align: center;
    }
    .metric-card .value { font-size: 2rem; font-weight: 700; color: #2563eb; }
    .metric-card .label { font-size: 0.85rem; color: #64748b; margin-top: 0.2rem; }
    
    .success-badge {
        background: #dcfce7; color: #166534;
        border-radius: 6px; padding: 2px 10px;
        font-size: 0.8rem; font-weight: 600;
    }
    .fail-badge {
        background: #fee2e2; color: #991b1b;
        border-radius: 6px; padding: 2px 10px;
        font-size: 0.8rem; font-weight: 600;
    }
</style>

“””, unsafe_allow_html=True)

# ── Algorithme ──────────────────────────────────────────────────────────────

def lire_excel(fichier):
xl = pd.ExcelFile(fichier)
df_terrain = pd.read_excel(xl, sheet_name=“Terrain”, header=None)
terrain = df_terrain.fillna(0).astype(int).values

```
df_bat = pd.read_excel(xl, sheet_name="Batiments")
df_bat.columns = [c.strip() for c in df_bat.columns]
batiments = []
for _, row in df_bat.iterrows():
    batiments.append({
        "nom": str(row["Nom"]),
        "longueur": int(row["Longueur"]),
        "largeur": int(row["Largeur"]),
    })
return terrain, batiments
```

def peut_placer(terrain_bool, ligne, col, longueur, largeur):
nb_lignes, nb_cols = terrain_bool.shape
if ligne + largeur > nb_lignes or col + longueur > nb_cols:
return False
return not np.any(terrain_bool[ligne:ligne + largeur, col:col + longueur])

def placer_batiments(terrain, batiments):
terrain_work = terrain.copy()
placements = []

```
# Trier par surface décroissante (grands bâtiments en premier)
batiments_tries = sorted(batiments, key=lambda b: b["longueur"] * b["largeur"], reverse=True)

for idx, bat in enumerate(batiments_tries, start=2):
    nom, longueur, largeur = bat["nom"], bat["longueur"], bat["largeur"]
    place = False
    nb_lignes, nb_cols = terrain_work.shape

    for l in range(nb_lignes):
        for c in range(nb_cols):
            if peut_placer(terrain_work != 0, l, c, longueur, largeur):
                terrain_work[l:l + largeur, c:c + longueur] = idx
                placements.append({
                    "nom": nom, "ligne": l, "colonne": c,
                    "longueur": longueur, "largeur": largeur,
                    "statut": "Placé", "index": idx,
                })
                place = True
                break
        if place:
            break

    if not place:
        placements.append({
            "nom": nom, "ligne": None, "colonne": None,
            "longueur": longueur, "largeur": largeur,
            "statut": "Échec", "index": idx,
        })

return terrain_work, placements
```

# ── Visualisation ────────────────────────────────────────────────────────────

PALETTE = [
“#3b82f6”, “#10b981”, “#f59e0b”, “#8b5cf6”,
“#ef4444”, “#06b6d4”, “#84cc16”, “#f97316”,
“#ec4899”, “#14b8a6”,
]

def dessiner_terrain(terrain_original, terrain_result, placements):
nb_lignes, nb_cols = terrain_result.shape
fig, ax = plt.subplots(figsize=(max(8, nb_cols * 0.7), max(5, nb_lignes * 0.7)))

```
# Fond
for l in range(nb_lignes):
    for c in range(nb_cols):
        val = terrain_result[l, c]
        orig = terrain_original[l, c]

        if orig == 1:
            color = "#374151"  # occupé initialement
        elif val == 0:
            color = "#e5e7eb"  # libre
        else:
            color = PALETTE[(val - 2) % len(PALETTE)]

        rect = mpatches.FancyBboxPatch(
            (c + 0.05, nb_lignes - l - 1 + 0.05),
            0.9, 0.9,
            boxstyle="round,pad=0.02",
            facecolor=color,
            edgecolor="white",
            linewidth=1.5,
        )
        ax.add_patch(rect)

# Labels bâtiments
placed = {p["index"]: p for p in placements if p["statut"] == "Placé"}
for idx, p in placed.items():
    cx = p["colonne"] + p["longueur"] / 2
    cy = nb_lignes - p["ligne"] - p["largeur"] / 2
    ax.text(cx, cy, p["nom"], ha="center", va="center",
            fontsize=7, fontweight="bold", color="white",
            wrap=True)

# Grille
for l in range(nb_lignes + 1):
    ax.axhline(l, color="white", linewidth=0.5, alpha=0.3)
for c in range(nb_cols + 1):
    ax.axvline(c, color="white", linewidth=0.5, alpha=0.3)

ax.set_xlim(0, nb_cols)
ax.set_ylim(0, nb_lignes)
ax.set_aspect("equal")
ax.axis("off")

# Légende
legend_elements = [
    mpatches.Patch(facecolor="#374151", label="Occupé (initial)"),
    mpatches.Patch(facecolor="#e5e7eb", label="Libre"),
]
for p in placements:
    if p["statut"] == "Placé":
        color = PALETTE[(p["index"] - 2) % len(PALETTE)]
        legend_elements.append(mpatches.Patch(facecolor=color, label=p["nom"]))

ax.legend(handles=legend_elements, loc="upper left",
          bbox_to_anchor=(1.01, 1), frameon=True, fontsize=8)

plt.tight_layout()
return fig
```

# ── Export Excel résultat ────────────────────────────────────────────────────

def exporter_excel(terrain_original, terrain_result, placements):
wb = openpyxl.Workbook()
ws = wb.active
ws.title = “Terrain_Résultat”

```
nb_lignes, nb_cols = terrain_result.shape
placed = {p["index"]: p for p in placements if p["statut"] == "Placé"}

OPENPYXL_PALETTE = [
    "3B82F6", "10B981", "F59E0B", "8B5CF6",
    "EF4444", "06B6D4", "84CC16", "F97316",
    "EC4899", "14B8A6",
]

for l in range(nb_lignes):
    for c in range(nb_cols):
        val = terrain_result[l, c]
        orig = terrain_original[l, c]
        cell = ws.cell(row=l + 1, column=c + 1)

        if orig == 1:
            fill_color = "374151"
            cell.value = "X"
            cell.font = Font(color="FFFFFF", bold=True)
        elif val == 0:
            fill_color = "E5E7EB"
            cell.value = ""
        else:
            fill_color = OPENPYXL_PALETTE[(val - 2) % len(OPENPYXL_PALETTE)]
            p = placed.get(val)
            cell.value = p["nom"] if p else str(val)
            cell.font = Font(color="FFFFFF", bold=True, size=7)

        cell.fill = PatternFill("solid", start_color=fill_color, end_color=fill_color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[l + 1].height = 22

for c in range(nb_cols):
    ws.column_dimensions[openpyxl.utils.get_column_letter(c + 1)].width = 10

# Récapitulatif
ws2 = wb.create_sheet("Récapitulatif")
headers = ["Bâtiment", "Longueur", "Largeur", "Ligne", "Colonne", "Statut"]
header_fill = PatternFill("solid", start_color="1E3A5F", end_color="1E3A5F")
for c, h in enumerate(headers, 1):
    cell = ws2.cell(row=1, column=c, value=h)
    cell.fill = header_fill
    cell.font = Font(bold=True, color="FFFFFF")
    cell.alignment = Alignment(horizontal="center")

for r, p in enumerate(placements, 2):
    ws2.cell(row=r, column=1, value=p["nom"])
    ws2.cell(row=r, column=2, value=p["longueur"])
    ws2.cell(row=r, column=3, value=p["largeur"])
    ws2.cell(row=r, column=4, value=p["ligne"] if p["ligne"] is not None else "-")
    ws2.cell(row=r, column=5, value=p["colonne"] if p["colonne"] is not None else "-")
    ws2.cell(row=r, column=6, value=p["statut"])

for col in ["A", "B", "C", "D", "E", "F"]:
    ws2.column_dimensions[col].width = 15

buf = io.BytesIO()
wb.save(buf)
buf.seek(0)
return buf
```

# ── Interface ────────────────────────────────────────────────────────────────

st.title(“🏗️ Placement de Bâtiments”)
st.markdown(“Uploadez votre fichier Excel pour placer automatiquement vos bâtiments sur le terrain.”)

with st.expander(“📋 Format du fichier Excel attendu”, expanded=False):
col1, col2 = st.columns(2)
with col1:
st.markdown(”**Onglet « Terrain »**”)
st.markdown(“Grille de `0` (case libre) et `1` (case occupée).”)
st.dataframe(pd.DataFrame([
[0, 0, 1, 0],
[0, 0, 1, 0],
[0, 0, 0, 0],
]), hide_index=True)
with col2:
st.markdown(”**Onglet « Batiments »**”)
st.dataframe(pd.DataFrame({
“Nom”: [“Maison A”, “Entrepôt”, “Garage”],
“Longueur”: [3, 4, 2],
“Largeur”: [2, 3, 2],
}), hide_index=True)

fichier = st.file_uploader(“Choisir un fichier Excel (.xlsx)”, type=[“xlsx”])

if fichier:
try:
terrain, batiments = lire_excel(fichier)

```
    nb_lignes, nb_cols = terrain.shape
    cases_libres = int(np.sum(terrain == 0))
    surface_bats = sum(b["longueur"] * b["largeur"] for b in batiments)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f'<div class="metric-card"><div class="value">{nb_lignes}×{nb_cols}</div><div class="label">Taille du terrain</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="metric-card"><div class="value">{cases_libres}</div><div class="label">Cases libres</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="metric-card"><div class="value">{len(batiments)}</div><div class="label">Bâtiments</div></div>', unsafe_allow_html=True)
    with col4:
        st.markdown(f'<div class="metric-card"><div class="value">{surface_bats}</div><div class="label">Surface totale bâtiments</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("🚀 Lancer le placement"):
        with st.spinner("Placement en cours..."):
            terrain_result, placements = placer_batiments(terrain, batiments)

        places = [p for p in placements if p["statut"] == "Placé"]
        echecs = [p for p in placements if p["statut"] == "Échec"]

        st.markdown("### Résultats")

        col1, col2 = st.columns(2)
        with col1:
            st.success(f"✅ {len(places)} bâtiment(s) placé(s)")
        with col2:
            if echecs:
                st.error(f"❌ {len(echecs)} bâtiment(s) non placé(s) (pas d'espace)")
            else:
                st.success("Tous les bâtiments ont été placés !")

        st.markdown("### Visualisation du terrain")
        fig = dessiner_terrain(terrain, terrain_result, placements)
        st.pyplot(fig)

        st.markdown("### Détail des placements")
        df_recap = pd.DataFrame([{
            "Bâtiment": p["nom"],
            "Longueur": p["longueur"],
            "Largeur": p["largeur"],
            "Ligne": p["ligne"] if p["ligne"] is not None else "-",
            "Colonne": p["colonne"] if p["colonne"] is not None else "-",
            "Statut": p["statut"],
        } for p in placements])
        st.dataframe(df_recap, hide_index=True, use_container_width=True)

        excel_buf = exporter_excel(terrain, terrain_result, placements)
        st.download_button(
            label="⬇️ Télécharger le résultat Excel",
            data=excel_buf,
            file_name="resultat_placement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

except Exception as e:
    st.error(f"Erreur lors de la lecture du fichier : {e}")
    st.info("Vérifiez que votre fichier contient bien les onglets 'Terrain' et 'Batiments'.")
```