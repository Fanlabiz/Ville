import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import copy

st.set_page_config(page_title=“Placement de Bâtiments”, layout=“wide”)
st.title(“🏗️ Placement de Bâtiments sur Terrain”)

# ─────────────────────────────────────────────

# LECTURE DU FICHIER EXCEL

# ─────────────────────────────────────────────

def lire_terrain(df):
terrain = []
for _, row in df.iterrows():
ligne = []
for val in row:
if pd.notna(val):
ligne.append(int(val))
if ligne:
terrain.append(ligne)
return terrain

def lire_batiments(df):
batiments = []
df.columns = [str(c).strip() for c in df.columns]
for _, row in df.iterrows():
if pd.isna(row.iloc[0]):
continue
b = {
“Nom”: str(row.get(“Nom”, “”)).strip(),
“Longueur”: int(row.get(“Longueur”, 1)),
“Largeur”: int(row.get(“Largeur”, 1)),
“Quantite”: int(row.get(“Quantite”, 1)),
“Type”: str(row.get(“Type”, “Neutre”)).strip(),
“Culture”: float(row.get(“Culture”, 0) or 0),
“Rayonnement”: int(row.get(“Rayonnement”, 0) or 0),
“Boost25”: float(row.get(“Boost 25%”, 0) or 0),
“Boost50”: float(row.get(“Boost 50%”, 0) or 0),
“Boost100”: float(row.get(“Boost 100%”, 0) or 0),
“Production”: str(row.get(“Production”, “”) or “”).strip(),
}
batiments.append(b)
return batiments

# ─────────────────────────────────────────────

# LOGIQUE DE PLACEMENT

# ─────────────────────────────────────────────

def cases_bord(terrain, L, l):
“”“Retourne les positions de bord valides pour un bâtiment L x l.”””
rows = len(terrain)
cols = len(terrain[0]) if rows > 0 else 0
positions = []
for r in range(rows):
for c in range(cols):
if est_bord(r, c, L, l, rows, cols) and peut_placer(terrain, r, c, L, l):
positions.append((r, c, L, l))
if L != l and est_bord(r, c, l, L, rows, cols) and peut_placer(terrain, r, c, l, L):
positions.append((r, c, l, L))
return positions

def est_bord(r, c, lon, lar, rows, cols):
return r == 0 or c == 0 or r + lon == rows or c + lar == cols

def peut_placer(terrain, r, c, lon, lar):
rows = len(terrain)
cols = len(terrain[0])
if r + lon > rows or c + lar > cols:
return False
for i in range(r, r + lon):
for j in range(c, c + lar):
if terrain[i][j] != 1:
return False
return True

def placer(terrain, r, c, lon, lar, id_bat):
for i in range(r, r + lon):
for j in range(c, c + lar):
terrain[i][j] = id_bat

def enlever(terrain, r, c, lon, lar):
for i in range(r, r + lon):
for j in range(c, c + lar):
terrain[i][j] = 1

def cases_libres(terrain):
return sum(1 for row in terrain for v in row if v == 1)

def toutes_positions(terrain, lon, lar):
rows = len(terrain)
cols = len(terrain[0]) if rows > 0 else 0
positions = []
for r in range(rows):
for c in range(cols):
if peut_placer(terrain, r, c, lon, lar):
positions.append((r, c, lon, lar))
if lon != lar and peut_placer(terrain, r, c, lar, lon):
positions.append((r, c, lar, lon))
return positions

def zone_rayonnement(r, c, lon, lar, ray, rows, cols):
“”“Cases dans la bande de rayonnement autour du bâtiment.”””
cases = set()
for dr in range(-ray, lon + ray):
for dc in range(-ray, lar + ray):
ri, ci = r + dr, c + dc
if 0 <= ri < rows and 0 <= ci < cols:
if not (0 <= dr < lon and 0 <= dc < lar):
cases.add((ri, ci))
return cases

def calcul_culture(placed, batiments_info, rows, cols):
“”“Calcule la culture reçue par chaque bâtiment Producteur placé.”””
culture_recue = {}  # id_bat -> culture totale
# Pour chaque producteur placé, vérifier les culturels
for bid, info in placed.items():
b = batiments_info[info[“nom”]]
if b[“Type”] == “Producteur”:
culture_recue[bid] = 0

```
for bid_cult, info_cult in placed.items():
    b_cult = batiments_info[info_cult["nom"]]
    if b_cult["Type"] != "Culturel":
        continue
    ray = b_cult["Rayonnement"]
    r, c, lon, lar = info_cult["r"], info_cult["c"], info_cult["lon"], info_cult["lar"]
    zone = zone_rayonnement(r, c, lon, lar, ray, rows, cols)
    # Chercher les producteurs dont au moins une case est dans la zone
    for bid_prod, info_prod in placed.items():
        if bid_prod not in culture_recue:
            continue
        rp, cp, lonp, larp = info_prod["r"], info_prod["c"], info_prod["lon"], info_prod["lar"]
        for ir in range(rp, rp + lonp):
            for ic in range(cp, cp + larp):
                if (ir, ic) in zone:
                    culture_recue[bid_prod] += b_cult["Culture"]
                    break
            else:
                continue
            break
return culture_recue
```

def taille_batiment(b):
return b[“Longueur”] * b[“Largeur”]

def preparer_liste(batiments):
“”“Expand les bâtiments selon leur quantité, retourner liste ordonnée.”””
liste = []
for b in batiments:
for _ in range(b[“Quantite”]):
liste.append(copy.deepcopy(b))
return liste

def trier_liste(liste):
ordre = {“Neutre”: 0, “Culturel”: 1, “Producteur”: 2}
# D’abord Neutres (bord), puis alternance Culturel/Producteur/Neutre du plus grand au plus petit
neutres = sorted([b for b in liste if b[“Type”] == “Neutre”], key=taille_batiment, reverse=True)
autres = sorted([b for b in liste if b[“Type”] != “Neutre”], key=taille_batiment, reverse=True)
return neutres + autres

def plus_grand_restant(liste_restante):
if not liste_restante:
return 0
return max(taille_batiment(b) for b in liste_restante)

def algorithme_placement(terrain_init, batiments):
terrain = copy.deepcopy(terrain_init)
rows = len(terrain)
cols = len(terrain[0]) if rows > 0 else 0

```
liste = trier_liste(preparer_liste(batiments))
batiments_info = {b["Nom"]: b for b in batiments}

placed = {}       # id -> {nom, r, c, lon, lar}
non_places = []
journal = []
id_counter = [2]  # terrain utilise 0 et 1

MAX_JOURNAL = 1000

def log(msg):
    if len(journal) < MAX_JOURNAL:
        journal.append(msg)
    return len(journal) >= MAX_JOURNAL

def next_id():
    i = id_counter[0]
    id_counter[0] += 1
    return i

# Stack pour le backtracking : (index dans liste, bid, tentatives_déjà_essayées)
stack = []  # liste des bâtiments placés avec leur état
i = 0

# Essais par bâtiment (index dans liste) -> set de positions déjà essayées
essais = {}

while i < len(liste):
    if len(journal) >= MAX_JOURNAL:
        break

    b = liste[i]
    nom = b["Nom"]
    lon, lar = b["Longueur"], b["Largeur"]

    if i not in essais:
        essais[i] = set()
        stop = log(f"[ÉVALUATION] Bâtiment '{nom}' ({b['Type']}, {lon}x{lar})")
        if stop: break

    # Trouver les positions possibles non encore essayées
    if b["Type"] == "Neutre" and not stack:
        positions = cases_bord(terrain, lon, lar)
    else:
        positions = toutes_positions(terrain, lon, lar)

    positions_nouvelles = [p for p in positions if p not in essais[i]]

    # Vérifier contrainte : assez de cases pour le plus grand restant APRÈS placement
    restants = liste[i+1:]
    taille_max = plus_grand_restant(restants)

    position_choisie = None
    for pos in positions_nouvelles:
        essais[i].add(pos)
        r, c, pl, pl2 = pos
        placer(terrain, r, c, pl, pl2, 999)  # test
        libres = cases_libres(terrain)
        enlever(terrain, r, c, pl, pl2)
        terrain_test_ok = libres >= taille_max or taille_max == 0
        # Compter les cases de la forme max
        if terrain_test_ok:
            position_choisie = pos
            break

    if position_choisie:
        r, c, pl, pl2 = position_choisie
        bid = next_id()
        placer(terrain, r, c, pl, pl2, bid)
        placed[bid] = {"nom": nom, "r": r, "c": c, "lon": pl, "lar": pl2, "type": b["Type"]}
        stack.append((i, bid))
        stop = log(f"[PLACÉ] '{nom}' en ({r},{c}) orientation {pl}x{pl2}")
        if stop: break
        i += 1
    else:
        # Backtrack
        if stack:
            prev_i, prev_bid = stack.pop()
            prev_info = placed.pop(prev_bid)
            enlever(terrain, prev_info["r"], prev_info["c"], prev_info["lon"], prev_info["lar"])
            stop = log(f"[ENLEVÉ] '{prev_info['nom']}' de ({prev_info['r']},{prev_info['c']})")
            if stop: break
            # On retente le bâtiment précédent à une autre position
            i = prev_i
        else:
            # Aucun backtrack possible : bâtiment non plaçable
            stop = log(f"[NON PLACÉ] '{nom}' ne peut pas être placé")
            if stop: break
            non_places.append(b)
            i += 1

# Bâtiments restants non placés
non_places += liste[i:]

return terrain, placed, non_places, journal, rows, cols, batiments_info
```

# ─────────────────────────────────────────────

# GÉNÉRATION DU FICHIER EXCEL RÉSULTAT

# ─────────────────────────────────────────────

ORANGE = “FFAA44”
VERT = “44BB44”
GRIS = “AAAAAA”
ROUGE = “FF4444”
BLEU_CLAIR = “DDEEFF”
JAUNE = “FFFFAA”

def couleur_type(t):
if t == “Culturel”: return ORANGE
if t == “Producteur”: return VERT
if t == “Neutre”: return GRIS
return “FFFFFF”

def generer_excel(terrain_init, terrain_final, placed, non_places, journal,
batiments_info, rows, cols, batiments_orig):
wb = Workbook()

```
# ── Onglet 1 : Journal ──────────────────────────
ws_j = wb.active
ws_j.title = "Journal"
ws_j.append(["#", "Entrée"])
ws_j["A1"].font = Font(bold=True)
ws_j["B1"].font = Font(bold=True)
for idx, ligne in enumerate(journal, 1):
    ws_j.append([idx, ligne])
ws_j.column_dimensions["A"].width = 6
ws_j.column_dimensions["B"].width = 80

# ── Onglet 2 : Résultats culture ────────────────
ws_r = wb.create_sheet("Résultats")

culture_recue = calcul_culture(placed, batiments_info, rows, cols)
total_culture = sum(culture_recue.values())

# Regrouper par nom de bâtiment
culture_par_prod = {}
for bid, info in placed.items():
    if info["type"] == "Producteur":
        nom = info["nom"]
        b = batiments_info[nom]
        c_val = culture_recue.get(bid, 0)
        if nom not in culture_par_prod:
            culture_par_prod[nom] = {"culture": 0, "count": 0, "b": b}
        culture_par_prod[nom]["culture"] += c_val
        culture_par_prod[nom]["count"] += 1

ws_r.append(["Bâtiment (Producteur)", "Nb placés", "Culture totale reçue", "Boost atteint", "Production"])
for cell in ws_r[1]:
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", start_color=BLEU_CLAIR)

# Par production type
culture_guerison = 0
culture_nourriture = 0
culture_or = 0

for nom, data in culture_par_prod.items():
    b = data["b"]
    cult = data["culture"]
    boost = ""
    if b["Boost100"] and cult >= b["Boost100"]: boost = "100%"
    elif b["Boost50"] and cult >= b["Boost50"]: boost = "50%"
    elif b["Boost25"] and cult >= b["Boost25"]: boost = "25%"
    else: boost = "0%"
    ws_r.append([nom, data["count"], cult, boost, b["Production"]])

    prod = b["Production"].lower()
    if "guérison" in prod or "guerison" in prod:
        culture_guerison += cult
    elif "nourriture" in prod:
        culture_nourriture += cult
    elif "or" in prod:
        culture_or += cult

ws_r.append([])
ws_r.append(["Culture totale reçue (tous Producteurs)", "", total_culture])
ws_r.append([])
ws_r.append(["Culture reçue – Guérison", "", culture_guerison])
ws_r.append(["Culture reçue – Nourriture", "", culture_nourriture])
ws_r.append(["Culture reçue – Or", "", culture_or])

for col in ["A", "B", "C", "D", "E"]:
    ws_r.column_dimensions[col].width = 28

# ── Onglet 3 : Terrain ──────────────────────────
ws_t = wb.create_sheet("Terrain")

# Construire mapping bid -> info
bid_to_info = {}
for bid, info in placed.items():
    bid_to_info[bid] = info

thin = Side(style="thin", color="888888")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# Afficher le terrain
for ri in range(rows):
    for ci in range(cols):
        cell = ws_t.cell(row=ri + 1, column=ci + 1)
        val = terrain_final[ri][ci]
        cell.border = border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if val == 0:
            cell.fill = PatternFill("solid", start_color="222222")
            cell.value = ""
        elif val == 1:
            cell.fill = PatternFill("solid", start_color="EEEEEE")
            cell.value = ""
        else:
            info = bid_to_info.get(val)
            if info:
                b = batiments_info.get(info["nom"], {})
                couleur = couleur_type(info["type"])
                cell.fill = PatternFill("solid", start_color=couleur)
                # Afficher le nom uniquement sur la case haut-gauche
                if ri == info["r"] and ci == info["c"]:
                    cell.value = info["nom"]
                    cell.font = Font(bold=True, size=7)
                else:
                    cell.value = ""

# Taille des cellules
for ci in range(1, cols + 1):
    ws_t.column_dimensions[get_column_letter(ci)].width = 12
for ri in range(1, rows + 1):
    ws_t.row_dimensions[ri].height = 20

# Légende
ws_t.cell(row=rows + 2, column=1, value="Légende :").font = Font(bold=True)
for idx, (label, couleur) in enumerate([("Culturel", ORANGE), ("Producteur", VERT), ("Neutre", GRIS), ("Case occupée", "222222"), ("Case libre", "EEEEEE")]):
    c = ws_t.cell(row=rows + 3 + idx, column=1, value=label)
    c.fill = PatternFill("solid", start_color=couleur)
    c.font = Font(color="FFFFFF" if couleur == "222222" else "000000", bold=True)

# ── Onglet 4 : Non placés & stats ───────────────
ws_n = wb.create_sheet("Non Placés")
ws_n.append(["Bâtiment non placé", "Type", "Longueur", "Largeur", "Cases"])
for cell in ws_n[1]:
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", start_color=JAUNE)

cases_non_places = 0
for b in non_places:
    taille = b["Longueur"] * b["Largeur"]
    cases_non_places += taille
    ws_n.append([b["Nom"], b["Type"], b["Longueur"], b["Largeur"], taille])

libres = cases_libres(terrain_final)
ws_n.append([])
ws_n.append(["Cases non utilisées (terrain)", "", "", "", libres])
ws_n.append(["Cases représentant les bâtiments non placés", "", "", "", cases_non_places])

for col in ["A", "B", "C", "D", "E"]:
    ws_n.column_dimensions[col].width = 22

output = io.BytesIO()
wb.save(output)
output.seek(0)
return output
```

# ─────────────────────────────────────────────

# INTERFACE STREAMLIT

# ─────────────────────────────────────────────

uploaded = st.file_uploader(“📂 Choisir le fichier Excel d’entrée (.xlsx)”, type=[“xlsx”])

if uploaded:
try:
xls = pd.ExcelFile(uploaded)
st.success(f”Fichier chargé. Onglets détectés : {xls.sheet_names}”)

```
    df_terrain = pd.read_excel(xls, sheet_name=0, header=None)
    df_batiments = pd.read_excel(xls, sheet_name=1, header=0)

    terrain_init = lire_terrain(df_terrain)
    batiments = lire_batiments(df_batiments)

    rows = len(terrain_init)
    cols = len(terrain_init[0]) if rows > 0 else 0

    st.write(f"**Terrain :** {rows} lignes × {cols} colonnes — {sum(v for row in terrain_init for v in row if v == 1)} cases libres")
    st.write(f"**Bâtiments :** {sum(b['Quantite'] for b in batiments)} au total ({len(batiments)} types)")

    with st.expander("Aperçu des bâtiments"):
        st.dataframe(df_batiments)

    if st.button("🚀 Lancer le placement"):
        with st.spinner("Calcul en cours..."):
            terrain_final, placed, non_places, journal, rows, cols, batiments_info = algorithme_placement(terrain_init, batiments)

        st.success(f"✅ Placement terminé — {len(placed)} bâtiments placés, {len(non_places)} non placés")

        if len(journal) >= 1000:
            st.warning("⚠️ Le journal a atteint 1000 entrées — le script s'est arrêté.")

        with st.expander(f"📋 Journal ({len(journal)} entrées)"):
            for entry in journal:
                st.text(entry)

        # Générer Excel
        excel_output = generer_excel(
            terrain_init, terrain_final, placed, non_places, journal,
            batiments_info, rows, cols, batiments
        )

        st.download_button(
            label="⬇️ Télécharger le fichier résultat (.xlsx)",
            data=excel_output,
            file_name="resultat_placement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

except Exception as e:
    st.error(f"Erreur lors du traitement : {e}")
    import traceback
    st.code(traceback.format_exc())
```

with st.expander(“ℹ️ Format attendu du fichier Excel d’entrée”):
st.markdown(”””
**Onglet 1 – Terrain**

- Chaque ligne = une rangée de cases
- `1` = case libre, `0` = case occupée

**Onglet 2 – Bâtiments**  
Colonnes : `Nom`, `Longueur`, `Largeur`, `Quantite`, `Type`, `Culture`, `Rayonnement`, `Boost 25%`, `Boost 50%`, `Boost 100%`, `Production`

- `Type` = `Culturel`, `Producteur` ou `Neutre`
- Les bâtiments Neutres sont placés en priorité sur les bords
  “””)
