# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import copy

st.set_page_config(page_title=“Placement de Batiments”, layout=“wide”)
st.title(“Placement de Batiments sur Terrain”)

def lire_terrain(df):
terrain = []
for _, row in df.iterrows():
ligne = [int(v) for v in row if pd.notna(v)]
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

def est_bord(r, c, lon, lar, rows, cols):
return r == 0 or c == 0 or r + lon == rows or c + lar == cols

def placer(terrain, r, c, lon, lar, bid):
for i in range(r, r + lon):
for j in range(c, c + lar):
terrain[i][j] = bid

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

def cases_bord_positions(terrain, lon, lar):
rows = len(terrain)
cols = len(terrain[0]) if rows > 0 else 0
return [p for p in toutes_positions(terrain, lon, lar)
if est_bord(p[0], p[1], p[2], p[3], rows, cols)]

def zone_rayonnement(r, c, lon, lar, ray, rows, cols):
cases = set()
for dr in range(-ray, lon + ray):
for dc in range(-ray, lar + ray):
ri, ci = r + dr, c + dc
if 0 <= ri < rows and 0 <= ci < cols:
if not (0 <= dr < lon and 0 <= dc < lar):
cases.add((ri, ci))
return cases

def calcul_culture(placed, batiments_info, rows, cols):
culture_recue = {bid: 0 for bid, info in placed.items() if info[“type”] == “Producteur”}
for bid_c, info_c in placed.items():
b_c = batiments_info[info_c[“nom”]]
if b_c[“Type”] != “Culturel”:
continue
ray = b_c[“Rayonnement”]
zone = zone_rayonnement(info_c[“r”], info_c[“c”], info_c[“lon”], info_c[“lar”], ray, rows, cols)
for bid_p, info_p in placed.items():
if bid_p not in culture_recue:
continue
found = False
for ir in range(info_p[“r”], info_p[“r”] + info_p[“lon”]):
for ic in range(info_p[“c”], info_p[“c”] + info_p[“lar”]):
if (ir, ic) in zone:
culture_recue[bid_p] += b_c[“Culture”]
found = True
break
if found:
break
return culture_recue

def taille_batiment(b):
return b[“Longueur”] * b[“Largeur”]

def preparer_liste(batiments):
liste = []
for b in batiments:
for _ in range(b[“Quantite”]):
liste.append(copy.deepcopy(b))
return liste

def trier_liste(liste):
neutres = sorted([b for b in liste if b[“Type”] == “Neutre”], key=taille_batiment, reverse=True)
autres  = sorted([b for b in liste if b[“Type”] != “Neutre”], key=taille_batiment, reverse=True)
return neutres + autres

def algorithme_placement(terrain_init, batiments):
terrain = copy.deepcopy(terrain_init)
rows = len(terrain)
cols = len(terrain[0]) if rows > 0 else 0
liste = trier_liste(preparer_liste(batiments))
batiments_info = {b[“Nom”]: b for b in batiments}
placed, non_places, journal = {}, [], []
id_counter = [2]

```
def log(msg):
    if len(journal) < 1000:
        journal.append(msg)
    return len(journal) >= 1000

def next_id():
    val = id_counter[0]
    id_counter[0] += 1
    return val

stack, essais, i = [], {}, 0

while i < len(liste):
    if len(journal) >= 1000:
        break
    b = liste[i]
    nom = b["Nom"]
    lon = b["Longueur"]
    lar = b["Largeur"]
    if i not in essais:
        essais[i] = set()
        if log("[EVALUATION] Batiment '" + nom + "' (" + b["Type"] + ", " + str(lon) + "x" + str(lar) + ")"):
            break

    if b["Type"] == "Neutre" and not stack:
        positions = cases_bord_positions(terrain, lon, lar)
    else:
        positions = toutes_positions(terrain, lon, lar)

    positions_nouvelles = [p for p in positions if p not in essais[i]]
    restants = liste[i + 1:]
    taille_max = max((taille_batiment(b2) for b2 in restants), default=0)

    position_choisie = None
    for pos in positions_nouvelles:
        essais[i].add(pos)
        r, c, pl, pl2 = pos
        placer(terrain, r, c, pl, pl2, 999)
        ok = cases_libres(terrain) >= taille_max or taille_max == 0
        enlever(terrain, r, c, pl, pl2)
        if ok:
            position_choisie = pos
            break

    if position_choisie:
        r, c, pl, pl2 = position_choisie
        bid = next_id()
        placer(terrain, r, c, pl, pl2, bid)
        placed[bid] = {"nom": nom, "r": r, "c": c, "lon": pl, "lar": pl2, "type": b["Type"]}
        stack.append((i, bid))
        if log("[PLACE] '" + nom + "' en (" + str(r) + "," + str(c) + ") " + str(pl) + "x" + str(pl2)):
            break
        i += 1
    else:
        if stack:
            prev_i, prev_bid = stack.pop()
            prev_info = placed.pop(prev_bid)
            enlever(terrain, prev_info["r"], prev_info["c"], prev_info["lon"], prev_info["lar"])
            if log("[ENLEVE] '" + prev_info["nom"] + "' de (" + str(prev_info["r"]) + "," + str(prev_info["c"]) + ")"):
                break
            i = prev_i
        else:
            if log("[NON PLACE] '" + nom + "' ne peut pas etre place"):
                break
            non_places.append(b)
            i += 1

non_places += liste[i:]
return terrain, placed, non_places, journal, rows, cols, batiments_info
```

ORANGE = “FFAA44”
VERT   = “44BB44”
GRIS   = “AAAAAA”
JAUNE  = “FFFFAA”
BLEU   = “DDEEFF”

def couleur_type(t):
if t == “Culturel”:   return ORANGE
if t == “Producteur”: return VERT
if t == “Neutre”:     return GRIS
return “FFFFFF”

def generer_excel(terrain_final, placed, non_places, journal, batiments_info, rows, cols):
wb = Workbook()

```
ws_j = wb.active
ws_j.title = "Journal"
ws_j.append(["#", "Entree"])
ws_j["A1"].font = Font(bold=True)
ws_j["B1"].font = Font(bold=True)
for idx, ligne in enumerate(journal, 1):
    ws_j.append([idx, ligne])
ws_j.column_dimensions["A"].width = 6
ws_j.column_dimensions["B"].width = 80

ws_r = wb.create_sheet("Resultats")
culture_recue = calcul_culture(placed, batiments_info, rows, cols)
total_culture = sum(culture_recue.values())
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

ws_r.append(["Batiment (Producteur)", "Nb places", "Culture totale recue", "Boost atteint", "Production"])
for cell in ws_r[1]:
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", start_color=BLEU)

culture_guerison = 0
culture_nourriture = 0
culture_or = 0

for nom, data in culture_par_prod.items():
    b = data["b"]
    cult = data["culture"]
    boost = "0%"
    if b["Boost100"] and cult >= b["Boost100"]:  boost = "100%"
    elif b["Boost50"] and cult >= b["Boost50"]:  boost = "50%"
    elif b["Boost25"] and cult >= b["Boost25"]:  boost = "25%"
    ws_r.append([nom, data["count"], cult, boost, b["Production"]])
    prod = b["Production"].lower()
    if "guerison" in prod or "gu" in prod and "rison" in prod:
        culture_guerison += cult
    elif "nourriture" in prod:
        culture_nourriture += cult
    elif "or" in prod:
        culture_or += cult

ws_r.append([])
ws_r.append(["Culture totale recue (tous Producteurs)", "", total_culture])
ws_r.append([])
ws_r.append(["Culture recue - Guerison", "", culture_guerison])
ws_r.append(["Culture recue - Nourriture", "", culture_nourriture])
ws_r.append(["Culture recue - Or", "", culture_or])
for col in ["A","B","C","D","E"]:
    ws_r.column_dimensions[col].width = 28

ws_t = wb.create_sheet("Terrain")
thin = Side(style="thin", color="888888")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
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
            info = placed.get(val)
            if info:
                cell.fill = PatternFill("solid", start_color=couleur_type(info["type"]))
                if ri == info["r"] and ci == info["c"]:
                    cell.value = info["nom"]
                    cell.font = Font(bold=True, size=7)

for ci in range(1, cols + 1):
    ws_t.column_dimensions[get_column_letter(ci)].width = 12
for ri in range(1, rows + 1):
    ws_t.row_dimensions[ri].height = 20

ws_t.cell(row=rows + 2, column=1, value="Legende :").font = Font(bold=True)
for idx, (label, couleur) in enumerate([
    ("Culturel (orange)", ORANGE), ("Producteur (vert)", VERT),
    ("Neutre (gris)", GRIS), ("Case occupee", "222222"), ("Case libre", "EEEEEE")
]):
    c = ws_t.cell(row=rows + 3 + idx, column=1, value=label)
    c.fill = PatternFill("solid", start_color=couleur)
    c.font = Font(color="FFFFFF" if couleur == "222222" else "000000", bold=True)

ws_n = wb.create_sheet("Non Places")
ws_n.append(["Batiment non place", "Type", "Longueur", "Largeur", "Cases"])
for cell in ws_n[1]:
    cell.font = Font(bold=True)
    cell.fill = PatternFill("solid", start_color=JAUNE)
cases_np = 0
for b in non_places:
    t = b["Longueur"] * b["Largeur"]
    cases_np += t
    ws_n.append([b["Nom"], b["Type"], b["Longueur"], b["Largeur"], t])
libres = cases_libres(terrain_final)
ws_n.append([])
ws_n.append(["Cases non utilisees (terrain)", "", "", "", libres])
ws_n.append(["Cases des batiments non places", "", "", "", cases_np])
for col in ["A","B","C","D","E"]:
    ws_n.column_dimensions[col].width = 24

output = io.BytesIO()
wb.save(output)
output.seek(0)
return output
```

uploaded = st.file_uploader(“Choisir le fichier Excel d’entree (.xlsx)”, type=[“xlsx”])

if uploaded:
try:
xls = pd.ExcelFile(uploaded)
st.success(“Fichier charge. Onglets : “ + str(xls.sheet_names))
df_terrain   = pd.read_excel(xls, sheet_name=0, header=None)
df_batiments = pd.read_excel(xls, sheet_name=1, header=0)
terrain_init = lire_terrain(df_terrain)
batiments    = lire_batiments(df_batiments)
rows = len(terrain_init)
cols = len(terrain_init[0]) if rows > 0 else 0
nb_libres = sum(v for row in terrain_init for v in row if v == 1)
st.write(“Terrain : “ + str(rows) + “ x “ + str(cols) + “ - “ + str(nb_libres) + “ cases libres”)
st.write(“Batiments : “ + str(sum(b[“Quantite”] for b in batiments)) + “ au total”)
with st.expander(“Apercu des batiments”):
st.dataframe(df_batiments)

```
    if st.button("Lancer le placement"):
        with st.spinner("Calcul en cours..."):
            terrain_final, placed, non_places, journal, rows, cols, batiments_info = \
                algorithme_placement(terrain_init, batiments)
        st.success(str(len(placed)) + " batiments places, " + str(len(non_places)) + " non places")
        if len(journal) >= 1000:
            st.warning("Journal a atteint 1000 entrees - arret du script.")
        with st.expander("Journal (" + str(len(journal)) + " entrees)"):
            for entry in journal:
                st.text(entry)
        excel_output = generer_excel(terrain_final, placed, non_places, journal, batiments_info, rows, cols)
        st.download_button(
            label="Telecharger le fichier resultat (.xlsx)",
            data=excel_output,
            file_name="resultat_placement.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
except Exception as e:
    st.error("Erreur : " + str(e))
    import traceback
    st.code(traceback.format_exc())
```

with st.expander(“Format attendu du fichier Excel”):
st.markdown(
“Onglet 1 - Terrain : 1 = case libre, 0 = occupee.\n\n”
“Onglet 2 - Batiments : colonnes Nom, Longueur, Largeur, Quantite, “
“Type (Culturel/Producteur/Neutre), Culture, Rayonnement, “
“Boost 25%, Boost 50%, Boost 100%, Production.”
)
