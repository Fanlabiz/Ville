import streamlit as st
import pandas as pd
import numpy as np

# =========================
# LECTURE
# =========================
def load_excel(file):
    terrain = pd.read_excel(file, sheet_name=0, header=None)
    batiments = pd.read_excel(file, sheet_name=1)

    try:
        existant = pd.read_excel(file, sheet_name=2, header=None)
    except:
        existant = None

    return terrain, batiments, existant


# =========================
# CLASSE
# =========================
class Batiment:
    def __init__(self, row):
        self.nom = str(row.get("Nom", "Inconnu"))

        self.L = int(row.get("Longueur", 1) or 1)
        self.l = int(row.get("Largeur", 1) or 1)
        self.nombre = int(row.get("Nombre", 1) or 1)

        self.type = str(row.get("Type", "Neutre"))

        self.culture = float(row.get("Culture", 0) or 0)
        self.rayonnement = int(row.get("Rayonnement", 0) or 0)

        self.boost25 = float(row.get("Boost 25%", 0) or 0)
        self.boost50 = float(row.get("Boost 50%", 0) or 0)
        self.boost100 = float(row.get("Boost 100%", 0) or 0)

        self.production = str(row.get("Production", "Rien"))
        self.quantite = float(row.get("Quantite", 0) or 0)


# =========================
# PRIORITE PROD
# =========================
def priorite_prod(p):
    ordre = {"Guérison": 0, "Nourriture": 1, "Or": 2}
    return ordre.get(p, 3)


# =========================
# OUTILS
# =========================
def peut_placer(grille, x, y, L, l):
    h, w = grille.shape
    if x + L > h or y + l > w:
        return False
    return np.all(grille[x:x+L, y:y+l] == 1)


def placer(grille, x, y, L, l, val):
    grille[x:x+L, y:y+l] = val


def score_position(grille, x, y, L, l, type_bat):
    h, w = grille.shape
    score = 0

    if type_bat == "Culturel":
        if x == 0 or y == 0 or x+L == h or y+l == w:
            score -= 100

    score -= abs(h/2 - x) + abs(w/2 - y)

    return score


# =========================
# PLACEMENT
# =========================
def placer_batiments(terrain, batiments):
    grille = terrain.copy().values

    objets = []
    for _, row in batiments.iterrows():
        for _ in range(int(row.get("Nombre", 1) or 1)):
            objets.append(Batiment(row))

    objets.sort(key=lambda b: (
        0 if b.type == "Culturel" else 1,
        priorite_prod(b.production),
        -(b.L * b.l)
    ))

    placements = []
    non_places = []

    for b in objets:
        best = None
        best_score = -1e9

        for i in range(grille.shape[0]):
            for j in range(grille.shape[1]):
                for (L, l) in [(b.L, b.l), (b.l, b.L)]:

                    if peut_placer(grille, i, j, L, l):
                        s = score_position(grille, i, j, L, l, b.type)

                        if s > best_score:
                            best_score = s
                            best = (i, j, L, l)

        if best:
            i, j, L, l = best
            placer(grille, i, j, L, l, 2)
            placements.append((b, i, j, L, l))
        else:
            non_places.append(b)

    return grille, placements, non_places


# =========================
# CALCUL CULTURE
# =========================
def zone_rayonnement(x, y, L, l, r, shape):
    cells = set()
    for i in range(x - r, x + L + r):
        for j in range(y - r, y + l + r):
            if 0 <= i < shape[0] and 0 <= j < shape[1]:
                if not (x <= i < x+L and y <= j < y+l):
                    cells.add((i, j))
    return cells


def calcul(placements, shape):
    culture_sources = []
    producteurs = []

    for b, x, y, L, l in placements:
        if b.type == "Culturel":
            cells = zone_rayonnement(x, y, L, l, b.rayonnement, shape)
            culture_sources.append((cells, b.culture))
        elif b.type == "Producteur":
            producteurs.append((b, x, y, L, l))

    rows = []
    total_prod = {}

    for b, x, y, L, l in producteurs:
        culture = 0

        for cells, val in culture_sources:
            if any((i, j) in cells for i in range(x, x+L) for j in range(y, y+l)):
                culture += val

        boost = 0
        if culture >= b.boost100:
            boost = 100
        elif culture >= b.boost50:
            boost = 50
        elif culture >= b.boost25:
            boost = 25

        prod = b.quantite * (1 + boost / 100)

        total_prod.setdefault(b.production, 0)
        total_prod[b.production] += prod

        rows.append([
            b.nom, b.type, b.production, x, y, culture, boost, prod
        ])

    df = pd.DataFrame(rows, columns=[
        "Nom", "Type", "Production", "X", "Y",
        "Culture", "Boost", "Prod/h"
    ])

    return df, total_prod


# =========================
# EXPORT EXCEL
# =========================
def export_excel(grille, placements, df, total_prod, non_places):

    output_file = "resultat.xlsx"

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

        df.to_excel(writer, sheet_name="Batiments", index=False)

        pd.DataFrame(list(total_prod.items()),
                     columns=["Production", "Total/h"]
                     ).to_excel(writer, sheet_name="Production", index=False)

        sheet = workbook.add_worksheet("Carte")

        fmt_cult = workbook.add_format({'bg_color': 'orange'})
        fmt_prod = workbook.add_format({'bg_color': 'green'})
        fmt_neutre = workbook.add_format({'bg_color': 'gray'})

        for b, x, y, L, l in placements:
            fmt = fmt_neutre
            if b.type == "Culturel":
                fmt = fmt_cult
            elif b.type == "Producteur":
                fmt = fmt_prod

            for i in range(x, x+L):
                for j in range(y, y+l):
                    sheet.write(i, j, f"{b.nom}", fmt)

        pd.DataFrame([b.nom for b in non_places],
                     columns=["Non placés"]
                     ).to_excel(writer, sheet_name="Non_places", index=False)

        total_cases = grille.size
        used = np.sum(grille != 1)

        stats = pd.DataFrame({
            "Cases utilisées": [used],
            "Cases libres": [np.sum(grille == 1)]
        })

        stats.to_excel(writer, sheet_name="Stats", index=False)

    return output_file


# =========================
# MAIN LOGIQUE
# =========================
def run(file):
    terrain, batiments, existant = load_excel(file)

    grille, placements, non_places = placer_batiments(terrain, batiments)

    df, total_prod = calcul(placements, grille.shape)

    output_file = export_excel(grille, placements, df, total_prod, non_places)

    return output_file


# =========================
# STREAMLIT UI
# =========================
st.title("Optimisation de ville")

file = st.file_uploader("Upload ton fichier Excel")

if file:
    try:
        output_file = run(file)

        with open(output_file, "rb") as f:
            st.download_button(
                "Télécharger le résultat",
                f,
                "resultat.xlsx"
            )

        st.success("Optimisation terminée ✅")

    except Exception as e:
        st.error("Erreur détectée ❌")
        st.exception(e)
