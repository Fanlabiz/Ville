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
        self.nom = row["Nom"]
        self.L = int(row["Longueur"])
        self.l = int(row["Largeur"])
        self.nombre = int(row["Nombre"])
        self.type = row["Type"]
        self.culture = row["Culture"]
        self.rayonnement = int(row["Rayonnement"])
        self.boost25 = row["Boost 25%"]
        self.boost50 = row["Boost 50%"]
        self.boost100 = row["Boost 100%"]
        self.production = row["Production"]
        self.quantite = row["Quantite"]

# =========================
# PRIORITE PROD
# =========================
def priorite_prod(p):
    ordre = {"Guérison":0, "Nourriture":1, "Or":2}
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

    # éviter bords pour culturels
    if type_bat == "Culturel":
        if x == 0 or y == 0 or x+L == h or y+l == w:
            score -= 100

    # favoriser centre
    score -= abs(h/2 - x) + abs(w/2 - y)

    return score


# =========================
# PLACEMENT AVANCE
# =========================
def placer_batiments(terrain, batiments):
    grille = terrain.copy().values

    objets = []
    for _, row in batiments.iterrows():
        for _ in range(int(row["Nombre"])):
            objets.append(Batiment(row))

    # tri intelligent
    objets.sort(key=lambda b: (
        0 if b.type == "Culturel" else 1,
        priorite_prod(b.production),
        -(b.L*b.l)
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
    for i in range(x-r, x+L+r):
        for j in range(y-r, y+l+r):
            if 0 <= i < shape[0] and 0 <= j < shape[1]:
                if not (x <= i < x+L and y <= j < y+l):
                    cells.add((i,j))
    return cells


def calcul(placements, shape):
    culture_sources = []
    producteurs = []

    for b, x, y, L, l in placements:
        if b.type == "Culturel":
            cells = zone_rayonnement(x,y,L,l,b.rayonnement,shape)
            culture_sources.append((cells,b.culture))
        elif b.type == "Producteur":
            producteurs.append((b,x,y,L,l))

    rows = []
    total_prod = {}

    for b,x,y,L,l in producteurs:
        culture = 0

        for cells,val in culture_sources:
            if any((i,j) in cells for i in range(x,x+L) for j in range(y,y+l)):
                culture += val

        boost = 0
        if culture >= b.boost100:
            boost = 100
        elif culture >= b.boost50:
            boost = 50
        elif culture >= b.boost25:
            boost = 25

        prod = b.quantite * (1 + boost/100)

        total_prod.setdefault(b.production, 0)
        total_prod[b.production] += prod

        rows.append([b.nom,b.type,b.production,x,y,culture,boost,prod])

    return pd.DataFrame(rows, columns=[
        "Nom","Type","Production","X","Y","Culture","Boost","Prod/h"
    ]), total_prod


# =========================
# EXPORT EXCEL AVANCE
# =========================
def export_excel(grille, placements, df, total_prod, non_places):

    with pd.ExcelWriter("resultat.xlsx", engine="xlsxwriter") as writer:
        workbook = writer.book

        # FEUILLE 1
        df.to_excel(writer, sheet_name="Batiments")

        # FEUILLE 2
        pd.DataFrame(list(total_prod.items()),
                     columns=["Production","Total/h"]
                     ).to_excel(writer, sheet_name="Production")

        # FEUILLE 3 - carte
        sheet = workbook.add_worksheet("Carte")

        fmt_cult = workbook.add_format({'bg_color':'orange'})
        fmt_prod = workbook.add_format({'bg_color':'green'})
        fmt_neutre = workbook.add_format({'bg_color':'gray'})

        for b,x,y,L,l in placements:
            fmt = fmt_neutre
            if b.type == "Culturel":
                fmt = fmt_cult
            elif b.type == "Producteur":
                fmt = fmt_prod

            for i in range(x,x+L):
                for j in range(y,y+l):
                    sheet.write(i,j,f"{b.nom}",fmt)

        # FEUILLE 4
        pd.DataFrame([b.nom for b in non_places],
                     columns=["Non placés"]
                     ).to_excel(writer, sheet_name="Non_places")

        # stats
        total_cases = grille.size
        used = np.sum(grille != 1)

        stats = pd.DataFrame({
            "Cases libres":[total_cases-used],
            "Cases non utilisées":[np.sum(grille==1)]
        })
        stats.to_excel(writer, sheet_name="Stats")


# =========================
# MAIN
# =========================
def run(file):
    terrain, batiments, existant = load_excel(file)

    grille, placements, non_places = placer_batiments(terrain, batiments)

    df, total_prod = calcul(placements, grille.shape)

    export_excel(grille, placements, df, total_prod, non_places)

    print("OK fichier généré")
