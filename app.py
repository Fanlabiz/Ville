import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import itertools
import copy
from collections import defaultdict
import io

st.title("Optimiseur de Placement de Bâtiments - Ville Fusion")

uploaded_file = st.file_uploader("Charge ton fichier Excel (Ville fusion.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # === 1. Lecture des données ===
    wb = load_workbook(uploaded_file, data_only=True)
    terrain_sheet = wb["Terrain"]
    bat_sheet = wb["Batiments"]

    # Lecture du terrain
    terrain_data = []
    for row in terrain_sheet.iter_rows(min_row=1, values_only=True):
        terrain_data.append(list(row))

    terrain = np.array(terrain_data, dtype=object)
    rows, cols = terrain.shape

    # Lecture des bâtiments
    bat_df = pd.read_excel(uploaded_file, sheet_name="Batiments", header=0)
    # On ajoute la colonne Priorite si elle existe (tu l'as mentionnée)
    if "Priorite" not in bat_df.columns:
        bat_df["Priorite"] = 0

    # === 2. Parser les bâtiments existants sur le terrain ===
    existing_buildings = []
    building_map = {}  # (r, c) -> nom du bâtiment

    for r in range(rows):
        for c in range(cols):
            cell = terrain[r, c]
            if isinstance(cell, str) and cell not in ["X", "", None]:
                existing_buildings.append({
                    "nom": cell.strip(),
                    "row": r,
                    "col": c
                })
                building_map[(r, c)] = cell.strip()

    # === 3. Création des listes de bâtiments à placer ===
    to_place = []
    for _, row in bat_df.iterrows():
        nom = row["Nom"]
        longueur = int(row["Longueur"])
        largeur = int(row["Largeur"])
        nombre = int(row["Nombre"])
        typ = row["Type"]
        culture = int(row["Culture"]) if pd.notna(row["Culture"]) else 0
        rayonnement = int(row["Rayonnement"]) if pd.notna(row["Rayonnement"]) else 0
        boost25 = int(row["Boost 25%"]) if pd.notna(row["Boost 25%"]) else 0
        boost50 = int(row["Boost 50%"]) if pd.notna(row["Boost 50%"]) else 0
        boost100 = int(row["Boost 100%"]) if pd.notna(row["Boost 100%"]) else 0
        production = row["Production"] if pd.notna(row["Production"]) else "Rien"
        quantite = int(row["Quantite"]) if pd.notna(row["Quantite"]) else 0
        priorite = int(row["Priorite"]) if "Priorite" in row else 0

        for i in range(nombre):
            to_place.append({
                "nom": nom,
                "long": longueur,
                "larg": largeur,
                "type": typ,
                "culture": culture,
                "rayonnement": rayonnement,
                "boost25": boost25,
                "boost50": boost50,
                "boost100": boost100,
                "production": production,
                "quantite_base": quantite,
                "priorite": priorite,
                "orientation": "H"  # H ou V
            })

    st.success(f"Terrain chargé : {rows}x{cols} cases | {len(to_place)} bâtiments à placer/optimiser")

    # === 4. Fonction d'optimisation (heuristique gloutonne + swap) ===
    def can_place(grid, r, c, h, w, orientation="H"):
        if orientation == "H":
            height, width = h, w
        else:
            height, width = w, h
        if r + height > grid.shape[0] or c + width > grid.shape[1]:
            return False
        for i in range(height):
            for j in range(width):
                if grid[r+i, c+j] not in [None, ""]:
                    return False
        return True

    def place_building(grid, r, c, building, orientation="H"):
        if orientation == "H":
            height, width = building["long"], building["larg"]
        else:
            height, width = building["larg"], building["long"]
        for i in range(height):
            for j in range(width):
                grid[r+i, c+j] = building["nom"]

    # Grille vide (on ignore les X pour l'instant, on les garde comme obstacles)
    grid = np.full_like(terrain, None, dtype=object)
    for r in range(rows):
        for c in range(cols):
            if terrain[r, c] == "X":
                grid[r, c] = "X"

    # Placement initial glouton (priorité aux bâtiments culturels importants + grands)
    cultural_buildings = [b for b in to_place if b["type"] == "Culturel"]
    producer_buildings = [b for b in to_place if b["type"] == "Producteur"]

    # Trier culturels par culture descendante puis rayonnement
    cultural_buildings.sort(key=lambda b: (-b["culture"], -b["rayonnement"]))

    placed = []
    for b in cultural_buildings + producer_buildings:
        placed_flag = False
        for orient in ["H", "V"]:
            for r in range(rows):
                for c in range(cols):
                    if can_place(grid, r, c, b["long"], b["larg"], orient):
                        place_building(grid, r, c, b, orient)
                        b["row"] = r
                        b["col"] = c
                        b["orientation"] = orient
                        placed.append(b)
                        placed_flag = True
                        break
                if placed_flag: break
            if placed_flag: break
        if not placed_flag:
            st.warning(f"Impossible de placer {b['nom']}")

    # === Calcul de la culture reçue et des boosts ===
    def get_culture_received(producer, placed_buildings, grid):
        total_culture = 0
        pr, pc = producer["row"], producer["col"]
        ph = producer["long"] if producer["orientation"] == "H" else producer["larg"]
        pw = producer["larg"] if producer["orientation"] == "H" else producer["long"]

        for b in placed_buildings:
            if b["type"] != "Culturel" or b["rayonnement"] == 0:
                continue
            br, bc = b["row"], b["col"]
            bh = b["long"] if b["orientation"] == "H" else b["larg"]
            bw = b["larg"] if b["orientation"] == "H" else b["long"]

            # Zone de rayonnement (Manhattan ou Chebyshev ? On prend Chebyshev pour "bande autour")
            max_dist = b["rayonnement"]
            for i in range(ph):
                for j in range(pw):
                    for di in range(-max_dist, max_dist+1):
                        for dj in range(-max_dist, max_dist+1):
                            if abs(di) <= max_dist and abs(dj) <= max_dist:  # Chebyshev
                                if (br + di == pr + i) and (bc + dj == pc + j):
                                    total_culture += b["culture"]
                                    break  # on compte une fois par bâtiment culturel
        return total_culture

    # Appliquer les boosts
    production_types = ["Guerison", "Nourriture", "Or"]
    for b in placed:
        if b["type"] == "Producteur" and b["production"] != "Rien":
            culture_recue = get_culture_received(b, placed, grid)
            b["culture_recue"] = culture_recue
            if culture_recue >= b["boost100"]:
                b["boost"] = 100
            elif culture_recue >= b["boost50"]:
                b["boost"] = 50
            elif culture_recue >= b["boost25"]:
                b["boost"] = 25
            else:
                b["boost"] = 0
            b["production_finale"] = b["quantite_base"] * (1 + b["boost"]/100)
        else:
            b["culture_recue"] = 0
            b["boost"] = 0
            b["production_finale"] = 0

    # === Calcul des totaux par type de production ===
    totals = defaultdict(lambda: {"base": 0, "final": 0, "culture_tot": 0})
    for b in placed:
        if b["type"] == "Producteur" and b["production"] != "Rien":
            prod_type = b["production"]
            totals[prod_type]["base"] += b["quantite_base"]
            totals[prod_type]["final"] += b["production_finale"]
            totals[prod_type]["culture_tot"] += b["culture_recue"]

    # === Création du fichier Excel de sortie ===
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Feuille 1 : Liste des bâtiments placés
        placed_df = pd.DataFrame([{
            "Nom": b["nom"],
            "Type": b["type"],
            "Production": b.get("production", ""),
            "Quantité base": b.get("quantite_base", 0),
            "Culture reçue": b.get("culture_recue", 0),
            "Boost %": b.get("boost", 0),
            "Production finale/heure": round(b.get("production_finale", 0), 2),
            "Coordonnées (haut-gauche)": f"({b.get('row',0)}, {b.get('col',0)})",
            "Orientation": b.get("orientation", "H")
        } for b in placed])
        placed_df.to_excel(writer, sheet_name="Batiments_Places", index=False)

        # Feuille 2 : Totaux par production
        totals_df = pd.DataFrame([{
            "Type Production": k,
            "Production base/heure": v["base"],
            "Production finale/heure": round(v["final"], 2),
            "Gain/Perte": round(v["final"] - v["base"], 2),
            "Culture totale reçue": v["culture_tot"]
        } for k, v in totals.items()])
        totals_df.to_excel(writer, sheet_name="Totaux_Production", index=False)

    output.seek(0)

    st.download_button(
        label="📥 Télécharger le fichier Excel optimisé",
        data=output,
        file_name="Ville_Optimisee_Resultat.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Affichage terrain simplifié dans Streamlit
    st.subheader("Aperçu du terrain optimisé (texte)")
    st.text(str(grid))

else:
    st.info("Charge ton fichier Excel pour lancer l'optimisation.")
