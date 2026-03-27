import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
import streamlit as st
import io
from collections import defaultdict

# ====================== CLASSE BATIMENT (corrigée) ======================
class Batiment:
    def __init__(self, row):
        self.nom = str(row['Nom']).strip() if pd.notna(row.get('Nom')) else "Inconnu"
        self.long = int(row['Longueur']) if pd.notna(row.get('Longueur')) else 1
        self.larg = int(row['Largeur']) if pd.notna(row.get('Largeur')) else 1
        self.nombre = int(row['Nombre']) if pd.notna(row.get('Nombre')) else 1
        self.typ = str(row['Type']).strip() if pd.notna(row.get('Type')) else "Neutre"
        
        self.culture = int(row.get('Culture', 0)) if pd.notna(row.get('Culture')) else 0
        self.rayon = int(row.get('Rayonnement', 0)) if pd.notna(row.get('Rayonnement')) else 0
        
        # Gestion robuste des colonnes de boost (plusieurs variantes possibles)
        def get_boost(col_names):
            for col in col_names:
                val = row.get(col)
                if val is not None and pd.notna(val):
                    try:
                        return int(val)
                    except:
                        return 0
            return 0
        
        self.boost25 = get_boost(['Boost 25%', 'Boost25', 'Boost 25 %', 'Boost25%'])
        self.boost50 = get_boost(['Boost 50%', 'Boost50', 'Boost 50 %', 'Boost50%'])
        self.boost100 = get_boost(['Boost 100%', 'Boost100', 'Boost 100 %', 'Boost100%'])
        
        self.prod_type = str(row.get('Production', 'Rien')).strip() if pd.notna(row.get('Production')) else 'Rien'
        self.quantite = int(row.get('Quantite', 0)) if pd.notna(row.get('Quantite')) else 0
        
        self.instances = []   # (r, c, orient, culture_recue, boost_pct)

    def surface(self):
        return self.long * self.larg

# ====================== LECTURE EXCEL ======================
def lire_excel(uploaded_file):
    wb = load_workbook(uploaded_file, data_only=True)
    df_terrain = pd.read_excel(uploaded_file, sheet_name='Terrain', header=None)
    terrain = df_terrain.to_numpy()

    df_bats = pd.read_excel(uploaded_file, sheet_name='Batiments')
    batiments = []
    for _, row in df_bats.iterrows():
        if pd.notna(row.get('Nom')) and str(row.get('Nom')).strip():
            batiments.append(Batiment(row))

    # Gestion du sheet "Actuel"
    if 'Actuel' in wb.sheetnames:
        df_actuel = pd.read_excel(uploaded_file, sheet_name='Actuel', header=None)
        actuel = df_actuel.to_numpy()
        # Marque comme occupé les cases déjà utilisées
        terrain = np.where((actuel.astype(str) == 'X') | (actuel == 0), 0, terrain)

    return terrain, batiments

# ====================== (Le reste du code reste identique aux améliorations précédentes) ======================
# ... [Je garde le code complet pour que tu puisses copier-coller tout d'un coup]

def peut_placer(terrain, r, c, h, w):
    if r < 0 or c < 0 or r + h > terrain.shape[0] or c + w > terrain.shape[1]:
        return False
    return np.all(terrain[r:r+h, c:c+w] == 1)

def placer(terrain, r, c, h, w):
    terrain[r:r+h, c:c+w] = 0
    return (r, c)

def score_position(terrain, r, c, h, w, bat, is_neutre=False, is_culturel=False):
    if not peut_placer(terrain, r, c, h, w):
        return -999999
    centre_r = r + h / 2
    centre_c = c + w / 2
    score = 1000 - (abs(centre_r - terrain.shape[0]/2) + abs(centre_c - terrain.shape[1]/2)) * 5

    if is_neutre:
        if r == 0 or r + h == terrain.shape[0] or c == 0 or c + w == terrain.shape[1]:
            score += 10000
    if is_culturel:
        free_in_ray = 0
        for i in range(max(0, r - bat.rayon), min(terrain.shape[0], r + h + bat.rayon)):
            for j in range(max(0, c - bat.rayon), min(terrain.shape[1], c + w + bat.rayon)):
                if terrain[i, j] == 1:
                    free_in_ray += 1
        score += free_in_ray * 10
    return score

def trouver_meilleure_position(terrain, bat, is_neutre=False, is_culturel=False, priorite_guerrison=False):
    best_score = -999999
    best_pos = None
    best_orient = 'H'
    for orient in ['H', 'V']:
        h = bat.larg if orient == 'V' else bat.long
        w = bat.long if orient == 'V' else bat.larg
        for i in range(terrain.shape[0] - h + 1):
            for j in range(terrain.shape[1] - w + 1):
                score = score_position(terrain, i, j, h, w, bat, is_neutre, is_culturel)
                if priorite_guerrison and bat.prod_type == "Guerison":
                    score += 20000
                if score > best_score:
                    best_score = score
                    best_pos = (i, j)
                    best_orient = orient
    return best_pos, best_orient

def optimiser_placement(terrain_orig, batiments):
    terrain = terrain_orig.copy()
    neutres = [b for b in batiments if b.typ == "Neutre"]
    culturels = [b for b in batiments if b.typ == "Culturel"]
    producteurs = [b for b in batiments if b.typ == "Producteur"]

    for grp in [neutres, culturels, producteurs]:
        grp.sort(key=lambda b: b.surface(), reverse=True)

    placed_list = []
    non_places = []

    # 1. Neutres sur les bords
    for bat in neutres:
        for _ in range(bat.nombre):
            pos, orient = trouver_meilleure_position(terrain, bat, is_neutre=True)
            if pos:
                h = bat.larg if orient == 'V' else bat.long
                w = bat.long if orient == 'V' else bat.larg
                placer(terrain, pos[0], pos[1], h, w)
                bat.instances.append((pos[0], pos[1], orient, 0, 0))
                placed_list.append((bat, pos[0], pos[1], orient, h, w))
            else:
                non_places.append(bat)
                break

    # 2. Placement alterné Culturel / Producteur
    all_non_neutres = sorted(culturels + producteurs, key=lambda b: b.surface(), reverse=True)
    i = 0
    while i < len(all_non_neutres):
        bat = all_non_neutres[i]
        is_cult = bat.typ == "Culturel"
        pos, orient = trouver_meilleure_position(terrain, bat, is_culturel=is_cult,
                                                 priorite_guerrison=(bat.prod_type == "Guerison"))
        if pos:
            h = bat.larg if orient == 'V' else bat.long
            w = bat.long if orient == 'V' else bat.larg
            placer(terrain, pos[0], pos[1], h, w)
            bat.instances.append((pos[0], pos[1], orient, 0, 0))
            placed_list.append((bat, pos[0], pos[1], orient, h, w))
        else:
            non_places.append(bat)
        i += 1

    return terrain, placed_list, non_places

def calcul_culture_boost(placed_list, batiments, terrain):
    culturels = [p for p in placed_list if p[0].typ == "Culturel"]

    for bat in batiments:
        if bat.typ != "Producteur":
            continue
        for idx, (r, c, orient, _, _) in enumerate(bat.instances):
            h = bat.larg if orient == 'V' else bat.long
            w = bat.long if orient == 'V' else bat.larg
            total_culture = 0
            seen = set()

            for cult_bat, cr, cc, corient, ch, cw in culturels:
                cult_key = id(cult_bat)
                if cult_key in seen:
                    continue
                for i in range(r, r + h):
                    for j in range(c, c + w):
                        for ci in range(cr - cult_bat.rayon, cr + ch + cult_bat.rayon):
                            for cj in range(cc - cult_bat.rayon, cc + cw + cult_bat.rayon):
                                if 0 <= ci < terrain.shape[0] and 0 <= cj < terrain.shape[1]:
                                    if abs(ci - i) <= cult_bat.rayon and abs(cj - j) <= cult_bat.rayon:
                                        total_culture += cult_bat.culture
                                        seen.add(cult_key)
                                        break
            culture_recue = total_culture

            boost = 0
            if culture_recue >= bat.boost100:
                boost = 100
            elif culture_recue >= bat.boost50:
                boost = 50
            elif culture_recue >= bat.boost25:
                boost = 25

            bat.instances[idx] = (r, c, orient, culture_recue, boost)

def creer_excel_resultat(batiments, non_places, terrain_final, placed_list):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='openpyxl')

    # Batiments places
    rows = []
    for bat in batiments:
        for inst in bat.instances:
            r, c, orient, cult, boost = inst
            h = bat.larg if orient == 'V' else bat.long
            w = bat.long if orient == 'V' else bat.larg
            rows.append({
                "Nom": bat.nom,
                "Type": bat.typ,
                "Production": bat.prod_type,
                "Ligne": r + 1,
                "Colonne": c + 1,
                "Hauteur": h,
                "Largeur": w,
                "Culture recue": cult,
                "Boost (%)": boost,
                "Quantite/h": bat.quantite,
                "Prod totale/h": round(bat.quantite * (1 + boost / 100.0), 1),
                "Origine": "Placé"
            })
    pd.DataFrame(rows).to_excel(writer, sheet_name="Batiments places", index=False)

    # Synthese
    prod_data = defaultdict(lambda: {"cult": 0, "nb": 0, "prod": 0.0})
    for bat in batiments:
        if bat.typ != "Producteur": continue
        for _, _, _, cult, boost in bat.instances:
            prod_data[bat.prod_type]["cult"] += cult
            prod_data[bat.prod_type]["nb"] += 1
            prod_data[bat.prod_type]["prod"] += bat.quantite * (1 + boost / 100.0)

    synthese = [
        ["Synthese par type de production", "", "", "", ""],
        ["Production", "Culture totale", "Boost moyen (%)", "Nb batiments", "Production/h"]
    ]
    for ptype, d in prod_data.items():
        boost_moy = int((d["prod"] / max(1, d["nb"] * list(batiments)[0].quantite if batiments else 1)) * 100) if d["nb"] > 0 else 0
        synthese.append([ptype, d["cult"], boost_moy, d["nb"], round(d["prod"], 1)])
    synthese += [["", "", "", "", ""],
                 ["Cases libres restantes", int(np.sum(terrain_final == 1)), "", "", ""],
                 ["Cases des batiments non places", sum(b.surface() * b.nombre for b in non_places), "", "", ""]]
    pd.DataFrame(synthese).to_excel(writer, sheet_name="Synthese", index=False, header=False)

    # Terrain
    pd.DataFrame(terrain_final).to_excel(writer, sheet_name="Terrain", index=False, header=False)

    # Non places
    non_data = [{"Nom": b.nom, "Type": b.typ, "Production": b.prod_type,
                 "Longueur": b.long, "Largeur": b.larg,
                 "Cases": b.surface() * b.nombre} for b in non_places]
    pd.DataFrame(non_data).to_excel(writer, sheet_name="Non places", index=False)

    writer.close()
    output.seek(0)

    # Coloration
    wb = load_workbook(output)
    ws = wb["Terrain"]

    orange = PatternFill(start_color="FFCC99", fill_type="solid")
    vert = PatternFill(start_color="99FF99", fill_type="solid")
    gris = PatternFill(start_color="CCCCCC", fill_type="solid")
    bold = Font(bold=True)

    for bat, r, c, orient, h, w in placed_list:
        fill = gris if bat.typ == "Neutre" else (orange if bat.typ == "Culturel" else vert)
        for i in range(r, r + h):
            for j in range(c, c + w):
                cell = ws.cell(row=i+1, column=j+1)
                cell.fill = fill
        cell = ws.cell(row=r+1, column=c+1)
        cell.value = f"{bat.nom[:10]} {bat.instances[0][4]}%" if bat.instances else bat.nom[:10]
        cell.font = bold
        cell.alignment = Alignment(horizontal="center")

    final = io.BytesIO()
    wb.save(final)
    final.seek(0)
    return final

# ====================== STREAMLIT ======================
def main():
    st.set_page_config(page_title="Optimiseur Ville", layout="wide")
    st.title("🚀 Optimiseur de Placement de Bâtiments")
    st.caption("Version corrigée – Gestion robuste des colonnes Boost")

    uploaded = st.file_uploader("Fichier Excel (Ville.xlsx)", type=["xlsx"])

    if uploaded:
        try:
            terrain, batiments = lire_excel(uploaded)
            st.success(f"Terrain {terrain.shape[0]}×{terrain.shape[1]} | {len(batiments)} bâtiments chargés")

            if st.button("🚀 Lancer l'optimisation complète", type="primary"):
                with st.spinner("Placement + calcul en cours..."):
                    terrain_final, placed_list, non_places = optimiser_placement(terrain, batiments)
                    calcul_culture_boost(placed_list, batiments, terrain_final)
                    excel_bytes = creer_excel_resultat(batiments, non_places, terrain_final, placed_list)

                    st.success("✅ Optimisation terminée !")
                    st.download_button(
                        label="📥 Télécharger resultats.xlsx",
                        data=excel_bytes,
                        file_name="resultats.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    st.subheader("Résumé")
                    st.write(f"Bâtiments placés : {len(placed_list)}")
                    st.write(f"Non placés : {len(non_places)}")
                    st.write(f"Cases libres restantes : {np.sum(terrain_final == 1)}")
        except Exception as e:
            st.error(f"Erreur : {str(e)}")
            st.info("Si l'erreur persiste, envoie-moi le message exact.")

if __name__ == "__main__":
    main()
