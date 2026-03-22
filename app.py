# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import itertools

# ────────────────────────────────────────────────
#  FONCTIONS UTILITAIRES
# ────────────────────────────────────────────────

def load_terrain(df_terrain):
    """Convertit la feuille Terrain en np.array (1=libre, 0=occupé, -1=mur)"""
    grid = df_terrain.fillna('').replace('X', -1).replace('', 0).astype(int).values
    # Supprimer lignes/colonnes entièrement vides à droite/bas
    while np.all(grid[-1] == 0): grid = grid[:-1]
    while np.all(grid[:, -1] == 0): grid = np.delete(grid, -1, axis=1)
    return grid


def can_place(grid, r, c, h, w, rotation=False):
    """Vérifie si on peut poser un bâtiment h×w (ou w×h si rotation) en (r,c)"""
    if rotation:
        h, w = w, h
    rows, cols = grid.shape
    if r + h > rows or c + w > cols:
        return False
    return np.all(grid[r:r+h, c:c+w] == 1)


def place(grid, r, c, h, w, value, rotation=False):
    if rotation:
        h, w = w, h
    grid[r:r+h, c:c+w] = value


def remove(grid, r, c, h, w, rotation=False):
    if rotation:
        h, w = w, h
    grid[r:r+h, c:c+w] = 1


def find_border_positions(grid, size_h, size_w, prefer_border=True):
    """Génère des positions candidates, de préférence sur les bords"""
    rows, cols = grid.shape
    positions = []
    
    for r in range(rows):
        for c in range(cols):
            if can_place(grid, r, c, size_h, size_w, rotation=False):
                positions.append((r, c, False))
            if size_h != size_w and can_place(grid, r, c, size_h, size_w, rotation=True):
                positions.append((r, c, True))
    
    if prefer_border:
        def border_score(pos):
            r, c, _ = pos
            dist_border = min(r, rows-1-r, c, cols-1-c)
            return -dist_border  # plus près du bord = meilleur score
        positions.sort(key=border_score)
    
    return positions


def compute_culture_coverage(grid, buildings_placed, cultural_buildings):
    """Calcule pour chaque case la somme de culture reçue"""
    rows, cols = grid.shape
    culture_map = np.zeros((rows, cols), dtype=float)
    
    for b in cultural_buildings:
        if b.get('placed', False):
            r, c, rot = b['row'], b['col'], b['rotation']
            bh, bw = (b['Largeur'], b['Longueur']) if rot else (b['Longueur'], b['Largeur'])
            rad = b['Rayonnement']
            cult = b['Culture']
            
            r_start = max(0, r - rad)
            r_end   = min(rows, r + bh + rad)
            c_start = max(0, c - rad)
            c_end   = min(cols, c + bw + rad)
            
            culture_map[r_start:r_end, c_start:c_end] += cult
    
    return culture_map


def get_boost(culture_received, thresholds):
    """Retourne le pourcentage de boost (0, 0.25, 0.5, 1.0)"""
    if pd.isna(thresholds[2]) or culture_received >= thresholds[2]:
        return 1.00
    if pd.isna(thresholds[1]) or culture_received >= thresholds[1]:
        return 0.50
    if pd.isna(thresholds[0]) or culture_received >= thresholds[0]:
        return 0.25
    return 0.00


# ────────────────────────────────────────────────
#  STREAMLIT APP
# ────────────────────────────────────────────────

st.set_page_config(page_title="Optimiseur de Ville", layout="wide")
st.title("Optimiseur de placement de bâtiments")

uploaded_file = st.file_uploader("Choisir le fichier Excel (Ville.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        df_terrain  = pd.read_excel(xls, sheet_name="Terrain", header=None)
        df_bat      = pd.read_excel(xls, sheet_name="Batiments")
        df_actuel   = pd.read_excel(xls, sheet_name="Actuel", header=None)
        
        grid = load_terrain(df_terrain)
        rows, cols = grid.shape
        
        st.write(f"Terrain détecté : {rows} lignes × {cols} colonnes")
        
        # ─── Préparation des bâtiments ──────────────────────────────
        batiments = []
        for _, row in df_bat.iterrows():
            if pd.isna(row['Nom']): continue
            bat = row.to_dict()
            bat['Longueur'] = int(bat['Longueur'])
            bat['Largeur']  = int(bat['Largeur'])
            bat['Nombre']   = int(bat['Nombre'])
            bat['placed']   = False
            batiments.extend([bat.copy() for _ in range(bat['Nombre'])])
        
        # ─── Priorisation ───────────────────────────────────────────
        neutres = [b for b in batiments if b['Type'] == 'Neutre']
        culturels = [b for b in batiments if b['Type'] == 'Culturel']
        producteurs = [b for b in batiments if b['Type'] == 'Producteur']
        
        # Tri global par surface décroissante
        for lst in [neutres, culturels, producteurs]:
            lst.sort(key=lambda b: b['Longueur'] * b['Largeur'], reverse=True)
        
        # Priorité forte Guérison au début
        prod_guerison = [p for p in producteurs if p['Production'] == 'Guerison']
        prod_autres   = [p for p in producteurs if p['Production'] != 'Guerison']
        producteurs = prod_guerison + prod_autres
        
        # ─── Placement ──────────────────────────────────────────────
        placed = []
        grid_copy = grid.copy()  # on travaille sur une copie
        
        # 1. Neutres → priorité bords
        for b in neutres:
            positions = find_border_positions(grid_copy, b['Longueur'], b['Largeur'])
            for r, c, rot in positions[:30]:  # limite pour ne pas trop ralentir
                if can_place(grid_copy, r, c, b['Longueur'], b['Largeur'], rot):
                    place(grid_copy, r, c, b['Longueur'], b['Largeur'], -2)  # gris = neutre
                    placed.append({**b, 'row':r, 'col':c, 'rotation':rot, 'placed':True})
                    b['placed'] = True
                    break
        
        # 2. Placement alterné culturels / producteurs prioritaires
        remaining = culturels + producteurs
        remaining.sort(key=lambda b: b['Longueur'] * b['Largeur'], reverse=True)
        
        i = 0
        while i < len(remaining):
            b = remaining[i]
            positions = find_border_positions(grid_copy, b['Longueur'], b['Largeur'], prefer_border=False)
            
            placed_ok = False
            for r, c, rot in positions[:80]:
                if can_place(grid_copy, r, c, b['Longueur'], b['Largeur'], rot):
                    place(grid_copy, r, c, b['Longueur'], b['Largeur'], -3 if b['Type']=='Culturel' else -4)
                    placed.append({**b, 'row':r, 'col':c, 'rotation':rot, 'placed':True})
                    b['placed'] = True
                    placed_ok = True
                    break
            
            if not placed_ok:
                i += 1
            else:
                # Après un placement, on réessaie depuis le début pour les gros
                remaining = [b for b in remaining if not b['placed']]
                remaining.sort(key=lambda x: x['Longueur']*x['Largeur'], reverse=True)
                i = 0
        
        # ─── Calculs finaux ─────────────────────────────────────────
        culture_map = compute_culture_coverage(grid_copy, placed, culturels)
        
        for b in placed:
            if b['Type'] != 'Producteur': continue
            r, c, rot = b['row'], b['col'], b['rotation']
            h = b['Largeur'] if rot else b['Longueur']
            w = b['Longueur'] if rot else b['Largeur']
            center_r, center_c = r + h//2, c + w//2
            cult_received = culture_map[center_r, center_c]  # approx centre
            b['culture_recue'] = cult_received
            b['boost'] = get_boost(cult_received, [
                b.get('Boost 25%'), b.get('Boost 50%'), b.get('Boost 100%')
            ])
            b['prod_reelle'] = b['Quantite'] * (1 + b['boost'])
        
        # ─── Statistiques globales ──────────────────────────────────
        stats = {}
        for b in placed:
            if b['Type'] != 'Producteur' or b['Production'] == 'Rien': continue
            res = b['Production']
            stats[res] = stats.get(res, 0) + b['prod_reelle']
        
        non_places = [b['Nom'] for b in batiments if not b.get('placed', False)]
        cases_libres = np.sum(grid_copy == 1)
        cases_bat_non_place = sum(b['Longueur']*b['Largeur'] for b in batiments if not b.get('placed', False))
        
        # ─── Préparation sortie Excel ───────────────────────────────
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 1. Bâtiments placés
            df_placed = pd.DataFrame([
                {
                    'Nom': b['Nom'],
                    'Type': b['Type'],
                    'Production': b.get('Production',''),
                    'row': b.get('row',''),
                    'col': b.get('col',''),
                    'rot': 'Oui' if b.get('rotation') else 'Non',
                    'Culture reçue': round(b.get('culture_recue',0),1),
                    'Boost': f"{int(b.get('boost',0)*100)}%",
                    'Prod/heure réelle': round(b.get('prod_reelle',0),1)
                } for b in placed
            ])
            df_placed.to_excel(writer, sheet_name='Bâtiments placés', index=False)
            
            # 2. Non placés
            pd.DataFrame({'Non placés': non_places}).to_excel(writer, sheet_name='Non placés', index=False)
            
            # 3. Stats
            pd.DataFrame({
                'Ressource': list(stats.keys()),
                'Production totale / h': [round(v,1) for v in stats.values()]
            }).to_excel(writer, sheet_name='Stats', index=False)
            
            # 4. Terrain visuel (simplifié texte + annotations)
            viz = grid_copy.astype(object)
            for b in placed:
                r,c,rot = b['row'],b['col'],b['rotation']
                h = b['Largeur'] if rot else b['Longueur']
                w = b['Longueur'] if rot else b['Largeur']
                short = b['Nom'][:12]
                boost_str = f"{int(b.get('boost',0)*100)}%" if b['Type']=='Producteur' else ""
                label = f"{short}\n{boost_str}".strip()
                for i in range(h):
                    for j in range(w):
                        viz[r+i, c+j] = label
            
            pd.DataFrame(viz).to_excel(writer, sheet_name='Terrain', index=False, header=False)
            
            # 5. Résumé
            pd.DataFrame({
                'Indicateur': ['Cases libres', 'Bâtiments non placés', 'Cases occupées par non-placés', 'Production totale Guérison', 'Production totale Nourriture', 'Production totale Or'],
                'Valeur': [
                    cases_libres,
                    len(non_places),
                    cases_bat_non_place,
                    round(stats.get('Guerison',0),1),
                    round(stats.get('Nourriture',0),1),
                    round(stats.get('Or',0),1)
                ]
            }).to_excel(writer, sheet_name='Résumé', index=False)
        
        output.seek(0)
        
        st.success("Placement terminé !")
        st.download_button(
            label="Télécharger résultat.xlsx",
            data=output,
            file_name="Ville_resultat.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        # Aperçu rapide
        st.subheader("Production totale estimée")
        st.write(stats)
        st.write(f"Cases libres restantes : {cases_libres}")
        st.write(f"Bâtiments non placés : {len(non_places)}")
        
    except Exception as e:
        st.error(f"Erreur lors du traitement : {e}")
        st.exception(e)

else:
    st.info("Chargez votre fichier Ville.xlsx pour commencer.")
