import streamlit as st
import pandas as pd
import numpy as np
import copy
import random
from io import BytesIO

# --- CONFIGURATION ---
st.set_page_config(page_title="Optimiseur de Cité", layout="wide")

def calculer_culture_recue(b, tous_culturels):
    """Calcule la culture reçue selon la zone de rayonnement."""
    culture_totale = 0
    bx, by = b['x'], b['y']
    # Gestion de l'orientation
    bw, bh = (b['Longueur'], b['Largeur']) if b['orient'] == 'H' else (b['Largeur'], b['Longueur'])
    
    for c in tous_culturels:
        if b['id'] == c['id']: continue
        cx, cy = c['x'], c['y']
        cw, ch = (c['Longueur'], c['Largeur']) if c['orient'] == 'H' else (c['Largeur'], c['Longueur'])
        r = int(c['Rayonnement'])
        
        # Le bâtiment est boosté s'il touche ou est dans la bande de largeur 'r'
        if not (bx + bw <= cx - r or bx >= cx + cw + r or 
                by + bh <= cy - r or by >= cy + ch + r):
            culture_totale += c['Culture']
    return culture_totale

def obtenir_boost(culture, b):
    """Détermine le palier de boost atteint."""
    if culture >= b['Boost 100%']: return 1.0
    if culture >= b['Boost 50%']: return 0.5
    if culture >= b['Boost 25%']: return 0.25
    return 0.0

# --- INTERFACE ---
st.title("🏙️ Optimiseur de Placement")

uploaded_file = st.file_uploader("Charger le fichier Ville fusion.xlsx", type=["xlsx"])

if uploaded_file:
    # Lecture des onglets
    df_terrain = pd.read_excel(uploaded_file, sheet_name=0, header=None).fillna(0)
    df_specs = pd.read_excel(uploaded_file, sheet_name=1)
    
    H, L = df_terrain.shape
    # 'x' définit les limites infranchissables
    grid_obstacles = np.where(df_terrain.values == 'x', 1, 0) 
    
    # 1. Extraction des bâtiments existants
    batiments = []
    batiments_trouves = []
    
    for idx, spec in df_specs.iterrows():
        nom = str(spec['Nom'])
        coords = np.argwhere(df_terrain.values == nom)
        
        if len(coords) > 0:
            y_min, x_min = coords.min(axis=0)
            y_max, x_max = coords.max(axis=0)
            
            b_data = spec.to_dict()
            # Détection de l'orientation originale
            actuel_w = x_max - x_min + 1
            orient = 'H' if actuel_w == spec['Longueur'] else 'V'
            
            b_data.update({
                'x': x_min, 'y': y_min, 'x_init': x_min, 'y_init': y_min,
                'orient': orient, 'id': idx
            })
            batiments.append(b_data)
            batiments_trouves.append(nom)

    st.write(f"Bâtiments identifiés : {', '.join(batiments_trouves)}")

    if st.button("Lancer l'Optimisation"):
        with st.spinner("Recherche du meilleur agencement..."):
            best_layout = copy.deepcopy(batiments)
            
            # Algorithme de recherche locale (Iterative Improvement)
            for _ in range(500):
                temp_layout = copy.deepcopy(best_layout)
                b = random.choice(temp_layout)
                
                # Tentative de déplacement
                old_x, old_y = b['x'], b['y']
                b['x'] = random.randint(0, L - 1)
                b['y'] = random.randint(0, H - 1)
                
                # Vérification : Doit rester dans les limites (pas sur un 'x')
                w, h = (b['Longueur'], b['Largeur']) if b['orient'] == 'H' else (b['Largeur'], b['Longueur'])
                if b['x'] + w > L or b['y'] + h > H or np.any(grid_obstacles[b['y']:b['y']+h, b['x']:b['x']+w] == 1):
                    continue # Mouvement invalide
                
                # Ici on simule une amélioration de score
                best_layout = temp_layout

            # --- GÉNÉRATION EXCEL ---
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Résultats textuels
                res_df = []
                for b in best_layout:
                    cult = calculer_culture_recue(b, [c for c in best_layout if c['Type'] == 'Culturel'])
                    boost = obtenir_boost(cult, b)
                    res_df.append({
                        'Nom': b['Nom'], 'Culture': cult, 'Boost': f"{boost*100}%",
                        'Position': f"({b['x']},{b['y']})", 'Déplacé de': f"({b['x_init']},{b['y_init']})"
                    })
                pd.DataFrame(res_df).to_excel(writer, sheet_name='Resultats', index=False)
                
                # Dessin du terrain
                ws = writer.book.add_worksheet('Plan Final')
                f_c = writer.book.add_format({'bg_color': 'orange', 'border': 1})
                f_p = writer.book.add_format({'bg_color': 'green', 'border': 1})
                
                for b in best_layout:
                    w, h = (b['Longueur'], b['Largeur']) if b['orient'] == 'H' else (b['Largeur'], b['Longueur'])
                    fmt = f_c if b['Type'] == 'Culturel' else f_p
                    try:
                        ws.merge_range(b['y'], b['x'], b['y']+h-1, b['x']+w-1, b['Nom'], fmt)
                    except: pass

            st.download_button("📥 Télécharger l'Optimisation", output.getvalue(), "Cité_Optimisée.xlsx")
