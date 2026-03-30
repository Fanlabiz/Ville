import streamlit as st
import pandas as pd
import numpy as np
import copy
import random
from io import BytesIO

# --- CONFIGURATION ET CONSTANTES ---
PRIORITES_PROD = {'Guérison': 4, 'Nourriture': 3, 'Or': 2} # Plus haut = plus prioritaire

class Batiment:
    def __init__(self, data):
        self.nom = data['Nom']
        self.longueur = int(data['Longueur'])
        self.largeur = int(data['Largeur'])
        self.type = data['Type'] # Culturel / Producteur
        self.culture = data.get('Culture', 0)
        self.rayonnement = data.get('Rayonnement', 0)
        self.boosts = {0.25: data['Boost 25%'], 0.5: data['Boost 50%'], 1.0: data['Boost 100%']}
        self.production_type = data['Production']
        self.quantite_base = data['Quantite']
        self.priorite = data.get('Priorite', 0)
        self.x, self.y = None, None
        self.orientation = 'H' # 'H' ou 'V'

    def get_dims(self):
        return (self.longueur, self.largeur) if self.orientation == 'H' else (self.largeur, self.longueur)

# --- FONCTIONS DE CALCUL ---

def calculer_score(batiments, terrain_dims):
    total_prod = 0
    prod_details = {}
    
    producteurs = [b for b in batiments if b.type == 'Producteur']
    culturels = [b for b in batiments if b.type == 'Culturel']
    
    for p in producteurs:
        culture_recue = 0
        px, py = p.x, p.y
        pl, ph = p.get_dims()
        
        for c in culturels:
            cl, ch = c.get_dims()
            # Zone de rayonnement : bande autour du bâtiment
            r = c.rayonnement
            if not (px + pl <= c.x - r or px >= c.x + cl + r or 
                    py + ph <= c.y - r or py >= c.y + ch + r):
                culture_recue += c.culture
        
        # Calcul du boost
        boost = 0
        if culture_recue >= p.boosts[1.0]: boost = 1.0
        elif culture_recue >= p.boosts[0.5]: boost = 0.5
        elif culture_recue >= p.boosts[0.25]: boost = 0.25
        
        prod_finale = p.quantite_base * (1 + boost)
        p.current_boost = boost
        p.current_culture = culture_recue
        
        # Pondération par priorité de production
        poids = PRIORITES_PROD.get(p.production_type, 1)
        total_prod += prod_finale * poids
        
    return total_prod

# --- INTERFACE STREAMLIT ---

st.set_page_config(page_title="Optimiseur de Terrain", layout="wide")
st.title("🏙️ Optimiseur de Placement de Bâtiments")

uploaded_file = st.file_uploader("Charger le fichier Excel (iPad compatible)", type=["xlsx"])

if uploaded_file:
    # Lecture des onglets
    df_terrain = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    df_data = pd.read_excel(uploaded_file, sheet_name=1)
    
    # 1. Analyse du terrain
    terrain_array = df_terrain.values
    H, L = terrain_array.shape
    cases_interdites = np.where(terrain_array == 'x', 1, 0)
    
    # 2. Préparation des bâtiments
    liste_batiments = []
    for _, row in df_data.iterrows():
        for _ in range(int(row['Nombre'])):
            liste_batiments.append(Batiment(row))

    st.sidebar.success(f"Terrain détecté : {L}x{H}")
    st.sidebar.info(f"Bâtiments à placer : {len(liste_batiments)}")

    if st.button("Lancer l'Optimisation"):
        with st.spinner("Recherche du meilleur agencement..."):
            # Algorithme simplifié de placement aléatoire itératif (Hill Climbing)
            # Note : Pour un usage réel, on implémenterait ici un Recuit Simulé complet
            
            best_score = -1
            best_layout = []
            
            for tentative in range(100): # Boucles d'essais
                current_layout = copy.deepcopy(liste_batiments)
                grid_occupancy = cases_interdites.copy()
                
                random.shuffle(current_layout)
                placed_count = 0
                
                for b in current_layout:
                    # Essayer de placer chaque bâtiment
                    dims = [(b.longueur, b.largeur), (b.largeur, b.longueur)]
                    placed = False
                    
                    # On cherche une place libre aléatoirement pour la démo
                    attempts = 0
                    while not placed and attempts < 50:
                        b.orientation = random.choice(['H', 'V'])
                        bl, bh = b.get_dims()
                        x = random.randint(0, L - bl)
                        y = random.randint(0, H - bh)
                        
                        if np.sum(grid_occupancy[y:y+bh, x:x+bl]) == 0:
                            grid_occupancy[y:y+bh, x:x+bl] = 1
                            b.x, b.y = x, y
                            placed = True
                            placed_count += 1
                        attempts += 1
                
                if placed_count == len(liste_batiments):
                    score = calculer_score(current_layout, (H, L))
                    if score > best_score:
                        best_score = score
                        best_layout = current_layout

            # --- GENERATION DU RESULTAT ---
            if best_layout:
                st.balloons()
                st.subheader("✅ Optimisation Terminée")
                
                # Création du fichier de sortie
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    # Onglet 1: Liste détaillée
                    res_data = []
                    for b in best_layout:
                        res_data.append([b.nom, b.type, b.production_type, f"({b.x},{b.y})", b.current_culture, b.current_boost])
                    
                    df_res = pd.DataFrame(res_data, columns=['Nom', 'Type', 'Production', 'Coords', 'Culture Reçue', 'Boost'])
                    df_res.to_excel(writer, sheet_name='Resultats_Details', index=False)
                    
                    # Onglet Visuel (Simplifié pour l'exemple)
                    workbook = writer.book
                    worksheet = workbook.add_worksheet('Terrain_Final')
                    format_cult = workbook.add_format({'bg_color': 'orange'})
                    format_prod = workbook.add_format({'bg_color': 'green'})
                    
                    for b in best_layout:
                        bl, bh = b.get_dims()
                        fmt = format_cult if b.type == 'Culturel' else format_prod
                        worksheet.merge_range(b.y, b.x, b.y + bh - 1, b.x + bl - 1, f"{b.nom}\n{b.current_boost:.0%}", fmt)

                st.download_button(
                    label="📥 Télécharger le plan optimisé (Excel)",
                    data=output.getvalue(),
                    file_name="resultat_optimisation.xlsx",
                    mime="application/vnd.ms-excel"
                )
            else:
                st.error("Impossible de placer tous les bâtiments sur ce terrain. Essayez d'agrandir l'espace ou de réduire le nombre de bâtiments.")
