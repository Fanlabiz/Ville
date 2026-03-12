import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# --- CONFIGURATION ---
COLORS = {
    'Culturel': 'FFFFA500', 
    'Producteur': 'FF008000', 
    'Neutre': 'FF808080'
}

class CityPlanner:
    def __init__(self, terrain_data):
        self.rows = len(terrain_data)
        self.cols = len(terrain_data[0])
        self.grid = np.zeros((self.rows, self.cols))
        self.border_mask = np.zeros((self.rows, self.cols), dtype=bool)
        self.initial_free_cells = 0
        
        for r in range(self.rows):
            for c in range(self.cols):
                val = str(terrain_data[r][c]).strip().upper()
                if val == '1': 
                    self.grid[r,c] = 1
                    self.initial_free_cells += 1
                elif val == 'X':
                    self.grid[r,c] = 0 
                    self.border_mask[r,c] = True
        
        self.journal = []
        self.placed_buildings = []
        self.max_entries = 10000 
        self.interrupted = False

    def log(self, msg):
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)

    def can_place(self, r, c, w, h):
        if r + h > self.rows or c + w > self.cols: return False
        return np.all(self.grid[r:r+h, c:c+w] == 1)

    def place_building(self, b, r, c, w, h):
        self.grid[r:r+h, c:c+w] = 0
        self.placed_buildings.append({
            'info': b,
            'r': r, 'c': c, 'w': w, 'h': h,
            'pos': (r, c)
        })
        self.log(f"Placé : {b['Nom']} en ({r},{c})")

    def solve_simple(self, buildings):
        """ Logique de placement d'origine de votre script """
        for b in buildings:
            w, h = int(b['Largeur']), int(b['Longueur'])
            placed = False
            for r in range(self.rows - h + 1):
                for c in range(self.cols - w + 1):
                    if self.can_place(r, c, w, h):
                        self.place_building(b, r, c, w, h)
                        placed = True
                        break
                if placed: break
            if not placed:
                self.log(f"Échec placement : {b['Nom']}")

def generate_excel(planner, full_queue):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # --- CORRECTION : CALCUL DE LA CULTURE CUMULATIVE ---
        # On crée une carte de culture pour tout le terrain
        culture_map = np.zeros((planner.rows, planner.cols))
        
        for pb in planner.placed_buildings:
            if pb['info'].get('Type') == 'Culturel':
                ray = int(pb['info'].get('Rayonnement', 0))
                val = float(pb['info'].get('Culture', 0))
                
                # Zone d'influence : Bâtiment + Rayonnement tout autour
                r_start = max(0, pb['r'] - ray)
                r_end = min(planner.rows, pb['r'] + pb['h'] + ray)
                c_start = max(0, pb['c'] - ray)
                c_end = min(planner.cols, pb['c'] + pb['w'] + ray)
                
                # Addition cumulative de la culture sur la zone
                culture_map[r_start:r_end, c_start:c_end] += val

        # --- CALCUL DES RÉSULTATS DE PRODUCTION ---
        prod_data = []
        synthese = {"Guerison": 0, "Nourriture": 0, "Or": 0}
        
        for pb in planner.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                # On récupère la culture cumulée sous le bâtiment
                culture_reçue = culture_map[pb['r'], pb['c']]
                
                # Calcul du boost atteint
                boost_label = "0%"
                if culture_reçue >= pb['info'].get('Boost 100%', 999999):
                    boost_label = "100%"
                elif culture_reçue >= pb['info'].get('Boost 50%', 999999):
                    boost_label = "50%"
                elif culture_reçue >= pb['info'].get('Boost 25%', 999999):
                    boost_label = "25%"
                
                prod_data.append([pb['info']['Nom'], culture_reçue, boost_label])
                
                # Ajout à la synthèse par type de production
                p_type = pb['info'].get('Production')
                if p_type in synthese:
                    synthese[p_type] += culture_reçue

        # Création des onglets
        pd.DataFrame(prod_data, columns=["Bâtiment", "Culture", "Boost"]).to_excel(writer, sheet_name="Production", index=False)
        pd.DataFrame(list(synthese.items()), columns=["Type", "Culture Totale"]).to_excel(writer, sheet_name="Synthese", index=False)
        
        # Plan visuel
        ws = writer.book.create_sheet("Plan_Terrain")
        for pb in planner.placed_buildings:
            fill = PatternFill(start_color=COLORS.get(pb['info']['Type'], 'FFFFFF'), fill_type='solid')
            for r in range(pb['r'], pb['r']+pb['h']):
                for c in range(pb['c'], pb['c']+pb['w']):
                    ws.cell(row=r+1, column=c+1, value=pb['info']['Nom']).fill = fill

        # Résumé
        not_placed = [b['Nom'] for b in full_queue if b not in [p['info'] for p in planner.placed_buildings]]
        summary = [
            ["Cases libres initiales", planner.initial_free_cells],
            ["Cases non utilisées", planner.initial_free_cells - sum(p['w']*p['h'] for p in planner.placed_buildings)],
            ["Bâtiments placés", len(planner.placed_buildings)],
            ["Bâtiments non placés", len(not_placed)]
        ]
        pd.DataFrame(summary).to_excel(writer, sheet_name="Resume", index=False, header=False)
        pd.DataFrame(planner.journal, columns=["Journal"]).to_excel(writer, sheet_name="Journal", index=False)

    return output.getvalue()

# --- INTERFACE ---
st.title("Optimiseur de Cité - Correction Culture 🏗️")
uploaded = st.file_uploader("Charger Ville.xlsx", type="xlsx")

if uploaded:
    t_df = pd.read_excel(uploaded, sheet_name=0, header=None)
    b_df = pd.read_excel(uploaded, sheet_name=1)
    b_df.columns = b_df.columns.str.strip()
    
    # Nettoyage des données numériques
    for col in ['Culture', 'Rayonnement', 'Boost 25%', 'Boost 50%', 'Boost 100%']:
        if col in b_df.columns:
            b_df[col] = pd.to_numeric(b_df[col], errors='coerce').fillna(0)
    
    # Préparation de la file d'attente (votre logique de tri originale)
    full_queue = []
    # (Note : On peut trier par Type ou Taille ici selon votre préférence)
    for _, r in b_df.iterrows():
        for _ in range(int(r['Quantite'])):
            full_queue.append(r.to_dict())
    
    planner = CityPlanner(t_df.values)
    with st.spinner("Placement en cours..."):
        planner.solve_simple(full_queue)
    
    st.success(f"Terminé : {len(planner.placed_buildings)} bâtiments placés.")
    st.download_button("Télécharger Résultat", generate_excel(planner, full_queue), "Resultat_Ville.xlsx")
