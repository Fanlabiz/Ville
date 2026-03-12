import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# --- CONFIGURATION COULEURS ---
COLORS = {
    'Culturel': 'FFFFA500',  # Orange
    'Producteur': 'FF008000', # Vert
    'Neutre': 'FF808080'      # Gris
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
        self.max_entries = 1000 # Limite demandée dans votre consigne
        self.interrupted = False

    def log(self, msg):
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)
        else:
            self.interrupted = True

    def can_place(self, r, c, w, h, remaining_queue):
        if r + h > self.rows or c + w > self.cols: return False
        if not np.all(self.grid[r:r+h, c:c+w] == 1): return False
        
        # Règle de vérification du plus grand bâtiment restant
        if remaining_queue:
            self.grid[r:r+h, c:c+w] = 0 # Simulation placement
            biggest = remaining_queue[0]
            bw, bh = int(biggest['Largeur']), int(biggest['Longueur'])
            can_fit_biggest = False
            for ow, oh in [(bw, bh), (bh, bw)]:
                for rr in range(self.rows - oh + 1):
                    for cc in range(self.cols - ow + 1):
                        if np.all(self.grid[rr:rr+oh, cc:cc+ow] == 1):
                            can_fit_biggest = True
                            break
                    if can_fit_biggest: break
                if can_fit_biggest: break
            self.grid[r:r+h, c:c+w] = 1 # Annulation simulation
            if not can_fit_biggest: return False
        return True

    def solve(self, buildings):
        if not buildings or self.interrupted: return True
        b = buildings[0]
        self.log(f"Évaluation de : {b['Nom']}")
        
        dims = list(set([(int(b['Largeur']), int(b['Longueur'])), (int(b['Longueur']), int(b['Largeur']))]))

        for w, h in dims:
            for r in range(self.rows - h + 1):
                for c in range(self.cols - w + 1):
                    if self.can_place(r, c, w, h, buildings[1:]):
                        self.grid[r:r+h, c:c+w] = 0
                        self.placed_buildings.append({'info': b, 'r': r, 'c': c, 'w': w, 'h': h})
                        self.log(f"Placé : {b['Nom']} en ({r},{c})")
                        
                        if self.solve(buildings[1:]): return True
                        
                        if not self.interrupted:
                            self.log(f"Enlevé : {b['Nom']} de ({r},{c})")
                            self.grid[r:r+h, c:c+w] = 1
                            self.placed_buildings.pop()
                if self.interrupted: return True
        return False

def generate_excel(planner, full_queue):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # --- CALCUL DE LA CULTURE CORRIGÉ (CUMULATIF) ---
        culture_map = np.zeros((planner.rows, planner.cols))
        for pb in planner.placed_buildings:
            if pb['info']['Type'] == 'Culturel':
                ray = int(pb['info'].get('Rayonnement', 0))
                val = float(pb['info'].get('Culture', 0))
                # Zone d'influence
                r_s, r_e = max(0, pb['r']-ray), min(planner.rows, pb['r']+pb['h']+ray)
                c_s, c_e = max(0, pb['c']-ray), min(planner.cols, pb['c']+pb['w']+ray)
                # On AJOUTE la culture (+=) pour permettre le cumul (ex: 301 + 615 + ...)
                culture_map[r_s:r_e, c_s:c_e] += val

        # 1. Synthèse Productions (Guerison, Nourriture, Or)
        stats_prod = []
        synthese_globale = {"Guerison": 0, "Nourriture": 0, "Or": 0}
        
        for pb in planner.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                # On prend la culture sur la case d'origine du bâtiment
                c_recue = culture_map[pb['r'], pb['c']]
                
                # Calcul des boosts
                boost = 0
                for b_val, b_perc in [(pb['info'].get('Boost 100%'), 100), 
                                      (pb['info'].get('Boost 50%'), 50), 
                                      (pb['info'].get('Boost 25%'), 25)]:
                    if pd.notnull(b_val) and b_val > 0 and c_recue >= b_val:
                        boost = b_perc
                        break
                
                stats_prod.append([pb['info']['Nom'], c_recue, f"{boost}%"])
                prod_type = pb['info'].get('Production')
                if prod_type in synthese_globale:
                    synthese_globale[prod_type] += c_recue

        pd.DataFrame(stats_prod, columns=["Bâtiment", "Culture Recue", "Boost"]).to_excel(writer, sheet_name="Production", index=False)
        pd.DataFrame(list(synthese_globale.items()), columns=["Type", "Culture Totale"]).to_excel(writer, sheet_name="Synthese", index=False)

        # 2. Plan visuel
        ws = writer.book.create_sheet("Plan_Terrain")
        for pb in planner.placed_buildings:
            fill = PatternFill(start_color=COLORS.get(pb['info']['Type'], 'FFFFFF'), fill_type='solid')
            for r in range(pb['r'], pb['r']+pb['h']):
                for c in range(pb['c'], pb['c']+pb['w']):
                    ws.cell(row=r+1, column=c+1, value=pb['info']['Nom']).fill = fill

        # 3. Journal et Résumé
        pd.DataFrame(planner.journal, columns=["Journal"]).to_excel(writer, sheet_name="Journal", index=False)
        
        not_placed = [b['Nom'] for b in full_queue if b not in [p['info'] for p in planner.placed_buildings]]
        summary = [
            ["Cases libres initiales", planner.initial_free_cells],
            ["Cases non utilisées", planner.initial_free_cells - sum(p['w']*p['h'] for p in planner.placed_buildings)],
            ["Bâtiments non placés", len(not_placed)],
            ["Statut", "LIMITE JOURNAL ATTEINTE" if planner.interrupted else "OK"]
        ]
        pd.DataFrame(summary).to_excel(writer, sheet_name="Resume", index=False, header=False)
        pd.DataFrame(not_placed, columns=["Nom"]).to_excel(writer, sheet_name="Non_Places", index=False)

    return output.getvalue()

# --- STREAMLIT ---
st.title("Optimiseur de Ville - Fix Culture 🏗️")
uploaded = st.file_uploader("Charger Ville.xlsx", type="xlsx")

if uploaded:
    t_df = pd.read_excel(uploaded, sheet_name=0, header=None)
    b_df = pd.read_excel(uploaded, sheet_name=1)
    
    # Prétraitement pour forcer les nombres
    for c in ['Culture', 'Rayonnement', 'Boost 25%', 'Boost 50%', 'Boost 100%']:
        if c in b_df.columns: b_df[c] = pd.to_numeric(b_df[c], errors='coerce').fillna(0)

    # Tri selon vos critères (Grand -> Petit)
    b_df['Surface'] = b_df['Longueur'] * b_df['Largeur']
    b_df = b_df.sort_values('Surface', ascending=False)
    
    full_queue = []
    for _, r in b_df.iterrows():
        for _ in range(int(r['Quantite'])): full_queue.append(r.to_dict())

    planner = CityPlanner(t_df.values)
    with st.spinner("Calcul en cours..."):
        planner.solve(full_queue)
    
    st.success("Terminé !")
    st.download_button("Télécharger Résultats", generate_excel(planner, full_queue), "Resultat_Placement.xlsx")
