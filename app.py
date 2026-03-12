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
        self.initial_free_cells = 0
        
        for r in range(self.rows):
            for c in range(self.cols):
                val = str(terrain_data[r][c]).strip().upper()
                if val == '1': 
                    self.grid[r,c] = 1
                    self.initial_free_cells += 1
                else:
                    self.grid[r,c] = 0 
        
        self.journal = []
        self.placed_buildings = []
        self.max_entries = 1000 # Retour à la limite stricte de ta demande

    def log(self, msg):
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)

    def can_place(self, r, c, w, h):
        if r + h > self.rows or c + w > self.cols: return False
        return np.all(self.grid[r:r+h, c:c+w] == 1)

    def solve(self, buildings):
        if not buildings or len(self.journal) >= self.max_entries:
            return True
        
        b = buildings[0]
        self.log(f"Évaluation de : {b['Nom']}")
        
        # On teste les deux orientations
        dims = list(set([(int(b['Largeur']), int(b['Longueur'])), (int(b['Longueur']), int(b['Largeur']))]))

        for w, h in dims:
            for r in range(self.rows - h + 1):
                for c in range(self.cols - w + 1):
                    if self.can_place(r, c, w, h):
                        # Placement
                        self.grid[r:r+h, c:c+w] = 0
                        self.placed_buildings.append({'info': b, 'r': r, 'c': c, 'w': w, 'h': h})
                        self.log(f"Placé : {b['Nom']} en ({r},{c})")
                        
                        if self.solve(buildings[1:]): return True
                        
                        # Backtrack
                        self.grid[r:r+h, c:c+w] = 1
                        self.placed_buildings.pop()
                        self.log(f"Enlevé : {b['Nom']} de ({r},{c})")
                        
                        if len(self.journal) >= self.max_entries: return True
        return False

def generate_excel(planner, full_queue):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        
        # --- CALCUL DE LA CULTURE CUMULATIVE ---
        # On crée une carte vide pour additionner les rayonnements
        culture_map = np.zeros((planner.rows, planner.cols))
        for pb in planner.placed_buildings:
            if pb['info']['Type'] == 'Culturel':
                ray = int(pb['info'].get('Rayonnement', 0))
                val = float(pb['info'].get('Culture', 0))
                r_s, r_e = max(0, pb['r']-ray), min(planner.rows, pb['r']+pb['h']+ray)
                c_s, c_e = max(0, pb['c']-ray), min(planner.cols, pb['c']+pb['w']+ray)
                culture_map[r_s:r_e, c_s:c_e] += val # L'addition cumulative est ici

        # 1. Onglet Production
        stats_prod = []
        totals = {"Guerison": 0, "Nourriture": 0, "Or": 0}
        for pb in planner.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                c_recue = culture_map[pb['r'], pb['c']]
                boost = 0
                for b_val, b_perc in [(pb['info'].get('Boost 100%'), 100), 
                                      (pb['info'].get('Boost 50%'), 50), 
                                      (pb['info'].get('Boost 25%'), 25)]:
                    if pd.notnull(b_val) and b_val > 0 and c_recue >= b_val:
                        boost = b_perc
                        break
                stats_prod.append([pb['info']['Nom'], c_recue, f"{boost}%"])
                ptype = pb['info'].get('Production')
                if ptype in totals: totals[ptype] += c_recue

        pd.DataFrame(stats_prod, columns=["Bâtiment", "Culture Recue", "Boost"]).to_excel(writer, sheet_name="Production", index=False)
        pd.DataFrame(list(totals.items()), columns=["Type Production", "Culture Totale"]).to_excel(writer, sheet_name="Synthese", index=False)

        # 2. Plan visuel
        ws = writer.book.create_sheet("Plan_Terrain")
        for pb in planner.placed_buildings:
            fill = PatternFill(start_color=COLORS.get(pb['info']['Type'], 'FFFFFF'), fill_type='solid')
            for r in range(pb['r'], pb['r']+pb['h']):
                for c in range(pb['c'], pb['c']+pb['w']):
                    ws.cell(row=r+1, column=c+1, value=pb['info']['Nom']).fill = fill

        # 3. Résumé et Non placés
        placed_names = [p['info']['Nom'] for p in planner.placed_buildings]
        # On compare par index pour gérer les doublons de noms
        not_placed = full_queue[len(planner.placed_buildings):]
        
        summary = [
            ["Cases libres initiales", planner.initial_free_cells],
            ["Cases non utilisées", planner.initial_free_cells - sum(p['w']*p['h'] for p in planner.placed_buildings)],
            ["Bâtiments non placés", len(not_placed)],
            ["Statut", "STOP: LIMITE JOURNAL" if len(planner.journal) >= planner.max_entries else "OK"]
        ]
        pd.DataFrame(summary).to_excel(writer, sheet_name="Resume", index=False, header=False)
        pd.DataFrame(planner.journal, columns=["Journal"]).to_excel(writer, sheet_name="Journal", index=False)

    return output.getvalue()

# --- INTERFACE STREAMLIT ---
st.title("Optimiseur de Ville - Version Stable 🏗️")
file = st.file_uploader("Charger Ville.xlsx", type="xlsx")

if file:
    t_df = pd.read_excel(file, sheet_name=0, header=None)
    b_df = pd.read_excel(file, sheet_name=1)
    
    # Forçage des types numériques
    for col in ['Culture', 'Rayonnement', 'Boost 25%', 'Boost 50%', 'Boost 100%']:
        if col in b_df.columns:
            b_df[col] = pd.to_numeric(b_df[col], errors='coerce').fillna(0)

    # Tri par taille (Surface) pour maximiser le remplissage
    b_df['Surface'] = b_df['Longueur'] * b_df['Largeur']
    b_df = b_df.sort_values('Surface', ascending=False)
    
    full_queue = []
    for _, r in b_df.iterrows():
        for _ in range(int(r['Quantite'])):
            full_queue.append(r.to_dict())

    planner = CityPlanner(t_df.values)
    with st.spinner("Calcul en cours..."):
        planner.solve(full_queue)
    
    st.success(f"Calcul terminé ! {len(planner.placed_buildings)} bâtiments placés.")
    st.download_button("Télécharger Résultats", generate_excel(planner, full_queue), "Resultat_Placement.xlsx")
