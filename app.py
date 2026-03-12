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
        # 1 = Libre, 0 = Occupé, X = Obstacle (non libre)
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
        self.max_entries = 10000  # Mise à jour de la limite à 10 000
        self.interrupted = False

    def log(self, msg):
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)
        else:
            self.interrupted = True

    def is_adjacent_to_X(self, r, c, w, h):
        # Vérifie si le bâtiment touche une case X
        r_min, r_max = max(0, r-1), min(self.rows, r+h+1)
        c_min, c_max = max(0, c-1), min(self.cols, c+w+1)
        return np.any(self.border_mask[r_min:r_max, c_min:c_max])

    def can_place(self, r, c, w, h, remaining_queue):
        if r + h > self.rows or c + w > self.cols:
            return False
        if not np.all(self.grid[r:r+h, c:c+w] == 1):
            return False
        
        # Règle de l'espace pour le plus grand bâtiment restant
        if remaining_queue:
            biggest = remaining_queue[0]
            needed = biggest['Longueur'] * biggest['Largeur']
            if (np.sum(self.grid == 1) - (w * h)) < needed:
                return False
        return True

    def solve(self, buildings):
        if not buildings or self.interrupted:
            return True

        b = buildings[0]
        self.log(f"Évaluation de : {b['Nom']}")

        # Test des orientations
        dims = [(b['Largeur'], b['Longueur']), (b['Longueur'], b['Largeur'])]
        if b['Largeur'] == b['Longueur']: dims = [dims[0]]

        for w, h in dims:
            for r in range(self.rows - h + 1):
                for c in range(self.cols - w + 1):
                    # Règle Neutre : doit être contre un 'X'
                    if b['Type'] == 'Neutre' and not self.is_adjacent_to_X(r, c, w, h):
                        continue

                    if self.can_place(r, c, w, h, buildings[1:]):
                        self.grid[r:r+h, c:c+w] = 0
                        self.placed_buildings.append({'info': b, 'r': r, 'c': c, 'w': w, 'h': h})
                        self.log(f"Placé : {b['Nom']} en ({r},{c})")
                        
                        if self.solve(buildings[1:]): return True
                        
                        if self.interrupted: return False
                        self.log(f"Enlevé : {b['Nom']} de ({r},{c})")
                        self.grid[r:r+h, c:c+w] = 1
                        self.placed_buildings.pop()
        return False

# --- EXPORT EXCEL ---
def generate_excel(planner, full_queue):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Journal
        pd.DataFrame(planner.journal, columns=["Journal des opérations"]).to_excel(writer, sheet_name="Journal", index=False)
        
        # 2 & 3. Culture et Boosts
        culture_map = np.zeros((planner.rows, planner.cols))
        for pb in planner.placed_buildings:
            if pb['info']['Type'] == 'Culturel':
                ray = int(pb['info'].get('Rayonnement', 0))
                val = pb['info'].get('Culture', 0)
                r_s, r_e = max(0, pb['r']-ray), min(planner.rows, pb['r']+pb['h']+ray)
                c_s, c_e = max(0, pb['c']-ray), min(planner.cols, pb['c']+pb['w']+ray)
                culture_map[r_s:r_e, c_s:c_e] += val

        stats_prod = []
        totals = {"Guerison": 0, "Nourriture": 0, "Or": 0}
        for pb in planner.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                c_recue = np.max(culture_map[pb['r']:pb['r']+pb['h'], pb['c']:pb['c']+pb['w']])
                boost = 0
                if c_recue >= pb['info'].get('Boost 100%', 1e9): boost = 100
                elif c_recue >= pb['info'].get('Boost 50%', 1e9): boost = 50
                elif c_recue >= pb['info'].get('Boost 25%', 1e9): boost = 25
                stats_prod.append([pb['info']['Nom'], c_recue, f"{boost}%"])
                
                ptype = str(pb['info'].get('Production', ''))
                if ptype in totals: totals[ptype] += c_recue

        pd.DataFrame(stats_prod, columns=["Bâtiment", "Culture", "Boost"]).to_excel(writer, sheet_name="Production", index=False)
        pd.DataFrame(list(totals.items()), columns=["Type", "Culture Totale"]).to_excel(writer, sheet_name="Synthese_Par_Type", index=False)

        # 4. Plan visuel
        ws = writer.book.create_sheet("Plan_Terrain")
        for pb in planner.placed_buildings:
            fill = PatternFill(start_color=COLORS.get(pb['info']['Type'], 'FFFFFF'), fill_type='solid')
            for r in range(pb['r'], pb['r']+pb['h']):
                for c in range(pb['c'], pb['c']+pb['w']):
                    cell = ws.cell(row=r+1, column=c+1, value=pb['info']['Nom'])
                    cell.fill = fill

        # 5, 6, 7. Stats Bâtiments Non Placés
        placed_ids = [id(p['info']) for p in planner.placed_buildings]
        not_placed = [b for b in full_queue if id(b) not in placed_ids]
        
        cases_occupees = sum(p['w'] * p['h'] for p in planner.placed_buildings)
        summary = [
            ["Nombre de cases initiales libres", planner.initial_free_cells],
            ["Nombre de cases restant non utilisées", planner.initial_free_cells - cases_occupees],
            ["Cases totales des bâtiments non placés", sum(b['Longueur']*b['Largeur'] for b in not_placed)],
            ["Statut de l'algorithme", "STOP: LIMITE 10000 LIGNES" if planner.interrupted else "TERMINE"]
        ]
        pd.DataFrame(summary).to_excel(writer, sheet_name="Resume_Global", index=False, header=False)
        pd.DataFrame(not_placed).to_excel(writer, sheet_name="Batiments_Non_Places", index=False)

    return output.getvalue()

# --- STREAMLIT ---
st.title("Optimiseur de Ville - Limite 10k 🏗️")
uploaded = st.file_uploader("Fichier Ville.xlsx", type="xlsx")

if uploaded:
    t_df = pd.read_excel(uploaded, sheet_name=0, header=None)
    b_df = pd.read_excel(uploaded, sheet_name=1)
    b_df.columns = b_df.columns.str.strip() # Correction des espaces dans les titres
    
    # Préparation de la file
    neutres = b_df[b_df['Type'] == 'Neutre'].sort_values('Longueur', ascending=False)
    autres = b_df[b_df['Type'] != 'Neutre'].sort_values('Longueur', ascending=False)
    
    full_queue = []
    # Expansion des quantités
    for _, r in neutres.iterrows():
        for _ in range(int(r['Quantite'])): full_queue.append(r.to_dict())
    
    c_list = autres[autres['Type'] == 'Culturel'].to_dict('records')
    p_list = autres[autres['Type'] == 'Producteur'].to_dict('records')
    
    for i in range(max(len(c_list), len(p_list))):
        if i < len(c_list):
            for _ in range(int(c_list[i]['Quantite'])): full_queue.append(c_list[i].copy())
        if i < len(p_list):
            for _ in range(int(p_list[i]['Quantite'])): full_queue.append(p_list[i].copy())

    planner = CityPlanner(t_df.values)
    planner.solve(full_queue)
    
    result_data = generate_excel(planner, full_queue)
    st.success(f"Traitement fini ({len(planner.journal)} lignes au journal)")
    st.download_button("Télécharger Resultat_Placement.xlsx", result_data, "Resultat_Placement.xlsx")
