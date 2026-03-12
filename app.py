import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# -[span_5](start_span)-- CONFIGURATION DES COULEURS[span_5](end_span) ---
COLORS = {
    'Culturel': 'FFFFA500',  # Orange
    'Producteur': 'FF008000', # Vert
    'Neutre': 'FF808080'      # Gris
}

class CityPlanner:
    def __init__(self, terrain_data):
        self.rows = len(terrain_data)
        self.cols = len(terrain_data[0])
        # [span_6](start_span)1 = Libre, 0 = Occupé, X = Bordure (Obstacle servant d'ancrage)[span_6](end_span)
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
                    self.grid[r,c] = 0 # X n'est pas une case libre
                    self.border_mask[r,c] = True
        
        self.journal = []
        self.placed_buildings = []
        self.max_entries = 1000
        self.interrupted = False

    def log(self, msg):
        [span_7](start_span)"""Ajoute une entrée au journal si la limite n'est pas atteinte[span_7](end_span)."""
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)
        else:
            self.interrupted = True

    def is_adjacent_to_X(self, r, c, w, h):
        [span_8](start_span)"""Vérifie si le bâtiment touche une case X[span_8](end_span)."""
        r_min, r_max = max(0, r-1), min(self.rows, r+h+1)
        c_min, c_max = max(0, c-1), min(self.cols, c+w+1)
        return np.any(self.border_mask[r_min:r_max, c_min:c_max])

    def can_place(self, r, c, w, h, remaining_queue):
        [span_9](start_span)"""Vérifie la disponibilité et la règle de l'espace restant[span_9](end_span)."""
        if r + h > self.rows or c + w > self.cols:
            return False
        if not np.all(self.grid[r:r+h, c:c+w] == 1):
            return False
        
        if remaining_queue:
            biggest = remaining_queue[0]
            needed = biggest['Longueur'] * biggest['Largeur']
            if (np.sum(self.grid == 1) - (w * h)) < needed:
                return False
        return True

    def solve(self, buildings):
        [span_10](start_span)"""Algorithme récursif avec backtracking[span_10](end_span)."""
        if not buildings or len(self.journal) >= self.max_entries:
            return True

        b = buildings[0]
        self.log(f"Évaluation de : {b['Nom']}")

        # [span_11](start_span)Test des deux orientations[span_11](end_span)
        dims = [(b['Largeur'], b['Longueur']), (b['Longueur'], b['Largeur'])]
        if b['Largeur'] == b['Longueur']: dims = [dims[0]]

        for w, h in dims:
            for r in range(self.rows - h + 1):
                for c in range(self.cols - w + 1):
                    # [span_12](start_span)Contrainte : Neutres à côté des 'X'[span_12](end_span)
                    if b['Type'] == 'Neutre' and not self.is_adjacent_to_X(r, c, w, h):
                        continue

                    if self.can_place(r, c, w, h, buildings[1:]):
                        # [span_13](start_span)Placement[span_13](end_span)
                        self.grid[r:r+h, c:c+w] = 0
                        self.placed_buildings.append({'info': b, 'r': r, 'c': c, 'w': w, 'h': h})
                        self.log(f"Placé : {b['Nom']} en ({r},{c})")
                        
                        if self.solve(buildings[1:]): return True
                        
                        # [span_14](start_span)Retour en arrière[span_14](end_span)
                        if len(self.journal) >= self.max_entries: return False
                        self.log(f"Enlevé : {b['Nom']} de ({r},{c})")
                        self.grid[r:r+h, c:c+w] = 1
                        self.placed_buildings.pop()
        return False

# --- LOGIQUE DE CALCUL ET EXPORT ---
def generate_excel(planner, full_queue):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. [span_15](start_span)Journal[span_15](end_span)
        pd.DataFrame(planner.journal, columns=["Journal"]).to_excel(writer, sheet_name="Journal", index=False)
        
        # [span_16](start_span)[span_17](start_span)2 & 3. Culture et Boosts[span_16](end_span)[span_17](end_span)
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
        pd.DataFrame(list(totals.items()), columns=["Type", "Culture Totale"]).to_excel(writer, sheet_name="Synthese_Types", index=False)

        # 4. [span_18](start_span)Plan visuel[span_18](end_span)
        ws = writer.book.create_sheet("Plan_Terrain")
        for pb in planner.placed_buildings:
            fill = PatternFill(start_color=COLORS.get(pb['info']['Type'], 'FFFFFF'), fill_type='solid')
            for r in range(pb['r'], pb['r']+pb['h']):
                for c in range(pb['c'], pb['c']+pb['w']):
                    cell = ws.cell(row=r+1, column=c+1, value=pb['info']['Nom'])
                    cell.fill = fill

        # [span_19](start_span)5, 6 & 7. Bâtiments non placés et statistiques de cases[span_19](end_span)
        placed_names = [p['info']['Nom'] for p in planner.placed_buildings]
        not_placed = []
        temp_placed = placed_names.copy()
        for b in full_queue:
            if b['Nom'] in temp_placed: temp_placed.remove(b['Nom'])
            else: not_placed.append(b)

        cases_occ = sum(p['w'] * p['h'] for p in planner.placed_buildings)
        summary = [
            ["Nombre de cases non utilisées", planner.initial_free_cells - cases_occ],
            ["Cases des bâtiments non placés", sum(b['Longueur']*b['Largeur'] for b in not_placed)],
            ["Statut", "LIMITE ATTEINTE" if planner.interrupted else "OK"]
        ]
        pd.DataFrame(summary).to_excel(writer, sheet_name="Resume", index=False, header=False)
        pd.DataFrame(not_placed).to_excel(writer, sheet_name="Non_Places", index=False)

    return output.getvalue()

# --- INTERFACE STREAMLIT ---
st.title("Générateur de Ville Optimisée 🏗️")
uploaded = st.file_uploader("Charger le fichier Ville.xlsx", type="xlsx")

if uploaded:
    t_df = pd.read_excel(uploaded, sheet_name=0, header=None)
    b_df = pd.read_excel(uploaded, sheet_name=1)
    
    # [span_20](start_span)Tri et préparation de la file[span_20](end_span)
    neutres = b_df[b_df['Type'] == 'Neutre'].sort_values('Longueur', ascending=False)
    autres = b_df[b_df['Type'] != 'Neutre'].sort_values('Longueur', ascending=False)
    
    full_queue = []
    for _, r in neutres.iterrows(): full_queue.extend([r.to_dict()] * int(r['Quantite']))
    c_list = autres[autres['Type'] == 'Culturel'].to_dict('records')
    p_list = avec_prod = autres[autres['Type'] == 'Producteur'].to_dict('records')
    
    for i in range(max(len(c_list), len(p_list))):
        if i < len(c_list): full_queue.append(c_list[i])
        if i < len(p_list): full_queue.append(p_list[i])

    planner = CityPlanner(t_df.values)
    planner.solve(full_queue)
    
    result_file = generate_excel(planner, full_queue)
    st.success("Calcul terminé (ou limite atteinte).")
    st.download_button("Télécharger Resultat_Placement.xlsx", result_file, "Resultat_Placement.xlsx")
