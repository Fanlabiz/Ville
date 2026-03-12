import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# --- CONFIGURATION ---
COLORS = {
    'Culturel': 'FFFFA500',  # Orange
    'Producteur': 'FF008000', # Vert
    'Neutre': 'FF808080'      # Gris
}

class CityPlanner:
    def __init__(self, grid_init):
        self.grid = grid_init.copy()
        self.rows, self.cols = self.grid.shape
        self.journal = []
        self.placed_buildings = [] # Liste de dicts: {info, r, c, w, h}
        self.max_journal = 1000

    def log(self, msg):
        if len(self.journal) < self.max_journal:
            self.journal.append(msg)

    def is_on_border(self, r, c, w, h):
        return r == 0 or c == 0 or (r + h) == self.rows or (c + w) == self.cols

    def can_place(self, r, c, w, h, remaining_buildings):
        # [span_4](start_span)Vérification des limites et collision[span_4](end_span)
        if r + h > self.rows or c + w > self.cols:
            return False
        if not np.all(self.grid[r:r+h, c:c+w] == 1):
            return False
            
        # [span_5](start_span)Règle du plus grand bâtiment restant[span_5](end_span)
        if remaining_buildings:
            max_b = remaining_buildings[0]
            max_area = max_b['Longueur'] * max_b['Largeur']
            if np.sum(self.grid) - (w * h) < max_area:
                return False
        return True

    def solve(self, buildings):
        if not buildings or len(self.journal) >= self.max_journal:
            return True

        b = buildings[0]
        self.log(f"Évaluation de : {b['Nom']}")

        # [span_6](start_span)Essais d'orientations[span_6](end_span)
        orientations = [(b['Largeur'], b['Longueur']), (b['Longueur'], b['Largeur'])]
        if b['Largeur'] == b['Longueur']: orientations = [orientations[0]]

        for w, h in orientations:
            for r in range(self.rows - h + 1):
                for c in range(self.cols - w + 1):
                    # [span_7](start_span)Contrainte spécifique Neutre sur les bords[span_7](end_span)
                    if b['Type'] == 'Neutre' and not self.is_on_border(r, c, w, h):
                        continue

                    if self.can_place(r, c, w, h, buildings[1:]):
                        # [span_8](start_span)Placement[span_8](end_span)
                        self.grid[r:r+h, c:c+w] = 0
                        self.placed_buildings.append({'info': b, 'r': r, 'c': c, 'w': w, 'h': h})
                        self.log(f"Placé : {b['Nom']} en ({r},{c})")

                        if self.solve(buildings[1:]):
                            return True
                        
                        # [span_9](start_span)Backtrack[span_9](end_span)
                        self.log(f"Enlevé : {b['Nom']} de ({r},{c})")
                        self.grid[r:r+h, c:c+w] = 1
                        self.placed_buildings.pop()
                        
                        if len(self.journal) >= self.max_journal: return False

        return False

    def calculate_stats(self):
        # [span_10](start_span)Calcul de la culture reçue par bâtiment[span_10](end_span)
        stats = []
        culture_map = np.zeros((self.rows, self.cols))
        
        # 1. [span_11](start_span)Remplir la carte des rayonnements[span_11](end_span)
        for pb in self.placed_buildings:
            if pb['info']['Type'] == 'Culturel':
                ray = int(pb['info'].get('Rayonnement', 0))
                val = pb['info'].get('Culture', 0)
                r_start, r_end = max(0, pb['r']-ray), min(self.rows, pb['r']+pb['h']+ray)
                c_start, c_end = max(0, pb['c']-ray), min(self.cols, pb['c']+pb['w']+ray)
                culture_map[r_start:r_end, c_start:c_end] += val

        # 2. [span_12](start_span)Calculer pour chaque Producteur[span_12](end_span)
        total_guerison = total_nourriture = total_or = 0
        for pb in self.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                # [span_13](start_span)Somme de la culture sur les cases occupées[span_13](end_span)
                culture_recue = np.max(culture_map[pb['r']:pb['r']+pb['h'], pb['c']:pb['c']+pb['w']])
                
                # [span_14](start_span)Calcul des boosts[span_14](end_span)
                boost = 0
                if culture_recue >= pb['info'].get('Boost 100%', 999999): boost = 100
                elif culture_recue >= pb['info'].get('Boost 50%', 999999): boost = 50
                elif culture_recue >= pb['info'].get('Boost 25%', 999999): boost = 25
                
                prod_type = str(pb['info'].get('Production', '')).lower()
                if 'guerison' in prod_type: total_guerison += culture_recue
                elif 'nourriture' in prod_type: total_nourriture += culture_recue
                elif 'or' in prod_type: total_or += culture_recue
                
                stats.append({
                    'Nom': pb['info']['Nom'],
                    'Culture Recue': culture_recue,
                    'Boost Atteint (%)': boost
                })
        
        return stats, (total_guerison, total_nourriture, total_or)

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="City Planner")
st.title("Optimiseur de Terrain 🏗️")

file = st.file_uploader("Charger Ville.xlsx", type="xlsx")

if file:
    # [span_15](start_span)Lecture[span_15](end_span)
    terrain_df = pd.read_excel(file, sheet_name=0, header=None)
    bat_df = pd.read_excel(file, sheet_name=1)
    
    # [span_16](start_span)Tri selon l'algorithme[span_16](end_span)
    neutres = bat_df[bat_df['Type'] == 'Neutre'].sort_values(by=['Longueur'], ascending=False)
    autres = bat_df[bat_df['Type'] != 'Neutre'].sort_values(by=['Longueur'], ascending=False)
    
    # [span_17](start_span)Expansion des quantités[span_17](end_span)
    bnts_list = []
    for _, r in neutres.iterrows(): bnts_list.extend([r.to_dict()] * int(r['Quantite']))
    # Note: On alterne Culturel/Producteur ici
    cults = autres[autres['Type'] == 'Culturel'].to_dict('records')
    prods = autres[autres['Type'] == 'Producteur'].to_dict('records')
    # ... logique d'alternance ...
    final_queue = bnts_list + cults + prods # Simplifié pour l'exemple

    planner = CityPlanner(terrain_df.values)
    planner.solve(final_queue)
    
    # [span_18](start_span)Export Excel[span_18](end_span)
    output = io.BytesIO()
    wb = Workbook()
    
    # 1. [span_19](start_span)Journal[span_19](end_span)
    ws1 = wb.active
    ws1.title = "Journal"
    for i, entry in enumerate(planner.journal, 1): ws1.cell(row=i, column=1, value=entry)
    
    # 2. [span_20](start_span)Stats & Boosts[span_20](end_span)
    stats, totals = planner.calculate_stats()
    ws2 = wb.create_sheet("Resultats_Production")
    ws2.append(["Bâtiment", "Culture Reçue", "Boost %"])
    for s in stats: ws2.append([s['Nom'], s['Culture Recue'], s['Boost Atteint (%)']])
    ws2.append([])
    ws2.append(["Total Guérison", totals[0], "Total Nourriture", totals[1], "Total Or", totals[2]])

    # 3. [span_21](start_span)Terrain Visuel[span_21](end_span)
    ws3 = wb.create_sheet("Plan_Terrain")
    for pb in planner.placed_buildings:
        fill = PatternFill(start_color=COLORS.get(pb['info']['Type'], 'FFFFFF'), fill_type='solid')
        for r in range(pb['r']+1, pb['r']+pb['h']+1):
            for c in range(pb['c']+1, pb['c']+pb['w']+1):
                cell = ws3.cell(row=r, column=c, value=pb['info']['Nom'])
                cell.fill = fill

    wb.save(output)
    st.download_button("Télécharger Resultat_Placement.xlsx", output.getvalue(), "Resultat_Placement.xlsx")
