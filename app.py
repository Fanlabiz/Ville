import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# --- CONFIGURATION ---
COLORS = {
    [span_1](start_span)'Culturel': 'FFFFA500',  # Orange[span_1](end_span)
    [span_2](start_span)'Producteur': 'FF008000', # Vert[span_2](end_span)
    [span_3](start_span)'Neutre': 'FF808080'      # Gris[span_3](end_span)
}

class CityPlanner:
    def __init__(self, terrain_data):
        self.rows = len(terrain_data)
        self.cols = len(terrain_data[0])
        # [span_4](start_span)1 = Libre, 0 = Occupé/Obstacle, X = Bord (non libre)[span_4](end_span)
        self.grid = np.zeros((self.rows, self.cols))
        self.border_mask = np.zeros((self.rows, self.cols), dtype=bool)
        self.initial_free_cases = 0
        
        for r in range(self.rows):
            for c in range(self.cols):
                val = str(terrain_data[r][c]).strip().upper()
                if val == '1': 
                    self.grid[r,c] = 1
                    self.initial_free_cases += 1
                elif val == 'X':
                    self.grid[r,c] = 0 # X n'est PAS libre
                    self.border_mask[r,c] = True
        
        self.journal = []
        self.placed_buildings = []
        self.max_entries = 1000
        self.interrupted = False

    def log(self, msg):
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)
        else:
            self.interrupted = True

    def is_adjacent_to_X(self, r, c, w, h):
        # [span_5](start_span)Vérifie si le bâtiment touche une case X[span_5](end_span)
        r_min, r_max = max(0, r-1), min(self.rows, r+h+1)
        c_min, c_max = max(0, c-1), min(self.cols, c+w+1)
        return np.any(self.border_mask[r_min:r_max, c_min:c_max])

    def can_place(self, r, c, w, h, remaining_queue):
        if r + h > self.rows or c + w > self.cols:
            return False
        if not np.all(self.grid[r:r+h, c:c+w] == 1):
            return False
        
        # [span_6](start_span)Règle de l'espace pour le plus grand bâtiment restant[span_6](end_span)
        if remaining_queue:
            biggest = remaining_queue[0]
            if np.sum(self.grid == 1) - (w * h) < (biggest['Longueur'] * biggest['Largeur']):
                return False
        return True

    def solve(self, buildings):
        if not buildings or len(self.journal) >= self.max_entries:
            return True

        b = buildings[0]
        [span_7](start_span)self.log(f"Évaluation de : {b['Nom']}") #[span_7](end_span)

        [span_8](start_span)dims = [(b['Largeur'], b['Longueur']), (b['Longueur'], b['Largeur'])] #[span_8](end_span)
        if b['Largeur'] == b['Longueur']: dims = [dims[0]]

        for w, h in dims:
            for r in range(self.rows - h + 1):
                for c in range(self.cols - w + 1):
                    if b['Type'] == 'Neutre' and not self.is_adjacent_to_X(r, c, w, h):
                        continue

                    if self.can_place(r, c, w, h, buildings[1:]):
                        self.grid[r:r+h, c:c+w] = 0
                        self.placed_buildings.append({'info': b, 'r': r, 'c': c, 'w': w, 'h': h})
                        [span_9](start_span)self.log(f"Placé : {b['Nom']} en ({r},{c})") #[span_9](end_span)
                        
                        if self.solve(buildings[1:]): return True
                        
                        if len(self.journal) >= self.max_entries: return False

                        [span_10](start_span)self.log(f"Enlevé : {b['Nom']} de ({r},{c})") #[span_10](end_span)
                        self.grid[r:r+h, c:c+w] = 1
                        self.placed_buildings.pop()
        return False

# --- GESTION DES RÉSULTATS ---
def generate_excel(planner, full_queue):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. [span_11](start_span)Journal[span_11](end_span)
        pd.DataFrame(planner.journal, columns=["Journal des opérations"]).to_excel(writer, sheet_name="Journal", index=False)
        
        # [span_12](start_span)2 & 3. Culture et Boosts[span_12](end_span)
        stats_data = []
        culture_map = np.zeros((planner.rows, planner.cols))
        for pb in planner.placed_buildings:
            if pb['info']['Type'] == 'Culturel':
                ray = int(pb['info'].get('Rayonnement', 0))
                v = pb['info'].get('Culture', 0)
                r_s, r_e = max(0, pb['r']-ray), min(planner.rows, pb['r']+pb['h']+ray)
                c_s, c_e = max(0, pb['c']-ray), min(planner.cols, pb['c']+pb['w']+ray)
                culture_map[r_s:r_e, c_s:c_e] += v

        totals = {"Guerison": 0, "Nourriture": 0, "Or": 0}
        for pb in planner.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                [span_13](start_span)c_recue = np.max(culture_map[pb['r']:pb['r']+pb['h'], pb['c']:pb['c']+pb['w']]) #[span_13](end_span)
                boost = 0
                if c_recue >= pb['info'].get('Boost 100%', 1e9): boost = 100
                elif c_recue >= pb['info'].get('Boost 50%', 1e9): boost = 50
                elif c_recue >= pb['info'].get('Boost 25%', 1e9): boost = 25
                stats_data.append([pb['info']['Nom'], c_recue, f"{boost}%"])
                
                prod_type = str(pb['info'].get('Production', ''))
                if prod_type in totals: totals[prod_type] += c_recue

        pd.DataFrame(stats_data, columns=["Bâtiment", "Culture Reçue", "Boost"]).to_excel(writer, sheet_name="Stats_Culture", index=False)
        pd.DataFrame(list(totals.items()), columns=["Type Production", "Culture Totale"]).to_excel(writer, sheet_name="Totaux_Culturels", index=False)

        # 4. [span_14](start_span)Le Terrain[span_14](end_span)
        ws = writer.book.create_sheet("Plan_Terrain")
        for pb in planner.placed_buildings:
            fill = PatternFill(start_color=COLORS.get(pb['info']['Type'], 'FFFFFF'), fill_type='solid')
            for r in range(pb['r'], pb['r']+pb['h']):
                for c in range(pb['c'], pb['c']+pb['w']):
                    cell = ws.cell(row=r+1, column=c+1, value=pb['info']['Nom'])
                    cell.fill = fill

        # [span_15](start_span)5, 6 & 7. Bâtiments non placés et stats cases[span_15](end_span)
        placed_names = [p['info']['Nom'] for p in planner.placed_buildings]
        # On compare la file initiale avec ce qui a été réellement placé
        not_placed = []
        temp_placed = placed_names.copy()
        for b in full_queue:
            if b['Nom'] in temp_placed:
                temp_placed.remove(b['Nom'])
            else:
                not_placed.append(b)

        cases_occupees = sum(p['w'] * p['h'] for p in planner.placed_buildings)
        cases_non_placees = sum(b['Longueur'] * b['Largeur'] for b in not_placed)
        
        summary = [
            ["Statistique", "Valeur"],
            ["Nombre de cases non utilisées", planner.initial_free_cases - cases_occupees],
            ["Nombre de cases des bâtiments non placés", cases_non_placees],
            ["Statut de l'algorithme", "Interrompu (1000 entrées)" if planner.interrupted else "Terminé"]
        ]
        pd.DataFrame(summary).to_excel(writer, sheet_name="Resume_Global", index=False, header=False)
        pd.DataFrame(not_placed).to_excel(writer, sheet_name="Non_Places", index=False)

    return output.getvalue()

# --- STREAMLIT APP ---
st.title("City Planner Pro 🏗️")
file = st.file_uploader("Importer Ville.xlsx", type="xlsx")

if file:
    t_df = pd.read_excel(file, sheet_name=0, header=None)
    b_df = pd.read_excel(file, sheet_name=1)
    
    # [span_16](start_span)Préparation de la file d'attente[span_16](end_span)
    neutres = b_df[b_df['Type'] == 'Neutre'].sort_values('Longueur', ascending=False)
    autres = b_df[b_df['Type'] != 'Neutre'].sort_values('Longueur', ascending=False)
    
    full_queue = []
    for _, r in neutres.iterrows(): full_queue.extend([r.to_dict()] * int(r['Quantite']))
    # Alternance simple Culturel / Producteur
    c_list = autres[autres['Type'] == 'Culturel'].to_dict('records')
    p_list = autres[autres['Type'] == 'Producteur'].to_dict('records')
    for i in range(max(len(c_list), len(p_list))):
        if i < len(c_list): full_queue.append(c_list[i])
        if i < len(p_list): full_queue.append(p_list[i])

    planner = CityPlanner(t_df.values)
    planner.solve(full_queue)
    
    excel_data = generate_excel(planner, full_queue)
    st.download_button("Télécharger le fichier de résultats", excel_data, "Resultat_Placement.xlsx")
