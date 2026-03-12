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
        else:
            self.interrupted = True

    def is_adjacent_to_X(self, r, c, w, h):
        r_min, r_max = max(0, r-1), min(self.rows, r+h+1)
        c_min, c_max = max(0, c-1), min(self.cols, c+w+1)
        return np.any(self.border_mask[r_min:r_max, c_min:c_max])

    def can_place(self, r, c, w, h, remaining_queue):
        # 1. Vérification limites et collision
        if r + h > self.rows or c + w > self.cols:
            return False
        if not np.all(self.grid[r:r+h, c:c+w] == 1):
            return False
        
        # 2. RÈGLE DU POINT 2 : Vérifier qu'il reste de la place pour le plus grand restant
        if remaining_queue:
            # On simule le placement pour voir l'espace restant
            self.grid[r:r+h, c:c+w] = 0
            biggest = remaining_queue[0]
            bw, bh = biggest['Largeur'], biggest['Longueur']
            
            can_fit_biggest = False
            # On cherche au moins UNE place pour le plus grand (en testant les 2 sens)
            for orient in [(bw, bh), (bh, bw)]:
                ow, oh = orient
                for rr in range(self.rows - oh + 1):
                    for cc in range(self.cols - ow + 1):
                        if np.all(self.grid[rr:rr+oh, cc:cc+ow] == 1):
                            can_fit_biggest = True
                            break
                    if can_fit_biggest: break
                if can_fit_biggest: break
            
            # On remet la grille en état
            self.grid[r:r+h, c:c+w] = 1
            if not can_fit_biggest:
                return False

        return True

    def solve(self, buildings):
        if not buildings or self.interrupted:
            return True

        b = buildings[0]
        self.log(f"Évaluation de : {b['Nom']}")
        
        dims = [(b['Largeur'], b['Longueur']), (b['Longueur'], b['Largeur'])]
        if b['Largeur'] == b['Longueur']: dims = [dims[0]]

        # PHASE 1 : Tentative Prioritaire (Bords pour Neutres)
        for w, h in dims:
            for r in range(self.rows - h + 1):
                for c in range(self.cols - w + 1):
                    if b['Type'] == 'Neutre' and not self.is_adjacent_to_X(r, c, w, h):
                        continue
                    if self.can_place(r, c, w, h, buildings[1:]):
                        if self.try_placement(b, r, c, w, h, buildings):
                            return True

        # PHASE 2 : Tentative de secours (Si Neutre n'a pas trouvé de bord)
        if b['Type'] == 'Neutre':
            self.log(f"Bords saturés, essai interne pour : {b['Nom']}")
            for w, h in dims:
                for r in range(self.rows - h + 1):
                    for c in range(self.cols - w + 1):
                        if self.can_place(r, c, w, h, buildings[1:]):
                            if self.try_placement(b, r, c, w, h, buildings):
                                return True
        return False

    def try_placement(self, b, r, c, w, h, buildings):
        self.grid[r:r+h, c:c+w] = 0
        self.placed_buildings.append({'info': b, 'r': r, 'c': c, 'w': w, 'h': h})
        self.log(f"Placé : {b['Nom']} en ({r},{c})")
        
        if self.solve(buildings[1:]): return True
        
        if not self.interrupted:
            self.log(f"Enlevé : {b['Nom']} de ({r},{c})")
            self.grid[r:r+h, c:c+w] = 1
            self.placed_buildings.pop()
        return False

# --- LOGIQUE D'EXPORT EXCEL ---
def generate_excel(planner, full_queue):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame(planner.journal, columns=["Journal"]).to_excel(writer, sheet_name="Journal", index=False)
        
        # Culture & Production
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
        pd.DataFrame(list(totals.items()), columns=["Type", "Culture Totale"]).to_excel(writer, sheet_name="Synthese", index=False)

        # Plan
        ws = writer.book.create_sheet("Plan_Terrain")
        for pb in planner.placed_buildings:
            fill = PatternFill(start_color=COLORS.get(pb['info']['Type'], 'FFFFFF'), fill_type='solid')
            for r in range(pb['r'], pb['r']+pb['h']):
                for c in range(pb['c'], pb['c']+pb['w']):
                    cell = ws.cell(row=r+1, column=c+1, value=pb['info']['Nom'])
                    cell.fill = fill

        # Resume
        placed_ids = [id(p['info']) for p in planner.placed_buildings]
        not_placed = [b for b in full_queue if id(b) not in placed_ids]
        cases_occ = sum(p['w'] * p['h'] for p in planner.placed_buildings)
        summary = [
            ["Cases libres initiales", planner.initial_free_cells],
            ["Cases non utilisées", planner.initial_free_cells - cases_occ],
            ["Volume non placé", sum(b['Longueur']*b['Largeur'] for b in not_placed)],
            ["Statut", "STOP: LIMITE JOURNAL" if planner.interrupted else "OK"]
        ]
        pd.DataFrame(summary).to_excel(writer, sheet_name="Resume", index=False, header=False)
        pd.DataFrame(not_placed).to_excel(writer, sheet_name="Non_Places", index=False)

    return output.getvalue()

st.title("Optimiseur de Cité - Règle du Plus Grand 🏗️")
uploaded = st.file_uploader("Charger Ville.xlsx", type="xlsx")

if uploaded:
    t_df = pd.read_excel(uploaded, sheet_name=0, header=None)
    b_df = pd.read_excel(uploaded, sheet_name=1)
    b_df.columns = b_df.columns.str.strip()
    
    # Tri par taille et par type
    neutres = b_df[b_df['Type'] == 'Neutre'].sort_values(['Longueur', 'Largeur'], ascending=False)
    autres = b_df[b_df['Type'] != 'Neutre'].sort_values(['Longueur', 'Largeur'], ascending=False)
    
    full_queue = []
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
    st.download_button("Télécharger le résultat final", generate_excel(planner, full_queue), "Resultat_Placement.xlsx")
