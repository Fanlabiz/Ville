import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill

# --- CONFIGURATION ---
COLORS = {
    'Culturel': 'FFFFA500',  # Orange [source 16]
    'Producteur': 'FF008000', # Vert [source 16]
    'Neutre': 'FF808080'      # Gris [source 16]
}

class CityPlanner:
    def __init__(self, terrain_data):
        self.rows = len(terrain_data)
        self.cols = len(terrain_data[0])
        # 1 = Libre, 0 = Occupé/Obstacle [source 3]
        self.grid = np.zeros((self.rows, self.cols))
        self.border_mask = np.zeros((self.rows, self.cols), dtype=bool)
        
        for r in range(self.rows):
            for c in range(self.cols):
                val = str(terrain_data[r][c]).strip().upper()
                if val == '1': 
                    self.grid[r,c] = 1
                elif val == 'X':
                    self.grid[r,c] = 0 # X n'est PAS libre
                    self.border_mask[r,c] = True # Mais c'est un bord pour les Neutres
                else:
                    self.grid[r,c] = 0
        
        self.journal = []
        self.placed_buildings = []
        self.max_entries = 1000

    def log(self, msg):
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)

    def is_adjacent_to_X(self, r, c, w, h):
        # Vérifie si le bâtiment touche une case X (nord, sud, est, ouest)
        r_min, r_max = max(0, r-1), min(self.rows, r+h+1)
        c_min, c_max = max(0, c-1), min(self.cols, c+w+1)
        return np.any(self.border_mask[r_min:r_max, c_min:c_max])

    def can_place(self, r, c, w, h, remaining_buildings):
        if r + h > self.rows or c + w > self.cols:
            return False
        # Doit être sur des cases 1 (Libres) [source 3]
        if not np.all(self.grid[r:r+h, c:c+w] == 1):
            return False
        
        # Vérification espace pour le plus grand restant [source 14]
        if remaining_buildings:
            biggest = remaining_buildings[0]
            if np.sum(self.grid == 1) - (w * h) < (biggest['Longueur'] * biggest['Largeur']):
                return False
        return True

    def solve(self, buildings):
        if not buildings or len(self.journal) >= self.max_entries:
            return True

        b = buildings[0]
        self.log(f"Évaluation de : {b['Nom']}")

        dims = [(b['Largeur'], b['Longueur']), (b['Longueur'], b['Largeur'])]
        if b['Largeur'] == b['Longueur']: dims = [dims[0]]

        for w, h in dims:
            for r in range(self.rows - h + 1):
                for c in range(self.cols - w + 1):
                    # Neutre : Doit être adjacent à un 'X' [source 13]
                    if b['Type'] == 'Neutre' and not self.is_adjacent_to_X(r, c, w, h):
                        continue

                    if self.can_place(r, c, w, h, buildings[1:]):
                        self.grid[r:r+h, c:c+w] = 0 # Occupe l'espace
                        self.placed_buildings.append({'info': b, 'r': r, 'c': c, 'w': w, 'h': h})
                        self.log(f"Placé : {b['Nom']} en ({r},{c})")
                        
                        if self.solve(buildings[1:]): return True
                        
                        self.log(f"Enlevé : {b['Nom']} de ({r},{c})") # Backtrack [source 15]
                        self.grid[r:r+h, c:c+w] = 1
                        self.placed_buildings.pop()
                        if len(self.journal) >= self.max_entries: return False
        return False

    def get_stats(self):
        # Calcul culture [source 10, 11]
        culture_map = np.zeros((self.rows, self.cols))
        for pb in self.placed_buildings:
            if pb['info']['Type'] == 'Culturel':
                ray = int(pb['info'].get('Rayonnement', 0))
                val = pb['info'].get('Culture', 0)
                r_s, r_e = max(0, pb['r']-ray), min(self.rows, pb['r']+pb['h']+ray)
                c_s, c_e = max(0, pb['c']-ray), min(self.cols, pb['c']+pb['w']+ray)
                culture_map[r_s:r_e, c_s:c_e] += val

        res = []
        totals = {"Guerison": 0, "Nourriture": 0, "Or": 0}
        for pb in self.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                c_recue = np.max(culture_map[pb['r']:pb['r']+pb['h'], pb['c']:pb['c']+pb['w']])
                boost = 0
                if c_recue >= pb['info'].get('Boost 100%', 1e9): boost = 100
                elif c_recue >= pb['info'].get('Boost 50%', 1e9): boost = 50
                elif c_recue >= pb['info'].get('Boost 25%', 1e9): boost = 25
                
                res.append([pb['info']['Nom'], c_recue, boost])
                p_type = str(pb['info'].get('Production', ''))
                if p_type in totals: totals[p_type] += c_recue
        return res, totals

# --- STREAMLIT UI ---
st.title("Optimiseur de Cité")
uploaded = st.file_uploader("Charger Ville.xlsx", type="xlsx")

if uploaded:
    t_df = pd.read_excel(uploaded, sheet_name=0, header=None)
    b_df = pd.read_excel(uploaded, sheet_name=1)
    
    # Préparation file [source 13]
    neutres = b_df[b_df['Type'] == 'Neutre'].sort_values('Longueur', ascending=False)
    others = b_df[b_df['Type'] != 'Neutre'].sort_values('Longueur', ascending=False)
    
    queue = []
    for _, r in neutres.iterrows(): queue.extend([r.to_dict()] * int(r['Quantite']))
    # Alternance simple pour l'exemple
    for _, r in others.iterrows(): queue.extend([r.to_dict()] * int(r['Quantite']))

    planner = CityPlanner(t_df.values)
    planner.solve(queue)
    
    # Export Excel [source 17]
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Journal [source 1]
        pd.DataFrame(planner.journal, columns=["Journal"]).to_excel(writer, sheet_name="Journal", index=False)
        
        # Stats [source 16]
        stats, totals = planner.get_stats()
        pd.DataFrame(stats, columns=["Nom", "Culture", "Boost %"]).to_excel(writer, sheet_name="Stats_Production", index=False)
        pd.DataFrame(list(totals.items()), columns=["Type", "Total Culture"]).to_excel(writer, sheet_name="Totaux_Speciaux", index=False)
        
        # Plan [source 16]
        ws = writer.book.create_sheet("Terrain_Place")
        for pb in planner.placed_buildings:
            color = COLORS.get(pb['info']['Type'], 'FFFFFF')
            for r in range(pb['r'], pb['r']+pb['h']):
                for c in range(pb['c'], pb['c']+pb['w']):
                    cell = ws.cell(row=r+1, column=c+1, value=pb['info']['Nom'])
                    cell.fill = PatternFill(start_color=color, fill_type='solid')

    st.download_button("Télécharger le Résultat", output.getvalue(), "Resultat_Ville.xlsx")
