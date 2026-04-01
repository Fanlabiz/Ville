import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter
from copy import deepcopy

# --- CONFIGURATION ---
COLORS = {
    'Culturel': 'FFA500',     # Orange
    'Producteur': '008000',   # Vert
    'Neutre': '808080'        # Gris
}

class CityPlanner:
    def __init__(self, terrain_data):
        # Convertir en DataFrame pour manipuler les noms
        self.terrain_df = pd.DataFrame(terrain_data)
        self.rows = len(terrain_data)
        self.cols = len(terrain_data[0])
        
        # Grille pour les placements
        self.grid = np.ones((self.rows, self.cols), dtype=bool)  # True = libre
        self.border_mask = np.zeros((self.rows, self.cols), dtype=bool)
        self.initial_buildings = {}  # Stocke les bâtiments initiaux
        
        # Analyser le terrain
        for r in range(self.rows):
            for c in range(self.cols):
                val = str(terrain_data[r][c]).strip().upper()
                if val == 'X':
                    self.grid[r, c] = False
                    self.border_mask[r, c] = True
                elif val and val != 'NAN' and val != '' and val != 'N/A':
                    # C'est un bâtiment existant
                    self.grid[r, c] = False
                    self.initial_buildings[(r, c)] = val
        
        self.journal = []
        self.placed_buildings = []  # Bâtiments placés (nouveaux ou déplacés)
        self.existing_buildings = []  # Bâtiments existants conservés
        self.moved_buildings = []  # Bâtiments déplacés
        self.max_entries = 10000
        self.interrupted = False
        
        # Initialiser les bâtiments existants
        for (r, c), name in self.initial_buildings.items():
            # Chercher les dimensions du bâtiment (nécessite de connaître sa forme)
            # Pour simplifier, on suppose que chaque bâtiment existant a été placé manuellement
            self.existing_buildings.append({
                'Nom': name,
                'r': r,
                'c': c,
                'w': 1,  # À déterminer
                'h': 1,
                'info': {'Type': 'Neutre'}  # Type par défaut
            })

    def log(self, msg):
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)
        else:
            self.interrupted = True

    def find_building_dimensions(self, name, buildings_df):
        """Trouve les dimensions d'un bâtiment à partir de son nom"""
        for _, row in buildings_df.iterrows():
            if row['Nom'] == name:
                return int(row['Longueur']), int(row['Largeur'])
        return 1, 1

    def can_place(self, r, c, w, h):
        """Vérifie si on peut placer un bâtiment"""
        if r + h > self.rows or c + w > self.cols:
            return False
        return np.all(self.grid[r:r+h, c:c+w])

    def place_building(self, r, c, w, h, info):
        """Place un bâtiment sur la grille"""
        self.grid[r:r+h, c:c+w] = False
        return {
            'info': info,
            'r': r,
            'c': c,
            'w': w,
            'h': h
        }

    def calculate_culture_for_position(self, r, c, w, h, cultural_buildings):
        """Calcule la culture reçue par un bâtiment producteur à une position donnée"""
        total_culture = 0
        
        for cb in cultural_buildings:
            rayon = int(cb['info'].get('Rayonnement', 0))
            culture_value = cb['info'].get('Culture', 0)
            
            if rayon > 0 and culture_value > 0:
                # Vérifier si le producteur est dans la zone du culturel
                prod_r_start, prod_r_end = r, r + h
                prod_c_start, prod_c_end = c, c + w
                cult_r_start, cult_r_end = cb['r'], cb['r'] + cb['h']
                cult_c_start, cult_c_end = cb['c'], cb['c'] + cb['w']
                
                # Distance de Manhattan minimale
                min_dist = float('inf')
                for pr in range(prod_r_start, prod_r_end):
                    for pc in range(prod_c_start, prod_c_end):
                        for cr in range(cult_r_start, cult_r_end):
                            for cc in range(cult_c_start, cult_c_end):
                                dist = abs(pr - cr) + abs(pc - cc)
                                min_dist = min(min_dist, dist)
                
                if min_dist <= rayon:
                    total_culture += culture_value
        
        return total_culture

    def calculate_boost(self, culture, info):
        """Calcule le boost en fonction de la culture reçue"""
        boost_100 = info.get('Boost 100%')
        boost_50 = info.get('Boost 50%')
        boost_25 = info.get('Boost 25%')
        
        if pd.notna(boost_100) and culture >= boost_100:
            return 100
        elif pd.notna(boost_50) and culture >= boost_50:
            return 50
        elif pd.notna(boost_25) and culture >= boost_25:
            return 25
        return 0

    def get_best_position(self, building_info, cultural_buildings, placed_positions):
        """Trouve la meilleure position pour un bâtiment producteur"""
        w, h = building_info['Largeur'], building_info['Longueur']
        
        best_pos = None
        best_culture = -1
        best_orientation = None
        
        # Tester les deux orientations
        for orientation in [(w, h), (h, w)]:
            ow, oh = orientation
            for r in range(self.rows - oh + 1):
                for c in range(self.cols - ow + 1):
                    if self.can_place(r, c, ow, oh):
                        # Vérifier que la position n'est pas déjà utilisée
                        pos_key = (r, c, ow, oh)
                        if pos_key in placed_positions:
                            continue
                        
                        culture = self.calculate_culture_for_position(r, c, ow, oh, cultural_buildings)
                        if culture > best_culture:
                            best_culture = culture
                            best_pos = (r, c, ow, oh)
                            best_orientation = orientation
        
        return best_pos, best_culture

    def get_best_cultural_position(self, building_info, existing_positions):
        """Trouve la meilleure position pour un bâtiment culturel"""
        w, h = building_info['Largeur'], building_info['Longueur']
        
        # Priorité : proche des zones où il y a de la place pour les producteurs
        best_pos = None
        best_score = -1
        
        for orientation in [(w, h), (h, w)]:
            ow, oh = orientation
            for r in range(self.rows - oh + 1):
                for c in range(self.cols - ow + 1):
                    if self.can_place(r, c, ow, oh):
                        pos_key = (r, c, ow, oh)
                        if pos_key in existing_positions:
                            continue
                        
                        # Calculer un score basé sur l'espace disponible autour
                        score = self.calculate_available_space_score(r, c, ow, oh)
                        if score > best_score:
                            best_score = score
                            best_pos = (r, c, ow, oh)
        
        return best_pos

    def calculate_available_space_score(self, r, c, w, h):
        """Calcule un score pour un emplacement basé sur l'espace disponible autour"""
        score = 0
        # Zone d'influence (rayon 5)
        for dr in range(-5, 6):
            for dc in range(-5, 6):
                nr, nc = r + dr, c + dc
                if 0 <= nr < self.rows and 0 <= nc < self.cols:
                    if self.grid[nr, nc]:
                        score += 1
        return score

    def solve_with_priority(self, buildings_list):
        """Algorithme de placement optimisé"""
        # Séparer par type
        cultural = [b for b in buildings_list if b['Type'] == 'Culturel']
        producers = [b for b in buildings_list if b['Type'] == 'Producteur']
        
        # Trier par rayonnement décroissant (les plus grands rayonnements d'abord)
        cultural.sort(key=lambda x: x.get('Rayonnement', 0), reverse=True)
        
        # Trier les producteurs par priorité (1 = plus prioritaire)
        producers.sort(key=lambda x: x.get('Priorite', 0))
        
        # Trier par production (Guérison > Nourriture > Or > autres)
        prod_order = {'Guérison': 1, 'Nourriture': 2, 'Or': 3}
        producers.sort(key=lambda x: prod_order.get(x.get('Production', ''), 99))
        
        placed_positions = set()
        placed_buildings = []
        
        # 1. Placer d'abord les bâtiments existants comme référence
        for eb in self.existing_buildings:
            r, c = eb['r'], eb['c']
            w, h = eb.get('w', 1), eb.get('h', 1)
            placed_positions.add((r, c, w, h))
            placed_buildings.append(eb)
        
        # 2. Placer les bâtiments culturels
        for building in cultural:
            for _ in range(int(building.get('Quantite', 1))):
                pos = self.get_best_cultural_position(building, placed_positions)
                if pos:
                    r, c, w, h = pos
                    new_building = self.place_building(r, c, w, h, building)
                    placed_positions.add((r, c, w, h))
                    placed_buildings.append(new_building)
                    self.log(f"Culturel placé: {building['Nom']} à ({r},{c})")
        
        # 3. Placer les bâtiments producteurs
        for building in producers:
            for _ in range(int(building.get('Quantite', 1))):
                cultural_list = [pb for pb in placed_buildings if pb['info']['Type'] == 'Culturel']
                pos, culture = self.get_best_position(building, cultural_list, placed_positions)
                
                if pos:
                    r, c, w, h = pos
                    new_building = self.place_building(r, c, w, h, building)
                    new_building['culture_recue'] = culture
                    new_building['boost'] = self.calculate_boost(culture, building)
                    placed_positions.add((r, c, w, h))
                    placed_buildings.append(new_building)
                    self.log(f"Producteur placé: {building['Nom']} à ({r},{c}) - Culture: {culture}")
        
        self.placed_buildings = placed_buildings
        return True

    def calculate_all_cultures(self):
        """Calcule la culture pour tous les producteurs placés"""
        cultural_buildings = [pb for pb in self.placed_buildings if pb['info']['Type'] == 'Culturel']
        prod_stats = []
        prod_by_type = {"Guérison": 0, "Nourriture": 0, "Or": 0}
        
        for pb in self.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                culture = self.calculate_culture_for_position(
                    pb['r'], pb['c'], pb['w'], pb['h'], cultural_buildings
                )
                pb['culture_recue'] = culture
                pb['boost'] = self.calculate_boost(culture, pb['info'])
                
                prod_stats.append({
                    'Nom': pb['info']['Nom'],
                    'Type': pb['info']['Type'],
                    'Production': pb['info'].get('Production', ''),
                    'X': pb['c'],
                    'Y': pb['r'],
                    'Orientation': 'Horizontal' if pb['w'] > pb['h'] else 'Vertical',
                    'Culture recue': culture,
                    'Boost': f"{pb['boost']}%",
                    'Production/heure': pb['info'].get('Quantite', 0) * (1 + pb['boost']/100)
                })
                
                prod_type = pb['info'].get('Production', '')
                if prod_type in prod_by_type:
                    prod_by_type[prod_type] += culture
        
        # Ajouter les culturels
        for pb in self.placed_buildings:
            if pb['info']['Type'] == 'Culturel':
                prod_stats.append({
                    'Nom': pb['info']['Nom'],
                    'Type': pb['info']['Type'],
                    'Production': 'N/A',
                    'X': pb['c'],
                    'Y': pb['r'],
                    'Orientation': 'Horizontal' if pb['w'] > pb['h'] else 'Vertical',
                    'Culture recue': 0,
                    'Boost': '0%',
                    'Production/heure': 0
                })
        
        return prod_stats, prod_by_type


# --- LOGIQUE D'EXPORT EXCEL ---
def generate_excel(planner, original_buildings_df):
    """Génère le fichier Excel de résultat"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Liste des bâtiments placés
        prod_stats, prod_by_type = planner.calculate_all_cultures()
        placed_df = pd.DataFrame(prod_stats)
        placed_df.to_excel(writer, sheet_name="Batiments_places", index=False)
        
        # 2. Statistiques de production
        production_stats = []
        for pb in planner.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                prod_type = pb['info'].get('Production', '')
                production_stats.append({
                    'Type production': prod_type,
                    'Quantité produite/heure': pb['info'].get('Quantite', 0) * (1 + pb['boost']/100),
                    'Quantité de base': pb['info'].get('Quantite', 0),
                    'Gain': pb['info'].get('Quantite', 0) * (pb['boost']/100),
                    'Boost': pb['boost']
                })
        
        if production_stats:
            prod_summary = pd.DataFrame(production_stats)
            prod_summary = prod_summary.groupby('Type production').agg({
                'Quantité produite/heure': 'sum',
                'Quantité de base': 'sum',
                'Gain': 'sum'
            }).reset_index()
            prod_summary.to_excel(writer, sheet_name="Stats_production", index=False)
        
        # 3. Journal
        pd.DataFrame(planner.journal, columns=["Journal"]).to_excel(writer, sheet_name="Journal", index=False)
        
        # 4. Terrain final
        ws = writer.book.create_sheet("Terrain_final")
        
        # Remplir le terrain
        for r in range(planner.rows):
            for c in range(planner.cols):
                cell = ws.cell(row=r+1, column=c+1)
                
                if planner.border_mask[r, c]:
                    cell.value = "X"
                    cell.fill = PatternFill(start_color='000000', fill_type='solid')
                    cell.font = Font(color='FFFFFF')
                else:
                    # Vérifier si c'est un bâtiment placé
                    cell_value = ""
                    cell_color = None
                    
                    for pb in planner.placed_buildings:
                        if pb['r'] <= r < pb['r'] + pb['h'] and pb['c'] <= c < pb['c'] + pb['w']:
                            cell_value = pb['info']['Nom']
                            cell_color = COLORS.get(pb['info']['Type'], '808080')
                            
                            # Ajouter le boost pour les producteurs
                            if pb['info']['Type'] == 'Producteur' and pb.get('boost', 0) > 0:
                                cell_value += f"\nBoost: {pb['boost']}%"
                            elif pb['info']['Type'] == 'Culturel' and pb['info'].get('Culture', 0) > 0:
                                cell_value += f"\nCulture: {pb['info']['Culture']}"
                            break
                    
                    if cell_value:
                        cell.value = cell_value
                        if cell_color:
                            cell.fill = PatternFill(start_color=cell_color, fill_type='solid')
                        cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        
        # Ajuster les colonnes
        for col in range(1, planner.cols + 1):
            column_letter = get_column_letter(col)
            ws.column_dimensions[column_letter].width = 20
        
        # 5. Résumé
        summary_data = [
            ["Bâtiments placés", len(planner.placed_buildings)],
            ["Cases utilisées", sum(p['w'] * p['h'] for p in planner.placed_buildings)],
            ["Bâtiments déplacés", len(planner.moved_buildings)],
            ["Entrées journal", len(planner.journal)]
        ]
        summary_df = pd.DataFrame(summary_data, columns=["Métrique", "Valeur"])
        summary_df.to_excel(writer, sheet_name="Resume", index=False)
    
    return output.getvalue()


# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Optimiseur de Cité - V2", page_icon="🏗️", layout="wide")

st.title("🏗️ Optimiseur de Placement de Bâtiments V2")
st.markdown("""
**Version améliorée** avec regroupement intelligent des bâtiments culturels et producteurs.
- 🟠 **Orange** : Bâtiments Culturels (placer en priorité)
- 🟢 **Vert** : Bâtiments Producteurs (optimisés pour recevoir le maximum de culture)
- ⚪ **Gris** : Bâtiments Neutres
- ⬛ **Noir** : Bords du terrain
""")

uploaded = st.file_uploader("📂 Charger le fichier Excel", type="xlsx")

if uploaded:
    with st.spinner("Analyse du fichier..."):
        try:
            # Lecture des données
            terrain_df = pd.read_excel(uploaded, sheet_name=0, header=None)
            buildings_df = pd.read_excel(uploaded, sheet_name=1)
            buildings_df.columns = buildings_df.columns.str.strip()
            
            st.success("✅ Fichier chargé avec succès!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Aperçu du terrain")
                st.dataframe(terrain_df.head(10))
            
            with col2:
                st.subheader("Aperçu des bâtiments")
                st.dataframe(buildings_df.head(10))
            
            # Préparer la liste des bâtiments
            buildings_list = []
            for _, row in buildings_df.iterrows():
                for _ in range(int(row['Quantite'])):
                    buildings_list.append(row.to_dict())
            
            st.info(f"📊 {len(buildings_list)} bâtiments à placer")
            
            # Lancer l'optimisation
            if st.button("🚀 Lancer l'optimisation", type="primary"):
                with st.spinner("Optimisation du placement en cours..."):
                    planner = CityPlanner(terrain_df.values)
                    
                    # Ajouter les dimensions des bâtiments
                    for b in buildings_list:
                        b['Longueur'] = int(b['Longueur'])
                        b['Largeur'] = int(b['Largeur'])
                        if pd.isna(b.get('Rayonnement', 0)):
                            b['Rayonnement'] = 0
                        if pd.isna(b.get('Culture', 0)):
                            b['Culture'] = 0
                        if pd.isna(b.get('Quantite', 0)):
                            b['Quantite'] = 0
                    
                    planner.solve_with_priority(buildings_list)
                    
                    # Calculer les résultats
                    prod_stats, prod_by_type = planner.calculate_all_cultures()
                    
                    # Afficher les résultats
                    st.success("✅ Optimisation terminée!")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Bâtiments placés", len(planner.placed_buildings))
                    with col2:
                        culture_total = sum(ps['Culture recue'] for ps in prod_stats if ps['Type'] == 'Producteur')
                        st.metric("Culture totale reçue", f"{culture_total:.0f}")
                    with col3:
                        st.metric("Producteurs placés", len([p for p in prod_stats if p['Type'] == 'Producteur']))
                    
                    # Afficher les boosts par type
                    st.subheader("📊 Production par type")
                    prod_df = pd.DataFrame([p for p in prod_stats if p['Type'] == 'Producteur'])
                    if not prod_df.empty:
                        st.dataframe(prod_df[['Nom', 'Production', 'Culture recue', 'Boost', 'Production/heure']])
                    
                    # Bouton de téléchargement
                    excel_data = generate_excel(planner, buildings_df)
                    st.download_button(
                        label="📥 Télécharger le résultat (Excel)",
                        data=excel_data,
                        file_name="Resultat_Placement_V2.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")
            st.exception(e)

else:
    st.info("👆 Veuillez charger un fichier Excel pour commencer")
