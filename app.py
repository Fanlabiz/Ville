import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils import get_column_letter

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
        
        # Grille pour les placements (True = libre)
        self.grid = np.ones((self.rows, self.cols), dtype=bool)
        self.border_mask = np.zeros((self.rows, self.cols), dtype=bool)
        self.existing_buildings = []  # Bâtiments déjà présents
        
        # Analyser le terrain
        for r in range(self.rows):
            for c in range(self.cols):
                val = str(terrain_data[r][c]).strip().upper()
                if val == 'X' or val == 'NAN' or val == '':
                    if val == 'X':
                        self.grid[r, c] = False
                        self.border_mask[r, c] = True
                else:
                    # C'est un bâtiment existant
                    self.grid[r, c] = False
                    self.existing_buildings.append({
                        'Nom': terrain_data[r][c],
                        'r': r,
                        'c': c,
                        'w': 1,
                        'h': 1,
                        'info': {'Type': 'Neutre', 'Nom': terrain_data[r][c]}
                    })
        
        self.journal = []
        self.placed_buildings = []  # Bâtiments placés (nouveaux)
        self.moved_buildings = []  # Bâtiments déplacés
        self.max_entries = 10000
        self.interrupted = False

    def log(self, msg):
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)
        else:
            self.interrupted = True

    def can_place(self, r, c, w, h):
        """Vérifie si on peut placer un bâtiment"""
        if r + h > self.rows or c + w > self.cols:
            return False
        if r < 0 or c < 0:
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
                
                # Distance de Manhattan minimale entre les deux bâtiments
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

    def get_best_position_for_producer(self, building_info, cultural_buildings, placed_positions):
        """Trouve la meilleure position pour un bâtiment producteur"""
        w = int(building_info['Largeur'])
        h = int(building_info['Longueur'])
        
        best_pos = None
        best_culture = -1
        best_orientation = None
        
        # Tester les deux orientations
        for orientation in [(w, h), (h, w)]:
            ow, oh = orientation
            for r in range(self.rows - oh + 1):
                for c in range(self.cols - ow + 1):
                    if self.can_place(r, c, ow, oh):
                        pos_key = (r, c, ow, oh)
                        if pos_key in placed_positions:
                            continue
                        
                        culture = self.calculate_culture_for_position(r, c, ow, oh, cultural_buildings)
                        if culture > best_culture:
                            best_culture = culture
                            best_pos = (r, c, ow, oh)
        
        return best_pos, best_culture

    def get_best_position_for_cultural(self, building_info, placed_positions):
        """Trouve une position pour un bâtiment culturel"""
        w = int(building_info['Largeur'])
        h = int(building_info['Longueur'])
        
        # Tester les deux orientations
        for orientation in [(w, h), (h, w)]:
            ow, oh = orientation
            for r in range(self.rows - oh + 1):
                for c in range(self.cols - ow + 1):
                    if self.can_place(r, c, ow, oh):
                        pos_key = (r, c, ow, oh)
                        if pos_key not in placed_positions:
                            return (r, c, ow, oh)
        return None

    def solve(self, buildings_list):
        """Algorithme de placement optimisé"""
        # Séparer par type
        cultural = [b for b in buildings_list if b['Type'] == 'Culturel']
        producers = [b for b in buildings_list if b['Type'] == 'Producteur']
        
        # Trier les culturels par rayonnement décroissant
        cultural.sort(key=lambda x: x.get('Rayonnement', 0), reverse=True)
        
        # Trier les producteurs par priorité (1 = plus prioritaire)
        prod_order = {'Guérison': 1, 'Nourriture': 2, 'Or': 3}
        producers.sort(key=lambda x: prod_order.get(x.get('Production', ''), 99))
        
        placed_positions = set()
        self.placed_buildings = []
        
        # 1. Conserver les bâtiments existants
        for eb in self.existing_buildings:
            r, c = eb['r'], eb['c']
            w, h = eb.get('w', 1), eb.get('h', 1)
            placed_positions.add((r, c, w, h))
            self.placed_buildings.append(eb)
        
        # 2. Placer les bâtiments culturels
        for building in cultural:
            for _ in range(int(building['Nombre'])):  # Utiliser Nombre, pas Quantite
                pos = self.get_best_position_for_cultural(building, placed_positions)
                if pos:
                    r, c, w, h = pos
                    new_building = self.place_building(r, c, w, h, building)
                    placed_positions.add((r, c, w, h))
                    self.placed_buildings.append(new_building)
                    self.log(f"Culturel placé: {building['Nom']} à ({r},{c})")
                else:
                    self.log(f"Impossible de placer {building['Nom']} (pas de place)")
        
        # 3. Placer les bâtiments producteurs
        cultural_list = [pb for pb in self.placed_buildings if pb['info']['Type'] == 'Culturel']
        
        for building in producers:
            for _ in range(int(building['Nombre'])):  # Utiliser Nombre, pas Quantite
                pos, culture = self.get_best_position_for_producer(building, cultural_list, placed_positions)
                
                if pos:
                    r, c, w, h = pos
                    new_building = self.place_building(r, c, w, h, building)
                    new_building['culture_recue'] = culture
                    new_building['boost'] = self.calculate_boost(culture, building)
                    placed_positions.add((r, c, w, h))
                    self.placed_buildings.append(new_building)
                    self.log(f"Producteur placé: {building['Nom']} à ({r},{c}) - Culture: {culture}")
                else:
                    self.log(f"Impossible de placer {building['Nom']} (pas de place)")
        
        return True

    def calculate_all_cultures(self):
        """Calcule la culture pour tous les producteurs placés"""
        cultural_buildings = [pb for pb in self.placed_buildings if pb['info']['Type'] == 'Culturel']
        prod_stats = []
        prod_by_type = {"Guérison": 0, "Nourriture": 0, "Or": 0}
        production_by_type = {}
        
        for pb in self.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                culture = self.calculate_culture_for_position(
                    pb['r'], pb['c'], pb['w'], pb['h'], cultural_buildings
                )
                pb['culture_recue'] = culture
                pb['boost'] = self.calculate_boost(culture, pb['info'])
                
                prod_type = pb['info'].get('Production', '')
                prod_quantite = pb['info'].get('Quantite', 0)
                prod_hour = prod_quantite * (1 + pb['boost']/100)
                
                prod_stats.append({
                    'Nom': pb['info']['Nom'],
                    'Type': pb['info']['Type'],
                    'Production': prod_type,
                    'X': pb['c'],
                    'Y': pb['r'],
                    'Orientation': 'Horizontal' if pb['w'] > pb['h'] else 'Vertical',
                    'Culture recue': culture,
                    'Boost': f"{pb['boost']}%",
                    'Production/heure': prod_hour
                })
                
                if prod_type in prod_by_type:
                    prod_by_type[prod_type] += culture
                
                if prod_type not in production_by_type:
                    production_by_type[prod_type] = {'total': 0, 'base': 0}
                production_by_type[prod_type]['total'] += prod_hour
                production_by_type[prod_type]['base'] += prod_quantite
        
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
        
        return prod_stats, prod_by_type, production_by_type


# --- LOGIQUE D'EXPORT EXCEL ---
def generate_excel(planner, original_buildings_df):
    """Génère le fichier Excel de résultat"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Liste des bâtiments placés
        prod_stats, prod_by_type, production_by_type = planner.calculate_all_cultures()
        placed_df = pd.DataFrame(prod_stats)
        placed_df.to_excel(writer, sheet_name="Batiments_places", index=False)
        
        # 2. Statistiques de production
        prod_summary = []
        for prod_type, data in production_by_type.items():
            prod_summary.append({
                'Type production': prod_type,
                'Quantité totale produite/heure': data['total'],
                'Quantité de base': data['base'],
                'Gain': data['total'] - data['base'],
                'Augmentation %': ((data['total'] - data['base']) / data['base'] * 100) if data['base'] > 0 else 0
            })
        
        if prod_summary:
            prod_summary_df = pd.DataFrame(prod_summary)
            prod_summary_df.to_excel(writer, sheet_name="Stats_production", index=False)
        
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
                    cell_value = ""
                    cell_color = None
                    
                    for pb in planner.placed_buildings:
                        if pb['r'] <= r < pb['r'] + pb['h'] and pb['c'] <= c < pb['c'] + pb['w']:
                            cell_value = pb['info']['Nom']
                            cell_color = COLORS.get(pb['info']['Type'], '808080')
                            
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
            ws.column_dimensions[column_letter].width = 22
        
        # 5. Résumé
        summary_data = [
            ["Bâtiments placés", len(planner.placed_buildings)],
            ["Cases utilisées", sum(p['w'] * p['h'] for p in planner.placed_buildings)],
            ["Entrées journal", len(planner.journal)]
        ]
        summary_df = pd.DataFrame(summary_data, columns=["Métrique", "Valeur"])
        summary_df.to_excel(writer, sheet_name="Resume", index=False)
    
    return output.getvalue()


# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Optimiseur de Cité", page_icon="🏗️", layout="wide")

st.title("🏗️ Optimiseur de Placement de Bâtiments")
st.markdown("""
**Placement intelligent** des bâtiments pour maximiser les boosts de production.
- 🟠 **Orange** : Bâtiments Culturels
- 🟢 **Vert** : Bâtiments Producteurs  
- ⚪ **Gris** : Bâtiments Neutres / Existants
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
            
            # Préparer la liste des bâtiments (utiliser Nombre, pas Quantite)
            buildings_list = []
            for _, row in buildings_df.iterrows():
                nombre = int(row['Nombre'])
                for _ in range(nombre):
                    buildings_list.append(row.to_dict())
            
            st.info(f"📊 {len(buildings_list)} bâtiments à placer (dont {len([b for b in buildings_list if b['Type']=='Culturel'])} culturels, {len([b for b in buildings_list if b['Type']=='Producteur'])} producteurs)")
            
            # Lancer l'optimisation
            if st.button("🚀 Lancer l'optimisation", type="primary"):
                with st.spinner("Optimisation du placement en cours..."):
                    planner = CityPlanner(terrain_df.values)
                    
                    # Convertir les types numériques
                    for b in buildings_list:
                        b['Longueur'] = int(b['Longueur'])
                        b['Largeur'] = int(b['Largeur'])
                        b['Nombre'] = int(b['Nombre'])
                        if pd.isna(b.get('Rayonnement', 0)):
                            b['Rayonnement'] = 0
                        if pd.isna(b.get('Culture', 0)):
                            b['Culture'] = 0
                        if pd.isna(b.get('Quantite', 0)):
                            b['Quantite'] = 0
                    
                    planner.solve(buildings_list)
                    
                    # Calculer les résultats
                    prod_stats, prod_by_type, production_by_type = planner.calculate_all_cultures()
                    
                    # Afficher les résultats
                    st.success("✅ Optimisation terminée!")
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Bâtiments placés", len([p for p in planner.placed_buildings if p['info']['Type'] != 'Neutre']))
                    with col2:
                        culture_total = sum(ps['Culture recue'] for ps in prod_stats if ps['Type'] == 'Producteur')
                        st.metric("Culture totale reçue", f"{culture_total:.0f}")
                    with col3:
                        st.metric("Producteurs placés", len([p for p in prod_stats if p['Type'] == 'Producteur']))
                    
                    # Afficher les résultats de production
                    st.subheader("📊 Production par type")
                    for prod_type, data in production_by_type.items():
                        st.metric(f"{prod_type}", f"{data['total']:,.0f}/h", delta=f"+{data['total'] - data['base']:,.0f}")
                    
                    # Afficher les boosts obtenus
                    st.subheader("🏆 Bâtiments producteurs")
                    prod_df = pd.DataFrame([p for p in prod_stats if p['Type'] == 'Producteur'])
                    if not prod_df.empty:
                        st.dataframe(prod_df[['Nom', 'Production', 'Culture recue', 'Boost', 'Production/heure']])
                    
                    # Bouton de téléchargement
                    excel_data = generate_excel(planner, buildings_df)
                    st.download_button(
                        label="📥 Télécharger le résultat (Excel)",
                        data=excel_data,
                        file_name="Resultat_Placement.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")
            st.exception(e)

else:
    st.info("👆 Veuillez charger un fichier Excel pour commencer")
