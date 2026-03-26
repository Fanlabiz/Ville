import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st
import io
from typing import List, Tuple, Dict, Optional
import copy

class Building:
    def __init__(self, nom, longueur, largeur, nombre, type_bat, culture, rayonnement,
                 boost_25, boost_50, boost_100, production, quantite):
        self.nom = nom
        self.longueur = int(longueur) if pd.notna(longueur) else 0
        self.largeur = int(largeur) if pd.notna(largeur) else 0
        self.nombre = int(nombre) if pd.notna(nombre) else 0
        self.type = type_bat
        self.culture = float(culture) if pd.notna(culture) else 0
        self.rayonnement = int(rayonnement) if pd.notna(rayonnement) else 0
        self.boost_25 = float(boost_25) if pd.notna(boost_25) else 0
        self.boost_50 = float(boost_50) if pd.notna(boost_50) else 0
        self.boost_100 = float(boost_100) if pd.notna(boost_100) else 0
        self.production = production if pd.notna(production) else "Rien"
        self.quantite = float(quantite) if pd.notna(quantite) else 0
        
    def get_area(self):
        return self.longueur * self.largeur
    
    def get_boost_percentage(self, culture_recue):
        if culture_recue >= self.boost_100:
            return 100
        elif culture_recue >= self.boost_50:
            return 50
        elif culture_recue >= self.boost_25:
            return 25
        return 0
    
    def get_production_per_hour(self, culture_recue):
        boost = self.get_boost_percentage(culture_recue)
        return self.quantite * (1 + boost / 100)

class PlacedBuilding:
    def __init__(self, building, row, col, orientation):
        self.building = building
        self.row = row
        self.col = col
        self.orientation = orientation
        self.culture_recue = 0
        
    def get_cells(self):
        cells = []
        if self.orientation == "horizontal":
            for i in range(self.building.longueur):
                for j in range(self.building.largeur):
                    cells.append((self.row + i, self.col + j))
        else:
            for i in range(self.building.largeur):
                for j in range(self.building.longueur):
                    cells.append((self.row + i, self.col + j))
        return cells
    
    def update_culture(self, cultural_buildings):
        total_culture = 0
        for cultural in cultural_buildings:
            if cultural != self and cultural.building.type == "Culturel":
                cells = self.get_cells()
                cultural_cells = cultural.get_cells()
                cultural_radius = cultural.building.rayonnement
                
                for cell in cells:
                    for cult_cell in cultural_cells:
                        distance = max(abs(cell[0] - cult_cell[0]), abs(cell[1] - cult_cell[1]))
                        if distance <= cultural_radius:
                            total_culture += cultural.building.culture
                            break
        self.culture_recue = total_culture

class Terrain:
    def __init__(self, grid):
        self.grid = np.array(grid)
        self.height, self.width = self.grid.shape
        self.buildings = []
        
    def is_cell_free(self, row, col):
        if row < 0 or row >= self.height or col < 0 or col >= self.width:
            return False
        return self.grid[row, col] == 1
    
    def can_place_building(self, building, row, col, orientation):
        if orientation == "horizontal":
            for i in range(building.longueur):
                for j in range(building.largeur):
                    r, c = row + i, col + j
                    if r >= self.height or c >= self.width:
                        return False
                    if not self.is_cell_free(r, c):
                        return False
        else:
            for i in range(building.largeur):
                for j in range(building.longueur):
                    r, c = row + i, col + j
                    if r >= self.height or c >= self.width:
                        return False
                    if not self.is_cell_free(r, c):
                        return False
        return True
    
    def place_building(self, building, row, col, orientation):
        placed = PlacedBuilding(building, row, col, orientation)
        cells = placed.get_cells()
        for r, c in cells:
            self.grid[r, c] = 2
        self.buildings.append(placed)
        return placed
    
    def remove_building(self, building):
        """Retire un bâtiment du terrain"""
        if building in self.buildings:
            cells = building.get_cells()
            for r, c in cells:
                self.grid[r, c] = 1
            self.buildings.remove(building)
            return True
        return False
    
    def get_free_cells_count(self):
        return np.sum(self.grid == 1)
    
    def get_cultural_buildings(self):
        return [b for b in self.buildings if b.building.type == "Culturel"]
    
    def get_producer_buildings(self):
        return [b for b in self.buildings if b.building.type == "Producteur"]

class BuildingPlacer:
    def __init__(self, terrain, buildings):
        self.terrain = terrain
        self.buildings = buildings
        self.neutral_buildings = [b for b in buildings if b.type == "Neutre"]
        self.cultural_buildings = [b for b in buildings if b.type == "Culturel"]
        self.producer_buildings = [b for b in buildings if b.type == "Producteur"]
        self.unplaced = []
        
    def get_border_cells(self):
        border_cells = []
        for col in range(self.terrain.width):
            if self.terrain.is_cell_free(0, col):
                border_cells.append((0, col))
            if self.terrain.is_cell_free(self.terrain.height - 1, col):
                border_cells.append((self.terrain.height - 1, col))
        for row in range(1, self.terrain.height - 1):
            if self.terrain.is_cell_free(row, 0):
                border_cells.append((row, 0))
            if self.terrain.is_cell_free(row, self.terrain.width - 1):
                border_cells.append((row, self.terrain.width - 1))
        return border_cells
    
    def find_position(self, building, prefer_border=False, avoid_border=False, near_cultural=None):
        best_position = None
        best_score = -1
        
        border_cells = set(self.get_border_cells())
        
        for orientation in ["horizontal", "vertical"]:
            for row in range(self.terrain.height):
                for col in range(self.terrain.width):
                    if self.terrain.can_place_building(building, row, col, orientation):
                        score = 0
                        
                        if avoid_border:
                            cells = []
                            if orientation == "horizontal":
                                for i in range(building.longueur):
                                    for j in range(building.largeur):
                                        cells.append((row + i, col + j))
                            else:
                                for i in range(building.largeur):
                                    for j in range(building.longueur):
                                        cells.append((row + i, col + j))
                            
                            touches_border = False
                            for r, c in cells:
                                if r == 0 or r == self.terrain.height - 1 or c == 0 or c == self.terrain.width - 1:
                                    touches_border = True
                                    break
                            if touches_border:
                                continue
                        
                        if prefer_border:
                            cells = []
                            if orientation == "horizontal":
                                for i in range(building.longueur):
                                    for j in range(building.largeur):
                                        cells.append((row + i, col + j))
                            else:
                                for i in range(building.largeur):
                                    for j in range(building.longueur):
                                        cells.append((row + i, col + j))
                            
                            border_count = sum(1 for r, c in cells if (r, c) in border_cells)
                            score += border_count * 10
                        
                        # Si on cherche à placer près d'un bâtiment culturel
                        if near_cultural is not None:
                            cells = []
                            if orientation == "horizontal":
                                for i in range(building.longueur):
                                    for j in range(building.largeur):
                                        cells.append((row + i, col + j))
                            else:
                                for i in range(building.largeur):
                                    for j in range(building.longueur):
                                        cells.append((row + i, col + j))
                            
                            cultural_cells = near_cultural.get_cells()
                            cultural_radius = near_cultural.building.rayonnement
                            
                            for cell in cells:
                                for cult_cell in cultural_cells:
                                    distance = max(abs(cell[0] - cult_cell[0]), abs(cell[1] - cult_cell[1]))
                                    if distance <= cultural_radius:
                                        score += 100 - distance * 10  # Plus on est proche, plus le score est élevé
                                        break
                        
                        if score > best_score:
                            best_score = score
                            best_position = (row, col, orientation)
        
        return best_position
    
    def find_any_position(self, building):
        """Trouve n'importe quelle position disponible pour un bâtiment"""
        for orientation in ["horizontal", "vertical"]:
            for row in range(self.terrain.height):
                for col in range(self.terrain.width):
                    if self.terrain.can_place_building(building, row, col, orientation):
                        return (row, col, orientation)
        return None
    
    def place_all_buildings(self):
        # 1. Placer les bâtiments neutres sur les bords
        self.neutral_buildings.sort(key=lambda x: x.get_area(), reverse=True)
        for building in self.neutral_buildings:
            for _ in range(building.nombre):
                pos = self.find_position(building, prefer_border=True)
                if pos:
                    row, col, orientation = pos
                    self.terrain.place_building(building, row, col, orientation)
                else:
                    self.unplaced.append(building)
        
        # 2. Placer les bâtiments de production par priorité
        production_priority = {"Guerison": 0, "Nourriture": 1, "Or": 2, "Autre": 3}
        
        all_to_place = []
        
        for b in self.producer_buildings:
            for _ in range(b.nombre):
                priority = production_priority.get(b.production, 3)
                all_to_place.append((b, "producteur", priority))
        
        for b in self.cultural_buildings:
            for _ in range(b.nombre):
                all_to_place.append((b, "culturel", 0))
        
        all_to_place.sort(key=lambda x: (x[1] == "culturel", -x[0].get_area()))
        
        for building, type_bat, _ in all_to_place:
            if type_bat == "culturel":
                pos = self.find_position(building, avoid_border=True)
                if not pos:
                    pos = self.find_position(building)
            else:
                pos = self.find_position(building)
            
            if pos:
                row, col, orientation = pos
                self.terrain.place_building(building, row, col, orientation)
            else:
                self.unplaced.append(building)
    
    def optimize_boosts(self):
        """Phase d'optimisation : entrelacer culturels et producteurs pour maximiser les boosts"""
        cultural_buildings = self.terrain.get_cultural_buildings()
        producer_buildings = self.terrain.get_producer_buildings()
        
        if not cultural_buildings or not producer_buildings:
            return
        
        # Trier les producteurs par priorité de production (Guérison d'abord)
        production_priority = {"Guerison": 0, "Nourriture": 1, "Or": 2, "Autre": 3}
        producer_buildings.sort(key=lambda x: production_priority.get(x.building.production, 3))
        
        # Trier les culturels par rayonnement décroissant (ceux qui donnent le plus de culture)
        cultural_buildings.sort(key=lambda x: x.building.culture * x.building.rayonnement, reverse=True)
        
        improvements_made = True
        max_iterations = 50
        iteration = 0
        
        while improvements_made and iteration < max_iterations:
            improvements_made = False
            iteration += 1
            
            # Pour chaque producteur, vérifier s'il peut être mieux positionné près des culturels
            for producer in producer_buildings[:]:  # Copie pour éviter les modifications pendant l'itération
                if producer not in self.terrain.buildings:
                    continue
                    
                current_culture = producer.culture_recue
                
                # Chercher une meilleure position près des culturels
                best_culture = current_culture
                best_position = None
                best_orientation = None
                
                for cultural in cultural_buildings:
                    # Essayer de placer le producteur près de ce culturel
                    pos = self.find_position(producer.building, near_cultural=cultural)
                    if pos:
                        row, col, orientation = pos
                        # Tester cette position
                        temp_producer = PlacedBuilding(producer.building, row, col, orientation)
                        temp_culture = 0
                        for cult in cultural_buildings:
                            if cult != cultural:
                                temp_cells = temp_producer.get_cells()
                                cult_cells = cult.get_cells()
                                cult_radius = cult.building.rayonnement
                                for cell in temp_cells:
                                    for cult_cell in cult_cells:
                                        distance = max(abs(cell[0] - cult_cell[0]), abs(cell[1] - cult_cell[1]))
                                        if distance <= cult_radius:
                                            temp_culture += cult.building.culture
                                            break
                        
                        # Ajouter la culture du culturel cible
                        temp_cells = temp_producer.get_cells()
                        cult_cells = cultural.get_cells()
                        cult_radius = cultural.building.rayonnement
                        for cell in temp_cells:
                            for cult_cell in cult_cells:
                                distance = max(abs(cell[0] - cult_cell[0]), abs(cell[1] - cult_cell[1]))
                                if distance <= cult_radius:
                                    temp_culture += cultural.building.culture
                                    break
                        
                        if temp_culture > best_culture:
                            best_culture = temp_culture
                            best_position = (row, col, orientation)
                
                # Si on a trouvé une meilleure position
                if best_position and best_culture > current_culture:
                    row, col, orientation = best_position
                    # Vérifier que l'emplacement est toujours libre
                    if self.terrain.can_place_building(producer.building, row, col, orientation):
                        # Retirer l'ancien emplacement
                        self.terrain.remove_building(producer)
                        # Placer à la nouvelle position
                        self.terrain.place_building(producer.building, row, col, orientation)
                        improvements_made = True
        
        # Recalculer les cultures après optimisation
        cultural_buildings = self.terrain.get_cultural_buildings()
        producer_buildings = self.terrain.get_producer_buildings()
        for producer in producer_buildings:
            producer.update_culture(cultural_buildings)
    
    def try_place_remaining(self):
        """Tente de placer les bâtiments restants dans l'espace libre"""
        if not self.unplaced:
            return
        
        self.unplaced.sort(key=lambda x: -x.get_area())
        
        remaining = []
        for building in self.unplaced:
            pos = self.find_any_position(building)
            if pos:
                row, col, orientation = pos
                self.terrain.place_building(building, row, col, orientation)
            else:
                remaining.append(building)
        
        self.unplaced = remaining
    
    def calculate_culture_and_production(self):
        producers = [b for b in self.terrain.buildings if b.building.type == "Producteur"]
        cultural = [b for b in self.terrain.buildings if b.building.type == "Culturel"]
        
        for producer in producers:
            producer.update_culture(cultural)
        
        production_stats = {}
        for producer in producers:
            prod_type = producer.building.production
            if prod_type not in production_stats:
                production_stats[prod_type] = {
                    "total_culture": 0,
                    "total_production": 0,
                    "count": 0,
                    "boost_total": 0
                }
            
            boost = producer.building.get_boost_percentage(producer.culture_recue)
            prod_per_hour = producer.building.get_production_per_hour(producer.culture_recue)
            
            production_stats[prod_type]["total_culture"] += producer.culture_recue
            production_stats[prod_type]["total_production"] += prod_per_hour
            production_stats[prod_type]["count"] += 1
            production_stats[prod_type]["boost_total"] += boost
        
        for prod_type in production_stats:
            if production_stats[prod_type]["count"] > 0:
                production_stats[prod_type]["avg_boost"] = production_stats[prod_type]["boost_total"] / production_stats[prod_type]["count"]
            else:
                production_stats[prod_type]["avg_boost"] = 0
        
        return production_stats

def read_input_excel(file):
    """Lecture du fichier Excel d'entrée"""
    xls = pd.ExcelFile(file)
    
    terrain_df = pd.read_excel(xls, sheet_name="Terrain", header=None)
    terrain_grid = terrain_df.fillna(0).values
    
    terrain_grid = np.where(terrain_grid == 'X', 0, terrain_grid)
    terrain_grid = np.where(terrain_grid == 1, 1, terrain_grid)
    terrain_grid = np.where(terrain_grid == 0, 0, terrain_grid)
    terrain_grid = terrain_grid.astype(int)
    
    buildings_df = pd.read_excel(xls, sheet_name="Batiments")
    buildings = []
    for _, row in buildings_df.iterrows():
        b = Building(
            row['Nom'], row['Longueur'], row['Largeur'], row['Nombre'],
            row['Type'], row['Culture'], row['Rayonnement'],
            row['Boost 25%'], row['Boost 50%'], row['Boost 100%'],
            row['Production'], row['Quantite']
        )
        buildings.append(b)
    
    return terrain_grid, buildings

def create_terrain_visualization(terrain):
    """Créer une représentation visuelle du terrain avec les noms des bâtiments"""
    visual = [['' for _ in range(terrain.width)] for _ in range(terrain.height)]
    
    for i in range(terrain.height):
        for j in range(terrain.width):
            if terrain.grid[i, j] == 1:
                visual[i][j] = ''
            elif terrain.grid[i, j] == 0:
                visual[i][j] = 'X'
    
    for pb in terrain.buildings:
        cells = pb.get_cells()
        boost = pb.building.get_boost_percentage(pb.culture_recue)
        boost_text = f" +{boost}%" if boost > 0 else ""
        
        for idx, (r, c) in enumerate(cells):
            if idx == 0:
                visual[r][c] = f"{pb.building.nom}{boost_text}"
            else:
                visual[r][c] = ''
    
    return visual

def create_output_excel(terrain, production_stats, unplaced_buildings):
    """Créer le fichier Excel de sortie avec mise en forme colorée"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Feuille des bâtiments placés
        placed_data = []
        for pb in terrain.buildings:
            boost = pb.building.get_boost_percentage(pb.culture_recue)
            placed_data.append({
                'Nom': pb.building.nom,
                'Type': pb.building.type,
                'Production': pb.building.production,
                'Ligne': pb.row + 1,
                'Colonne': pb.col + 1,
                'Hauteur': pb.building.longueur if pb.orientation == "horizontal" else pb.building.largeur,
                'Largeur': pb.building.largeur if pb.orientation == "horizontal" else pb.building.longueur,
                'Culture recue': pb.culture_recue,
                'Boost (%)': boost,
                'Quantite/h': pb.building.quantite,
                'Prod totale/h': pb.building.get_production_per_hour(pb.culture_recue),
                'Origine': 'Placé'
            })
        
        placed_df = pd.DataFrame(placed_data)
        placed_df.to_excel(writer, sheet_name='Batiments places', index=False)
        
        # 2. Feuille de synthèse
        synthesis_data = []
        for prod_type, stats in production_stats.items():
            if prod_type != "Rien":
                synthesis_data.append({
                    'Production': prod_type,
                    'Culture totale': stats['total_culture'],
                    'Boost moyen (%)': round(stats['avg_boost'], 1),
                    'Nb batiments': stats['count'],
                    'Production/h': round(stats['total_production'], 2)
                })
        
        if synthesis_data:
            synthesis_df = pd.DataFrame(synthesis_data)
            synthesis_df.to_excel(writer, sheet_name='Synthese', index=False)
        
        total_unplaced_area = sum(b.get_area() for b in unplaced_buildings)
        extra_data = pd.DataFrame({
            'Statistiques': ['Cases libres restantes', 'Cases des batiments non places', 'Nombre de batiments non places'],
            'Valeur': [terrain.get_free_cells_count(), total_unplaced_area, len(unplaced_buildings)]
        })
        extra_data.to_excel(writer, sheet_name='Synthese', startrow=len(synthesis_data) + 3, index=False)
        
        # 3. Feuille du terrain avec bâtiments et mise en forme
        terrain_visual = create_terrain_visualization(terrain)
        terrain_visual_df = pd.DataFrame(terrain_visual)
        terrain_visual_df.to_excel(writer, sheet_name='Terrain', index=False, header=False)
        
        # Appliquer la mise en forme colorée
        workbook = writer.book
        sheet = workbook['Terrain']
        
        color_culturel = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
        color_producteur_guerison = PatternFill(start_color="228B22", end_color="228B22", fill_type="solid")  # Vert forêt pour Guérison
        color_producteur_normal = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # Vert clair pour les autres
        color_neutre = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        color_libre = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        color_bordure = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        
        for pb in terrain.buildings:
            cells = pb.get_cells()
            for r, c in cells:
                cell = sheet.cell(row=r+1, column=c+1)
                if pb.building.type == "Culturel":
                    cell.fill = color_culturel
                elif pb.building.type == "Producteur":
                    if pb.building.production == "Guerison":
                        cell.fill = color_producteur_guerison
                    else:
                        cell.fill = color_producteur_normal
                else:
                    cell.fill = color_neutre
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.font = Font(size=9)
        
        for i in range(terrain.height):
            for j in range(terrain.width):
                cell = sheet.cell(row=i+1, column=j+1)
                if terrain.grid[i, j] == 1:
                    cell.fill = color_libre
                elif terrain.grid[i, j] == 0:
                    cell.fill = color_bordure
                    cell.font = Font(color="FFFFFF", bold=True)
                sheet.column_dimensions[get_column_letter(j+1)].width = 22
        
        # 4. Feuille des bâtiments non placés
        if unplaced_buildings:
            unplaced_data = []
            for b in unplaced_buildings:
                unplaced_data.append({
                    'Nom': b.nom,
                    'Type': b.type,
                    'Production': b.production,
                    'Longueur': b.longueur,
                    'Largeur': b.largeur,
                    'Cases': b.get_area()
                })
            
            unplaced_df = pd.DataFrame(unplaced_data)
            unplaced_df.to_excel(writer, sheet_name='Non places', index=False)
            
            total_cases = sum(b.get_area() for b in unplaced_buildings)
            total_row = pd.DataFrame({'Nom': ['TOTAL'], 'Type': [''], 'Production': [''], 
                                      'Longueur': [''], 'Largeur': [''], 'Cases': [total_cases]})
            total_row.to_excel(writer, sheet_name='Non places', startrow=len(unplaced_data) + 1, index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Placeur de Bâtiments", layout="wide")
    
    st.title("🏗️ Placeur de Bâtiments")
    st.markdown("Chargez un fichier Excel pour placer automatiquement les bâtiments sur le terrain")
    
    uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            terrain_grid, buildings = read_input_excel(uploaded_file)
            
            total_buildings_to_place = sum(b.nombre for b in buildings)
            st.success(f"✅ Fichier chargé avec succès! Terrain: {terrain_grid.shape[0]}x{terrain_grid.shape[1]}, {len(buildings)} types de bâtiments")
            st.info(f"📦 Nombre total de bâtiments à placer : {total_buildings_to_place}")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                max_placement = st.checkbox("🔧 Mode placement maximal", value=True, 
                                            help="Tente de placer TOUS les bâtiments, y compris ceux sans boost")
            with col2:
                optimize_boosts = st.checkbox("⚡ Optimiser les boosts", value=True,
                                              help="Réorganise les bâtiments pour maximiser les boosts des producteurs")
            
            if st.button("🚀 Lancer le placement", type="primary"):
                with st.spinner("Placement des bâtiments en cours..."):
                    terrain = Terrain(terrain_grid)
                    placer = BuildingPlacer(terrain, buildings)
                    
                    # Phase 1: Placement initial
                    placer.place_all_buildings()
                    
                    # Phase 2: Optimisation des boosts (entrelacement culturels/producteurs)
                    if optimize_boosts and len(placer.terrain.buildings) > 0:
                        st.info("⚡ Optimisation des boosts en cours...")
                        placer.optimize_boosts()
                    
                    # Phase 3: Placement des bâtiments restants (si mode maximal)
                    if max_placement and placer.unplaced:
                        st.info(f"🔧 Tentative de placement des {len(placer.unplaced)} bâtiments restants...")
                        placer.try_place_remaining()
                    
                    # Recalcul final des productions
                    production_stats = placer.calculate_culture_and_production()
                    output_file = create_output_excel(terrain, production_stats, placer.unplaced)
                    
                    # Affichage des résultats
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Bâtiments placés", len(terrain.buildings))
                    with col2:
                        st.metric("Bâtiments non placés", len(placer.unplaced))
                    with col3:
                        st.metric("Cases libres restantes", terrain.get_free_cells_count())
                    with col4:
                        cases_non_placees = sum(b.get_area() for b in placer.unplaced)
                        st.metric("Cases des bâtiments non placés", cases_non_placees)
                    
                    if len(placer.unplaced) == 0:
                        st.success("🎉 **SUCCÈS!** Tous les bâtiments ont été placés!")
                    else:
                        st.warning(f"⚠️ **Attention:** {len(placer.unplaced)} bâtiments n'ont pas pu être placés (manque d'espace).")
                    
                    # Résultats par type de production
                    if production_stats:
                        st.subheader("📊 Production par type")
                        
                        priority_prods = ["Guerison", "Nourriture", "Or"]
                        other_prods = [p for p in production_stats.keys() if p not in priority_prods and p != "Rien"]
                        
                        for prod_type in priority_prods:
                            if prod_type in production_stats:
                                stats = production_stats[prod_type]
                                with st.expander(f"**{prod_type}**", expanded=True):
                                    col_a, col_b, col_c = st.columns(3)
                                    with col_a:
                                        st.metric("Culture totale", f"{stats['total_culture']:.0f}")
                                        st.metric("Boost moyen", f"{stats['avg_boost']:.1f}%")
                                    with col_b:
                                        st.metric("Nombre de bâtiments", stats['count'])
                                    with col_c:
                                        st.metric("Production/heure", f"{stats['total_production']:.2f}")
                        
                        if other_prods:
                            st.subheader("📦 Autres productions")
                            for prod_type in other_prods:
                                stats = production_stats[prod_type]
                                with st.expander(f"{prod_type}"):
                                    col_a, col_b, col_c = st.columns(3)
                                    with col_a:
                                        st.metric("Culture totale", f"{stats['total_culture']:.0f}")
                                        st.metric("Boost moyen", f"{stats['avg_boost']:.1f}%")
                                    with col_b:
                                        st.metric("Nombre de bâtiments", stats['count'])
                                    with col_c:
                                        st.metric("Production/heure", f"{stats['total_production']:.2f}")
                    
                    # Bâtiments non placés
                    if placer.unplaced:
                        st.subheader("⚠️ Bâtiments non placés")
                        unplaced_summary = {}
                        for b in placer.unplaced:
                            if b.nom not in unplaced_summary:
                                unplaced_summary[b.nom] = 0
                            unplaced_summary[b.nom] += 1
                        
                        cols = st.columns(4)
                        col_idx = 0
                        for nom, count in sorted(unplaced_summary.items(), key=lambda x: -x[1]):
                            with cols[col_idx % 4]:
                                st.write(f"- **{nom}**: {count} exemplaire(s)")
                            col_idx += 1
                    
                    st.divider()
                    st.download_button(
                        label="📥 Télécharger le fichier de résultats",
                        data=output_file,
                        file_name="resultats_placement.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()
