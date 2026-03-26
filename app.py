import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st
import io
from typing import List, Tuple, Dict, Optional
import copy
import random

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
        cells = []
        if orientation == "horizontal":
            for i in range(building.longueur):
                for j in range(building.largeur):
                    r, c = row + i, col + j
                    if r >= self.height or c >= self.width:
                        return False
                    if not self.is_cell_free(r, c):
                        return False
                    cells.append((r, c))
        else:
            for i in range(building.largeur):
                for j in range(building.longueur):
                    r, c = row + i, col + j
                    if r >= self.height or c >= self.width:
                        return False
                    if not self.is_cell_free(r, c):
                        return False
                    cells.append((r, c))
        return True
    
    def place_building(self, building, row, col, orientation):
        placed = PlacedBuilding(building, row, col, orientation)
        cells = placed.get_cells()
        for r, c in cells:
            self.grid[r, c] = 2
        self.buildings.append(placed)
        return placed
    
    def get_free_cells_count(self):
        return np.sum(self.grid == 1)
    
    def get_all_free_cells(self):
        """Retourne la liste de toutes les cellules libres"""
        free_cells = []
        for i in range(self.height):
            for j in range(self.width):
                if self.grid[i, j] == 1:
                    free_cells.append((i, j))
        return free_cells

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
    
    def find_all_positions(self, building):
        """Trouve toutes les positions possibles pour un bâtiment"""
        positions = []
        for orientation in ["horizontal", "vertical"]:
            for row in range(self.terrain.height):
                for col in range(self.terrain.width):
                    if self.terrain.can_place_building(building, row, col, orientation):
                        positions.append((row, col, orientation))
        return positions
    
    def find_best_position_by_area(self, building, prefer_border=False, avoid_border=False):
        """Trouve la meilleure position basée sur la surface disponible autour"""
        positions = self.find_all_positions(building)
        if not positions:
            return None
        
        best_pos = None
        best_score = -1
        
        border_cells = set(self.get_border_cells())
        
        for row, col, orientation in positions:
            score = 0
            
            # Pénaliser les positions qui touchent les bords si on veut éviter
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
                
                touches_border = any(r == 0 or r == self.terrain.height - 1 or 
                                    c == 0 or c == self.terrain.width - 1 for r, c in cells)
                if touches_border:
                    score -= 100
            
            # Favoriser les bords si demandé
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
            
            # Favoriser les positions qui laissent plus d'espace pour les autres bâtiments
            # (simulation simple : compter les cases libres adjacentes)
            score += self._calculate_space_score(row, col, orientation, building)
            
            if score > best_score:
                best_score = score
                best_pos = (row, col, orientation)
        
        return best_pos
    
    def _calculate_space_score(self, row, col, orientation, building):
        """Calcule un score basé sur l'espace disponible autour"""
        score = 0
        cells = []
        if orientation == "horizontal":
            for i in range(building.longueur):
                for j in range(building.largeur):
                    cells.append((row + i, col + j))
        else:
            for i in range(building.largeur):
                for j in range(building.longueur):
                    cells.append((row + i, col + j))
        
        # Vérifier les cellules adjacentes
        for r, c in cells:
            for dr in [-1, 0, 1]:
                for dc in [-1, 0, 1]:
                    if dr == 0 and dc == 0:
                        continue
                    nr, nc = r + dr, c + dc
                    if 0 <= nr < self.terrain.height and 0 <= nc < self.terrain.width:
                        if self.terrain.is_cell_free(nr, nc):
                            score += 1
        return score
    
    def place_all_buildings_with_priority(self, force_place_all=True):
        """Place tous les bâtiments avec priorité mais en essayant de tout placer"""
        
        # 1. Créer une liste de tous les bâtiments à placer
        all_instances = []
        
        # Définir la priorité des productions
        production_priority = {"Guerison": 0, "Nourriture": 1, "Or": 2}
        
        # Ajouter les bâtiments neutres (priorité basse pour le placement)
        for b in self.neutral_buildings:
            for _ in range(b.nombre):
                all_instances.append({
                    'building': b,
                    'type': 'neutre',
                    'priority': 4,
                    'area': b.get_area()
                })
        
        # Ajouter les bâtiments culturels (priorité moyenne)
        for b in self.cultural_buildings:
            for _ in range(b.nombre):
                all_instances.append({
                    'building': b,
                    'type': 'culturel',
                    'priority': 2,
                    'area': b.get_area()
                })
        
        # Ajouter les bâtiments producteurs
        for b in self.producer_buildings:
            priority = production_priority.get(b.production, 3)
            for _ in range(b.nombre):
                all_instances.append({
                    'building': b,
                    'type': 'producteur',
                    'priority': priority,
                    'area': b.get_area()
                })
        
        # Trier par priorité (plus petit = plus prioritaire) et par taille (plus grand d'abord)
        all_instances.sort(key=lambda x: (x['priority'], -x['area']))
        
        # 2. Premier passage : placement avec contraintes
        placed_count = 0
        for instance in all_instances:
            building = instance['building']
            building_type = instance['type']
            
            if building_type == 'neutre':
                # Neutres : sur les bords
                pos = self.find_best_position_by_area(building, prefer_border=True)
            elif building_type == 'culturel':
                # Culturels : éviter les bords si possible
                pos = self.find_best_position_by_area(building, avoid_border=True)
                if not pos and force_place_all:
                    # Si pas trouvé, essayer sans contrainte
                    pos = self.find_best_position_by_area(building)
            else:
                # Producteurs : placement normal
                pos = self.find_best_position_by_area(building)
            
            if pos:
                row, col, orientation = pos
                self.terrain.place_building(building, row, col, orientation)
                placed_count += 1
            else:
                self.unplaced.append(building)
        
        # 3. Si force_place_all est True et qu'il reste des bâtiments, 
        #    essayer de replacer en assouplissant les contraintes
        if force_place_all and self.unplaced:
            # Sauvegarder l'état actuel
            original_terrain = copy.deepcopy(self.terrain)
            original_buildings = copy.deepcopy(self.terrain.buildings)
            
            # Essayer de replacer en mode "tout accepter"
            # On va d'abord essayer de supprimer les bâtiments les moins prioritaires
            # pour faire de la place
            
            # Regrouper les bâtiments par type
            placed_by_type = {'neutre': [], 'culturel': [], 'producteur': []}
            for pb in self.terrain.buildings:
                if pb.building.type == 'Neutre':
                    placed_by_type['neutre'].append(pb)
                elif pb.building.type == 'Culturel':
                    placed_by_type['culturel'].append(pb)
                else:
                    placed_by_type['producteur'].append(pb)
            
            # Trier les producteurs par priorité (les moins prioritaires d'abord)
            producteurs_tries = sorted(placed_by_type['producteur'], 
                                      key=lambda x: production_priority.get(x.building.production, 3),
                                      reverse=True)
            
            # Essayer de retirer des bâtiments pour faire de la place
            for removed_building in producteurs_tries:
                if not self.unplaced:
                    break
                
                # Retirer ce bâtiment
                self.terrain = copy.deepcopy(original_terrain)
                self.terrain.buildings = [b for b in original_buildings if b != removed_building]
                # Reconstruire la grille
                self.terrain.grid = copy.deepcopy(original_terrain.grid)
                for b in self.terrain.buildings:
                    for r, c in b.get_cells():
                        self.terrain.grid[r, c] = 2
                
                # Essayer de placer les bâtiments non placés
                unplaced_temp = self.unplaced.copy()
                self.unplaced = []
                
                for unplaced_building in unplaced_temp:
                    pos = self.find_best_position_by_area(unplaced_building)
                    if pos:
                        row, col, orientation = pos
                        self.terrain.place_building(unplaced_building, row, col, orientation)
                    else:
                        self.unplaced.append(unplaced_building)
                
                # Si on a réussi à placer plus de bâtiments, garder cette configuration
                if len(self.unplaced) < len(unplaced_temp):
                    original_terrain = copy.deepcopy(self.terrain)
                    original_buildings = copy.deepcopy(self.terrain.buildings)
                else:
                    # Annuler les changements
                    self.terrain = copy.deepcopy(original_terrain)
                    self.terrain.buildings = copy.deepcopy(original_buildings)
                    self.unplaced = unplaced_temp
        
        # 4. Dernier essai : placement sans aucune contrainte
        if force_place_all and self.unplaced:
            remaining = self.unplaced.copy()
            self.unplaced = []
            
            for building in remaining:
                # Essayer toutes les orientations et positions
                positions = self.find_all_positions(building)
                if positions:
                    # Prendre la première position disponible
                    row, col, orientation = positions[0]
                    self.terrain.place_building(building, row, col, orientation)
                else:
                    self.unplaced.append(building)
    
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
    
    def place_all_buildings(self):
        """Méthode principale de placement"""
        # 1. Placer les bâtiments neutres sur les bords
        self.neutral_buildings.sort(key=lambda x: x.get_area(), reverse=True)
        for building in self.neutral_buildings:
            for _ in range(building.nombre):
                pos = self.find_best_position_by_area(building, prefer_border=True)
                if pos:
                    row, col, orientation = pos
                    self.terrain.place_building(building, row, col, orientation)
                else:
                    self.unplaced.append(building)
        
        # 2. Placer tous les bâtiments restants avec priorité mais en forçant le placement
        self.place_all_buildings_with_priority(force_place_all=True)

def read_input_excel(file):
    """Lecture du fichier Excel d'entrée"""
    xls = pd.ExcelFile(file)
    
    # Lire le terrain
    terrain_df = pd.read_excel(xls, sheet_name="Terrain", header=None)
    terrain_grid = terrain_df.fillna(0).values
    
    # Convertir les X en 0 et les 1 en 1
    terrain_grid = np.where(terrain_grid == 'X', 0, terrain_grid)
    terrain_grid = np.where(terrain_grid == 1, 1, terrain_grid)
    terrain_grid = np.where(terrain_grid == 0, 0, terrain_grid)
    terrain_grid = terrain_grid.astype(int)
    
    # Lire les bâtiments
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

def create_output_excel(terrain, production_stats, unplaced_buildings):
    """Créer le fichier Excel de sortie"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Feuille des bâtiments placés
        placed_data = []
        for pb in terrain.buildings:
            placed_data.append({
                'Nom': pb.building.nom,
                'Type': pb.building.type,
                'Production': pb.building.production,
                'Ligne': pb.row + 1,
                'Colonne': pb.col + 1,
                'Hauteur': pb.building.longueur if pb.orientation == "horizontal" else pb.building.largeur,
                'Largeur': pb.building.largeur if pb.orientation == "horizontal" else pb.building.longueur,
                'Culture recue': pb.culture_recue,
                'Boost (%)': pb.building.get_boost_percentage(pb.culture_recue),
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
                    'Boost moyen (%)': stats['avg_boost'],
                    'Nb batiments': stats['count'],
                    'Production/h': stats['total_production']
                })
        
        if synthesis_data:
            synthesis_df = pd.DataFrame(synthesis_data)
            synthesis_df.to_excel(writer, sheet_name='Synthese', index=False)
        
        # Ajouter les cases libres et non placées
        total_unplaced_area = sum(b.get_area() for b in unplaced_buildings)
        extra_data = pd.DataFrame({
            'Statistiques': ['Cases libres restantes', 'Cases des batiments non places', 'Nombre de batiments non places'],
            'Valeur': [terrain.get_free_cells_count(), total_unplaced_area, len(unplaced_buildings)]
        })
        extra_data.to_excel(writer, sheet_name='Synthese', startrow=len(synthesis_data) + 3, index=False)
        
        # 3. Feuille du terrain avec bâtiments
        terrain_visual = create_terrain_visualization(terrain)
        terrain_visual_df = pd.DataFrame(terrain_visual)
        terrain_visual_df.to_excel(writer, sheet_name='Terrain', index=False, header=False)
        
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

def create_terrain_visualization(terrain):
    """Créer une représentation visuelle du terrain avec les noms des bâtiments"""
    visual = [['' for _ in range(terrain.width)] for _ in range(terrain.height)]
    
    # Remplir avec le terrain de base (1 = libre)
    for i in range(terrain.height):
        for j in range(terrain.width):
            if terrain.grid[i, j] == 1:
                visual[i][j] = 'Libre'
            elif terrain.grid[i, j] == 0:
                visual[i][j] = 'X'
    
    # Ajouter les bâtiments
    for pb in terrain.buildings:
        cells = pb.get_cells()
        boost = pb.building.get_boost_percentage(pb.culture_recue)
        boost_text = f"+{boost}%" if boost > 0 else ""
        
        for idx, (r, c) in enumerate(cells):
            if idx == 0:
                if pb.building.type == "Culturel":
                    visual[r][c] = f"{pb.building.nom} (C)\n{boost_text}".strip()
                elif pb.building.type == "Producteur":
                    visual[r][c] = f"{pb.building.nom} (P)\n{boost_text}".strip()
                else:
                    visual[r][c] = f"{pb.building.nom}\n{boost_text}".strip()
            else:
                visual[r][c] = ''
    
    return visual

def main():
    st.set_page_config(page_title="Placeur de Bâtiments - Version Optimisée", layout="wide")
    
    st.title("🏗️ Placeur de Bâtiments - Version Maximale")
    st.markdown("Chargez un fichier Excel pour placer **TOUS** les bâtiments (si techniquement possible)")
    
    uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # Lecture du fichier
            terrain_grid, buildings = read_input_excel(uploaded_file)
            
            st.success(f"✅ Fichier chargé avec succès! Terrain: {terrain_grid.shape[0]}x{terrain_grid.shape[1]}, {len(buildings)} types de bâtiments")
            
            # Afficher le nombre total de bâtiments à placer
            total_buildings = sum(b.nombre for b in buildings)
            st.info(f"📦 Nombre total de bâtiments à placer : {total_buildings}")
            
            if st.button("🚀 Lancer le placement maximal", type="primary"):
                with st.spinner("Placement des bâtiments en cours (cela peut prendre quelques secondes)..."):
                    # Initialisation
                    terrain = Terrain(terrain_grid)
                    placer = BuildingPlacer(terrain, buildings)
                    
                    # Placement
                    placer.place_all_buildings()
                    
                    # Calculs
                    production_stats = placer.calculate_culture_and_production()
                    
                    # Création du fichier de sortie
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
                    
                    # Message sur le résultat
                    if len(placer.unplaced) == 0:
                        st.success("🎉 **SUCCÈS!** Tous les bâtiments ont été placés!")
                    else:
                        st.warning(f"⚠️ **Attention:** {len(placer.unplaced)} bâtiments n'ont pas pu être placés (manque d'espace).")
                    
                    # Résultats par type de production
                    if production_stats:
                        st.subheader("📊 Production par type")
                        for prod_type, stats in production_stats.items():
                            if prod_type != "Rien":
                                with st.expander(f"**{prod_type}**"):
                                    col1, col2, col3 = st.columns(3)
                                    with col1:
                                        st.metric("Culture totale", f"{stats['total_culture']:.0f}")
                                        st.metric("Boost moyen", f"{stats['avg_boost']:.1f}%")
                                    with col2:
                                        st.metric("Nombre de bâtiments", stats['count'])
                                    with col3:
                                        st.metric("Production/heure", f"{stats['total_production']:.2f}")
                    
                    # Bâtiments non placés
                    if placer.unplaced:
                        st.subheader("⚠️ Bâtiments non placés")
                        unplaced_summary = {}
                        for b in placer.unplaced:
                            if b.nom not in unplaced_summary:
                                unplaced_summary[b.nom] = 0
                            unplaced_summary[b.nom] += 1
                        
                        cols = st.columns(3)
                        col_idx = 0
                        for nom, count in unplaced_summary.items():
                            with cols[col_idx % 3]:
                                st.write(f"- **{nom}**: {count} exemplaire(s)")
                            col_idx += 1
                    
                    # Téléchargement
                    st.divider()
                    st.download_button(
                        label="📥 Télécharger le fichier de résultats",
                        data=output_file,
                        file_name="resultats_placement_maximal.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()
