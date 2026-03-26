import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.cell.cell import MergedCell
import streamlit as st
import io
from typing import List, Tuple, Dict, Optional, Set
import copy
import heapq
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
    
    def get_rayonnement_value(self):
        return self.culture * (self.rayonnement + 1)

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
    
    def get_all_cells_set(self):
        return set(self.get_cells())
    
    def get_rayonnement_cells(self):
        cells = set()
        building_cells = self.get_cells()
        radius = self.building.rayonnement
        
        for r, c in building_cells:
            for dr in range(-radius, radius + 1):
                for dc in range(-radius, radius + 1):
                    nr, nc = r + dr, c + dc
                    if nr >= 0 and nc >= 0:
                        cells.add((nr, nc))
        
        cells -= set(building_cells)
        return cells
    
    def update_culture(self, cultural_buildings):
        total_culture = 0
        producer_cells = self.get_all_cells_set()
        
        for cultural in cultural_buildings:
            if cultural == self or cultural.building.type != "Culturel":
                continue
                
            cultural_cells = cultural.get_all_cells_set()
            cultural_radius = cultural.building.rayonnement
            
            for prod_cell in producer_cells:
                for cult_cell in cultural_cells:
                    distance = max(abs(prod_cell[0] - cult_cell[0]), abs(prod_cell[1] - cult_cell[1]))
                    if distance <= cultural_radius:
                        total_culture += cultural.building.culture
                        break
                else:
                    continue
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
    
    def get_neutral_buildings(self):
        return [b for b in self.buildings if b.building.type == "Neutre"]
    
    def get_free_cells_in_radius(self, row, col, radius):
        count = 0
        for dr in range(-radius, radius + 1):
            for dc in range(-radius, radius + 1):
                nr, nc = row + dr, col + dc
                if 0 <= nr < self.height and 0 <= nc < self.width:
                    if self.is_cell_free(nr, nc):
                        count += 1
        return count
    
    def find_largest_free_rectangle(self):
        """Trouve le plus grand rectangle libre dans le terrain"""
        max_area = 0
        best_rect = None
        
        # Pour chaque cellule, calculer la hauteur maximale de cellules libres
        height_matrix = np.zeros((self.height, self.width), dtype=int)
        
        for i in range(self.height):
            for j in range(self.width):
                if self.grid[i, j] == 1:
                    height_matrix[i, j] = height_matrix[i-1, j] + 1 if i > 0 else 1
                else:
                    height_matrix[i, j] = 0
        
        # Pour chaque ligne, trouver le plus grand rectangle
        for i in range(self.height):
            stack = []
            for j in range(self.width):
                while stack and height_matrix[i, stack[-1]] > height_matrix[i, j]:
                    h = height_matrix[i, stack.pop()]
                    w = j if not stack else j - stack[-1] - 1
                    area = h * w
                    if area > max_area:
                        max_area = area
                        best_rect = (i - h + 1, stack[-1] + 1 if stack else 0, h, w)
                stack.append(j)
            
            while stack:
                h = height_matrix[i, stack.pop()]
                w = self.width if not stack else self.width - stack[-1] - 1
                area = h * w
                if area > max_area:
                    max_area = area
                    best_rect = (i - h + 1, stack[-1] + 1 if stack else 0, h, w)
        
        return best_rect

class BuildingPlacer:
    def __init__(self, terrain, buildings):
        self.terrain = terrain
        self.buildings = buildings
        self.neutral_buildings = [b for b in buildings if b.type == "Neutre"]
        self.cultural_buildings = [b for b in buildings if b.type == "Culturel"]
        self.producer_buildings = [b for b in buildings if b.type == "Producteur"]
        self.unplaced = []
        
        self.neutral_instances = []
        for b in self.neutral_buildings:
            for _ in range(b.nombre):
                self.neutral_instances.append(b)
        
        self.cultural_instances = []
        for b in self.cultural_buildings:
            for _ in range(b.nombre):
                self.cultural_instances.append(b)
        
        self.producer_instances = []
        for b in self.producer_buildings:
            for _ in range(b.nombre):
                self.producer_instances.append(b)
        
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
    
    def find_position_with_priority(self, building, prefer_border=False, avoid_border=False):
        best_position = None
        best_score = -1e9
        
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
                        
                        if building.type == "Producteur":
                            cultural_buildings = self.terrain.get_cultural_buildings()
                            temp_producer = PlacedBuilding(building, row, col, orientation)
                            temp_cells = temp_producer.get_all_cells_set()
                            
                            for cultural in cultural_buildings:
                                cultural_cells = cultural.get_all_cells_set()
                                cultural_radius = cultural.building.rayonnement
                                
                                for prod_cell in temp_cells:
                                    for cult_cell in cultural_cells:
                                        dist = max(abs(prod_cell[0] - cult_cell[0]), abs(prod_cell[1] - cult_cell[1]))
                                        if dist <= cultural_radius:
                                            score += cultural.building.culture * 50
                                            break
                        
                        if building.type == "Culturel":
                            cells = []
                            if orientation == "horizontal":
                                for i in range(building.longueur):
                                    for j in range(building.largeur):
                                        cells.append((row + i, col + j))
                            else:
                                for i in range(building.largeur):
                                    for j in range(building.longueur):
                                        cells.append((row + i, col + j))
                            
                            free_space = 0
                            for r, c in cells:
                                free_space += self.terrain.get_free_cells_in_radius(r, c, building.rayonnement + 2)
                            score += free_space
                        
                        if score > best_score:
                            best_score = score
                            best_position = (row, col, orientation)
        
        return best_position
    
    def find_any_position(self, building):
        for orientation in ["horizontal", "vertical"]:
            for row in range(self.terrain.height):
                for col in range(self.terrain.width):
                    if self.terrain.can_place_building(building, row, col, orientation):
                        return (row, col, orientation)
        return None
    
    def find_position_in_rectangle(self, building, rect):
        """Trouve une position pour le bâtiment dans un rectangle donné"""
        rect_row, rect_col, rect_height, rect_width = rect
        
        for orientation in ["horizontal", "vertical"]:
            for row in range(rect_row, rect_row + rect_height):
                for col in range(rect_col, rect_col + rect_width):
                    if self.terrain.can_place_building(building, row, col, orientation):
                        # Vérifier que tout le bâtiment reste dans le rectangle
                        if orientation == "horizontal":
                            if row + building.longueur <= rect_row + rect_height and col + building.largeur <= rect_col + rect_width:
                                return (row, col, orientation)
                        else:
                            if row + building.largeur <= rect_row + rect_height and col + building.longueur <= rect_col + rect_width:
                                return (row, col, orientation)
        return None
    
    def consolidate_free_space(self):
        """Tente de regrouper les espaces libres en déplaçant des bâtiments"""
        # Trier les bâtiments par taille décroissante
        all_buildings = sorted(self.terrain.buildings, key=lambda x: -x.building.get_area())
        
        # Pour chaque bâtiment, essayer de le déplacer pour libérer un grand espace
        for building in all_buildings[:50]:  # Limiter pour performance
            if building in self.terrain.buildings:
                # Sauvegarder l'ancienne position
                old_cells = building.get_cells()
                old_row, old_col = building.row, building.col
                old_orientation = building.orientation
                
                # Chercher une nouvelle position
                new_pos = self.find_position_with_priority(building.building)
                
                if new_pos:
                    new_row, new_col, new_orientation = new_pos
                    
                    # Vérifier que la nouvelle position est différente
                    if (new_row, new_col, new_orientation) != (old_row, old_col, old_orientation):
                        # Tester le déplacement
                        self.terrain.remove_building(building)
                        
                        if self.terrain.can_place_building(building.building, new_row, new_col, new_orientation):
                            self.terrain.place_building(building.building, new_row, new_col, new_orientation)
                        else:
                            # Revenir à l'ancienne position
                            self.terrain.place_building(building.building, old_row, old_col, old_orientation)
    
    def try_place_remaining_with_consolidation(self):
        """Tente de placer les bâtiments restants après consolidation des espaces"""
        if not self.unplaced:
            return
        
        # D'abord, essayer de consolider l'espace libre
        self.consolidate_free_space()
        
        # Recalculer les espaces libres
        free_cells = self.terrain.get_free_cells_count()
        
        # Trier les bâtiments non placés par taille décroissante
        self.unplaced.sort(key=lambda x: -x.get_area())
        
        remaining = []
        for building in self.unplaced:
            # Chercher le plus grand rectangle libre
            largest_rect = self.terrain.find_largest_free_rectangle()
            
            if largest_rect:
                rect_row, rect_col, rect_height, rect_width = largest_rect
                
                # Essayer de placer dans ce rectangle
                pos = self.find_position_in_rectangle(building, largest_rect)
                
                if pos:
                    row, col, orientation = pos
                    self.terrain.place_building(building, row, col, orientation)
                else:
                    # Essayer n'importe où
                    pos = self.find_any_position(building)
                    if pos:
                        row, col, orientation = pos
                        self.terrain.place_building(building, row, col, orientation)
                    else:
                        remaining.append(building)
            else:
                # Pas de rectangle, essayer n'importe où
                pos = self.find_any_position(building)
                if pos:
                    row, col, orientation = pos
                    self.terrain.place_building(building, row, col, orientation)
                else:
                    remaining.append(building)
        
        self.unplaced = remaining
    
    def place_all_buildings_complete(self):
        # Phase 1: Placement initial
        self.neutral_instances.sort(key=lambda x: -x.get_area())
        remaining_neutral = []
        for building in self.neutral_instances:
            pos = self.find_position_with_priority(building, prefer_border=True)
            if pos:
                row, col, orientation = pos
                self.terrain.place_building(building, row, col, orientation)
            else:
                remaining_neutral.append(building)
        
        self.cultural_instances.sort(key=lambda x: -x.get_rayonnement_value())
        
        production_priority = {"Guerison": 0, "Nourriture": 1, "Or": 2, "Autre": 3}
        self.producer_instances.sort(key=lambda x: (production_priority.get(x.production, 3), -x.get_area()))
        
        used_cultural = set()
        used_producers = set()
        remaining_producers = list(self.producer_instances)
        
        for cultural in self.cultural_instances:
            if cultural in used_cultural:
                continue
            
            pos = self.find_position_with_priority(cultural, avoid_border=True)
            if not pos:
                pos = self.find_position_with_priority(cultural)
            
            if pos:
                row, col, orientation = pos
                placed_cultural = self.terrain.place_building(cultural, row, col, orientation)
                used_cultural.add(cultural)
                
                rayonnement_cells = list(placed_cultural.get_rayonnement_cells())
                
                producers_placed = 0
                for producer in remaining_producers[:]:
                    if producers_placed >= 4:
                        break
                    
                    best_prod_pos = None
                    best_prod_score = -1
                    
                    for r, c in rayonnement_cells[:100]:
                        for orientation_prod in ["horizontal", "vertical"]:
                            if self.terrain.can_place_building(producer, r, c, orientation_prod):
                                temp_prod = PlacedBuilding(producer, r, c, orientation_prod)
                                temp_cells = temp_prod.get_all_cells_set()
                                
                                in_rayonnement = False
                                for prod_cell in temp_cells:
                                    for cult_cell in placed_cultural.get_all_cells_set():
                                        dist = max(abs(prod_cell[0] - cult_cell[0]), abs(prod_cell[1] - cult_cell[1]))
                                        if dist <= cultural.rayonnement:
                                            in_rayonnement = True
                                            break
                                    if in_rayonnement:
                                        break
                                
                                if in_rayonnement:
                                    score = cultural.culture
                                    if score > best_prod_score:
                                        best_prod_score = score
                                        best_prod_pos = (r, c, orientation_prod)
                    
                    if best_prod_pos:
                        r, c, orientation_prod = best_prod_pos
                        self.terrain.place_building(producer, r, c, orientation_prod)
                        remaining_producers.remove(producer)
                        used_producers.add(producer)
                        producers_placed += 1
            else:
                pass
        
        for producer in remaining_producers:
            if producer not in used_producers:
                pos = self.find_position_with_priority(producer)
                if pos:
                    row, col, orientation = pos
                    self.terrain.place_building(producer, row, col, orientation)
                    used_producers.add(producer)
        
        for cultural in self.cultural_instances:
            if cultural not in used_cultural:
                pos = self.find_position_with_priority(cultural)
                if pos:
                    row, col, orientation = pos
                    self.terrain.place_building(cultural, row, col, orientation)
                    used_cultural.add(cultural)
        
        for building in remaining_neutral:
            pos = self.find_any_position(building)
            if pos:
                row, col, orientation = pos
                self.terrain.place_building(building, row, col, orientation)
            else:
                self.unplaced.append(building)
        
        # Phase 2: Consolidation et placement des bâtiments restants
        if self.unplaced:
            self.try_place_remaining_with_consolidation()
        
        # Phase 3: Vérification finale
        all_placed = set([b.building for b in self.terrain.buildings])
        for b in self.cultural_instances:
            if b not in all_placed and b not in self.unplaced:
                self.unplaced.append(b)
        for b in self.producer_instances:
            if b not in all_placed and b not in self.unplaced:
                self.unplaced.append(b)
        for b in self.neutral_instances:
            if b not in all_placed and b not in self.unplaced:
                self.unplaced.append(b)
    
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
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
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
        
        terrain_visual = create_terrain_visualization(terrain)
        terrain_visual_df = pd.DataFrame(terrain_visual)
        terrain_visual_df.to_excel(writer, sheet_name='Terrain', index=False, header=False)
        
        workbook = writer.book
        sheet = workbook['Terrain']
        
        color_culturel = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
        color_producteur_guerison = PatternFill(start_color="228B22", end_color="228B22", fill_type="solid")
        color_producteur_normal = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        color_neutre = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        color_libre = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        color_bordure = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
        
        row_height = 45
        for i in range(1, terrain.height + 1):
            sheet.row_dimensions[i].height = row_height
        
        col_width = 7.14
        for j in range(1, terrain.width + 1):
            sheet.column_dimensions[get_column_letter(j)].width = col_width
        
        for pb in terrain.buildings:
            cells = pb.get_cells()
            if not cells:
                continue
            
            cells_sorted = sorted(cells, key=lambda x: (x[0], x[1]))
            start_cell = cells_sorted[0]
            
            rows = [r for r, c in cells]
            cols = [c for r, c in cells]
            height = max(rows) - min(rows) + 1
            width = max(cols) - min(cols) + 1
            
            start_row = start_cell[0] + 1
            start_col = start_cell[1] + 1
            end_row = start_row + height - 1
            end_col = start_col + width - 1
            
            start_cell_ref = f"{get_column_letter(start_col)}{start_row}"
            end_cell_ref = f"{get_column_letter(end_col)}{end_row}"
            
            try:
                sheet.merge_cells(f"{start_cell_ref}:{end_cell_ref}")
            except:
                pass
            
            cell = sheet.cell(row=start_row, column=start_col)
            boost = pb.building.get_boost_percentage(pb.culture_recue)
            boost_text = f" +{boost}%" if boost > 0 else ""
            cell.value = f"{pb.building.nom}{boost_text}"
            
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
            cell.font = Font(size=10, bold=True)
        
        for i in range(terrain.height):
            for j in range(terrain.width):
                cell = sheet.cell(row=i+1, column=j+1)
                if isinstance(cell, MergedCell):
                    continue
                
                if terrain.grid[i, j] == 1:
                    cell.fill = color_libre
                    cell.value = ''
                elif terrain.grid[i, j] == 0:
                    cell.fill = color_bordure
                    cell.font = Font(color="FFFFFF", bold=True)
                    if cell.value != 'X':
                        cell.value = 'X'
        
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
        
        if unplaced_data:
            total_cases = sum(b.get_area() for b in unplaced_buildings)
            total_row = pd.DataFrame({'Nom': ['TOTAL'], 'Type': [''], 'Production': [''], 
                                      'Longueur': [''], 'Largeur': [''], 'Cases': [total_cases]})
            total_row.to_excel(writer, sheet_name='Non places', startrow=len(unplaced_data) + 1, index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Placeur de Bâtiments", layout="wide")
    
    st.title("🏗️ Placeur de Bâtiments - Version Consolidée")
    st.markdown("Chargez un fichier Excel pour placer les bâtiments avec consolidation des espaces")
    
    uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            terrain_grid, buildings = read_input_excel(uploaded_file)
            
            total_buildings_to_place = sum(b.nombre for b in buildings)
            st.success(f"✅ Fichier chargé avec succès! Terrain: {terrain_grid.shape[0]}x{terrain_grid.shape[1]}, {len(buildings)} types de bâtiments")
            st.info(f"📦 Nombre total de bâtiments à placer : {total_buildings_to_place}")
            
            if st.button("🚀 Lancer le placement", type="primary"):
                with st.spinner("Placement et consolidation des espaces..."):
                    terrain = Terrain(terrain_grid)
                    placer = BuildingPlacer(terrain, buildings)
                    
                    placer.place_all_buildings_complete()
                    
                    production_stats = placer.calculate_culture_and_production()
                    output_file = create_output_excel(terrain, production_stats, placer.unplaced)
                    
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
                        st.warning(f"⚠️ **Attention:** {len(placer.unplaced)} bâtiments n'ont pas pu être placés.")
                        
                        st.subheader("📋 Bâtiments non placés")
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
                    
                    st.divider()
                    st.download_button(
                        label="📥 Télécharger le fichier de résultats",
                        data=output_file,
                        file_name="resultats_placement_consolide.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement: {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()
