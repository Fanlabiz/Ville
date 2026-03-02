import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import io
from typing import List, Tuple, Dict, Optional
import re
from collections import defaultdict
import math
import itertools

class Building:
    def __init__(self, nom, longueur, largeur, quantite, type_bat, culture, rayonnement, 
                 boost_25, boost_50, boost_100, production):
        self.nom = nom
        self.longueur = int(float(longueur)) if pd.notna(longueur) and longueur != '' else 1
        self.largeur = int(float(largeur)) if pd.notna(largeur) and largeur != '' else 1
        self.quantite = int(float(quantite)) if pd.notna(quantite) and quantite != '' else 0
        self.type = str(type_bat).lower() if pd.notna(type_bat) and type_bat != '' else ""
        self.culture = float(culture) if pd.notna(culture) and culture != '' else 0
        self.rayonnement = int(float(rayonnement)) if pd.notna(rayonnement) and rayonnement != '' else 0
        self.boost_25 = float(boost_25) if pd.notna(boost_25) and boost_25 != '' else 0
        self.boost_50 = float(boost_50) if pd.notna(boost_50) and boost_50 != '' else 0
        self.boost_100 = float(boost_100) if pd.notna(boost_100) and boost_100 != '' else 0
        self.production = str(production) if pd.notna(production) and production != '' else ""
        
        # Pour le placement
        self.placed = 0
        self.positions = []  # Liste de tuples (x, y, orientation)
        self.failed_attempts = []  # Raisons des échecs
        self.id = f"{nom}_{id(self)}"
        
    def get_dimensions(self, orientation='H'):
        """Retourne (longueur, largeur) selon l'orientation"""
        if orientation == 'H':
            return self.longueur, self.largeur
        else:  # Vertical
            return self.largeur, self.longueur
            
    def get_area(self):
        """Retourne la surface du bâtiment"""
        return self.longueur * self.largeur
            
    def __repr__(self):
        return f"{self.nom} ({self.longueur}x{self.largeur})"

class Terrain:
    def __init__(self, grid):
        self.grid = np.array(grid)
        self.height, self.width = self.grid.shape
        self.occupied = np.zeros_like(self.grid, dtype=bool)
        self.buildings = []  # Liste des bâtiments placés
        self.cultural_buildings = []  # Liste des bâtiments culturels pour référence
        self.cultural_zones = np.empty((self.height, self.width), dtype=object)
        self.building_type_map = np.full_like(self.grid, '', dtype=object)
        self.building_objects = []  # Liste des objets Building placés
        self.reset_cultural_zones()
        
    def reset_cultural_zones(self):
        """Réinitialise les zones culturelles"""
        for i in range(self.height):
            for j in range(self.width):
                self.cultural_zones[i, j] = set()
        
    def can_place(self, x, y, longueur, largeur):
        """Vérifie si un bâtiment peut être placé à la position (x,y)"""
        if x + longueur > self.width or y + largeur > self.height:
            return False, "Hors limites"
            
        for i in range(longueur):
            for j in range(largeur):
                if self.grid[y + j, x + i] == 0:
                    return False, "Case occupée (0 dans le terrain)"
                if self.occupied[y + j, x + i]:
                    return False, "Case déjà occupée par un bâtiment"
        return True, "OK"
    
    def place_building(self, building, x, y, orientation):
        """Place un bâtiment sur le terrain"""
        longueur, largeur = building.get_dimensions(orientation)
        
        for i in range(longueur):
            for j in range(largeur):
                self.occupied[y + j, x + i] = True
                self.building_type_map[y + j, x + i] = building.type
                
        building.placed += 1
        building.positions.append((x, y, orientation))
        self.buildings.append((building, x, y, orientation, longueur, largeur))
        self.building_objects.append(building)
        
        if building.type == "culturel" and building.culture > 0:
            self.cultural_buildings.append((building, x, y, orientation, longueur, largeur))
            self.update_cultural_zones_for_building(building, x, y, orientation, longueur, largeur)
    
    def update_cultural_zones_for_building(self, building, x, y, orientation, longueur, largeur):
        """Met à jour les zones culturelles pour un nouveau bâtiment culturel"""
        if building.culture == 0:
            return
            
        center_x = x + longueur // 2
        center_y = y + largeur // 2
        
        for i in range(max(0, center_x - building.rayonnement), 
                     min(self.width, center_x + building.rayonnement + 1)):
            for j in range(max(0, center_y - building.rayonnement), 
                         min(self.height, center_y + building.rayonnement + 1)):
                distance = max(abs(i - center_x), abs(j - center_y))
                if distance <= building.rayonnement:
                    self.cultural_zones[j, i].add(building.id)
    
    def calculate_cultural_effect(self):
        """Calcule l'effet culturel de tous les bâtiments"""
        self.reset_cultural_zones()
        
        for building, x, y, orientation, longueur, largeur in self.cultural_buildings:
            self.update_cultural_zones_for_building(building, x, y, orientation, longueur, largeur)
        
        return self.cultural_zones
    
    def get_culture_for_position(self, x, y, longueur, largeur):
        """Calcule la culture reçue par un bâtiment à une position donnée"""
        affecting_cultural_ids = set()
        
        for i in range(longueur):
            for j in range(largeur):
                if 0 <= y+j < self.height and 0 <= x+i < self.width:
                    affecting_cultural_ids.update(self.cultural_zones[y + j, x + i])
        
        total_culture = 0
        for cultural_id in affecting_cultural_ids:
            for b, _, _, _, _, _ in self.cultural_buildings:
                if b.id == cultural_id:
                    total_culture += b.culture
                    break
        
        return total_culture
    
    def get_all_possible_positions(self, building):
        """Retourne toutes les positions possibles pour un bâtiment"""
        positions = []
        
        for orientation in ['H', 'V']:
            longueur, largeur = building.get_dimensions(orientation)
            
            if longueur > self.width or largeur > self.height:
                continue
                
            for y in range(self.height - largeur + 1):
                for x in range(self.width - longueur + 1):
                    can_place, _ = self.can_place(x, y, longueur, largeur)
                    if can_place:
                        positions.append((x, y, orientation))
        
        return positions
    
    def find_largest_free_area(self):
        """Trouve la plus grande zone continue de cases libres"""
        visited = np.zeros_like(self.grid, dtype=bool)
        largest_area = []
        
        for y in range(self.height):
            for x in range(self.width):
                if self.grid[y, x] == 1 and not self.occupied[y, x] and not visited[y, x]:
                    # BFS pour trouver la zone
                    area = []
                    queue = [(x, y)]
                    visited[y, x] = True
                    
                    while queue:
                        cx, cy = queue.pop(0)
                        area.append((cx, cy))
                        
                        for dx, dy in [(0,1), (1,0), (0,-1), (-1,0)]:
                            nx, ny = cx + dx, cy + dy
                            if (0 <= nx < self.width and 0 <= ny < self.height and 
                                self.grid[ny, nx] == 1 and not self.occupied[ny, nx] and 
                                not visited[ny, nx]):
                                visited[ny, nx] = True
                                queue.append((nx, ny))
                    
                    if len(area) > len(largest_area):
                        largest_area = area
        
        return largest_area
    
    def get_production_boosts(self):
        """Calcule les boosts de production pour tous les bâtiments producteurs"""
        self.calculate_cultural_effect()
        
        results = []
        total_culture_by_type = defaultdict(float)
        boost_counts = defaultdict(lambda: {0: 0, 25: 0, 50: 0, 100: 0})
        
        for building, x, y, orientation, longueur, largeur in self.buildings:
            if building.type == "producteur" and building.production:
                prod_type = building.production.strip()
                if not prod_type:
                    continue
                
                total_culture = self.get_culture_for_position(x, y, longueur, largeur)
                
                total_culture_by_type[prod_type] += total_culture
                
                boost = 0
                if building.boost_100 > 0 and total_culture >= building.boost_100:
                    boost = 100
                elif building.boost_50 > 0 and total_culture >= building.boost_50:
                    boost = 50
                elif building.boost_25 > 0 and total_culture >= building.boost_25:
                    boost = 25
                    
                boost_counts[prod_type][boost] += 1
                
                results.append({
                    "Nom": building.nom,
                    "Production": building.production,
                    "Culture reçue": round(total_culture, 2),
                    "Boost": f"{boost}%",
                    "Seuil 25%": building.boost_25,
                    "Seuil 50%": building.boost_50,
                    "Seuil 100%": building.boost_100
                })
        
        return results, dict(total_culture_by_type), dict(boost_counts)
    
    def copy(self):
        """Crée une copie du terrain pour simulation"""
        new_terrain = Terrain(self.grid.tolist())
        new_terrain.occupied = self.occupied.copy()
        new_terrain.buildings = self.buildings.copy()
        new_terrain.cultural_buildings = self.cultural_buildings.copy()
        new_terrain.building_objects = self.building_objects.copy()
        new_terrain.cultural_zones = np.empty_like(self.cultural_zones, dtype=object)
        for i in range(self.height):
            for j in range(self.width):
                new_terrain.cultural_zones[i, j] = self.cultural_zones[i, j].copy()
        return new_terrain

class BuildingPlacer:
    def __init__(self, terrain, buildings):
        self.terrain = terrain
        self.buildings = buildings
        self.placement_log = []
        
    def get_priority_score(self, building):
        """Calcule un score de priorité pour un bâtiment"""
        if building.type == "producteur" and building.production:
            prod = building.production.lower()
            if 'guerison' in prod or 'guérison' in prod:
                return 1
            elif 'nourriture' in prod:
                return 2
            elif 'or' in prod:
                return 3
            return 4
        elif building.type == "culturel":
            return 5
        return 6
    
    def evaluate_placement_score(self, terrain, producer_buildings):
        """Évalue la qualité d'un placement en fonction des boosts obtenus"""
        # Recalculer l'effet culturel
        terrain.calculate_cultural_effect()
        
        # Calculer les boosts pour chaque producteur
        score_by_priority = {1: 0, 2: 0, 3: 0, 4: 0}
        
        for building, x, y, orientation, longueur, largeur in terrain.buildings:
            if building.type == "producteur" and building.production:
                prod_type = building.production.lower()
                
                # Déterminer la priorité
                priority = 4
                if 'guerison' in prod_type:
                    priority = 1
                elif 'nourriture' in prod_type:
                    priority = 2
                elif 'or' in prod_type:
                    priority = 3
                
                # Calculer la culture reçue
                total_culture = terrain.get_culture_for_position(x, y, longueur, largeur)
                
                # Déterminer le boost
                boost = 0
                if building.boost_100 > 0 and total_culture >= building.boost_100:
                    boost = 100
                elif building.boost_50 > 0 and total_culture >= building.boost_50:
                    boost = 50
                elif building.boost_25 > 0 and total_culture >= building.boost_25:
                    boost = 25
                
                # Ajouter au score (les boosts 100% valent plus)
                score_by_priority[priority] += boost * 100
        
        # Score pondéré par priorité
        total_score = (score_by_priority[1] * 10000 + 
                      score_by_priority[2] * 1000 + 
                      score_by_priority[3] * 100 + 
                      score_by_priority[4])
        
        return total_score
    
    def find_optimal_placement_for_building(self, building, remaining_buildings, current_terrain):
        """Trouve la meilleure position pour un bâtiment en considérant les placements futurs"""
        best_score = -1
        best_position = None
        best_terrain = None
        
        positions = current_terrain.get_all_possible_positions(building)
        
        # Si pas de positions, retourner None
        if not positions:
            return None, None, -1
        
        # Pour chaque position possible
        for x, y, orientation in positions:
            # Créer une copie du terrain
            test_terrain = current_terrain.copy()
            
            # Placer le bâtiment
            test_building = Building(
                building.nom, building.longueur, building.largeur, 1,
                building.type, building.culture, building.rayonnement,
                building.boost_25, building.boost_50, building.boost_100,
                building.production
            )
            test_building.id = building.id
            test_terrain.place_building(test_building, x, y, orientation)
            
            # Si c'est un bâtiment culturel, mettre à jour les zones
            if test_building.type == "culturel":
                test_terrain.calculate_cultural_effect()
            
            # Évaluer ce placement
            score = self.evaluate_placement_score(test_terrain, remaining_buildings)
            
            if score > best_score:
                best_score = score
                best_position = (x, y, orientation)
                best_terrain = test_terrain
        
        return best_position, best_terrain, best_score
    
    def place_all_optimized(self):
        """Place tous les bâtiments en optimisant globalement"""
        
        # Filtrer les bâtiments avec quantité > 0
        valid_buildings = []
        for b in self.buildings:
            for i in range(b.quantite):
                new_b = Building(
                    b.nom, b.longueur, b.largeur, 1,
                    b.type, b.culture, b.rayonnement,
                    b.boost_25, b.boost_50, b.boost_100,
                    b.production
                )
                new_b.id = f"{b.nom}_{i}"
                valid_buildings.append(new_b)
        
        st.subheader("📊 Plan de placement optimisé")
        st.write(f"Total à placer: {len(valid_buildings)} bâtiments")
        
        # Séparer par type
        cultural_buildings = [b for b in valid_buildings if b.type == "culturel"]
        producer_buildings = [b for b in valid_buildings if b.type == "producteur" and b.production]
        other_buildings = [b for b in valid_buildings if b not in cultural_buildings and b not in producer_buildings]
        
        st.write(f"  - Culturels: {len(cultural_buildings)}")
        st.write(f"  - Producteurs: {len(producer_buildings)}")
        st.write(f"  - Autres: {len(other_buildings)}")
        
        # Trier les culturels par rayon décroissant
        cultural_buildings.sort(key=lambda b: (-b.rayonnement, -b.culture))
        
        # Trier les producteurs par priorité
        producer_buildings.sort(key=lambda b: (self.get_priority_score(b), -b.get_area()))
        
        # Combiner dans l'ordre de placement
        all_buildings = cultural_buildings + producer_buildings + other_buildings
        
        # Placement séquentiel avec optimisation locale
        current_terrain = self.terrain
        progress_bar = st.progress(0)
        
        for i, building in enumerate(all_buildings):
            st.write(f"\n📦 Placement de {building.nom} ({i+1}/{len(all_buildings)})")
            
            # Trouver la meilleure position
            remaining = all_buildings[i+1:]
            best_pos, best_terrain, score = self.find_optimal_placement_for_building(
                building, remaining, current_terrain
            )
            
            if best_pos:
                x, y, orientation = best_pos
                current_terrain.place_building(building, x, y, orientation)
                
                # Afficher la culture reçue si c'est un producteur
                if building.type == "producteur":
                    culture = current_terrain.get_culture_for_position(x, y, 
                        building.get_dimensions(orientation)[0], 
                        building.get_dimensions(orientation)[1])
                    st.write(f"  ✅ Placé à ({x},{y}) - reçoit {culture:.0f} culture")
                    
                    # Vérifier le boost
                    boost = 0
                    if building.boost_100 > 0 and culture >= building.boost_100:
                        boost = 100
                        st.write(f"  🎯 Boost 100% atteint!")
                    elif building.boost_50 > 0 and culture >= building.boost_50:
                        boost = 50
                        st.write(f"  🎯 Boost 50% atteint!")
                    elif building.boost_25 > 0 and culture >= building.boost_25:
                        boost = 25
                        st.write(f"  🎯 Boost 25% atteint!")
                else:
                    st.write(f"  ✅ Placé à ({x},{y})")
            else:
                st.write(f"  ❌ Impossible de placer {building.nom}")
                building.failed_attempts.append("Aucun emplacement disponible")
            
            progress_bar.progress((i + 1) / len(all_buildings))
        
        # Mettre à jour les quantités placées dans les bâtiments originaux
        for original_b in self.buildings:
            original_b.placed = 0
            for placed_b in all_buildings:
                if placed_b.nom == original_b.nom and not placed_b.failed_attempts:
                    original_b.placed += 1
                    original_b.positions.append(placed_b.positions[0] if placed_b.positions else None)
        
        return current_terrain
    
    def place_all(self):
        """Point d'entrée principal"""
        return self.place_all_optimized()

# Les fonctions normalize_column_name, read_input_file, create_buildings_from_df, 
# create_output_excel et main restent identiques aux versions précédentes