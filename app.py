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
        self.buildings = []  # Liste des bâtiments placés (building, x, y, orientation, longueur, largeur)
        self.cultural_buildings = []  # Liste des bâtiments culturels avec leurs infos
        self.cultural_coverage = {}  # Dictionnaire pour stocker quels culturels couvrent quelles cases
        
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
        
        # Marquer les cases comme occupées
        for i in range(longueur):
            for j in range(largeur):
                self.occupied[y + j, x + i] = True
                
        building.placed += 1
        building.positions.append((x, y, orientation))
        self.buildings.append((building, x, y, orientation, longueur, largeur))
        
        # Si c'est un bâtiment culturel, mettre à jour la couverture
        if building.type == "culturel" and building.culture > 0:
            self.cultural_buildings.append({
                'building': building,
                'x': x,
                'y': y,
                'orientation': orientation,
                'longueur': longueur,
                'largeur': largeur,
                'center_x': x + longueur // 2,
                'center_y': y + largeur // 2,
                'rayonnement': building.rayonnement,
                'culture': building.culture,
                'id': building.id
            })
    
    def get_culture_for_building(self, building, x, y, orientation):
        """Calcule la culture reçue par un bâtiment à une position donnée
        Chaque bâtiment culturel ne compte qu'une seule fois, peu importe le nombre de cases du producteur"""
        
        longueur, largeur = building.get_dimensions(orientation)
        
        # Ensemble des IDs des culturels qui touchent ce bâtiment
        affecting_cultural_ids = set()
        
        # Pour chaque bâtiment culturel
        for cb in self.cultural_buildings:
            # Vérifier si une des cases du producteur est dans le rayon du culturel
            for i in range(longueur):
                for j in range(largeur):
                    px, py = x + i, y + j
                    # Calculer la distance du centre du culturel à cette case
                    distance = max(abs(px - cb['center_x']), abs(py - cb['center_y']))
                    if distance <= cb['rayonnement']:
                        affecting_cultural_ids.add(cb['id'])
                        break  # Une case suffit, pas besoin de vérifier les autres
                if cb['id'] in affecting_cultural_ids:
                    break
        
        # Calculer la culture totale (chaque culturel ne compte qu'une fois)
        total_culture = 0
        for cultural_id in affecting_cultural_ids:
            for cb in self.cultural_buildings:
                if cb['id'] == cultural_id:
                    total_culture += cb['culture']
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
    
    def get_production_boosts(self):
        """Calcule les boosts de production pour tous les bâtiments producteurs"""
        
        results = []
        total_culture_by_type = defaultdict(float)
        boost_counts = defaultdict(lambda: {0: 0, 25: 0, 50: 0, 100: 0})
        
        for building, x, y, orientation, longueur, largeur in self.buildings:
            if building.type == "producteur" and building.production:
                prod_type = building.production.strip()
                if not prod_type:
                    continue
                
                # Calculer la culture reçue (chaque culturel ne compte qu'une fois)
                total_culture = self.get_culture_for_building(building, x, y, orientation)
                
                total_culture_by_type[prod_type] += total_culture
                
                # Déterminer le boost
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

class BuildingPlacer:
    def __init__(self, terrain, buildings):
        self.terrain = terrain
        self.buildings = buildings
        
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
            # Les culturels avec grand rayon en premier
            return 10 - building.rayonnement
        return 100
    
    def evaluate_position_score(self, terrain, building, x, y, orientation):
        """Évalue le score d'une position pour un bâtiment"""
        if building.type == "producteur" and building.production:
            # Pour un producteur, on veut maximiser la culture reçue
            culture = terrain.get_culture_for_building(building, x, y, orientation)
            
            # Bonus important si on atteint un seuil
            if building.boost_100 > 0 and culture >= building.boost_100:
                return 10000 + culture
            elif building.boost_50 > 0 and culture >= building.boost_50:
                return 5000 + culture
            elif building.boost_25 > 0 and culture >= building.boost_25:
                return 2500 + culture
            else:
                return culture
        
        elif building.type == "culturel":
            # Pour un culturel, on veut maximiser la couverture future
            longueur, largeur = building.get_dimensions(orientation)
            center_x = x + longueur // 2
            center_y = y + largeur // 2
            
            free_cases_in_radius = 0
            for i in range(max(0, center_x - building.rayonnement), 
                         min(terrain.width, center_x + building.rayonnement + 1)):
                for j in range(max(0, center_y - building.rayonnement), 
                             min(terrain.height, center_y + building.rayonnement + 1)):
                    if terrain.grid[j, i] == 1 and not terrain.occupied[j, i]:
                        free_cases_in_radius += 1
            
            return free_cases_in_radius * 100
        
        return 0
    
    def find_best_position(self, terrain, building):
        """Trouve la meilleure position pour un bâtiment"""
        best_score = -1
        best_position = None
        
        positions = terrain.get_all_possible_positions(building)
        
        for x, y, orientation in positions:
            score = self.evaluate_position_score(terrain, building, x, y, orientation)
            if score > best_score:
                best_score = score
                best_position = (x, y, orientation)
        
        return best_position
    
    def place_all(self):
        """Place tous les bâtiments en optimisant"""
        
        # Créer une liste de tous les bâtiments à placer (un par exemplaire)
        all_buildings = []
        building_counter = defaultdict(int)
        
        for b in self.buildings:
            if b.quantite > 0:
                for i in range(b.quantite):
                    # Créer une copie du bâtiment
                    new_b = Building(
                        b.nom, b.longueur, b.largeur, 1,
                        b.type, b.culture, b.rayonnement,
                        b.boost_25, b.boost_50, b.boost_100,
                        b.production
                    )
                    new_b.id = f"{b.nom}_{i}"
                    building_counter[b.nom] += 1
                    all_buildings.append(new_b)
        
        st.subheader("📊 Plan de placement")
        st.write(f"Total à placer: {len(all_buildings)} bâtiments")
        
        # Séparer par type
        cultural_buildings = [b for b in all_buildings if b.type == "culturel"]
        producer_buildings = [b for b in all_buildings if b.type == "producteur" and b.production]
        other_buildings = [b for b in all_buildings if b not in cultural_buildings and b not in producer_buildings]
        
        st.write(f"  - Culturels: {len(cultural_buildings)}")
        st.write(f"  - Producteurs: {len(producer_buildings)}")
        st.write(f"  - Autres: {len(other_buildings)}")
        
        # Trier les culturels par rayon décroissant
        cultural_buildings.sort(key=lambda b: (-b.rayonnement, -b.culture))
        
        # Trier les producteurs par priorité
        producer_buildings.sort(key=lambda b: (self.get_priority_score(b), -b.get_area()))
        
        # Ordre de placement: d'abord les culturels, puis les producteurs, puis les autres
        ordered_buildings = cultural_buildings + producer_buildings + other_buildings
        
        # Placer chaque bâtiment
        current_terrain = self.terrain
        progress_bar = st.progress(0)
        
        for i, building in enumerate(ordered_buildings):
            best_pos = self.find_best_position(current_terrain, building)
            
            if best_pos:
                x, y, orientation = best_pos
                current_terrain.place_building(building, x, y, orientation)
                
                if building.type == "producteur":
                    # Calculer la culture qu'il reçoit
                    culture = current_terrain.get_culture_for_building(building, x, y, orientation)
                    st.write(f"  ✅ {building.nom} placé à ({x},{y}) - reçoit {culture:.0f} culture")
                else:
                    st.write(f"  ✅ {building.nom} placé à ({x},{y})")
            else:
                st.write(f"  ❌ Impossible de placer {building.nom}")
                building.failed_attempts.append("Aucun emplacement disponible")
            
            progress_bar.progress((i + 1) / len(ordered_buildings))
        
        # Mettre à jour les quantités placées dans les bâtiments originaux
        for original_b in self.buildings:
            original_b.placed = 0
            original_b.positions = []
            for placed_b in ordered_buildings:
                if placed_b.nom == original_b.nom and not placed_b.failed_attempts:
                    original_b.placed += 1
                    if placed_b.positions:
                        original_b.positions.append(placed_b.positions[0])
        
        return current_terrain

# Les fonctions normalize_column_name, read_input_file, create_buildings_from_df,
# create_output_excel et main restent identiques aux versions précédentes
# (je les ai incluses dans le script complet mais je les omets ici pour la lisibilité)