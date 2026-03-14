import pandas as pd
import numpy as np
import streamlit as st
import io
import copy
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# --- CONFIGURATION ---
COLORS = {
    'Culturel': 'FFFFA500',  # Orange
    'Producteur': 'FF008000', # Vert
    'Neutre': 'FF808080'      # Gris
}

class CityPlanner:
    def __init__(self, terrain_data):
        """
        Initialise le planner avec le terrain de base
        """
        # Initialiser le journal et les paramètres AVANT tout
        self.journal = []
        self.max_entries = 10000 
        self.interrupted = False
        self.best_solution = None
        self.best_score = -1
        
        self.rows = len(terrain_data)
        self.cols = len(terrain_data[0])
        
        # Grille des cases libres (1 = libre, 0 = occupée par le terrain)
        self.base_grid = np.zeros((self.rows, self.cols))
        self.border_mask = np.zeros((self.rows, self.cols), dtype=bool)
        self.initial_free_cells = 0
        
        for r in range(self.rows):
            for c in range(self.cols):
                val = str(terrain_data[r][c]).strip().upper()
                if val == '1': 
                    self.base_grid[r, c] = 1
                    self.initial_free_cells += 1
                elif val == 'X':
                    self.base_grid[r, c] = 0 
                    self.border_mask[r, c] = True
        
        # Grille de travail (copie de la base)
        self.grid = self.base_grid.copy()
        
        # Bâtiments placés
        self.placed_buildings = []
        
        # Dictionnaire des dimensions possibles par nom
        self.possible_dimensions = {}

    def log(self, msg):
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)
        else:
            self.interrupted = True

    def set_building_dimensions(self, buildings_info):
        """
        Initialise le dictionnaire des dimensions possibles à partir de l'onglet Batiments
        """
        for b in buildings_info:
            nom = b['Nom']
            if nom not in self.possible_dimensions:
                self.possible_dimensions[nom] = set()
            # Ajouter les deux orientations possibles
            self.possible_dimensions[nom].add((b['Largeur'], b['Longueur']))
            self.possible_dimensions[nom].add((b['Longueur'], b['Largeur']))

    def load_existing_buildings(self, actuel_data):
        """
        Charge les bâtiments déjà placés depuis l'onglet "Actuel"
        Version qui détecte chaque bâtiment individuellement, même quand ils sont regroupés
        """
        # Créer une grille de visite pour marquer les cellules déjà traitées
        visited = np.zeros((self.rows, self.cols), dtype=bool)
        buildings_found = []
        
        # Convertir actuel_data en tableau de strings
        actuel_str = []
        for r in range(min(len(actuel_data), self.rows)):
            row = []
            for c in range(min(len(actuel_data[0]), self.cols)):
                val = str(actuel_data[r][c]).strip()
                row.append(val)
            actuel_str.append(row)
        
        # Afficher les premières cellules pour déboguer
        st.write("Aperçu des premières cellules de l'onglet Actuel:")
        sample = []
        for r in range(min(5, len(actuel_str))):
            for c in range(min(5, len(actuel_str[0]))):
                if actuel_str[r][c] not in ['', 'X', '1', '0', 'nan', 'None']:
                    sample.append(f"({r},{c}): '{actuel_str[r][c]}'")
        st.write(sample[:10])
        
        # Parcourir toutes les cellules
        for r in range(len(actuel_str)):
            for c in range(len(actuel_str[0])):
                if visited[r, c]:
                    continue
                    
                cell_value = actuel_str[r][c]
                
                # Ignorer les cases vides, les X et les 1/0
                if cell_value in ['', 'X', '1', '0', 'nan', 'None']:
                    continue
                
                # Chercher les dimensions possibles pour ce type de bâtiment
                if cell_value in self.possible_dimensions:
                    possible_dims = list(self.possible_dimensions[cell_value])
                    
                    # Trier par surface croissante pour prendre les plus petits d'abord
                    possible_dims.sort(key=lambda x: x[0] * x[1])
                    
                    # Prendre la plus petite dimension disponible
                    min_w, min_h = possible_dims[0]
                    
                    # Vérifier si on peut placer un bâtiment de cette dimension
                    if r + min_h <= self.rows and c + min_w <= self.cols:
                        valid = True
                        for dr in range(min_h):
                            for dc in range(min_w):
                                if r + dr >= len(actuel_str) or c + dc >= len(actuel_str[0]):
                                    valid = False
                                    break
                                if actuel_str[r + dr][c + dc] != cell_value:
                                    valid = False
                                    break
                            if not valid:
                                break
                        
                        if valid:
                            # Marquer les cellules comme visitées
                            for dr in range(min_h):
                                for dc in range(min_w):
                                    if r + dr < self.rows and c + dc < self.cols:
                                        visited[r + dr, c + dc] = True
                            
                            # Marquer la grille comme occupée
                            self.grid[r:r+min_h, c:c+min_w] = 0
                            
                            # Enregistrer le bâtiment
                            building_info = {
                                'nom_temp': cell_value,
                                'r': r,
                                'c': c,
                                'w': min_w,
                                'h': min_h,
                                'info': None
                            }
                            self.placed_buildings.append(building_info)
                            buildings_found.append(f"{cell_value} à ({r},{c}) dimensions {min_w}x{min_h}")
                            
                            self.log(f"Bâtiment existant trouvé: {cell_value} à ({r},{c}) dimensions {min_w}x{min_h}")
                            continue  # Passer à la cellule suivante
                    
                    # Si la plus petite dimension ne fonctionne pas, essayer les autres
                    found = False
                    for (w, h) in possible_dims:
                        if r + h <= self.rows and c + w <= self.cols:
                            valid = True
                            for dr in range(h):
                                for dc in range(w):
                                    if r + dr >= len(actuel_str) or c + dc >= len(actuel_str[0]):
                                        valid = False
                                        break
                                    if actuel_str[r + dr][c + dc] != cell_value:
                                        valid = False
                                        break
                                if not valid:
                                    break
                            
                            if valid:
                                # Marquer toutes les cellules comme visitées
                                for dr in range(h):
                                    for dc in range(w):
                                        if r + dr < self.rows and c + dc < self.cols:
                                            visited[r + dr, c + dc] = True
                                
                                # Marquer la grille comme occupée
                                self.grid[r:r+h, c:c+w] = 0
                                
                                # Enregistrer le bâtiment
                                building_info = {
                                    'nom_temp': cell_value,
                                    'r': r,
                                    'c': c,
                                    'w': w,
                                    'h': h,
                                    'info': None
                                }
                                self.placed_buildings.append(building_info)
                                buildings_found.append(f"{cell_value} à ({r},{c}) dimensions {w}x{h}")
                                
                                self.log(f"Bâtiment existant trouvé: {cell_value} à ({r},{c}) dimensions {w}x{h}")
                                found = True
                                break
                    
                    if not found:
                        # Si vraiment rien ne correspond, on prend la plus grande zone
                        st.warning(f"⚠️ Impossible de trouver des dimensions pour {cell_value} à ({r},{c})")
                        
                        # Chercher la plus grande zone rectangulaire
                        max_w = 1
                        for dc in range(1, self.cols - c):
                            if c + dc < len(actuel_str[0]) and actuel_str[r][c + dc] == cell_value and not visited[r, c + dc]:
                                max_w = dc + 1
                            else:
                                break
                        
                        max_h = 1
                        for dr in range(1, self.rows - r):
                            if r + dr < len(actuel_str) and actuel_str[r + dr][c] == cell_value and not visited[r + dr, c]:
                                max_h = dr + 1
                            else:
                                break
                        
                        # Prendre le plus grand rectangle valide
                        w, h = max_w, max_h
                        
                        # Marquer toutes les cellules comme visitées
                        for dr in range(h):
                            for dc in range(w):
                                if r + dr < self.rows and c + dc < self.cols:
                                    visited[r + dr, c + dc] = True
                        
                        # Marquer la grille comme occupée
                        self.grid[r:r+h, c:c+w] = 0
                        
                        # Enregistrer le bâtiment
                        building_info = {
                            'nom_temp': cell_value,
                            'r': r,
                            'c': c,
                            'w': w,
                            'h': h,
                            'info': None
                        }
                        self.placed_buildings.append(building_info)
                        buildings_found.append(f"{cell_value} à ({r},{c}) dimensions {w}x{h} (dimensions estimées)")
                        
                        self.log(f"Bâtiment existant trouvé (estimé): {cell_value} à ({r},{c}) dimensions {w}x{h}")
                else:
                    # Si le type de bâtiment n'est pas dans possible_dimensions, on le signale
                    st.warning(f"⚠️ Type de bâtiment inconnu: {cell_value} à ({r},{c})")
                    
                    # Chercher la plus grande zone rectangulaire
                    max_w = 1
                    for dc in range(1, self.cols - c):
                        if c + dc < len(actuel_str[0]) and actuel_str[r][c + dc] == cell_value and not visited[r, c + dc]:
                            max_w = dc + 1
                        else:
                            break
                    
                    max_h = 1
                    for dr in range(1, self.rows - r):
                        if r + dr < len(actuel_str) and actuel_str[r + dr][c] == cell_value and not visited[r + dr, c]:
                            max_h = dr + 1
                        else:
                            break
                    
                    # Prendre le plus grand rectangle valide
                    w, h = max_w, max_h
                    
                    # Marquer toutes les cellules comme visitées
                    for dr in range(h):
                        for dc in range(w):
                            if r + dr < self.rows and c + dc < self.cols:
                                visited[r + dr, c + dc] = True
                    
                    # Marquer la grille comme occupée
                    self.grid[r:r+h, c:c+w] = 0
                    
                    # Enregistrer le bâtiment
                    building_info = {
                        'nom_temp': cell_value,
                        'r': r,
                        'c': c,
                        'w': w,
                        'h': h,
                        'info': None
                    }
                    self.placed_buildings.append(building_info)
                    buildings_found.append(f"{cell_value} à ({r},{c}) dimensions {w}x{h} (type inconnu)")
                    
                    self.log(f"Bâtiment existant trouvé (type inconnu): {cell_value} à ({r},{c}) dimensions {w}x{h}")
        
        self.log(f"Chargement de {len(self.placed_buildings)} bâtiments depuis la configuration actuelle")
        
        # Afficher les bâtiments trouvés dans Streamlit
        st.write(f"🏗️ Bâtiments détectés: {len(buildings_found)}")
        with st.expander("Voir la liste des bâtiments détectés"):
            for b in buildings_found:
                st.text(b)
        
        return buildings_found

    # ... (le reste des méthodes reste inchangé)
    # is_adjacent_to_X, can_place, solve, try_placement, calculate_culture_for_position,
    # calculate_boost_from_culture, calculate_score_from_boost, calculate_culture_and_score,
    # optimize_placement, generate_excel