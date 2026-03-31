"""
Optimiseur de placement de bâtiments pour jeu de ville
Version avancée avec optimisation globale - CORRIGÉE
"""

import pandas as pd
import numpy as np
from collections import defaultdict
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

class Building:
    """Classe représentant un bâtiment"""
    def __init__(self, nom, longueur, largeur, nombre, type_bat, culture, rayonnement,
                 boost_25, boost_50, boost_100, production, quantite, priorite):
        self.nom = str(nom) if pd.notna(nom) else ""
        self.longueur = self._safe_int(longueur)
        self.largeur = self._safe_int(largeur)
        self.nombre = self._safe_int(nombre)
        self.type = str(type_bat) if pd.notna(type_bat) else ""
        self.culture = self._safe_int(culture) if pd.notna(culture) and str(culture) != "Rien" else 0
        self.rayonnement = self._safe_int(rayonnement)
        self.boost_25 = self._safe_int(boost_25)
        self.boost_50 = self._safe_int(boost_50)
        self.boost_100 = self._safe_int(boost_100)
        self.production = str(production) if pd.notna(production) and str(production) != "Rien" else ""
        self.quantite = self._safe_int(quantite)
        self.priorite = self._safe_int(priorite)
    
    def _safe_int(self, value):
        """Convertit une valeur en entier de façon sécurisée"""
        if pd.isna(value) or value == "" or value is None:
            return 0
        try:
            return int(float(value))
        except (ValueError, TypeError):
            return 0
    
    def get_dimensions(self):
        return self.longueur, self.largeur
    
    def get_culture(self):
        return self.culture
    
    def get_rayonnement(self):
        return self.rayonnement
    
    def get_production_type(self):
        return self.production
    
    def get_boost_thresholds(self):
        """Retourne les seuils de boost"""
        return self.boost_25, self.boost_50, self.boost_100

class BuildingInstance:
    """Classe représentant une instance placée d'un bâtiment"""
    def __init__(self, building, x, y, orientation='h'):
        self.building = building
        self.x = x
        self.y = y
        self.orientation = orientation  # 'h' horizontal, 'v' vertical
        self.culture_recue = 0
        self.boost_percentage = 0
        self.old_x = None
        self.old_y = None
        self.old_orientation = None
        
    def get_occupied_cells(self):
        """Retourne la liste des cases occupées par le bâtiment"""
        cells = []
        l, L = self.building.longueur, self.building.largeur
        if l <= 0 or L <= 0:
            return cells
        if self.orientation == 'h':
            for i in range(l):
                for j in range(L):
                    cells.append((self.x + i, self.y + j))
        else:  # vertical
            for i in range(L):
                for j in range(l):
                    cells.append((self.x + i, self.y + j))
        return cells
    
    def get_min_cell(self):
        """Retourne la cellule minimale (coin supérieur gauche)"""
        return (self.x, self.y)
    
    def get_rayonnement_cells(self, terrain_size):
        """Retourne les cases dans le rayonnement, limitées au terrain"""
        if self.building.rayonnement <= 0:
            return []
        cells = []
        l, L = self.building.longueur, self.building.largeur
        height, width = terrain_size
        
        if self.orientation == 'h':
            for i in range(-self.building.rayonnement, l + self.building.rayonnement):
                for j in range(-self.building.rayonnement, L + self.building.rayonnement):
                    if i < 0 or i >= l or j < 0 or j >= L:  # hors du bâtiment
                        x, y = self.x + i, self.y + j
                        if 0 <= x < height and 0 <= y < width:
                            cells.append((x, y))
        else:
            for i in range(-self.building.rayonnement, L + self.building.rayonnement):
                for j in range(-self.building.rayonnement, l + self.building.rayonnement):
                    if i < 0 or i >= L or j < 0 or j >= l:
                        x, y = self.x + i, self.y + j
                        if 0 <= x < height and 0 <= y < width:
                            cells.append((x, y))
        return cells
    
    def get_boost_score(self):
        """Calcule le score de boost (0-3)"""
        thresholds = self.building.get_boost_thresholds()
        if self.culture_recue >= thresholds[2] and thresholds[2] > 0:
            return 3
        elif self.culture_recue >= thresholds[1] and thresholds[1] > 0:
            return 2
        elif self.culture_recue >= thresholds[0] and thresholds[0] > 0:
            return 1
        return 0
    
    def set_culture(self, culture_value):
        self.culture_recue = culture_value
        self.boost_percentage = self.get_boost_score() * 25
    
    def get_production(self):
        """Calcule la production horaire"""
        return self.building.quantite * (1 + self.boost_percentage / 100)

class Terrain:
    """Classe représentant le terrain de jeu"""
    def __init__(self, df_terrain):
        self.df = df_terrain
        self.grid, self.height, self.width = self._parse_terrain(df_terrain)
        self.building_instances = []
        self.existing_buildings = {}  # Pour stocker les bâtiments déjà présents
        
    def _parse_terrain(self, df):
        """Parse le fichier terrain pour identifier les cases libres/occupées"""
        arr = df.fillna('').values
        x_rows, x_cols = [], []
        for i in range(arr.shape[0]):
            for j in range(arr.shape[1]):
                cell_value = str(arr[i, j]).strip().upper() if arr[i, j] != '' else ''
                if cell_value == 'X':
                    x_rows.append(i)
                    x_cols.append(j)
        
        if x_rows:
            min_row, max_row = min(x_rows), max(x_rows)
            min_col, max_col = min(x_cols), max(x_cols)
            interior = arr[min_row+1:max_row, min_col+1:max_col]
            return interior, interior.shape[0], interior.shape[1]
        return arr, arr.shape[0], arr.shape[1]
    
    def is_within_bounds(self, cells):
        """Vérifie si toutes les cases sont dans les limites du terrain"""
        for x, y in cells:
            if x < 0 or x >= self.height or y < 0 or y >= self.width:
                return False
        return True
    
    def is_occupied(self, cells):
        """Vérifie si les cases sont déjà occupées"""
        for x, y in cells:
            cell_value = str(self.grid[x, y]).strip() if self.grid[x, y] != '' else ''
            if cell_value != '' and cell_value != ' ':
                return True
        return False
    
    def place_building(self, instance):
        """Place un bâtiment sur le terrain"""
        cells = instance.get_occupied_cells()
        if not cells:
            return False
        if not self.is_within_bounds(cells):
            return False
        if self.is_occupied(cells):
            return False
        
        for x, y in cells:
            self.grid[x, y] = instance.building.nom
        self.building_instances.append(instance)
        return True
    
    def remove_building(self, instance):
        """Enlève un bâtiment du terrain"""
        cells = instance.get_occupied_cells()
        for x, y in cells:
            if str(self.grid[x, y]) == instance.building.nom:
                self.grid[x, y] = ''
        if instance in self.building_instances:
            self.building_instances.remove(instance)
    
    def calculate_culture(self):
        """Calcule la culture reçue par chaque bâtiment producteur"""
        cultural_buildings = [b for b in self.building_instances if b.building.type == 'Culturel']
        
        for producteur in self.building_instances:
            if producteur.building.type == 'Producteur':
                total_culture = 0
                producteur_cells = set(producteur.get_occupied_cells())
                
                for culturel in cultural_buildings:
                    rayonnement_cells = set(culturel.get_rayonnement_cells((self.height, self.width)))
                    if producteur_cells & rayonnement_cells:
                        total_culture += culturel.building.culture
                
                producteur.set_culture(total_culture)
    
    def get_production_stats(self):
        """Calcule les statistiques de production"""
        stats = defaultdict(lambda: {'total': 0, 'quantite': 0, 'boost_25': 0, 'boost_50': 0, 'boost_100': 0})
        
        for building in self.building_instances:
            if building.building.type == 'Producteur':
                prod_type = building.building.production
                if prod_type and prod_type != '':
                    production = building.get_production()
                    stats[prod_type]['total'] += production
                    stats[prod_type]['quantite'] += building.building.quantite
                    
                    if building.boost_percentage >= 100:
                        stats[prod_type]['boost_100'] += 1
                    elif building.boost_percentage >= 50:
                        stats[prod_type]['boost_50'] += 1
                    elif building.boost_percentage >= 25:
                        stats[prod_type]['boost_25'] += 1
        
        return stats

class AdvancedOptimizer:
    """Optimiseur avancé avec placement intelligent"""
    
    def __init__(self, terrain, buildings):
        self.terrain = terrain
        self.buildings = buildings
        self.moves = []
        
    def get_all_positions(self, building):
        """Génère toutes les positions possibles pour un bâtiment"""
        positions = []
        l, L = building.longueur, building.largeur
        
        if l <= 0 or L <= 0:
            return positions
        
        # Orientation horizontale
        for x in range(self.terrain.height - l + 1):
            for y in range(self.terrain.width - L + 1):
                instance = BuildingInstance(building, x, y, 'h')
                cells = instance.get_occupied_cells()
                if cells and self.terrain.is_within_bounds(cells) and not self.terrain.is_occupied(cells):
                    positions.append(('h', x, y))
        
        # Orientation verticale (si différent)
        if l != L:
            for x in range(self.terrain.height - L + 1):
                for y in range(self.terrain.width - l + 1):
                    instance = BuildingInstance(building, x, y, 'v')
                    cells = instance.get_occupied_cells()
                    if cells and self.terrain.is_within_bounds(cells) and not self.terrain.is_occupied(cells):
                        positions.append(('v', x, y))
        
        return positions
    
    def evaluate_cultural_position(self, cultural_instance, producteur_buildings, terrain_size):
        """
        Évalue une position pour un bâtiment culturel
        Calcule le score potentiel pour les producteurs
        """
        rayonnement_cells = set(cultural_instance.get_rayonnement_cells(terrain_size))
        score = 0
        
        # Pour chaque type de producteur, estimer le boost potentiel
        for prod_building in producteur_buildings:
            # Facteur de priorité pour ce type de production
            priority_map = {'Guérison': 100, 'Nourriture': 10, 'Or': 1}
            priority = priority_map.get(prod_building.production, 0)
            
            # Score basé sur la taille du rayonnement et la priorité
            score += len(rayonnement_cells) * (1 + priority / 100) * prod_building.quantite
        
        return score
    
    def evaluate_position_score(self, instance, cultural_instances, terrain_size):
        """Évalue la qualité d'une position pour un producteur"""
        if instance.building.type == 'Producteur':
            total_culture = 0
            producteur_cells = set(instance.get_occupied_cells())
            
            for culturel in cultural_instances:
                rayonnement_cells = set(culturel.get_rayonnement_cells(terrain_size))
                if producteur_cells & rayonnement_cells:
                    total_culture += culturel.building.culture
            
            thresholds = instance.building.get_boost_thresholds()
            if total_culture >= thresholds[2] and thresholds[2] > 0:
                boost_score = 100
            elif total_culture >= thresholds[1] and thresholds[1] > 0:
                boost_score = 50
            elif total_culture >= thresholds[0] and thresholds[0] > 0:
                boost_score = 25
            else:
                boost_score = 0
            
            # Priorité par type de production
            priority_map = {'Guérison': 10000, 'Nourriture': 1000, 'Or': 100}
            priority = priority_map.get(instance.building.production, 1)
            
            return boost_score * 1000 + priority * 10 + instance.building.priorite
        
        return 0
    
    def optimize(self):
        """Exécute l'optimisation globale"""
        
        # Filtrer les bâtiments valides
        valid_buildings = [b for b in self.buildings if b.longueur > 0 and b.largeur > 0 and b.nombre > 0]
        
        cultural_buildings = [b for b in valid_buildings if b.type == 'Culturel']
        producteur_buildings = [b for b in valid_buildings if b.type == 'Producteur']
        
        # Liste complète des producteurs à placer (répéter selon nombre)
        all_producteur_buildings = []
        for building in producteur_buildings:
            for _ in range(building.nombre):
                all_producteur_buildings.append(building)
        
        # Étape 1: Placer intelligemment les bâtiments culturels
        cultural_instances = []
        
        for building in cultural_buildings:
            for idx in range(building.nombre):
                best_score = -1
                best_position = None
                
                positions = self.get_all_positions(building)
                
                for orientation, x, y in positions:
                    instance = BuildingInstance(building, x, y, orientation)
                    score = self.evaluate_cultural_position(instance, all_producteur_buildings, 
                                                            (self.terrain.height, self.terrain.width))
                    
                    # Bonus pour être au centre (meilleure couverture)
                    center_x = self.terrain.height / 2
                    center_y = self.terrain.width / 2
                    building_center_x = x + building.longueur / 2
                    building_center_y = y + building.largeur / 2
                    distance_to_center = abs(building_center_x - center_x) + abs(building_center_y - center_y)
                    center_bonus = (self.terrain.height + self.terrain.width) / (distance_to_center + 1)
                    
                    score += center_bonus
                    
                    if score > best_score:
                        best_score = score
                        best_position = (orientation, x, y)
                
                if best_position:
                    orientation, x, y = best_position
                    instance = BuildingInstance(building, x, y, orientation)
                    if self.terrain.place_building(instance):
                        cultural_instances.append(instance)
        
        # Étape 2: Placer les producteurs en optimisant leur boost
        for building in producteur_buildings:
            for _ in range(building.nombre):
                best_score = -1
                best_position = None
                
                positions = self.get_all_positions(building)
                
                for orientation, x, y in positions:
                    instance = BuildingInstance(building, x, y, orientation)
                    score = self.evaluate_position_score(instance, cultural_instances, 
                                                         (self.terrain.height, self.terrain.width))
                    
                    if score > best_score:
                        best_score = score
                        best_position = (orientation, x, y)
                
                if best_position:
                    orientation, x, y = best_position
                    instance = BuildingInstance(building, x, y, orientation)
                    self.terrain.place_building(instance)
        
        # Étape 3: Amélioration itérative
        improved = True
        iterations = 0
        max_iterations = 100
        
        while improved and iterations < max_iterations:
            improved = False
            iterations += 1
            
            for idx, instance in enumerate(self.terrain.building_instances):
                if instance.building.type != 'Producteur':
                    continue
                
                old_x, old_y, old_orientation = instance.x, instance.y, instance.orientation
                old_score = self.evaluate_position_score(instance, cultural_instances, 
                                                         (self.terrain.height, self.terrain.width))
                
                self.terrain.remove_building(instance)
                
                best_score = -1
                best_position = None
                
                positions = self.get_all_positions(instance.building)
                for orientation, x, y in positions:
                    test_instance = BuildingInstance(instance.building, x, y, orientation)
                    score = self.evaluate_position_score(test_instance, cultural_instances, 
                                                         (self.terrain.height, self.terrain.width))
                    
                    if score > best_score:
                        best_score = score
                        best_position = (orientation, x, y)
                
                if best_position and best_score > old_score:
                    orientation, x, y = best_position
                    instance.x, instance.y, instance.orientation = x, y, orientation
                    if self.terrain.place_building(instance):
                        self.moves.append({
                            'nom': instance.building.nom,
                            'avant_x': old_x,
                            'avant_y': old_y,
                            'avant_orientation': 'Horizontal' if old_orientation == 'h' else 'Vertical',
                            'apres_x': x,
                            'apres_y': y,
                            'apres_orientation': 'Horizontal' if orientation == 'h' else 'Vertical'
                        })
                        improved = True
                else:
                    instance.x, instance.y, instance.orientation = old_x, old_y, old_orientation
                    self.terrain.place_building(instance)

class ExcelExporter:
    """Classe pour l'export Excel avec mise en forme"""
    
    @staticmethod
    def export_results(terrain, optimizer):
        output = io.BytesIO()
        wb = Workbook()
        wb.remove(wb.active)
        
        # 1. Bâtiments placés
        ws1 = wb.create_sheet("Batiments_places")
        buildings_data = []
        for instance in terrain.building_instances:
            buildings_data.append({
                'Nom': instance.building.nom,
                'Type': instance.building.type,
                'Production': instance.building.production if instance.building.production else 'N/A',
                'X': instance.x,
                'Y': instance.y,
                'Orientation': 'Horizontal' if instance.orientation == 'h' else 'Vertical',
                'Culture recue': instance.culture_recue,
                'Boost': f"{instance.boost_percentage}%",
                'Production/heure': f"{instance.get_production():.2f}"
            })
        
        if buildings_data:
            df = pd.DataFrame(buildings_data)
            for col_idx, col_name in enumerate(df.columns, 1):
                ws1.cell(row=1, column=col_idx, value=col_name)
                for row_idx, value in enumerate(df[col_name], 2):
                    ws1.cell(row=row_idx, column=col_idx, value=value)
        
        # 2. Statistiques
        ws2 = wb.create_sheet("Stats_production")
        stats = terrain.get_production_stats()
        stats_data = []
        for prod_type, data in stats.items():
            if prod_type and prod_type != '':
                stats_data.append({
                    'Type production': prod_type,
                    'Quantité totale produite/heure': round(data['total'], 2),
                    'Quantité de base': data['quantite'],
                    'Gain': round(data['total'] - data['quantite'], 2),
                    'Nb boost 25%': data['boost_25'],
                    'Nb boost 50%': data['boost_50'],
                    'Nb boost 100%': data['boost_100']
                })
        
        if stats_data:
            df = pd.DataFrame(stats_data)
            for col_idx, col_name in enumerate(df.columns, 1):
                ws2.cell(row=1, column=col_idx, value=col_name)
                for row_idx, value in enumerate(df[col_name], 2):
                    ws2.cell(row=row_idx, column=col_idx, value=value)
        
        # 3. Gains
        ws3 = wb.create_sheet("Gains_pertes")
        gains_data = []
        for prod_type, data in stats.items():
            if prod_type and prod_type != '':
                gain = data['total'] - data['quantite']
                percent_gain = (gain / data['quantite'] * 100) if data['quantite'] > 0 else 0
                gains_data.append({
                    'Type production': prod_type,
                    'Production avant': data['quantite'],
                    'Production après': round(data['total'], 2),
                    'Gain/Perte': round(gain, 2),
                    'Augmentation %': round(percent_gain, 1)
                })
        
        if gains_data:
            df = pd.DataFrame(gains_data)
            for col_idx, col_name in enumerate(df.columns, 1):
                ws3.cell(row=1, column=col_idx, value=col_name)
                for row_idx, value in enumerate(df[col_name], 2):
                    ws3.cell(row=row_idx, column=col_idx, value=value)
        
        # 4. Déplacements
        if optimizer.moves:
            ws4 = wb.create_sheet("Batiments_deplaces")
            df = pd.DataFrame(optimizer.moves)
            for col_idx, col_name in enumerate(df.columns, 1):
                ws4.cell(row=1, column=col_idx, value=col_name)
                for row_idx, value in enumerate(df[col_name], 2):
                    ws4.cell(row=row_idx, column=col_idx, value=value)
            
            # 5. Opérations
            ws5 = wb.create_sheet("Sequence_operations")
            operations = []
            for i, move in enumerate(optimizer.moves, 1):
                operations.append({
                    'Étape': i,
                    'Action': f"Déplacer {move['nom']}",
                    'De': f"({move['avant_x']}, {move['avant_y']})",
                    'Vers': f"({move['apres_x']}, {move['apres_y']})",
                    'Orientation': move['apres_orientation']
                })
            df = pd.DataFrame(operations)
            for col_idx, col_name in enumerate(df.columns, 1):
                ws5.cell(row=1, column=col_idx, value=col_name)
                for row_idx, value in enumerate(df[col_name], 2):
                    ws5.cell(row=row_idx, column=col_idx, value=value)
        
        # 6. Terrain final
        ws6 = wb.create_sheet("Terrain_final")
        ExcelExporter._create_terrain_sheet(ws6, terrain)
        
        # Ajuster les largeurs
        for ws in wb.worksheets:
            for column in ws.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 30)
                ws.column_dimensions[column_letter].width = adjusted_width
        
        wb.save(output)
        output.seek(0)
        return output.getvalue()
    
    @staticmethod
    def _create_terrain_sheet(ws: Worksheet, terrain: Terrain):
        COLOR_CULTUREL = "FFB84D"  # Orange
        COLOR_PRODUCTEUR = "85C77E"  # Vert
        COLOR_NEUTRE = "D3D3D3"  # Gris
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Créer un dictionnaire des bâtiments par cellule
        buildings_by_cell = {}
        for instance in terrain.building_instances:
            cells = instance.get_occupied_cells()
            for cell in cells:
                buildings_by_cell[cell] = instance
        
        # Parcourir toutes les cellules
        for i in range(terrain.height):
            for j in range(terrain.width):
                cell = ws.cell(row=i + 1, column=j + 1)
                
                if (i, j) in buildings_by_cell:
                    instance = buildings_by_cell[(i, j)]
                    # Vérifier si c'est la première cellule du bâtiment
                    first_cell = min(instance.get_occupied_cells())
                    if (i, j) == first_cell:
                        l, L = instance.building.longueur, instance.building.largeur
                        if instance.orientation == 'h':
                            height, width = l, L
                        else:
                            height, width = L, l
                        
                        if height > 1 or width > 1:
                            ws.merge_cells(start_row=i + 1, start_column=j + 1,
                                          end_row=i + height, end_column=j + width)
                        
                        color = COLOR_CULTUREL if instance.building.type == 'Culturel' else \
                                COLOR_PRODUCTEUR if instance.building.type == 'Producteur' else COLOR_NEUTRE
                        
                        cell_text = instance.building.nom
                        if instance.building.type == 'Producteur' and instance.boost_percentage > 0:
                            cell_text += f"\nBoost: {instance.boost_percentage}%"
                        elif instance.building.type == 'Culturel':
                            cell_text += f"\nCulture: {instance.building.culture}"
                        
                        cell.value = cell_text
                        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                        cell.font = Font(bold=True, size=10)
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                    else:
                        cell.value = ""
                else:
                    cell.value = ""
                
                cell.border = thin_border
                if not cell.alignment:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Ajuster les dimensions
        for i in range(1, terrain.height + 1):
            ws.row_dimensions[i].height = 45
        for j in range(1, terrain.width + 1):
            ws.column_dimensions[get_column_letter(j)].width = 15

class ExcelProcessor:
    @staticmethod
    def load_terrain(file):
        try:
            df = pd.read_excel(file, sheet_name='Terrain', header=None)
            return Terrain(df)
        except Exception as e:
            raise Exception(f"Erreur lors du chargement de l'onglet 'Terrain': {str(e)}")
    
    @staticmethod
    def load_buildings(file):
        try:
            df = pd.read_excel(file, sheet_name='Batiments')
            buildings = []
            
            required_columns = ['Nom', 'Longueur', 'Largeur', 'Nombre', 'Type', 'Culture', 
                               'Rayonnement', 'Boost 25%', 'Boost 50%', 'Boost 100%', 
                               'Production', 'Quantite', 'Priorite']
            
            missing_cols = [col for col in required_columns if col not in df.columns]
            if missing_cols:
                st.warning(f"Colonnes manquantes: {missing_cols}")
                # Continuer avec les colonnes disponibles
            
            for idx, row in df.iterrows():
                try:
                    building = Building(
                        row.get('Nom', ''), row.get('Longueur', 0), row.get('Largeur', 0),
                        row.get('Nombre', 0), row.get('Type', ''), row.get('Culture', 0),
                        row.get('Rayonnement', 0), row.get('Boost 25%', 0), row.get('Boost 50%', 0),
                        row.get('Boost 100%', 0), row.get('Production', ''), row.get('Quantite', 0),
                        row.get('Priorite', 0)
                    )
                    if building.longueur > 0 and building.largeur > 0 and building.nombre > 0:
                        buildings.append(building)
                except Exception as e:
                    st.warning(f"Ligne {idx + 2}: Impossible de charger - {str(e)}")
            
            return buildings
        except Exception as e:
            raise Exception(f"Erreur lors du chargement de l'onglet 'Batiments': {str(e)}")

def main():
    st.set_page_config(page_title="Optimiseur de placement de bâtiments", layout="wide")
    
    st.title("🏗️ Optimiseur avancé de placement de bâtiments")
    st.markdown("""
    ### Optimisation intelligente du placement
    
    **Améliorations :**
    - 🎯 Placement optimal des culturels au centre pour maximiser la couverture
    - 🔄 Amélioration itérative des positions
    - 📊 Priorisation des productions (Guérison > Nourriture > Or)
    
    **Légende :**
    - 🟠 **Orange** : Bâtiments culturels (avec valeur de culture)
    - 🟢 **Vert** : Bâtiments producteurs (avec % de boost)
    """)
    
    uploaded_file = st.file_uploader(
        "Choisissez votre fichier Excel",
        type=['xlsx', 'xls'],
        help="Le fichier doit contenir les onglets 'Terrain' et 'Batiments'"
    )
    
    if uploaded_file is not None:
        try:
            with st.spinner("Chargement des données..."):
                terrain = ExcelProcessor.load_terrain(uploaded_file)
                buildings = ExcelProcessor.load_buildings(uploaded_file)
                
                st.success(f"✅ Terrain: {terrain.height} x {terrain.width} cases")
                st.success(f"✅ {len(buildings)} types de bâtiments chargés")
                
                with st.expander("📋 Liste des bâtiments à placer"):
                    buildings_data = []
                    for b in buildings:
                        buildings_data.append({
                            'Nom': b.nom,
                            'Type': b.type,
                            'Dimensions': f"{b.longueur}x{b.largeur}",
                            'Nombre': b.nombre,
                            'Production': b.production if b.production else 'N/A',
                            'Quantité base': b.quantite,
                            'Culture': b.culture if b.culture > 0 else 'N/A',
                            'Rayonnement': b.rayonnement
                        })
                    st.dataframe(pd.DataFrame(buildings_data))
            
            with st.spinner("Optimisation avancée en cours..."):
                optimizer = AdvancedOptimizer(terrain, buildings)
                optimizer.optimize()
                terrain.calculate_culture()
                st.success("✅ Optimisation terminée!")
                
                stats = terrain.get_production_stats()
                
                if stats:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.subheader("📊 Résultats")
                        stats_df = pd.DataFrame([
                            {
                                'Type': k,
                                'Production/h': f"{v['total']:.2f}",
                                'Base': v['quantite'],
                                'Gain': f"{v['total'] - v['quantite']:.2f}",
                                'Boost 100%': v['boost_100']
                            }
                            for k, v in stats.items()
                        ])
                        st.dataframe(stats_df)
                    
                    with col2:
                        st.subheader("📈 Augmentation")
                        gains_df = pd.DataFrame([
                            {
                                'Type': k,
                                'Augmentation': f"{((v['total'] - v['quantite']) / v['quantite'] * 100):.1f}%" if v['quantite'] > 0 else "N/A",
                                'Nb bâtiments': len([b for b in terrain.building_instances if b.building.type == 'Producteur' and b.building.production == k])
                            }
                            for k, v in stats.items()
                        ])
                        st.dataframe(gains_df)
                
                with st.expander("🏠 Détail des bâtiments placés"):
                    buildings_placed = []
                    for instance in terrain.building_instances:
                        buildings_placed.append({
                            'Nom': instance.building.nom,
                            'Type': instance.building.type,
                            'Position': f"({instance.x}, {instance.y})",
                            'Orientation': 'Horizontal' if instance.orientation == 'h' else 'Vertical',
                            'Culture recue': instance.culture_recue,
                            'Boost': f"{instance.boost_percentage}%",
                            'Production/h': f"{instance.get_production():.2f}"
                        })
                    st.dataframe(pd.DataFrame(buildings_placed))
                
                st.subheader("📥 Télécharger les résultats")
                excel_data = ExcelExporter.export_results(terrain, optimizer)
                st.download_button(
                    label="📎 Télécharger le fichier Excel",
                    data=excel_data,
                    file_name="resultats_optimisation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.markdown("---")
                st.markdown("""
                ### 📝 Instructions
                1. Téléchargez le fichier Excel
                2. Consultez l'onglet **Terrain_final** pour voir le placement optimisé
                3. Utilisez l'onglet **Sequence_operations** pour les déplacements
                """)
                
        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
