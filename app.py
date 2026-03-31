"""
Optimiseur de placement de bâtiments pour jeu de ville
Version améliorée - Optimisation des culturels autour des producteurs
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
        return self.boost_25, self.boost_50, self.boost_100

class BuildingInstance:
    """Classe représentant une instance placée d'un bâtiment"""
    def __init__(self, building, x, y, orientation='h'):
        self.building = building
        self.x = x
        self.y = y
        self.orientation = orientation
        self.culture_recue = 0
        self.boost_percentage = 0
        self.old_x = None
        self.old_y = None
        self.old_orientation = None
        
    def get_occupied_cells(self):
        cells = []
        l, L = self.building.longueur, self.building.largeur
        if l <= 0 or L <= 0:
            return cells
        if self.orientation == 'h':
            for i in range(l):
                for j in range(L):
                    cells.append((self.x + i, self.y + j))
        else:
            for i in range(L):
                for j in range(l):
                    cells.append((self.x + i, self.y + j))
        return cells
    
    def get_min_cell(self):
        return (self.x, self.y)
    
    def get_rayonnement_cells(self, terrain_size):
        if self.building.rayonnement <= 0:
            return []
        cells = []
        l, L = self.building.longueur, self.building.largeur
        height, width = terrain_size
        
        if self.orientation == 'h':
            for i in range(-self.building.rayonnement, l + self.building.rayonnement):
                for j in range(-self.building.rayonnement, L + self.building.rayonnement):
                    if i < 0 or i >= l or j < 0 or j >= L:
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
    
    def calculate_boost_from_culture(self, total_culture):
        thresholds = self.building.get_boost_thresholds()
        if total_culture >= thresholds[2] and thresholds[2] > 0:
            return 100
        elif total_culture >= thresholds[1] and thresholds[1] > 0:
            return 50
        elif total_culture >= thresholds[0] and thresholds[0] > 0:
            return 25
        return 0
    
    def set_culture(self, culture_value):
        self.culture_recue = culture_value
        self.boost_percentage = self.calculate_boost_from_culture(culture_value)
    
    def get_production(self):
        return self.building.quantite * (1 + self.boost_percentage / 100)

class Terrain:
    """Classe représentant le terrain de jeu"""
    def __init__(self, df_terrain):
        self.df = df_terrain
        self.grid, self.height, self.width = self._parse_terrain(df_terrain)
        self.building_instances = []
        
    def _parse_terrain(self, df):
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
        for x, y in cells:
            if x < 0 or x >= self.height or y < 0 or y >= self.width:
                return False
        return True
    
    def is_occupied(self, cells):
        for x, y in cells:
            cell_value = str(self.grid[x, y]).strip() if self.grid[x, y] != '' else ''
            if cell_value != '' and cell_value != ' ':
                return True
        return False
    
    def place_building(self, instance):
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
        cells = instance.get_occupied_cells()
        for x, y in cells:
            if str(self.grid[x, y]) == instance.building.nom:
                self.grid[x, y] = ''
        if instance in self.building_instances:
            self.building_instances.remove(instance)
    
    def calculate_culture_for_all(self):
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
    
    def get_total_production_by_type(self):
        """Calcule la production totale par type"""
        production_by_type = defaultdict(float)
        for building in self.building_instances:
            if building.building.type == 'Producteur':
                prod_type = building.building.production
                if prod_type and prod_type != '':
                    production_by_type[prod_type] += building.get_production()
        return production_by_type
    
    def get_production_stats(self):
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
        
        # Orientation verticale
        if l != L:
            for x in range(self.terrain.height - L + 1):
                for y in range(self.terrain.width - l + 1):
                    instance = BuildingInstance(building, x, y, 'v')
                    cells = instance.get_occupied_cells()
                    if cells and self.terrain.is_within_bounds(cells) and not self.terrain.is_occupied(cells):
                        positions.append(('v', x, y))
        
        return positions
    
    def calculate_culture_for_producteur_at_position(self, producteur_instance, cultural_instances, terrain_size):
        """Calcule la culture qu'un producteur recevrait à une position donnée"""
        total_culture = 0
        producteur_cells = set(producteur_instance.get_occupied_cells())
        
        for culturel in cultural_instances:
            rayonnement_cells = set(culturel.get_rayonnement_cells(terrain_size))
            if producteur_cells & rayonnement_cells:
                total_culture += culturel.building.culture
        
        return total_culture
    
    def evaluate_producteur_position(self, producteur_instance, cultural_instances, terrain_size):
        """Évalue une position pour un producteur (retourne la production)"""
        if producteur_instance.building.type == 'Producteur':
            total_culture = self.calculate_culture_for_producteur_at_position(
                producteur_instance, cultural_instances, terrain_size)
            
            boost_percentage = producteur_instance.calculate_boost_from_culture(total_culture)
            production = producteur_instance.building.quantite * (1 + boost_percentage / 100)
            
            # Priorité par type de production
            priority_map = {'Guérison': 10000, 'Nourriture': 1000, 'Or': 100}
            priority = priority_map.get(producteur_instance.building.production, 1)
            
            return production * priority + producteur_instance.building.priorite
        
        return 0
    
    def evaluate_cultural_position(self, cultural_instance, producteur_instances, terrain_size):
        """
        Évalue une position pour un bâtiment culturel
        Calcule la production totale générée par tous les producteurs
        """
        total_production = 0
        
        for producteur in producteur_instances:
            # Calculer la culture que ce producteur recevrait de ce culturel
            producteur_cells = set(producteur.get_occupied_cells())
            rayonnement_cells = set(cultural_instance.get_rayonnement_cells(terrain_size))
            
            additional_culture = cultural_instance.building.culture if (producteur_cells & rayonnement_cells) else 0
            new_total_culture = producteur.culture_recue + additional_culture
            
            # Calculer le nouveau boost
            new_boost = producteur.calculate_boost_from_culture(new_total_culture)
            old_boost = producteur.boost_percentage
            
            # Calculer le gain de production
            if new_boost > old_boost:
                gain = producteur.building.quantite * (new_boost - old_boost) / 100
                total_production += gain
        
        return total_production
    
    def optimize(self):
        """Exécute l'optimisation globale"""
        
        # Filtrer les bâtiments valides
        valid_buildings = [b for b in self.buildings if b.longueur > 0 and b.largeur > 0 and b.nombre > 0]
        
        cultural_buildings = [b for b in valid_buildings if b.type == 'Culturel']
        producteur_buildings = [b for b in valid_buildings if b.type == 'Producteur']
        
        # D'abord, placer tous les producteurs temporairement pour avoir une base
        # On va les placer dans des positions disponibles (sans optimisation)
        temp_producteurs = []
        for building in producteur_buildings:
            for _ in range(building.nombre):
                positions = self.get_all_positions(building)
                if positions:
                    orientation, x, y = positions[0]
                    instance = BuildingInstance(building, x, y, orientation)
                    if self.terrain.place_building(instance):
                        temp_producteurs.append(instance)
        
        # Calculer la culture initiale (sans culturels)
        self.terrain.calculate_culture_for_all()
        
        # Maintenant, placer les culturels de manière optimale
        for building in cultural_buildings:
            for idx in range(building.nombre):
                best_score = -1
                best_position = None
                best_instance = None
                
                positions = self.get_all_positions(building)
                
                for orientation, x, y in positions:
                    test_instance = BuildingInstance(building, x, y, orientation)
                    score = self.evaluate_cultural_position(test_instance, temp_producteurs, 
                                                            (self.terrain.height, self.terrain.width))
                    
                    if score > best_score:
                        best_score = score
                        best_position = (orientation, x, y)
                
                if best_position:
                    orientation, x, y = best_position
                    cultural_instance = BuildingInstance(building, x, y, orientation)
                    if self.terrain.place_building(cultural_instance):
                        # Mettre à jour la culture de tous les producteurs
                        self.terrain.calculate_culture_for_all()
        
        # Maintenant, optimiser le placement des producteurs avec les culturels en place
        # Sauvegarder les positions actuelles des producteurs
        current_producteurs = [p for p in self.terrain.building_instances if p.building.type == 'Producteur']
        
        # Réinitialiser le terrain pour les producteurs (garder les culturels)
        cultural_instances = [c for c in self.terrain.building_instances if c.building.type == 'Culturel']
        
        # Enlever tous les producteurs
        for producteur in current_producteurs:
            self.terrain.remove_building(producteur)
        
        # Placer les producteurs de manière optimale
        new_producteurs = []
        for building in producteur_buildings:
            for _ in range(building.nombre):
                best_score = -1
                best_position = None
                
                positions = self.get_all_positions(building)
                
                for orientation, x, y in positions:
                    test_instance = BuildingInstance(building, x, y, orientation)
                    score = self.evaluate_producteur_position(test_instance, cultural_instances, 
                                                              (self.terrain.height, self.terrain.width))
                    
                    if score > best_score:
                        best_score = score
                        best_position = (orientation, x, y)
                
                if best_position:
                    orientation, x, y = best_position
                    instance = BuildingInstance(building, x, y, orientation)
                    if self.terrain.place_building(instance):
                        new_producteurs.append(instance)
        
        # Enregistrer les déplacements
        for old, new in zip(current_producteurs, new_producteurs):
            if (old.x, old.y, old.orientation) != (new.x, new.y, new.orientation):
                self.moves.append({
                    'nom': new.building.nom,
                    'avant_x': old.x,
                    'avant_y': old.y,
                    'avant_orientation': 'Horizontal' if old.orientation == 'h' else 'Vertical',
                    'apres_x': new.x,
                    'apres_y': new.y,
                    'apres_orientation': 'Horizontal' if new.orientation == 'h' else 'Vertical'
                })
        
        # Recalculer la culture finale
        self.terrain.calculate_culture_for_all()

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
        """Crée la feuille du terrain avec cellules fusionnées et couleurs"""
        COLOR_CULTUREL = "FFB84D"
        COLOR_PRODUCTEUR = "85C77E"
        COLOR_NEUTRE = "D3D3D3"
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Premier passage: Créer toutes les cellules de base
        for i in range(terrain.height):
            for j in range(terrain.width):
                cell = ws.cell(row=i + 1, column=j + 1)
                cell.value = ""
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Créer un dictionnaire pour regrouper les bâtiments par leur première cellule
        building_first_cells = {}
        for instance in terrain.building_instances:
            cells = instance.get_occupied_cells()
            if cells:
                min_cell = min(cells, key=lambda c: (c[0], c[1]))
                building_first_cells[min_cell] = instance
        
        # Deuxième passage: Appliquer les fusions et le contenu
        for (i, j), instance in building_first_cells.items():
            l, L = instance.building.longueur, instance.building.largeur
            if instance.orientation == 'h':
                height, width = l, L
            else:
                height, width = L, l
            
            if height > 1 or width > 1:
                try:
                    ws.merge_cells(start_row=i + 1, start_column=j + 1,
                                  end_row=i + height, end_column=j + width)
                except Exception:
                    pass
            
            color = COLOR_CULTUREL if instance.building.type == 'Culturel' else \
                    COLOR_PRODUCTEUR if instance.building.type == 'Producteur' else COLOR_NEUTRE
            
            cell_text = instance.building.nom
            if instance.building.type == 'Producteur' and instance.boost_percentage > 0:
                cell_text += f"\nBoost: {instance.boost_percentage}%"
            elif instance.building.type == 'Culturel' and instance.building.culture > 0:
                cell_text += f"\nCulture: {instance.building.culture}"
            
            main_cell = ws.cell(row=i + 1, column=j + 1)
            main_cell.value = cell_text
            main_cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
            main_cell.font = Font(bold=True, size=10)
            main_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            main_cell.border = thin_border
        
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
    
    **Nouvelle stratégie d'optimisation :**
    1. 📍 Placement temporaire des producteurs pour évaluer l'impact des culturels
    2. 🎯 Optimisation du placement des culturels pour maximiser les boosts des producteurs
    3. 🔄 Re-optimisation du placement des producteurs autour des culturels
    4. 💰 Calcul précis des gains de production
    
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
                
                stats = terrain.get_production_stats()
                
                st.success("✅ Optimisation terminée!")
                
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
                4. Vérifiez les gains par type de production dans l'onglet **Gains_pertes**
                """)
                
        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
