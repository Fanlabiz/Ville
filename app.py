"""
Placement optimal des bâtiments - Version corrigée
Respecte le nombre d'exemplaires, fusions de cellules, entrelacement
"""

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import streamlit as st
import io
from typing import List, Tuple, Dict, Optional
from collections import defaultdict, Counter

class Building:
    def __init__(self, name: str, longueur: int, largeur: int, nombre: int,
                 building_type: str, culture: int, rayonnement: int,
                 boost_25: int, boost_50: int, boost_100: int,
                 production: str, quantite: int):
        self.name = name
        self.longueur = longueur
        self.largeur = largeur
        self.nombre_original = nombre
        self.nombre = nombre  # Nombre restant à placer
        self.building_type = building_type
        self.culture = culture
        self.rayonnement = rayonnement
        self.boost_25 = boost_25
        self.boost_50 = boost_50
        self.boost_100 = boost_100
        self.production = production
        self.quantite = quantite
        self.placed = []  # Liste des (row, col, orientation)
        
    def get_dimensions(self, orientation: str = 'H') -> Tuple[int, int]:
        if orientation == 'H':
            return self.longueur, self.largeur
        return self.largeur, self.longueur
        
    def get_cells(self, row: int, col: int, orientation: str = 'H') -> List[Tuple[int, int]]:
        h, w = self.get_dimensions(orientation)
        return [(row + i, col + j) for i in range(h) for j in range(w)]
        
    def get_zone_rayonnement(self, row: int, col: int, orientation: str = 'H') -> List[Tuple[int, int]]:
        if self.rayonnement == 0:
            return []
        h, w = self.get_dimensions(orientation)
        cells = set()
        for i in range(-self.rayonnement, h + self.rayonnement):
            for j in range(-self.rayonnement, w + self.rayonnement):
                if i < 0 or i >= h or j < 0 or j >= w:
                    cells.add((row + i, col + j))
        return list(cells)

class Terrain:
    def __init__(self, grid: np.ndarray):
        self.grid = grid.copy()  # 0 = libre, 1 = occupé
        self.rows, self.cols = grid.shape
        self.buildings = []  # Liste des (building, row, col, orientation)
        
    def can_place(self, building: Building, row: int, col: int, orientation: str) -> bool:
        h, w = building.get_dimensions(orientation)
        if row < 0 or col < 0 or row + h > self.rows or col + w > self.cols:
            return False
        for i in range(h):
            for j in range(w):
                if self.grid[row + i, col + j] != 0:
                    return False
        return True
        
    def place(self, building: Building, row: int, col: int, orientation: str):
        h, w = building.get_dimensions(orientation)
        for i in range(h):
            for j in range(w):
                self.grid[row + i, col + j] = 1
        building.placed.append((row, col, orientation))
        self.buildings.append((building, row, col, orientation))
        
    def get_building_at(self, row: int, col: int) -> Optional[Tuple[Building, int, int, str]]:
        """Retourne le bâtiment qui occupe la case (row, col)"""
        for b, r, c, o in self.buildings:
            h, w = b.get_dimensions(o)
            if r <= row < r + h and c <= col < c + w:
                return (b, r, c, o)
        return None

def parse_ville_excel(file) -> Tuple[Terrain, List[Building]]:
    """Parse spécifique pour Ville.xlsx"""
    xl = pd.ExcelFile(file)
    
    # Lire le terrain (onglet Terrain)
    df_terrain = pd.read_excel(file, sheet_name='Terrain', header=None)
    grid = df_terrain.values
    
    terrain_grid = np.zeros(grid.shape, dtype=int)
    for i in range(grid.shape[0]):
        for j in range(grid.shape[1]):
            val = str(grid[i, j]) if pd.notna(grid[i, j]) else ''
            if val == 'X' or val == 'x':
                terrain_grid[i, j] = 1
            else:
                terrain_grid[i, j] = 0
    
    terrain = Terrain(terrain_grid)
    
    # Lire les bâtiments (onglet Batiments)
    df_buildings = pd.read_excel(file, sheet_name='Batiments')
    buildings = []
    
    for _, row in df_buildings.iterrows():
        quantite = row['Quantite']
        if isinstance(quantite, str) and quantite.startswith('='):
            try:
                quantite = eval(quantite[1:])
            except:
                quantite = 0
        elif pd.isna(quantite):
            quantite = 0
            
        building = Building(
            name=str(row['Nom']),
            longueur=int(row['Longueur']),
            largeur=int(row['Largeur']),
            nombre=int(row['Nombre']),
            building_type=str(row['Type']),
            culture=int(row['Culture']) if pd.notna(row['Culture']) else 0,
            rayonnement=int(row['Rayonnement']) if pd.notna(row['Rayonnement']) else 0,
            boost_25=int(row['Boost 25%']) if pd.notna(row['Boost 25%']) else 0,
            boost_50=int(row['Boost 50%']) if pd.notna(row['Boost 50%']) else 0,
            boost_100=int(row['Boost 100%']) if pd.notna(row['Boost 100%']) else 0,
            production=str(row['Production']),
            quantite=int(quantite) if quantite != 0 else 0
        )
        buildings.append(building)
    
    return terrain, buildings

def calculate_boost(culture_total: int, building: Building) -> float:
    if culture_total >= building.boost_100:
        return 2.0
    elif culture_total >= building.boost_50:
        return 1.5
    elif culture_total >= building.boost_25:
        return 1.25
    return 1.0

def get_production_score(production: str) -> int:
    scores = {'Guerison': 1000, 'Nourriture': 100, 'Or': 10}
    return scores.get(production, 1)

def find_best_spot_border(terrain: Terrain, building: Building) -> Optional[Tuple[int, int, str]]:
    """Place sur les bords de préférence"""
    best_score = -1
    best_pos = None
    
    for orientation in ['H', 'V']:
        h, w = building.get_dimensions(orientation)
        for row in range(terrain.rows - h + 1):
            for col in range(terrain.cols - w + 1):
                if terrain.can_place(building, row, col, orientation):
                    score = 0
                    if row == 0 or row + h == terrain.rows:
                        score += 1000
                    if col == 0 or col + w == terrain.cols:
                        score += 1000
                    
                    if score > best_score:
                        best_score = score
                        best_pos = (row, col, orientation)
    
    return best_pos

def find_best_spot_for_producer(terrain: Terrain, building: Building, 
                                 culture_map: np.ndarray) -> Optional[Tuple[int, int, str]]:
    """Trouve la meilleure position pour un producteur"""
    best_score = -1
    best_pos = None
    
    for orientation in ['H', 'V']:
        h, w = building.get_dimensions(orientation)
        for row in range(terrain.rows - h + 1):
            for col in range(terrain.cols - w + 1):
                if terrain.can_place(building, row, col, orientation):
                    culture_recue = 0
                    for i in range(h):
                        for j in range(w):
                            if 0 <= row + i < terrain.rows and 0 <= col + j < terrain.cols:
                                culture_recue += culture_map[row + i, col + j]
                    
                    boost = calculate_boost(culture_recue, building)
                    prod_value = building.quantite * boost
                    priority = get_production_score(building.production)
                    score = prod_value * priority * (1 + culture_recue / 10000)
                    
                    if score > best_score:
                        best_score = score
                        best_pos = (row, col, orientation)
    
    return best_pos

def find_best_spot_for_culturel(terrain: Terrain, building: Building,
                                  culture_map: np.ndarray,
                                  producers: List[Building]) -> Optional[Tuple[int, int, str]]:
    """Trouve la meilleure position pour un culturel"""
    best_score = -1
    best_pos = None
    
    for orientation in ['H', 'V']:
        h, w = building.get_dimensions(orientation)
        for row in range(terrain.rows - h + 1):
            for col in range(terrain.cols - w + 1):
                if terrain.can_place(building, row, col, orientation):
                    gain = 0
                    zone = building.get_zone_rayonnement(row, col, orientation)
                    
                    for r, c in zone:
                        if 0 <= r < terrain.rows and 0 <= c < terrain.cols:
                            building_at = terrain.get_building_at(r, c)
                            if building_at:
                                prod, pr, pc, po = building_at
                                if prod.building_type == 'Producteur':
                                    old_culture = culture_map[r, c]
                                    new_culture = old_culture + building.culture
                                    old_boost = calculate_boost(old_culture, prod)
                                    new_boost = calculate_boost(new_culture, prod)
                                    gain += prod.quantite * (new_boost - old_boost)
                    
                    if gain > best_score:
                        best_score = gain
                        best_pos = (row, col, orientation)
    
    return best_pos

def place_buildings_optimized(terrain: Terrain, buildings: List[Building]) -> Tuple[List[Building], List[Building]]:
    """Place tous les bâtiments selon la stratégie définie"""
    
    # Séparer par type et créer des listes d'instances uniques
    neutres = []
    producteurs = []
    culturels = []
    
    for b in buildings:
        # Créer des copies individuelles pour chaque exemplaire
        for _ in range(b.nombre):
            copy = Building(
                b.name, b.longueur, b.largeur, 1,
                b.building_type, b.culture, b.rayonnement,
                b.boost_25, b.boost_50, b.boost_100,
                b.production, b.quantite
            )
            if b.building_type == 'Neutre':
                neutres.append(copy)
            elif b.building_type == 'Producteur':
                producteurs.append(copy)
            elif b.building_type == 'Culturel':
                culturels.append(copy)
    
    # Trier par taille décroissante
    neutres.sort(key=lambda b: b.longueur * b.largeur, reverse=True)
    producteurs.sort(key=lambda b: (get_production_score(b.production), b.longueur * b.largeur), reverse=True)
    culturels.sort(key=lambda b: b.longueur * b.largeur, reverse=True)
    
    placed = []
    not_placed = []
    
    # Étape 1: Placer les neutres sur les bords
    for building in neutres:
        pos = find_best_spot_border(terrain, building)
        if pos:
            terrain.place(building, pos[0], pos[1], pos[2])
            placed.append(building)
        else:
            not_placed.append(building)
    
    # Initialiser la carte des boosts
    culture_map = np.zeros((terrain.rows, terrain.cols), dtype=int)
    
    # Mettre à jour avec les culturels déjà placés (dans les neutres)
    for b, row, col, orientation in terrain.buildings:
        if b.building_type == 'Culturel':
            zone = b.get_zone_rayonnement(row, col, orientation)
            for r, c in zone:
                if 0 <= r < terrain.rows and 0 <= c < terrain.cols:
                    culture_map[r, c] += b.culture
    
    # Étape 2: Placer les producteurs par priorité
    producers_list = []
    for building in producteurs:
        pos = find_best_spot_for_producer(terrain, building, culture_map)
        if pos:
            terrain.place(building, pos[0], pos[1], pos[2])
            placed.append(building)
            producers_list.append(building)
        else:
            not_placed.append(building)
    
    # Étape 3: Placer les culturels
    for building in culturels:
        pos = find_best_spot_for_culturel(terrain, building, culture_map, producers_list)
        if pos:
            terrain.place(building, pos[0], pos[1], pos[2])
            placed.append(building)
            # Mettre à jour la carte des boosts
            zone = building.get_zone_rayonnement(pos[0], pos[1], pos[2])
            for r, c in zone:
                if 0 <= r < terrain.rows and 0 <= c < terrain.cols:
                    culture_map[r, c] += building.culture
        else:
            not_placed.append(building)
    
    # Étape 4: Essayer de replacer les producteurs non placés
    remaining_producers = [b for b in not_placed if b.building_type == 'Producteur']
    for building in remaining_producers:
        pos = find_best_spot_for_producer(terrain, building, culture_map)
        if pos:
            terrain.place(building, pos[0], pos[1], pos[2])
            placed.append(building)
            not_placed.remove(building)
    
    return placed, not_placed

def generate_output_excel(terrain: Terrain, placed: List[Building], 
                          not_placed: List[Building]) -> io.BytesIO:
    """Génère le fichier Excel de résultat avec fusions de cellules"""
    output = io.BytesIO()
    
    # Recalculer la carte des boosts
    culture_map = np.zeros((terrain.rows, terrain.cols), dtype=int)
    for b, row, col, orientation in terrain.buildings:
        if b.building_type == 'Culturel':
            zone = b.get_zone_rayonnement(row, col, orientation)
            for r, c in zone:
                if 0 <= r < terrain.rows and 0 <= c < terrain.cols:
                    culture_map[r, c] += b.culture
    
    # 1. Liste des bâtiments placés
    placed_data = []
    for building in placed:
        for idx, (row, col, orientation) in enumerate(building.placed):
            h, w = building.get_dimensions(orientation)
            culture_recue = 0
            for i in range(h):
                for j in range(w):
                    if 0 <= row + i < terrain.rows and 0 <= col + j < terrain.cols:
                        culture_recue += culture_map[row + i, col + j]
            
            boost = calculate_boost(culture_recue, building)
            prod_h = building.quantite * boost
            
            placed_data.append({
                'Nom': building.name,
                'Type': building.building_type,
                'Production': building.production,
                'Ligne': row + 1,
                'Colonne': col + 1,
                'Hauteur': h,
                'Largeur': w,
                'Orientation': 'H' if orientation == 'H' else 'V',
                'Culture recue': culture_recue,
                'Boost (%)': int((boost - 1) * 100),
                'Quantite/h': building.quantite,
                'Prod totale/h': prod_h
            })
    
    # 2. Synthèse par type de production
    df_placed = pd.DataFrame(placed_data)
    synthesis = []
    
    for prod_type in ['Guerison', 'Nourriture', 'Or']:
        prod_df = df_placed[df_placed['Production'] == prod_type]
        if len(prod_df) > 0:
            total_quantite = prod_df['Quantite/h'].sum()
            total_prod = prod_df['Prod totale/h'].sum()
            boost_moyen = (total_prod / total_quantite - 1) * 100 if total_quantite > 0 else 0
            synthesis.append({
                'Production': prod_type,
                'Culture totale': int(prod_df['Culture recue'].sum()),
                'Boost moyen (%)': round(boost_moyen, 1),
                'Nb batiments': len(prod_df),
                'Production/h': total_prod
            })
    
    # Ajouter les autres productions
    autres = df_placed[~df_placed['Production'].isin(['Guerison', 'Nourriture', 'Or', 'Rien'])]
    if len(autres) > 0:
        for prod in autres['Production'].unique():
            prod_df = autres[autres['Production'] == prod]
            synthesis.append({
                'Production': prod,
                'Culture totale': int(prod_df['Culture recue'].sum()),
                'Boost moyen (%)': 0,
                'Nb batiments': len(prod_df),
                'Production/h': prod_df['Prod totale/h'].sum()
            })
    
    # 3. Statistiques
    total_cells = terrain.rows * terrain.cols
    used_cells = np.sum(terrain.grid)
    free_cells = total_cells - used_cells
    unplaced_cells = sum(b.longueur * b.largeur for b in not_placed)
    
    # Écrire le fichier Excel
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_placed.to_excel(writer, sheet_name='Batiments places', index=False)
        
        if synthesis:
            pd.DataFrame(synthesis).to_excel(writer, sheet_name='Synthese', index=False)
            
            stats_df = pd.DataFrame({
                'Statistiques': ['Cases libres restantes', 'Cases des batiments non places', 'Nombre de batiments non places'],
                'Valeur': [free_cells, unplaced_cells, len(not_placed)]
            })
            stats_df.to_excel(writer, sheet_name='Synthese', startrow=len(synthesis) + 3, index=False)
        
        # Terrain final avec fusions de cellules
        # Créer une feuille pour la visualisation
        workbook = writer.book
        ws = workbook.create_sheet('Terrain final')
        
        # Couleurs
        orange_fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
        green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
        grey_fill = PatternFill(start_color='808080', end_color='808080', fill_type='solid')
        border = Border(left=Side(style='thin'), right=Side(style='thin'),
                        top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Écrire les en-têtes de colonnes (numéros)
        for c in range(terrain.cols):
            cell = ws.cell(row=1, column=c+2, value=c+1)
            cell.fill = PatternFill(start_color='DDDDDD', end_color='DDDDDD', fill_type='solid')
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
            cell.border = border
        
        # Écrire les en-têtes de lignes (numéros)
        for r in range(terrain.rows):
            cell = ws.cell(row=r+2, column=1, value=r+1)
            cell.fill = PatternFill(start_color='DDDDDD', end_color='DDDDDD', fill_type='solid')
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
            cell.border = border
        
        # Placer les bâtiments avec fusions
        for building, row, col, orientation in terrain.buildings:
            h, w = building.get_dimensions(orientation)
            
            # Calculer le nom à afficher
            display_name = building.name
            if building.building_type == 'Producteur':
                culture_recue = 0
                for i in range(h):
                    for j in range(w):
                        if 0 <= row + i < terrain.rows and 0 <= col + j < terrain.cols:
                            culture_recue += culture_map[row + i, col + j]
                boost = calculate_boost(culture_recue, building)
                if boost > 1:
                    display_name = f"{building.name} +{int((boost-1)*100)}%"
            
            # Fusionner les cellules
            start_row = row + 2
            start_col = col + 2
            end_row = row + h + 1
            end_col = col + w + 1
            
            ws.merge_cells(start_row=start_row, start_column=start_col,
                          end_row=end_row, end_column=end_col)
            
            cell = ws.cell(row=start_row, column=start_col, value=display_name)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
            cell.font = Font(size=10)
            
            # Couleur selon type
            if building.building_type == 'Culturel':
                cell.fill = orange_fill
            elif building.building_type == 'Producteur':
                cell.fill = green_fill
            else:
                cell.fill = grey_fill
        
        # Ajuster les colonnes
        for col in range(2, terrain.cols + 2):
            ws.column_dimensions[get_column_letter(col)].width = 20
        
        # Bâtiments non placés
        if not_placed:
            unplaced_data = []
            for b in not_placed:
                unplaced_data.append({
                    'Nom': b.name,
                    'Longueur': b.longueur,
                    'Largeur': b.largeur,
                    'Type': b.building_type,
                    'Production': b.production
                })
            pd.DataFrame(unplaced_data).to_excel(writer, sheet_name='Non places', index=False)
        else:
            pd.DataFrame({'Message': ['Tous les bâtiments ont été placés']}).to_excel(writer, sheet_name='Non places', index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Optimiseur de Ville", layout="wide")
    st.title("🏗️ Optimiseur de Placement de Bâtiments")
    st.markdown("""
    ### Instructions
    1. Téléchargez votre fichier `Ville.xlsx`
    2. Cliquez sur "Lancer l'optimisation"
    3. Téléchargez le résultat
    
    **Stratégie utilisée:**
    - Priorité: Guérison > Nourriture > Or
    - Neutres placés sur les bords
    - Producteurs et culturels alternés par taille décroissante
    - Maximisation des boosts cumulés
    - Fusions de cellules pour visualisation
    """)
    
    uploaded_file = st.file_uploader("Choisissez le fichier Excel", type=['xlsx'])
    
    if uploaded_file is not None:
        if st.button("🚀 Lancer l'optimisation", type="primary"):
            try:
                with st.spinner("📂 Chargement du fichier..."):
                    terrain, buildings = parse_ville_excel(uploaded_file)
                
                total_buildings = sum(b.nombre for b in buildings)
                st.info(f"📊 Terrain: {terrain.rows}×{terrain.cols} cases")
                st.info(f"🏢 Bâtiments à placer: {total_buildings}")
                
                with st.spinner("🔄 Placement optimisé en cours..."):
                    placed, not_placed = place_buildings_optimized(terrain, buildings)
                
                # Statistiques
                total_cells = terrain.rows * terrain.cols
                used_cells = np.sum(terrain.grid)
                free_cells = total_cells - used_cells
                
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Bâtiments placés", len(placed))
                col2.metric("Bâtiments non placés", len(not_placed))
                col3.metric("Cases utilisées", int(used_cells))
                col4.metric("Cases libres", int(free_cells))
                
                # Afficher les comptes par type
                prod_count = len([b for b in placed if b.building_type == 'Producteur'])
                cult_count = len([b for b in placed if b.building_type == 'Culturel'])
                neutre_count = len([b for b in placed if b.building_type == 'Neutre'])
                st.write(f"📦 Détail: {prod_count} producteurs, {cult_count} culturels, {neutre_count} neutres")
                
                with st.spinner("📝 Génération du fichier résultat..."):
                    output_file = generate_output_excel(terrain, placed, not_placed)
                
                st.success("✅ Optimisation terminée!")
                
                st.download_button(
                    label="📥 Télécharger le résultat (Excel)",
                    data=output_file,
                    file_name="resultats_placement_optimise.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"❌ Erreur: {str(e)}")
                import traceback
                st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
