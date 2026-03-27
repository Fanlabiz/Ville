"""
Placement optimal de bâtiments sur un terrain
Auteur: Assistant
Date: 2026-03-27
"""

import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import streamlit as st
import io
from typing import List, Tuple, Dict, Optional

class Building:
    """Classe représentant un bâtiment"""
    def __init__(self, name: str, longueur: int, largeur: int, nombre: int, 
                 building_type: str, culture: int, rayonnement: int,
                 boost_25: int, boost_50: int, boost_100: int,
                 production: str, quantite: int):
        self.name = name
        self.longueur = longueur
        self.largeur = largeur
        self.nombre = nombre
        self.building_type = building_type  # 'Culturel', 'Producteur', 'Neutre'
        self.culture = culture
        self.rayonnement = rayonnement
        self.boost_25 = boost_25
        self.boost_50 = boost_50
        self.boost_100 = boost_100
        self.production = production
        self.quantite = quantite
        self.placed = []  # Liste des positions (ligne, colonne, orientation) pour chaque exemplaire
        
    def get_dimensions(self, orientation: str = 'H') -> Tuple[int, int]:
        """Retourne (hauteur, largeur) selon l'orientation"""
        if orientation == 'H':
            return self.longueur, self.largeur
        else:
            return self.largeur, self.longueur
            
    def get_cells(self, row: int, col: int, orientation: str = 'H') -> List[Tuple[int, int]]:
        """Retourne la liste des cases occupées par le bâtiment"""
        h, w = self.get_dimensions(orientation)
        cells = []
        for i in range(h):
            for j in range(w):
                cells.append((row + i, col + j))
        return cells
        
    def get_zone_rayonnement(self, row: int, col: int, orientation: str = 'H') -> List[Tuple[int, int]]:
        """Retourne la zone de rayonnement (cases autour du bâtiment)"""
        if self.rayonnement == 0:
            return []
        
        h, w = self.get_dimensions(orientation)
        cells = set()
        
        # Zone de rayonnement = cases dans un rectangle agrandi
        for i in range(-self.rayonnement, h + self.rayonnement):
            for j in range(-self.rayonnement, w + self.rayonnement):
                if i < 0 or i >= h or j < 0 or j >= w:  # Uniquement les cases extérieures
                    cells.add((row + i, col + j))
        return list(cells)

class Terrain:
    """Classe représentant le terrain"""
    def __init__(self, grid: np.ndarray):
        self.grid = grid  # 0 = libre, 1 = occupé
        self.rows, self.cols = grid.shape
        self.buildings = []  # Liste des bâtiments placés (avec positions)
        
    def can_place(self, building: Building, row: int, col: int, orientation: str) -> bool:
        """Vérifie si le bâtiment peut être placé à la position donnée"""
        h, w = building.get_dimensions(orientation)
        
        # Vérifier les limites
        if row < 0 or col < 0 or row + h > self.rows or col + w > self.cols:
            return False
            
        # Vérifier les collisions
        for i in range(h):
            for j in range(w):
                if self.grid[row + i, col + j] != 0:
                    return False
        return True
        
    def place(self, building: Building, row: int, col: int, orientation: str):
        """Place le bâtiment sur le terrain"""
        h, w = building.get_dimensions(orientation)
        
        for i in range(h):
            for j in range(w):
                self.grid[row + i, col + j] = 1
                
        building.placed.append((row, col, orientation))
        self.buildings.append((building, row, col, orientation))

def parse_excel_input(file) -> Tuple[Terrain, List[Building]]:
    """Parse le fichier Excel d'entrée"""
    # Lire le terrain (premier onglet)
    df_terrain = pd.read_excel(file, sheet_name=0, header=None)
    grid = df_terrain.values
    
    # Convertir: 1 = libre (0 dans notre grid), 0 = occupé (1 dans notre grid)
    # Si la valeur est 'X' ou autre, c'est un obstacle
    terrain_grid = np.zeros(grid.shape, dtype=int)
    for i in range(grid.shape[0]):
        for j in range(grid.shape[1]):
            val = grid[i, j]
            if val == 1 or val == '1':
                terrain_grid[i, j] = 0  # libre
            else:
                terrain_grid[i, j] = 1  # occupé
    
    terrain = Terrain(terrain_grid)
    
    # Lire les bâtiments (deuxième onglet)
    df_buildings = pd.read_excel(file, sheet_name=1)
    buildings = []
    
    for _, row in df_buildings.iterrows():
        # Gérer les valeurs NaN
        quantite_val = row['Quantite']
        if isinstance(quantite_val, str) and '=' in quantite_val:
            # Évaluer la formule simple comme =49980/2
            try:
                quantite_val = eval(quantite_val.replace('=', ''))
            except:
                quantite_val = 0
        elif pd.isna(quantite_val):
            quantite_val = 0
            
        b = Building(
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
            quantite=int(quantite_val) if quantite_val != 0 else 0
        )
        buildings.append(b)
    
    return terrain, buildings

def calculate_boost(culture_total: int, building: Building) -> float:
    """Calcule le boost de production en fonction de la culture totale"""
    if culture_total >= building.boost_100:
        return 2.0  # +100%
    elif culture_total >= building.boost_50:
        return 1.5  # +50%
    elif culture_total >= building.boost_25:
        return 1.25  # +25%
    else:
        return 1.0

def get_production_priority(production: str) -> int:
    """Retourne la priorité de production"""
    priorities = {'Guerison': 3, 'Nourriture': 2, 'Or': 1}
    return priorities.get(production, 0)

def find_best_position(terrain: Terrain, building: Building, prefer_border: bool = False) -> Optional[Tuple[int, int, str]]:
    """Trouve la meilleure position pour placer un bâtiment"""
    best_score = -1
    best_pos = None
    
    for orientation in ['H', 'V']:
        h, w = building.get_dimensions(orientation)
        for row in range(terrain.rows - h + 1):
            for col in range(terrain.cols - w + 1):
                if terrain.can_place(building, row, col, orientation):
                    score = 0
                    if prefer_border:
                        # Favoriser les positions sur le bord
                        if row == 0 or row + h == terrain.rows or col == 0 or col + w == terrain.cols:
                            score += 1000
                    
                    if score > best_score:
                        best_score = score
                        best_pos = (row, col, orientation)
    
    return best_pos

def find_best_position_for_producer(terrain: Terrain, building: Building, 
                                    culture_map: np.ndarray) -> Optional[Tuple[int, int, str]]:
    """Trouve la meilleure position pour un producteur (maximise la culture reçue)"""
    best_score = -1
    best_pos = None
    
    for orientation in ['H', 'V']:
        h, w = building.get_dimensions(orientation)
        for row in range(terrain.rows - h + 1):
            for col in range(terrain.cols - w + 1):
                if terrain.can_place(building, row, col, orientation):
                    # Calculer la culture totale sur toutes les cases du bâtiment
                    culture_total = 0
                    for i in range(h):
                        for j in range(w):
                            if 0 <= row + i < terrain.rows and 0 <= col + j < terrain.cols:
                                culture_total += culture_map[row + i, col + j]
                    
                    # Calculer le boost et la production
                    boost = calculate_boost(culture_total, building)
                    prod_value = building.quantite * boost
                    
                    # Priorité selon le type de production
                    priority = get_production_priority(building.production)
                    score = prod_value * (priority * 10 + 1)
                    
                    if score > best_score:
                        best_score = score
                        best_pos = (row, col, orientation)
    
    return best_pos

def find_best_position_for_culturel(terrain: Terrain, building: Building, 
                                     culture_map: np.ndarray) -> Optional[Tuple[int, int, str]]:
    """Trouve la meilleure position pour un culturel (maximise le gain de production futur)"""
    best_score = -1
    best_pos = None
    
    for orientation in ['H', 'V']:
        h, w = building.get_dimensions(orientation)
        for row in range(terrain.rows - h + 1):
            for col in range(terrain.cols - w + 1):
                if terrain.can_place(building, row, col, orientation):
                    # Calculer le gain potentiel
                    gain = 0
                    zone = building.get_zone_rayonnement(row, col, orientation)
                    
                    for r, c in zone:
                        if 0 <= r < terrain.rows and 0 <= c < terrain.cols:
                            # Vérifier si la case est occupée
                            if terrain.grid[r, c] == 1:
                                # Trouver le producteur sur cette case
                                for b, br, bc, bo in terrain.buildings:
                                    if b.building_type == 'Producteur':
                                        cells = b.get_cells(br, bc, bo)
                                        if (r, c) in cells:
                                            old_culture = culture_map[r, c]
                                            new_culture = old_culture + building.culture
                                            old_boost = calculate_boost(old_culture, b)
                                            new_boost = calculate_boost(new_culture, b)
                                            gain += b.quantite * (new_boost - old_boost)
                                            break
                    
                    score = gain
                    if score > best_score:
                        best_score = score
                        best_pos = (row, col, orientation)
    
    return best_pos

def place_buildings_optimized(terrain: Terrain, buildings: List[Building]) -> Tuple[Terrain, List[Building], List[Building]]:
    """Place les bâtiments de manière optimisée"""
    
    # Séparer les bâtiments par type
    neutres = [b for b in buildings if b.building_type == 'Neutre']
    producteurs = [b for b in buildings if b.building_type == 'Producteur']
    culturels = [b for b in buildings if b.building_type == 'Culturel']
    
    # Trier par priorité de production
    producteurs.sort(key=lambda b: get_production_priority(b.production), reverse=True)
    
    # Trier par taille décroissante
    neutres.sort(key=lambda b: b.longueur * b.largeur, reverse=True)
    producteurs.sort(key=lambda b: b.longueur * b.largeur, reverse=True)
    culturels.sort(key=lambda b: b.longueur * b.largeur, reverse=True)
    
    placed_buildings = []
    not_placed = []
    
    # Créer une liste d'attente pour chaque type
    neutres_to_place = []
    for b in neutres:
        for _ in range(b.nombre):
            neutres_to_place.append(b)
    
    producteurs_to_place = []
    for b in producteurs:
        for _ in range(b.nombre):
            producteurs_to_place.append(b)
    
    culturels_to_place = []
    for b in culturels:
        for _ in range(b.nombre):
            culturels_to_place.append(b)
    
    # Étape 1: Placer les neutres sur les bords
    for building in neutres_to_place:
        # Chercher une position sur le bord
        best_pos = find_best_position(terrain, building, prefer_border=True)
        if not best_pos:
            best_pos = find_best_position(terrain, building, prefer_border=False)
        
        if best_pos:
            row, col, orientation = best_pos
            terrain.place(building, row, col, orientation)
            placed_buildings.append(building)
        else:
            not_placed.append(building)
    
    # Calculer la carte des boosts initiaux
    culture_map = np.zeros((terrain.rows, terrain.cols), dtype=int)
    
    # Mettre à jour avec les culturels déjà placés
    for b, row, col, orientation in terrain.buildings:
        if b.building_type == 'Culturel':
            zone = b.get_zone_rayonnement(row, col, orientation)
            for r, c in zone:
                if 0 <= r < terrain.rows and 0 <= c < terrain.cols:
                    culture_map[r, c] += b.culture
    
    # Étape 2: Placer les producteurs par priorité
    for building in producteurs_to_place:
        best_pos = find_best_position_for_producer(terrain, building, culture_map)
        if best_pos:
            row, col, orientation = best_pos
            terrain.place(building, row, col, orientation)
            placed_buildings.append(building)
        else:
            not_placed.append(building)
    
    # Étape 3: Placer les culturels
    for building in culturels_to_place:
        best_pos = find_best_position_for_culturel(terrain, building, culture_map)
        if best_pos:
            row, col, orientation = best_pos
            terrain.place(building, row, col, orientation)
            placed_buildings.append(building)
            # Mettre à jour la carte des boosts
            zone = building.get_zone_rayonnement(row, col, orientation)
            for r, c in zone:
                if 0 <= r < terrain.rows and 0 <= c < terrain.cols:
                    culture_map[r, c] += building.culture
        else:
            not_placed.append(building)
    
    # Étape 4: Essayer de replacer les producteurs non placés
    remaining_producers = [b for b in not_placed if b.building_type == 'Producteur']
    for building in remaining_producers:
        best_pos = find_best_position_for_producer(terrain, building, culture_map)
        if best_pos:
            row, col, orientation = best_pos
            terrain.place(building, row, col, orientation)
            placed_buildings.append(building)
            not_placed.remove(building)
    
    return terrain, placed_buildings, not_placed

def generate_output(terrain: Terrain, placed_buildings: List[Building], 
                    not_placed: List[Building]) -> io.BytesIO:
    """Génère le fichier Excel de résultat"""
    output = io.BytesIO()
    
    # Calculer la carte des boosts
    culture_map = np.zeros((terrain.rows, terrain.cols), dtype=int)
    
    for building, row, col, orientation in terrain.buildings:
        if building.building_type == 'Culturel':
            zone = building.get_zone_rayonnement(row, col, orientation)
            for r, c in zone:
                if 0 <= r < terrain.rows and 0 <= c < terrain.cols:
                    culture_map[r, c] += building.culture
    
    # 1. Liste des bâtiments placés
    placed_data = []
    for building in placed_buildings:
        for idx, (row, col, orientation) in enumerate(building.placed):
            # Calculer la culture reçue
            culture_recue = 0
            h, w = building.get_dimensions(orientation)
            for i in range(h):
                for j in range(w):
                    if 0 <= row + i < terrain.rows and 0 <= col + j < terrain.cols:
                        culture_recue += culture_map[row + i, col + j]
            
            # Calculer le boost
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
                'Culture recue': culture_recue,
                'Boost (%)': int((boost - 1) * 100),
                'Quantite/h': building.quantite,
                'Prod totale/h': prod_h,
                'Origine': 'Placé'
            })
    
    # 2. Synthèse par type de production
    synthesis_data = []
    if placed_data:
        df_placed = pd.DataFrame(placed_data)
        
        for production in df_placed['Production'].unique():
            if production != 'Rien':
                prod_df = df_placed[df_placed['Production'] == production]
                if len(prod_df) > 0 and prod_df['Quantite/h'].sum() > 0:
                    total_culture = prod_df['Culture recue'].sum()
                    boost_moyen = (prod_df['Prod totale/h'].sum() / prod_df['Quantite/h'].sum() - 1) * 100
                    synthesis_data.append({
                        'Production': production,
                        'Culture totale': int(total_culture),
                        'Boost moyen (%)': round(boost_moyen, 1),
                        'Nb batiments': len(prod_df),
                        'Production/h': prod_df['Prod totale/h'].sum()
                    })
    
    # 3. Statistiques
    total_cells = terrain.rows * terrain.cols
    used_cells = np.sum(terrain.grid)
    free_cells = total_cells - used_cells
    unplaced_cells = sum(b.longueur * b.largeur for b in not_placed)
    
    # Écrire dans Excel avec pandas
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Feuille des bâtiments placés
        if placed_data:
            pd.DataFrame(placed_data).to_excel(writer, sheet_name='Batiments places', index=False)
        else:
            pd.DataFrame().to_excel(writer, sheet_name='Batiments places', index=False)
        
        # Feuille de synthèse
        if synthesis_data:
            df_synthesis = pd.DataFrame(synthesis_data)
            df_synthesis.to_excel(writer, sheet_name='Synthese', index=False)
            
            # Ajouter les stats
            stats_df = pd.DataFrame({
                'Statistiques': ['Cases libres restantes', 'Cases des batiments non places', 'Nombre de batiments non places'],
                'Valeur': [free_cells, unplaced_cells, len(not_placed)]
            })
            stats_df.to_excel(writer, sheet_name='Synthese', startrow=len(synthesis_data) + 3, index=False)
        else:
            stats_df = pd.DataFrame({
                'Statistiques': ['Cases libres restantes', 'Cases des batiments non places', 'Nombre de batiments non places'],
                'Valeur': [free_cells, unplaced_cells, len(not_placed)]
            })
            stats_df.to_excel(writer, sheet_name='Synthese', index=False)
        
        # Feuille terrain final (texte)
        terrain_display = np.full((terrain.rows, terrain.cols), '', dtype=object)
        
        for building, row, col, orientation in terrain.buildings:
            h, w = building.get_dimensions(orientation)
            name = building.name
            if building.building_type == 'Producteur':
                culture_recue = 0
                for i in range(h):
                    for j in range(w):
                        if 0 <= row + i < terrain.rows and 0 <= col + j < terrain.cols:
                            culture_recue += culture_map[row + i, col + j]
                boost = calculate_boost(culture_recue, building)
                if boost > 1:
                    name = f"{building.name} +{int((boost-1)*100)}%"
            
            for i in range(h):
                for j in range(w):
                    if 0 <= row + i < terrain.rows and 0 <= col + j < terrain.cols:
                        terrain_display[row + i, col + j] = name
        
        df_terrain = pd.DataFrame(terrain_display)
        df_terrain.to_excel(writer, sheet_name='Terrain final', index=False)
        
        # Feuille des bâtiments non placés
        if not_placed:
            unplaced_data = []
            for building in not_placed:
                unplaced_data.append({
                    'Nom': building.name,
                    'Longueur': building.longueur,
                    'Largeur': building.largeur,
                    'Nombre': 1,
                    'Type': building.building_type,
                    'Production': building.production
                })
            df_unplaced = pd.DataFrame(unplaced_data)
            df_unplaced.to_excel(writer, sheet_name='Non places', index=False)
        else:
            pd.DataFrame({'Message': ['Aucun bâtiment non placé']}).to_excel(writer, sheet_name='Non places', index=False)
        
        # Appliquer la mise en forme
        workbook = writer.book
        
        # Couleurs
        orange_fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
        green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
        grey_fill = PatternFill(start_color='808080', end_color='808080', fill_type='solid')
        
        # Formater le terrain
        if 'Terrain final' in workbook.sheetnames:
            ws = workbook['Terrain final']
            for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=terrain.rows, min_col=1, max_col=terrain.cols), 1):
                for col_idx, cell in enumerate(row, 1):
                    if cell.value and isinstance(cell.value, str):
                        # Trouver le bâtiment correspondant
                        for building, r, c, o in terrain.buildings:
                            if cell.value.startswith(building.name):
                                if building.building_type == 'Culturel':
                                    cell.fill = orange_fill
                                elif building.building_type == 'Producteur':
                                    cell.fill = green_fill
                                else:
                                    cell.fill = grey_fill
                                break
    
    output.seek(0)
    return output

def main():
    """Fonction principale Streamlit"""
    st.set_page_config(page_title="Placement de Bâtiments", layout="wide")
    st.title("🏗️ Optimiseur de Placement de Bâtiments")
    
    st.markdown("""
    ### Instructions
    1. Téléchargez un fichier Excel avec:
       - **Onglet 1**: Terrain (1=libre, autre=occupé)
       - **Onglet 2**: Liste des bâtiments à placer
    2. Lancez l'optimisation
    3. Téléchargez le résultat
    """)
    
    uploaded_file = st.file_uploader("Choisissez le fichier Excel", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        if st.button("🚀 Lancer l'optimisation"):
            try:
                with st.spinner("Chargement et analyse du fichier..."):
                    terrain, buildings = parse_excel_input(uploaded_file)
                
                st.info(f"📊 Terrain: {terrain.rows}×{terrain.cols} cases")
                total_buildings = sum(b.nombre for b in buildings)
                st.info(f"🏢 Bâtiments à placer: {total_buildings}")
                
                with st.spinner("Placement optimisé des bâtiments..."):
                    terrain, placed, not_placed = place_buildings_optimized(terrain, buildings)
                
                st.success(f"✅ Placement terminé! {len(placed)} bâtiments placés, {len(not_placed)} non placés")
                
                # Afficher les statistiques
                total_cells = terrain.rows * terrain.cols
                used_cells = np.sum(terrain.grid)
                free_cells = total_cells - used_cells
                
                col1, col2, col3 = st.columns(3)
                col1.metric("Cases libres", free_cells)
                col2.metric("Bâtiments placés", len(placed))
                col3.metric("Bâtiments non placés", len(not_placed))
                
                with st.spinner("Génération du fichier résultat..."):
                    output_file = generate_output(terrain, placed, not_placed)
                
                st.download_button(
                    label="📥 Télécharger le résultat",
                    data=output_file,
                    file_name="resultats_placement_optimise.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                st.success("🎉 Optimisation terminée! Cliquez sur le bouton pour télécharger.")
                
            except Exception as e:
                st.error(f"❌ Une erreur est survenue: {str(e)}")
                st.info("Vérifiez que votre fichier Excel a les bons onglets et le bon format.")

if __name__ == "__main__":
    main()
