import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import io
from typing import List, Tuple, Dict, Optional
import copy

class Building:
    def __init__(self, nom, longueur, largeur, quantite, type_bat, culture, rayonnement, 
                 boost_25, boost_50, boost_100, production):
        self.nom = nom
        self.longueur = int(longueur)
        self.largeur = int(largeur)
        self.quantite = int(quantite)
        self.type = type_bat
        self.culture = float(culture) if culture else 0
        self.rayonnement = int(rayonnement) if rayonnement else 0
        self.boost_25 = float(boost_25) if boost_25 else 0
        self.boost_50 = float(boost_50) if boost_50 else 0
        self.boost_100 = float(boost_100) if boost_100 else 0
        self.production = production if production else ""
        
        # Pour le placement
        self.placed = 0
        self.positions = []  # Liste de tuples (x, y, orientation)
        
    def get_dimensions(self, orientation='H'):
        """Retourne (longueur, largeur) selon l'orientation"""
        if orientation == 'H':
            return self.longueur, self.largeur
        else:  # Vertical
            return self.largeur, self.longueur
            
    def __repr__(self):
        return f"{self.nom} ({self.longueur}x{self.largeur})"

class Terrain:
    def __init__(self, grid):
        self.grid = np.array(grid)
        self.height, self.width = self.grid.shape
        self.occupied = np.zeros_like(self.grid, dtype=bool)
        self.buildings = []  # Liste des bâtiments placés
        self.cultural_zones = np.zeros_like(self.grid, dtype=float)  # Accumulation de culture
        
    def can_place(self, x, y, longueur, largeur):
        """Vérifie si un bâtiment peut être placé à la position (x,y)"""
        if x + longueur > self.width or y + largeur > self.height:
            return False
            
        # Vérifier que toutes les cases sont libres (1 dans la grille) et non occupées
        for i in range(longueur):
            for j in range(largeur):
                if self.grid[y + j, x + i] == 0 or self.occupied[y + j, x + i]:
                    return False
        return True
    
    def place_building(self, building, x, y, orientation):
        """Place un bâtiment sur le terrain"""
        longueur, largeur = building.get_dimensions(orientation)
        
        for i in range(longueur):
            for j in range(largeur):
                self.occupied[y + j, x + i] = True
                
        building.placed += 1
        building.positions.append((x, y, orientation))
        self.buildings.append((building, x, y, orientation, longueur, largeur))
        
    def calculate_cultural_effect(self):
        """Calcule l'effet culturel de tous les bâtiments"""
        self.cultural_zones = np.zeros_like(self.grid, dtype=float)
        
        # D'abord, ajouter la culture des bâtiments culturels
        for building, x, y, orientation, longueur, largeur in self.buildings:
            if building.type == "culturel" and building.culture > 0:
                # Créer une zone d'effet autour du bâtiment
                center_x = x + longueur // 2
                center_y = y + largeur // 2
                
                # Ajouter la culture dans le rayon
                for i in range(max(0, center_x - building.rayonnement), 
                             min(self.width, center_x + building.rayonnement + 1)):
                    for j in range(max(0, center_y - building.rayonnement), 
                                 min(self.height, center_y + building.rayonnement + 1)):
                        self.cultural_zones[j, i] += building.culture
        
        return self.cultural_zones
    
    def get_production_boosts(self):
        """Calcule les boosts de production pour tous les bâtiments producteurs"""
        self.calculate_cultural_effect()
        
        results = []
        total_culture_by_type = {"Guerison": 0, "Nourriture": 0, "Or": 0}
        boost_counts = {"Guerison": {0: 0, 25: 0, 50: 0, 100: 0},
                       "Nourriture": {0: 0, 25: 0, 50: 0, 100: 0},
                       "Or": {0: 0, 25: 0, 50: 0, 100: 0}}
        
        for building, x, y, orientation, longueur, largeur in self.buildings:
            if building.type == "producteur" and building.production:
                # Calculer la culture reçue (moyenne sur les cases du bâtiment)
                total_culture = 0
                for i in range(longueur):
                    for j in range(largeur):
                        total_culture += self.cultural_zones[y + j, x + i]
                
                avg_culture = total_culture / (longueur * largeur)
                total_culture_by_type[building.production] += avg_culture
                
                # Déterminer le boost
                boost = 0
                if avg_culture >= building.boost_100:
                    boost = 100
                elif avg_culture >= building.boost_50:
                    boost = 50
                elif avg_culture >= building.boost_25:
                    boost = 25
                    
                boost_counts[building.production][boost] += 1
                
                results.append({
                    "Nom": building.nom,
                    "Production": building.production,
                    "Culture reçue": round(avg_culture, 2),
                    "Boost": f"{boost}%"
                })
        
        return results, total_culture_by_type, boost_counts

class BuildingPlacer:
    def __init__(self, terrain, buildings):
        self.terrain = terrain
        self.buildings = buildings
        
        # Ordre de priorité pour les types de production
        self.production_priority = {"Guerison": 1, "Nourriture": 2, "Or": 3}
        
    def get_priority_score(self, building):
        """Calcule un score de priorité pour un bâtiment"""
        if building.type == "producteur" and building.production:
            return self.production_priority.get(building.production, 4)
        return 4  # Priorité plus basse pour les bâtiments culturels
    
    def place_all(self):
        """Place tous les bâtiments en optimisant les boosts"""
        
        # Trier les bâtiments par priorité
        sorted_buildings = sorted(self.buildings, 
                                 key=lambda b: (self.get_priority_score(b), 
                                              -b.longueur * b.largeur))
        
        # Placer chaque bâtiment
        for building in sorted_buildings:
            for _ in range(building.quantite - building.placed):
                best_position = self.find_best_position(building)
                if best_position:
                    x, y, orientation = best_position
                    self.terrain.place_building(building, x, y, orientation)
                else:
                    st.warning(f"Impossible de placer {building.nom}")
                    break
    
    def find_best_position(self, building):
        """Trouve la meilleure position pour un bâtiment"""
        best_score = -1
        best_position = None
        
        # Essayer les deux orientations
        for orientation in ['H', 'V']:
            longueur, largeur = building.get_dimensions(orientation)
            
            for y in range(self.terrain.height - largeur + 1):
                for x in range(self.terrain.width - longueur + 1):
                    if self.terrain.can_place(x, y, longueur, largeur):
                        score = self.evaluate_position(building, x, y, orientation)
                        if score > best_score:
                            best_score = score
                            best_position = (x, y, orientation)
        
        return best_position
    
    def evaluate_position(self, building, x, y, orientation):
        """Évalue la qualité d'une position pour un bâtiment"""
        longueur, largeur = building.get_dimensions(orientation)
        
        # Simuler le placement temporaire
        temp_occupied = self.terrain.occupied.copy()
        for i in range(longueur):
            for j in range(largeur):
                temp_occupied[y + j, x + i] = True
        
        # Calculer la culture potentielle si c'est un producteur
        if building.type == "producteur" and building.production:
            temp_cultural = np.zeros_like(self.terrain.cultural_zones)
            
            # Ajouter l'effet des bâtiments culturels existants
            for b, bx, by, bor, bl, bw in self.terrain.buildings:
                if b.type == "culturel" and b.culture > 0:
                    center_x = bx + bl // 2
                    center_y = by + bw // 2
                    
                    for i in range(max(0, center_x - b.rayonnement), 
                                 min(self.terrain.width, center_x + b.rayonnement + 1)):
                        for j in range(max(0, center_y - b.rayonnement), 
                                     min(self.terrain.height, center_y + b.rayonnement + 1)):
                            temp_cultural[j, i] += b.culture
            
            # Calculer la culture sur le nouveau bâtiment
            total_culture = 0
            for i in range(longueur):
                for j in range(largeur):
                    total_culture += temp_cultural[y + j, x + i]
            
            avg_culture = total_culture / (longueur * largeur)
            
            # Score basé sur le boost potentiel
            if avg_culture >= building.boost_100:
                return 100 + avg_culture
            elif avg_culture >= building.boost_50:
                return 50 + avg_culture
            elif avg_culture >= building.boost_25:
                return 25 + avg_culture
            else:
                return avg_culture
        
        return 0  # Pour les bâtiments culturels, juste les placer

def read_input_file(file):
    """Lit le fichier Excel d'entrée"""
    try:
        # Lire les onglets
        xl = pd.ExcelFile(file)
        
        # Premier onglet : terrain
        terrain_df = pd.read_excel(xl, sheet_name=0, header=None)
        terrain_grid = terrain_df.values.tolist()
        
        # Deuxième onglet : bâtiments
        buildings_df = pd.read_excel(xl, sheet_name=1)
        
        return terrain_grid, buildings_df
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier: {e}")
        return None, None

def create_buildings_from_df(df):
    """Crée les objets Building à partir du DataFrame"""
    buildings = []
    for _, row in df.iterrows():
        building = Building(
            nom=row['Nom'],
            longueur=row['Longueur'],
            largeur=row['Largeur'],
            quantite=row['Quantite'],
            type_bat=row['Type'],
            culture=row['Culture'],
            rayonnement=row['Rayonnement'],
            boost_25=row['Boost 25%'],
            boost_50=row['Boost 50%'],
            boost_100=row['Boost 100%'],
            production=row['Production']
        )
        buildings.append(building)
    return buildings

def create_output_excel(terrain, boost_results, total_culture, boost_counts):
    """Crée le fichier Excel de sortie"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Feuille 1: Terrain avec bâtiments placés
        terrain_display = np.array(terrain.grid, dtype=object)
        
        # Marquer les bâtiments sur le terrain
        building_map = {}
        color_index = 1
        
        for building, x, y, orientation, longueur, largeur in terrain.buildings:
            for i in range(longueur):
                for j in range(largeur):
                    if building.nom not in building_map:
                        building_map[building.nom] = color_index
                        color_index += 1
                    terrain_display[y + j, x + i] = f"{building.nom[:3]}_{building_map[building.nom]}"
        
        terrain_df = pd.DataFrame(terrain_display)
        terrain_df.to_excel(writer, sheet_name='Terrain avec batiments', index=False, header=False)
        
        # Ajuster la largeur des colonnes
        worksheet = writer.sheets['Terrain avec batiments']
        for col in range(terrain_df.shape[1]):
            worksheet.column_dimensions[chr(65 + col)].width = 15
        
        # Colorer les cellules des bâtiments
        for building, x, y, orientation, longueur, largeur in terrain.buildings:
            fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
            for i in range(longueur):
                for j in range(largeur):
                    cell = worksheet.cell(row=y + j + 1, column=x + i + 1)
                    cell.fill = fill
                    cell.font = Font(bold=True)
        
        # Feuille 2: Résultats des boosts
        boost_df = pd.DataFrame(boost_results)
        if not boost_df.empty:
            boost_df.to_excel(writer, sheet_name='Boosts de production', index=False)
            
            # Ajuster la largeur
            worksheet = writer.sheets['Boosts de production']
            for column in boost_df:
                column_width = max(boost_df[column].astype(str).map(len).max(), len(column))
                col_idx = boost_df.columns.get_loc(column)
                worksheet.column_dimensions[chr(65 + col_idx)].width = column_width + 2
        
        # Feuille 3: Statistiques
        stats_data = []
        for prod_type in ["Guerison", "Nourriture", "Or"]:
            stats_data.append({
                "Type de production": prod_type,
                "Culture totale reçue": total_culture.get(prod_type, 0),
                "Nb boost 0%": boost_counts[prod_type][0],
                "Nb boost 25%": boost_counts[prod_type][25],
                "Nb boost 50%": boost_counts[prod_type][50],
                "Nb boost 100%": boost_counts[prod_type][100]
            })
        
        stats_df = pd.DataFrame(stats_data)
        stats_df.to_excel(writer, sheet_name='Statistiques', index=False)
        
        # Feuille 4: Position des bâtiments
        positions_data = []
        for building, x, y, orientation, longueur, largeur in terrain.buildings:
            positions_data.append({
                "Nom": building.nom,
                "Position X": x,
                "Position Y": y,
                "Orientation": orientation,
                "Longueur": longueur,
                "Largeur": largeur,
                "Type": building.type,
                "Production": building.production
            })
        
        positions_df = pd.DataFrame(positions_data)
        positions_df.to_excel(writer, sheet_name='Positions', index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Placeur de bâtiments optimisé", layout="wide")
    
    st.title("🏗️ Placeur de bâtiments optimisé")
    st.markdown("""
    Cette application optimise le placement de bâtiments sur un terrain pour maximiser 
    les boosts de production selon l'ordre de priorité : Guérison > Nourriture > Or.
    """)
    
    # Sidebar pour le téléchargement
    with st.sidebar:
        st.header("📂 Chargement des données")
        uploaded_file = st.file_uploader(
            "Choisir le fichier Excel", 
            type=['xlsx', 'xls'],
            help="Le fichier doit contenir un onglet 'Terrain' et un onglet 'Batiments'"
        )
        
        if uploaded_file:
            st.success("Fichier chargé avec succès!")
    
    # Zone principale
    if uploaded_file:
        # Lire le fichier
        terrain_grid, buildings_df = read_input_file(uploaded_file)
        
        if terrain_grid and buildings_df is not None:
            # Créer les objets
            terrain = Terrain(terrain_grid)
            buildings = create_buildings_from_df(buildings_df)
            
            # Afficher les données chargées
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("🗺️ Terrain")
                terrain_preview = pd.DataFrame(terrain_grid)
                st.dataframe(terrain_preview, use_container_width=True)
            
            with col2:
                st.subheader("🏢 Bâtiments à placer")
                st.dataframe(buildings_df, use_container_width=True)
            
            # Bouton pour lancer l'optimisation
            if st.button("🚀 Lancer l'optimisation", type="primary"):
                with st.spinner("Optimisation en cours..."):
                    # Créer le placeur et placer les bâtiments
                    placer = BuildingPlacer(terrain, buildings)
                    placer.place_all()
                    
                    # Calculer les résultats
                    boost_results, total_culture, boost_counts = terrain.get_production_boosts()
                    
                    # Afficher les résultats
                    st.subheader("📊 Résultats de l'optimisation")
                    
                    # Métriques principales
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        total_buildings_placed = len([b for b in buildings if b.placed > 0])
                        total_buildings = sum(b.quantite for b in buildings)
                        st.metric("Bâtiments placés", f"{total_buildings_placed}/{total_buildings}")
                    
                    with col2:
                        total_culture_sum = sum(total_culture.values())
                        st.metric("Culture totale distribuée", f"{total_culture_sum:.0f}")
                    
                    with col3:
                        total_producteurs = len([b for b in buildings if b.type == "producteur" and b.production])
                        st.metric("Bâtiments producteurs", total_producteurs)
                    
                    # Visualisation du terrain avec les bâtiments
                    st.subheader("🗺️ Terrain avec bâtiments placés")
                    
                    # Créer une visualisation simplifiée
                    terrain_viz = np.array(terrain.grid, dtype=str)
                    building_colors = {}
                    color_idx = 0
                    colors = ['🟦', '🟥', '🟩', '🟨', '🟪', '🟫', '🟧']
                    
                    for building, x, y, orientation, longueur, largeur in terrain.buildings:
                        if building.nom not in building_colors:
                            building_colors[building.nom] = colors[color_idx % len(colors)]
                            color_idx += 1
                        
                        for i in range(longueur):
                            for j in range(largeur):
                                terrain_viz[y + j, x + i] = building_colors[building.nom]
                    
                    # Remplacer les 1 par des cases vides
                    terrain_viz[terrain_viz == '1'] = '⬜'
                    terrain_viz[terrain_viz == '0'] = '⬛'
                    
                    st.dataframe(pd.DataFrame(terrain_viz), use_container_width=True)
                    
                    # Légende des bâtiments
                    st.subheader("🏷️ Légende des bâtiments")
                    legend_cols = st.columns(4)
                    for i, (building_name, symbol) in enumerate(building_colors.items()):
                        with legend_cols[i % 4]:
                            st.markdown(f"{symbol} {building_name}")
                    
                    # Tableau des boosts
                    st.subheader("📈 Détail des boosts par bâtiment")
                    if boost_results:
                        boost_df = pd.DataFrame(boost_results)
                        st.dataframe(boost_df, use_container_width=True)
                    
                    # Statistiques par type
                    st.subheader("📊 Statistiques par type de production")
                    stats_data = []
                    for prod_type in ["Guerison", "Nourriture", "Or"]:
                        stats_data.append({
                            "Type": prod_type,
                            "Culture totale": total_culture.get(prod_type, 0),
                            "Boost 0%": boost_counts[prod_type][0],
                            "Boost 25%": boost_counts[prod_type][25],
                            "Boost 50%": boost_counts[prod_type][50],
                            "Boost 100%": boost_counts[prod_type][100]
                        })
                    
                    stats_df = pd.DataFrame(stats_data)
                    st.dataframe(stats_df, use_container_width=True)
                    
                    # Créer et proposer le téléchargement du fichier Excel
                    output_file = create_output_excel(terrain, boost_results, total_culture, boost_counts)
                    
                    st.download_button(
                        label="📥 Télécharger les résultats (Excel)",
                        data=output_file,
                        file_name="resultats_placement_batiments.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("✅ Optimisation terminée avec succès!")
    
    else:
        # Message d'accueil
        st.info("👆 Veuillez charger un fichier Excel pour commencer.")
        
        # Aide
        with st.expander("ℹ️ Format du fichier attendu"):
            st.markdown("""
            ### Premier onglet : Terrain
            - Matrice de 0 et 1
            - 1 = case libre
            - 0 = case occupée
            
            ### Second onglet : Bâtiments
            Colonnes attendues :
            - **Nom** : nom du bâtiment
            - **Longueur** : en nombre de cases
            - **Largeur** : en nombre de cases
            - **Quantite** : nombre à placer
            - **Type** : "culturel" ou "producteur"
            - **Culture** : quantité produite (pour culturel)
            - **Rayonnement** : portée en cases (pour culturel)
            - **Boost 25%** : seuil pour boost 25%
            - **Boost 50%** : seuil pour boost 50%
            - **Boost 100%** : seuil pour boost 100%
            - **Production** : "Guerison", "Nourriture", "Or" ou vide
            """)

if __name__ == "__main__":
    main()