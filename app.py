import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import io
from typing import List, Tuple, Dict, Optional
import copy
import re

class Building:
    def __init__(self, nom, longueur, largeur, quantite, type_bat, culture, rayonnement, 
                 boost_25, boost_50, boost_100, production):
        self.nom = nom
        self.longueur = int(float(longueur)) if pd.notna(longueur) else 0
        self.largeur = int(float(largeur)) if pd.notna(largeur) else 0
        self.quantite = int(float(quantite)) if pd.notna(quantite) else 1
        self.type = str(type_bat).lower() if pd.notna(type_bat) else ""
        self.culture = float(culture) if pd.notna(culture) else 0
        self.rayonnement = int(float(rayonnement)) if pd.notna(rayonnement) else 0
        self.boost_25 = float(boost_25) if pd.notna(boost_25) else 0
        self.boost_50 = float(boost_50) if pd.notna(boost_50) else 0
        self.boost_100 = float(boost_100) if pd.notna(boost_100) else 0
        self.production = str(production) if pd.notna(production) else ""
        
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
        total_culture_by_type = {}
        boost_counts = {}
        
        for building, x, y, orientation, longueur, largeur in self.buildings:
            if building.type == "producteur" and building.production:
                # Normaliser la production
                prod_type = building.production.strip()
                if prod_type not in total_culture_by_type:
                    total_culture_by_type[prod_type] = 0
                    boost_counts[prod_type] = {0: 0, 25: 0, 50: 0, 100: 0}
                
                # Calculer la culture reçue (moyenne sur les cases du bâtiment)
                total_culture = 0
                for i in range(longueur):
                    for j in range(largeur):
                        total_culture += self.cultural_zones[y + j, x + i]
                
                avg_culture = total_culture / (longueur * largeur) if (longueur * largeur) > 0 else 0
                
                # Ajouter à la catégorie correspondante
                total_culture_by_type[prod_type] += avg_culture
                
                # Déterminer le boost
                boost = 0
                if avg_culture >= building.boost_100:
                    boost = 100
                elif avg_culture >= building.boost_50:
                    boost = 50
                elif avg_culture >= building.boost_25:
                    boost = 25
                    
                # Incrémenter le compteur de boost
                boost_counts[prod_type][boost] += 1
                
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
        return 4  # Priorité plus basse pour les bâtiments culturels
    
    def place_all(self):
        """Place tous les bâtiments en optimisant les boosts"""
        
        # Séparer les bâtiments culturels et producteurs
        cultural_buildings = [b for b in self.buildings if b.type == "culturel"]
        producer_buildings = [b for b in self.buildings if b.type == "producteur" and b.production]
        
        # Trier les producteurs par priorité
        producer_buildings.sort(key=lambda b: (self.get_priority_score(b), -b.longueur * b.largeur))
        
        # Placer d'abord les bâtiments culturels
        st.write("🏛️ Placement des bâtiments culturels...")
        for building in cultural_buildings:
            for _ in range(building.quantite - building.placed):
                self.try_place_building(building)
        
        # Recalculer l'effet culturel après placement des culturels
        self.terrain.calculate_cultural_effect()
        
        # Placer ensuite les producteurs
        st.write("🏭 Placement des bâtiments producteurs...")
        for building in producer_buildings:
            for _ in range(building.quantite - building.placed):
                self.try_place_building(building)
    
    def try_place_building(self, building):
        """Essaie de placer un bâtiment et retourne True si réussi"""
        best_position = self.find_best_position(building)
        if best_position:
            x, y, orientation = best_position
            self.terrain.place_building(building, x, y, orientation)
            return True
        return False
    
    def find_best_position(self, building):
        """Trouve la meilleure position pour un bâtiment"""
        best_score = -1
        best_position = None
        
        # Essayer les deux orientations
        for orientation in ['H', 'V']:
            longueur, largeur = building.get_dimensions(orientation)
            
            # Essayer toutes les positions possibles
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
        
        # Calculer la culture potentielle si c'est un producteur
        if building.type == "producteur" and building.production:
            # Calculer la culture sur le nouveau bâtiment
            total_culture = 0
            for i in range(longueur):
                for j in range(largeur):
                    total_culture += self.terrain.cultural_zones[y + j, x + i]
            
            avg_culture = total_culture / (longueur * largeur) if (longueur * largeur) > 0 else 0
            
            # Score basé sur le boost potentiel
            if avg_culture >= building.boost_100:
                return 100 + avg_culture
            elif avg_culture >= building.boost_50:
                return 50 + avg_culture
            elif avg_culture >= building.boost_25:
                return 25 + avg_culture
            else:
                return avg_culture
        
        return 1  # Pour les bâtiments culturels

def normalize_column_name(name):
    """Normalise les noms de colonnes pour gérer les accents et variations"""
    if pd.isna(name):
        return ""
    
    # Convertir en string et enlever les espaces
    name = str(name).strip()
    
    # Remplacer les caractères accentués
    replacements = {
        'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
        'à': 'a', 'â': 'a', 'ä': 'a',
        'î': 'i', 'ï': 'i',
        'ô': 'o', 'ö': 'o',
        'ù': 'u', 'û': 'u', 'ü': 'u',
        'ç': 'c',
        'É': 'E', 'È': 'E', 'Ê': 'E', 'Ë': 'E',
        'À': 'A', 'Â': 'A', 'Ä': 'A',
        'Î': 'I', 'Ï': 'I',
        'Ô': 'O', 'Ö': 'O',
        'Ù': 'U', 'Û': 'U', 'Ü': 'U',
        'Ç': 'C'
    }
    
    for accented, unaccented in replacements.items():
        name = name.replace(accented, unaccented)
    
    # Enlever les caractères spéciaux et espaces
    name = re.sub(r'[^a-zA-Z0-9%]', '', name)
    
    return name.lower()

def read_input_file(file):
    """Lit le fichier Excel d'entrée avec gestion robuste des colonnes"""
    try:
        # Lire les onglets
        xl = pd.ExcelFile(file)
        
        # Premier onglet : terrain
        terrain_df = pd.read_excel(xl, sheet_name=0, header=None)
        terrain_grid = terrain_df.values.tolist()
        
        # Deuxième onglet : bâtiments
        buildings_df = pd.read_excel(xl, sheet_name=1)
        
        # Normaliser les noms de colonnes
        buildings_df.columns = [normalize_column_name(col) for col in buildings_df.columns]
        
        return terrain_grid, buildings_df
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier: {e}")
        return None, None

def create_buildings_from_df(df):
    """Crée les objets Building à partir du DataFrame avec mapping flexible"""
    
    # Mapping des noms de colonnes possibles
    column_mapping = {
        'nom': ['nom', 'name', 'nomdubatiment', 'batiment'],
        'longueur': ['longueur', 'long', 'length', 'l'],
        'largeur': ['largeur', 'larg', 'width', 'w'],
        'quantite': ['quantite', 'quantité', 'quantity', 'qte', 'qt', 'nb', 'nombre'],
        'type': ['type', 'typebat', 'typebatiment'],
        'culture': ['culture', 'cult', 'productionculture'],
        'rayonnement': ['rayonnement', 'rayon', 'radius', 'range', 'portee'],
        'boost25': ['boost25%', 'boost25', 'boost25%', '25%', 'seuil25'],
        'boost50': ['boost50%', 'boost50', 'boost50%', '50%', 'seuil50'],
        'boost100': ['boost100%', 'boost100', 'boost100%', '100%', 'seuil100'],
        'production': ['production', 'prod', 'output', 'typeproduction']
    }
    
    # Fonction pour trouver une colonne
    def find_column(possible_names):
        for name in possible_names:
            if name in df.columns:
                return name
        return None
    
    # Trouver les colonnes
    nom_col = find_column(column_mapping['nom'])
    longueur_col = find_column(column_mapping['longueur'])
    largeur_col = find_column(column_mapping['largeur'])
    quantite_col = find_column(column_mapping['quantite'])
    type_col = find_column(column_mapping['type'])
    culture_col = find_column(column_mapping['culture'])
    rayonnement_col = find_column(column_mapping['rayonnement'])
    boost25_col = find_column(column_mapping['boost25'])
    boost50_col = find_column(column_mapping['boost50'])
    boost100_col = find_column(column_mapping['boost100'])
    production_col = find_column(column_mapping['production'])
    
    # Vérifier que les colonnes essentielles sont trouvées
    if not nom_col:
        st.error("Colonne 'Nom' non trouvée")
        return []
    
    buildings = []
    for idx, row in df.iterrows():
        try:
            building = Building(
                nom=row[nom_col],
                longueur=row[longueur_col] if longueur_col else 1,
                largeur=row[largeur_col] if largeur_col else 1,
                quantite=row[quantite_col] if quantite_col else 1,
                type_bat=row[type_col] if type_col else "",
                culture=row[culture_col] if culture_col else 0,
                rayonnement=row[rayonnement_col] if rayonnement_col else 0,
                boost_25=row[boost25_col] if boost25_col else 0,
                boost_50=row[boost50_col] if boost50_col else 0,
                boost_100=row[boost100_col] if boost100_col else 0,
                production=row[production_col] if production_col else ""
            )
            buildings.append(building)
        except Exception as e:
            st.warning(f"Erreur sur la ligne {idx+1}: {e}")
            continue
    
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
        
        # Ajuster la largeur des colonnes (version corrigée)
        worksheet = writer.sheets['Terrain avec batiments']
        for i, col in enumerate(worksheet.columns):
            worksheet.column_dimensions[col[0].column_letter].width = 15
        
        # Colorer les cellules des bâtiments
        for building, x, y, orientation, longueur, largeur in terrain.buildings:
            fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
            for i in range(longueur):
                for j in range(largeur):
                    cell = worksheet.cell(row=y + j + 1, column=x + i + 1)
                    cell.fill = fill
                    cell.font = Font(bold=True)
        
        # Feuille 2: Résultats des boosts
        if boost_results:
            boost_df = pd.DataFrame(boost_results)
            boost_df.to_excel(writer, sheet_name='Boosts de production', index=False)
        
        # Feuille 3: Statistiques
        stats_data = []
        for prod_type, culture in total_culture.items():
            if prod_type in boost_counts:
                stats_data.append({
                    "Type de production": prod_type,
                    "Culture totale reçue": round(culture, 2),
                    "Nb boost 0%": boost_counts[prod_type][0],
                    "Nb boost 25%": boost_counts[prod_type][25],
                    "Nb boost 50%": boost_counts[prod_type][50],
                    "Nb boost 100%": boost_counts[prod_type][100]
                })
        
        if stats_data:
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
                "Production": building.production if building.production else ""
            })
        
        if positions_data:
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
            help="Le fichier doit contenir un onglet avec le terrain et un onglet avec les bâtiments"
        )
        
        if uploaded_file:
            st.success("Fichier chargé avec succès!")
    
    # Zone principale
    if uploaded_file:
        # Lire le fichier
        terrain_grid, buildings_df = read_input_file(uploaded_file)
        
        if terrain_grid is not None and buildings_df is not None:
            # Créer les objets
            terrain = Terrain(terrain_grid)
            buildings = create_buildings_from_df(buildings_df)
            
            if not buildings:
                st.error("Aucun bâtiment n'a pu être créé. Vérifiez le format des colonnes.")
                return
            
            # Afficher les données chargées
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("🗺️ Terrain")
                terrain_preview = pd.DataFrame(terrain_grid)
                st.dataframe(terrain_preview, use_container_width=True)
                st.caption(f"Dimensions: {terrain.height} lignes x {terrain.width} colonnes")
            
            with col2:
                st.subheader("🏢 Bâtiments à placer")
                buildings_display = []
                for b in buildings:
                    buildings_display.append({
                        "Nom": b.nom,
                        "Type": b.type,
                        "Production": b.production if b.production else "-",
                        "Dimensions": f"{b.longueur}x{b.largeur}",
                        "Quantité": b.quantite,
                        "Culture": b.culture if b.culture > 0 else "-",
                        "Rayon": b.rayonnement if b.rayonnement > 0 else "-"
                    })
                st.dataframe(pd.DataFrame(buildings_display), use_container_width=True)
            
            # Bouton pour lancer l'optimisation
            if st.button("🚀 Lancer l'optimisation", type="primary"):
                with st.spinner("Optimisation en cours..."):
                    # Créer le placeur et placer les bâtiments
                    placer = BuildingPlacer(terrain, buildings)
                    placer.place_all()
                    
                    # Afficher un résumé du placement
                    st.subheader("📋 Résumé du placement")
                    placement_summary = []
                    for b in buildings:
                        placement_summary.append({
                            "Bâtiment": b.nom,
                            "Type": b.type,
                            "Placés": b.placed,
                            "À placer": b.quantite,
                            "Statut": "✅ Complet" if b.placed == b.quantite else f"⚠️ {b.placed}/{b.quantite}"
                        })
                    st.dataframe(pd.DataFrame(placement_summary), use_container_width=True)
                    
                    # Calculer les résultats seulement si des bâtiments sont placés
                    if terrain.buildings:
                        boost_results, total_culture, boost_counts = terrain.get_production_boosts()
                        
                        # Afficher les résultats
                        st.subheader("📊 Résultats de l'optimisation")
                        
                        # Métriques principales
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            total_buildings_placed = sum(b.placed for b in buildings)
                            total_buildings = sum(b.quantite for b in buildings)
                            st.metric("Bâtiments placés", f"{total_buildings_placed}/{total_buildings}")
                        
                        with col2:
                            total_culture_sum = sum(total_culture.values())
                            st.metric("Culture totale distribuée", f"{total_culture_sum:.0f}")
                        
                        with col3:
                            total_producteurs_places = sum(1 for b in buildings if b.type == "producteur" and b.production and b.placed > 0)
                            st.metric("Bâtiments producteurs placés", total_producteurs_places)
                        
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
                        
                        terrain_viz[terrain_viz == '1'] = '⬜'
                        terrain_viz[terrain_viz == '0'] = '⬛'
                        
                        st.dataframe(pd.DataFrame(terrain_viz), use_container_width=True)
                        
                        # Légende des bâtiments
                        if building_colors:
                            st.subheader("🏷️ Légende des bâtiments")
                            legend_cols = st.columns(4)
                            for i, (building_name, symbol) in enumerate(building_colors.items()):
                                with legend_cols[i % 4]:
                                    st.markdown(f"{symbol} {building_name}")
                        
                        # Tableau des boosts
                        if boost_results:
                            st.subheader("📈 Détail des boosts par bâtiment")
                            boost_df = pd.DataFrame(boost_results)
                            st.dataframe(boost_df, use_container_width=True)
                        
                        # Statistiques par type
                        if total_culture:
                            st.subheader("📊 Statistiques par type de production")
                            stats_data = []
                            for prod_type, culture in total_culture.items():
                                if prod_type in boost_counts:
                                    stats_data.append({
                                        "Type": prod_type,
                                        "Culture totale": f"{culture:.2f}",
                                        "Boost 0%": boost_counts[prod_type][0],
                                        "Boost 25%": boost_counts[prod_type][25],
                                        "Boost 50%": boost_counts[prod_type][50],
                                        "Boost 100%": boost_counts[prod_type][100]
                                    })
                            
                            if stats_data:
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
                        st.error("Aucun bâtiment n'a pu être placé!")
    
    else:
        st.info("👆 Veuillez charger un fichier Excel pour commencer.")

if __name__ == "__main__":
    main()
    