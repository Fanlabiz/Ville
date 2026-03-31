"""
Optimiseur de placement de bâtiments pour jeu de ville
Compatible avec Streamlit et export Excel
Version corrigée - Gestion des valeurs NaN
"""

import pandas as pd
import numpy as np
from itertools import product
from collections import defaultdict
import copy
import streamlit as st
import io
import math

class Building:
    """Classe représentant un bâtiment"""
    def __init__(self, nom, longueur, largeur, nombre, type_bat, culture, rayonnement,
                 boost_25, boost_50, boost_100, production, quantite, priorite):
        self.nom = str(nom) if pd.notna(nom) else ""
        self.longueur = self._safe_int(longueur)
        self.largeur = self._safe_int(largeur)
        self.nombre = self._safe_int(nombre)
        self.type = str(type_bat) if pd.notna(type_bat) else ""
        self.culture = self._safe_int(culture) if pd.notna(culture) and culture != "Rien" else 0
        self.rayonnement = self._safe_int(rayonnement)
        self.boost_25 = self._safe_int(boost_25)
        self.boost_50 = self._safe_int(boost_50)
        self.boost_100 = self._safe_int(boost_100)
        self.production = str(production) if pd.notna(production) else ""
        self.quantite = self._safe_int(quantite)
        self.priorite = self._safe_int(priorite)
        
        # Pour le placement
        self.instances = []
    
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
    
    def get_rayonnement_cells(self):
        """Retourne les cases dans le rayonnement du bâtiment"""
        if self.building.rayonnement <= 0:
            return []
        cells = []
        l, L = self.building.longueur, self.building.largeur
        if l <= 0 or L <= 0:
            return cells
        if self.orientation == 'h':
            for i in range(-self.building.rayonnement, l + self.building.rayonnement):
                for j in range(-self.building.rayonnement, L + self.building.rayonnement):
                    if i < 0 or i >= l or j < 0 or j >= L:  # hors du bâtiment
                        cells.append((self.x + i, self.y + j))
        else:
            for i in range(-self.building.rayonnement, L + self.building.rayonnement):
                for j in range(-self.building.rayonnement, l + self.building.rayonnement):
                    if i < 0 or i >= L or j < 0 or j >= l:
                        cells.append((self.x + i, self.y + j))
        return cells
    
    def set_culture(self, culture_value):
        self.culture_recue = culture_value
        self._update_boost()
    
    def _update_boost(self):
        """Met à jour le pourcentage de boost en fonction de la culture reçue"""
        thresholds = self.building.get_boost_thresholds()
        if self.culture_recue >= thresholds[2] and thresholds[2] > 0:
            self.boost_percentage = 100
        elif self.culture_recue >= thresholds[1] and thresholds[1] > 0:
            self.boost_percentage = 50
        elif self.culture_recue >= thresholds[0] and thresholds[0] > 0:
            self.boost_percentage = 25
        else:
            self.boost_percentage = 0
    
    def get_production(self):
        """Calcule la production horaire"""
        return self.building.quantite * (1 + self.boost_percentage / 100)

class Terrain:
    """Classe représentant le terrain de jeu"""
    def __init__(self, df_terrain):
        self.df = df_terrain
        self.grid = self._parse_terrain(df_terrain)
        self.height, self.width = self.grid.shape
        self.building_instances = []
        
    def _parse_terrain(self, df):
        """Parse le fichier terrain pour identifier les cases libres/occupées"""
        # Convertir en numpy array
        arr = df.fillna('').values
        
        # Trouver les limites (cases avec X)
        x_rows = []
        x_cols = []
        for i in range(arr.shape[0]):
            for j in range(arr.shape[1]):
                cell_value = str(arr[i, j]).strip().upper() if arr[i, j] != '' else ''
                if cell_value == 'X':
                    x_rows.append(i)
                    x_cols.append(j)
        
        if x_rows:
            min_row, max_row = min(x_rows), max(x_rows)
            min_col, max_col = min(x_cols), max(x_cols)
            # Extraire la zone intérieure (sans les bordures X)
            interior = arr[min_row+1:max_row, min_col+1:max_col]
            return interior
        return arr
    
    def is_within_bounds(self, cells):
        """Vérifie si toutes les cases sont dans les limites du terrain"""
        for x, y in cells:
            if x < 0 or x >= self.height or y < 0 or y >= self.width:
                return False
        return True
    
    def is_occupied(self, cells, exclude_instance=None):
        """Vérifie si les cases sont déjà occupées"""
        for x, y in cells:
            cell_value = str(self.grid[x, y]).strip() if self.grid[x, y] != '' else ''
            if cell_value != '' and cell_value != ' ':
                # Vérifier si c'est le bâtiment qu'on exclut
                if exclude_instance:
                    for cell in exclude_instance.get_occupied_cells():
                        if (x, y) == cell:
                            continue
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
        
        # Marquer les cases comme occupées
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
        # Récupérer tous les bâtiments culturels
        cultural_buildings = [b for b in self.building_instances if b.building.type == 'Culturel']
        
        # Pour chaque bâtiment producteur, calculer la culture reçue
        for producteur in self.building_instances:
            if producteur.building.type == 'Producteur':
                total_culture = 0
                producteur_cells = set(producteur.get_occupied_cells())
                
                for culturel in cultural_buildings:
                    rayonnement_cells = set(culturel.get_rayonnement_cells())
                    # Vérifier l'intersection
                    if producteur_cells & rayonnement_cells:
                        total_culture += culturel.building.get_culture()
                
                producteur.set_culture(total_culture)
    
    def get_production_stats(self):
        """Calcule les statistiques de production"""
        stats = defaultdict(lambda: {'total': 0, 'quantite': 0, 'boost_25': 0, 'boost_50': 0, 'boost_100': 0})
        
        for building in self.building_instances:
            if building.building.type == 'Producteur':
                prod_type = building.building.get_production_type()
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

class Optimizer:
    """Classe d'optimisation du placement"""
    def __init__(self, terrain, buildings):
        self.terrain = terrain
        self.buildings = buildings
        self.moves = []  # Liste des déplacements effectués
        
    def get_all_positions(self, building):
        """Génère toutes les positions possibles pour un bâtiment"""
        positions = []
        l, L = building.get_dimensions()
        
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
        if l != L:  # Si ce n'est pas un carré
            for x in range(self.terrain.height - L + 1):
                for y in range(self.terrain.width - l + 1):
                    instance = BuildingInstance(building, x, y, 'v')
                    cells = instance.get_occupied_cells()
                    if cells and self.terrain.is_within_bounds(cells) and not self.terrain.is_occupied(cells):
                        positions.append(('v', x, y))
        
        return positions
    
    def evaluate_position(self, building_instance, cultural_instances):
        """Évalue la qualité d'une position pour un bâtiment"""
        # Pour les producteurs, on calcule la culture qu'ils recevraient
        if building_instance.building.type == 'Producteur':
            total_culture = 0
            producteur_cells = set(building_instance.get_occupied_cells())
            
            for culturel in cultural_instances:
                rayonnement_cells = set(culturel.get_rayonnement_cells())
                if producteur_cells & rayonnement_cells:
                    total_culture += culturel.building.get_culture()
            
            thresholds = building_instance.building.get_boost_thresholds()
            if total_culture >= thresholds[2] and thresholds[2] > 0:
                score = 3  # Boost 100%
            elif total_culture >= thresholds[1] and thresholds[1] > 0:
                score = 2  # Boost 50%
            elif total_culture >= thresholds[0] and thresholds[0] > 0:
                score = 1  # Boost 25%
            else:
                score = 0
            
            # Priorité pour les types de production importants
            priority_map = {'Guérison': 100, 'Nourriture': 10, 'Or': 1}
            priority = priority_map.get(building_instance.building.get_production_type(), 0)
            
            return score * 100 + priority + building_instance.building.priorite
        
        return 0
    
    def optimize(self):
        """Exécute l'optimisation du placement"""
        # Filtrer les bâtiments valides
        valid_buildings = [b for b in self.buildings if b.longueur > 0 and b.largeur > 0 and b.nombre > 0]
        
        # Séparer les bâtiments culturels et producteurs
        cultural_buildings = [b for b in valid_buildings if b.type == 'Culturel']
        producteur_buildings = [b for b in valid_buildings if b.type == 'Producteur']
        
        # D'abord placer les bâtiments culturels
        cultural_instances = []
        for building in cultural_buildings:
            for _ in range(building.nombre):
                positions = self.get_all_positions(building)
                if positions:
                    # Prendre la première position disponible
                    orientation, x, y = positions[0]
                    instance = BuildingInstance(building, x, y, orientation)
                    if self.terrain.place_building(instance):
                        cultural_instances.append(instance)
        
        # Ensuite optimiser le placement des producteurs
        for building in producteur_buildings:
            for i in range(building.nombre):
                best_score = -1
                best_position = None
                
                # Évaluer toutes les positions possibles
                positions = self.get_all_positions(building)
                for orientation, x, y in positions:
                    instance = BuildingInstance(building, x, y, orientation)
                    score = self.evaluate_position(instance, cultural_instances)
                    if score > best_score:
                        best_score = score
                        best_position = (orientation, x, y)
                
                # Placer le bâtiment à la meilleure position
                if best_position:
                    orientation, x, y = best_position
                    instance = BuildingInstance(building, x, y, orientation)
                    self.terrain.place_building(instance)
    
    def record_move(self, instance, old_x, old_y, old_orientation):
        """Enregistre un déplacement pour le rapport"""
        self.moves.append({
            'nom': instance.building.nom,
            'avant_x': old_x,
            'avant_y': old_y,
            'avant_orientation': old_orientation,
            'apres_x': instance.x,
            'apres_y': instance.y,
            'apres_orientation': instance.orientation
        })

class ExcelProcessor:
    """Classe pour traiter les fichiers Excel"""
    
    @staticmethod
    def load_terrain(file):
        """Charge le terrain depuis un fichier Excel"""
        try:
            df = pd.read_excel(file, sheet_name='Terrain', header=None)
            return Terrain(df)
        except Exception as e:
            raise Exception(f"Erreur lors du chargement de l'onglet 'Terrain': {str(e)}")
    
    @staticmethod
    def load_buildings(file):
        """Charge les bâtiments depuis un fichier Excel"""
        try:
            df = pd.read_excel(file, sheet_name='Batiments')
            buildings = []
            
            # Vérifier que les colonnes nécessaires existent
            required_columns = ['Nom', 'Longueur', 'Largeur', 'Nombre', 'Type', 'Culture', 
                               'Rayonnement', 'Boost 25%', 'Boost 50%', 'Boost 100%', 
                               'Production', 'Quantite', 'Priorite']
            
            missing_cols = [col for col in required_columns if col not in df.columns]
            if missing_cols:
                raise Exception(f"Colonnes manquantes dans l'onglet 'Batiments': {missing_cols}")
            
            for idx, row in df.iterrows():
                try:
                    building = Building(
                        row['Nom'], row['Longueur'], row['Largeur'], row['Nombre'],
                        row['Type'], row['Culture'], row['Rayonnement'],
                        row['Boost 25%'], row['Boost 50%'], row['Boost 100%'],
                        row['Production'], row['Quantite'], row['Priorite']
                    )
                    if building.longueur > 0 and building.largeur > 0:
                        buildings.append(building)
                except Exception as e:
                    st.warning(f"Ligne {idx + 2}: Impossible de charger le bâtiment - {str(e)}")
            
            return buildings
        except Exception as e:
            raise Exception(f"Erreur lors du chargement de l'onglet 'Batiments': {str(e)}")
    
    @staticmethod
    def export_results(terrain, optimizer, original_terrain):
        """Exporte les résultats dans un fichier Excel"""
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 1. Liste des bâtiments placés
            buildings_data = []
            for instance in terrain.building_instances:
                buildings_data.append({
                    'Nom': instance.building.nom,
                    'Type': instance.building.type,
                    'Production': instance.building.production,
                    'X': instance.x,
                    'Y': instance.y,
                    'Orientation': 'Horizontal' if instance.orientation == 'h' else 'Vertical',
                    'Culture recue': instance.culture_recue,
                    'Boost': f"{instance.boost_percentage}%",
                    'Production/heure': f"{instance.get_production():.2f}"
                })
            df_buildings = pd.DataFrame(buildings_data)
            df_buildings.to_excel(writer, sheet_name='Batiments_places', index=False)
            
            # 2. Statistiques de production
            stats = terrain.get_production_stats()
            stats_data = []
            for prod_type, data in stats.items():
                if prod_type and prod_type != '':
                    stats_data.append({
                        'Type production': prod_type,
                        'Quantité totale produite/heure': round(data['total'], 2),
                        'Quantité de base': data['quantite'],
                        'Nb boost 25%': data['boost_25'],
                        'Nb boost 50%': data['boost_50'],
                        'Nb boost 100%': data['boost_100']
                    })
            if stats_data:
                df_stats = pd.DataFrame(stats_data)
                df_stats.to_excel(writer, sheet_name='Stats_production', index=False)
            
            # 3. Gains/pertes par type de production
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
                df_gains = pd.DataFrame(gains_data)
                df_gains.to_excel(writer, sheet_name='Gains_pertes', index=False)
            
            # 4. Liste des bâtiments déplacés
            if optimizer.moves:
                df_moves = pd.DataFrame(optimizer.moves)
                df_moves.to_excel(writer, sheet_name='Batiments_deplaces', index=False)
            
            # 5. Séquence des opérations
            operations = []
            for i, move in enumerate(optimizer.moves, 1):
                operations.append({
                    'Étape': i,
                    'Action': f"Déplacer {move['nom']}",
                    'De': f"({move['avant_x']}, {move['avant_y']})",
                    'Vers': f"({move['apres_x']}, {move['apres_y']})",
                    'Orientation': move['apres_orientation']
                })
            if operations:
                df_ops = pd.DataFrame(operations)
                df_ops.to_excel(writer, sheet_name='Sequence_operations', index=False)
            
            # 6. Export du terrain final
            terrain_df = pd.DataFrame(terrain.grid)
            terrain_df.to_excel(writer, sheet_name='Terrain_final', index=False, header=False)
        
        output.seek(0)
        return output.getvalue()

def main():
    """Fonction principale pour Streamlit"""
    st.set_page_config(page_title="Optimiseur de placement de bâtiments", layout="wide")
    
    st.title("🏗️ Optimiseur de placement de bâtiments")
    st.markdown("""
    Cette application optimise le placement des bâtiments sur votre terrain pour maximiser la production.
    Les bâtiments culturels sont prioritaires pour le placement, suivis des producteurs optimisés pour recevoir le maximum de boosts.
    """)
    
    # Upload du fichier
    uploaded_file = st.file_uploader(
        "Choisissez votre fichier Excel",
        type=['xlsx', 'xls'],
        help="Le fichier doit contenir les onglets 'Terrain' et 'Batiments' avec le format attendu"
    )
    
    if uploaded_file is not None:
        try:
            # Chargement des données
            with st.spinner("Chargement des données..."):
                terrain = ExcelProcessor.load_terrain(uploaded_file)
                buildings = ExcelProcessor.load_buildings(uploaded_file)
                
                st.success(f"✅ Terrain chargé: {terrain.height} x {terrain.width} cases")
                st.success(f"✅ {len(buildings)} types de bâtiments chargés")
                
                # Afficher un résumé des bâtiments
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
            
            # Optimisation
            with st.spinner("Optimisation du placement en cours..."):
                optimizer = Optimizer(terrain, buildings)
                optimizer.optimize()
                
                # Calcul de la culture et des boosts
                terrain.calculate_culture()
                
                st.success("✅ Optimisation terminée!")
                
                # Afficher les résultats
                stats = terrain.get_production_stats()
                
                if stats:
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.subheader("📊 Résultats de l'optimisation")
                        stats_df = pd.DataFrame([
                            {
                                'Type': k,
                                'Production/h': f"{v['total']:.2f}",
                                'Base': v['quantite'],
                                'Boost 25%': v['boost_25'],
                                'Boost 50%': v['boost_50'],
                                'Boost 100%': v['boost_100']
                            }
                            for k, v in stats.items()
                        ])
                        st.dataframe(stats_df)
                    
                    with col2:
                        st.subheader("📈 Gains par type")
                        gains_df = pd.DataFrame([
                            {
                                'Type': k,
                                'Gain': f"{v['total'] - v['quantite']:.2f}",
                                '% d\'augmentation': f"{((v['total'] - v['quantite']) / v['quantite'] * 100):.1f}%" if v['quantite'] > 0 else "N/A"
                            }
                            for k, v in stats.items()
                        ])
                        st.dataframe(gains_df)
                else:
                    st.warning("Aucun bâtiment producteur trouvé dans les données.")
                
                # Afficher les bâtiments placés
                if terrain.building_instances:
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
                
                # Export des résultats
                st.subheader("📥 Export des résultats")
                
                excel_data = ExcelProcessor.export_results(terrain, optimizer, None)
                
                st.download_button(
                    label="📎 Télécharger le fichier Excel des résultats",
                    data=excel_data,
                    file_name="resultats_optimisation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Instructions
                st.markdown("---")
                st.markdown("""
                ### 📝 Instructions
                1. Téléchargez le fichier Excel des résultats
                2. Ouvrez-le sur votre iPad
                3. Consultez les différents onglets pour voir:
                   - La liste détaillée des bâtiments placés
                   - Les statistiques de production
                   - Les gains par type de production
                   - La séquence des déplacements à effectuer
                4. Utilisez la séquence d'opérations pour réaliser les déplacements dans votre jeu
                """)
                
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement: {str(e)}")
            st.info("""
            **Vérifications :**
            - Votre fichier doit contenir un onglet nommé 'Terrain'
            - Votre fichier doit contenir un onglet nommé 'Batiments'
            - L'onglet 'Batiments' doit contenir les colonnes: Nom, Longueur, Largeur, Nombre, Type, Culture, Rayonnement, Boost 25%, Boost 50%, Boost 100%, Production, Quantite, Priorite
            - Les cases vides sont acceptées (elles seront remplacées par 0)
            """)

if __name__ == "__main__":
    main()
