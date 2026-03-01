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
        self.cultural_zones = np.zeros_like(self.grid, dtype=float)
        self.building_type_map = np.full_like(self.grid, '', dtype=object)
        
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
        
    def calculate_cultural_effect(self):
        """Calcule l'effet culturel de tous les bâtiments"""
        self.cultural_zones = np.zeros_like(self.grid, dtype=float)
        
        for building, x, y, orientation, longueur, largeur in self.buildings:
            if building.type == "culturel" and building.culture > 0:
                center_x = x + longueur // 2
                center_y = y + largeur // 2
                
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
        total_culture_by_type = defaultdict(float)
        boost_counts = defaultdict(lambda: {0: 0, 25: 0, 50: 0, 100: 0})
        
        for building, x, y, orientation, longueur, largeur in self.buildings:
            if building.type == "producteur" and building.production:
                prod_type = building.production.strip()
                if not prod_type:
                    continue
                    
                # Calculer la culture reçue
                total_culture = 0
                for i in range(longueur):
                    for j in range(largeur):
                        total_culture += self.cultural_zones[y + j, x + i]
                
                avg_culture = total_culture / (longueur * largeur) if (longueur * largeur) > 0 else 0
                total_culture_by_type[prod_type] += avg_culture
                
                # Déterminer le boost
                boost = 0
                if building.boost_100 > 0 and avg_culture >= building.boost_100:
                    boost = 100
                elif building.boost_50 > 0 and avg_culture >= building.boost_50:
                    boost = 50
                elif building.boost_25 > 0 and avg_culture >= building.boost_25:
                    boost = 25
                    
                boost_counts[prod_type][boost] += 1
                
                results.append({
                    "Nom": building.nom,
                    "Production": building.production,
                    "Culture reçue": round(avg_culture, 2),
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
            elif any(x in prod for x in ['bijoux', 'onguents', 'cristal', 'epices', 'boiseries', 'scriberie']):
                return 5
            return 4
        elif building.type == "culturel":
            return 6
        return 7
    
    def place_all(self):
        """Place tous les bâtiments en optimisant les boosts"""
        
        # Filtrer les bâtiments avec quantité > 0
        valid_buildings = [b for b in self.buildings if b.quantite > 0]
        
        # Séparer par type
        cultural_buildings = [b for b in valid_buildings if b.type == "culturel"]
        producer_buildings = [b for b in valid_buildings if b.type == "producteur" and b.production]
        other_buildings = [b for b in valid_buildings if b not in cultural_buildings and b not in producer_buildings]
        
        # Trier par priorité
        cultural_buildings.sort(key=lambda b: (-b.culture, -b.rayonnement))
        producer_buildings.sort(key=lambda b: (self.get_priority_score(b), -b.longueur * b.largeur))
        
        # Statistiques
        total_a_placer = sum(b.quantite for b in valid_buildings)
        stats = {"culturels": 0, "producteurs": 0, "autres": 0, "total": 0}
        
        # Fonction pour placer plusieurs exemplaires
        def place_multiple(buildings_list, category):
            for building in buildings_list:
                if building.quantite == 0:
                    continue
                    
                st.write(f"📦 {building.nom} ({building.quantite} exemplaires)")
                
                for i in range(building.quantite):
                    if building.placed >= building.quantite:
                        break
                        
                    success, reason = self.try_place_building_with_reason(building)
                    if success:
                        stats[category] += 1
                        stats["total"] += 1
                        st.write(f"  ✅ Exemplaire {i+1}/{building.quantite} placé")
                    else:
                        building.failed_attempts.append(f"Exemplaire {i+1}: {reason}")
                        st.write(f"  ❌ Exemplaire {i+1}/{building.quantite} non placé - {reason}")
                        # On continue avec les autres bâtiments, on ne bloque pas
        
        # 1. Placer les culturels d'abord
        if cultural_buildings:
            st.subheader("🏛️ Placement des bâtiments culturels")
            place_multiple(cultural_buildings, "culturels")
            
            # Recalculer l'effet culturel
            self.terrain.calculate_cultural_effect()
        
        # 2. Placer les producteurs
        if producer_buildings:
            st.subheader("🏭 Placement des bâtiments producteurs")
            place_multiple(producer_buildings, "producteurs")
        
        # 3. Placer les autres
        if other_buildings:
            st.subheader("📦 Placement des autres bâtiments")
            place_multiple(other_buildings, "autres")
        
        # Afficher le résumé
        st.subheader("📊 Résumé du placement")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Bâtiments placés", f"{stats['total']}/{total_a_placer}")
        with col2:
            taux = (stats['total'] / total_a_placer * 100) if total_a_placer > 0 else 0
            st.metric("Taux de placement", f"{taux:.1f}%")
        with col3:
            st.metric("Bâtiments culturels", f"{stats['culturels']}/{sum(b.quantite for b in cultural_buildings)}")
        
        return stats
    
    def try_place_building_with_reason(self, building):
        """Essaie de placer un bâtiment et retourne (succès, raison)"""
        best_position = self.find_best_position(building)
        if best_position:
            x, y, orientation = best_position
            self.terrain.place_building(building, x, y, orientation)
            return True, "OK"
        
        # Si aucune position trouvée, essayer de donner une raison
        # Vérifier d'abord si c'est un problème de taille
        for orientation in ['H', 'V']:
            longueur, largeur = building.get_dimensions(orientation)
            if longueur > self.terrain.width or largeur > self.terrain.height:
                return False, f"Trop grand ({longueur}x{largeur}) pour le terrain ({self.terrain.width}x{self.terrain.height})"
        
        return False, "Aucun emplacement libre disponible"
    
    def find_best_position(self, building):
        """Trouve la meilleure position pour un bâtiment"""
        best_score = -1
        best_position = None
        
        # Essayer les deux orientations
        for orientation in ['H', 'V']:
            longueur, largeur = building.get_dimensions(orientation)
            
            # Vérifier si le bâtiment tient dans le terrain
            if longueur > self.terrain.width or largeur > self.terrain.height:
                continue
            
            for y in range(self.terrain.height - largeur + 1):
                for x in range(self.terrain.width - longueur + 1):
                    can_place, _ = self.terrain.can_place(x, y, longueur, largeur)
                    if can_place:
                        score = self.evaluate_position(building, x, y, orientation)
                        if score > best_score:
                            best_score = score
                            best_position = (x, y, orientation)
        
        return best_position
    
    def evaluate_position(self, building, x, y, orientation):
        """Évalue la qualité d'une position pour un bâtiment"""
        longueur, largeur = building.get_dimensions(orientation)
        
        if building.type == "producteur" and building.production:
            # Calculer la culture potentielle
            total_culture = 0
            for i in range(longueur):
                for j in range(largeur):
                    total_culture += self.terrain.cultural_zones[y + j, x + i]
            
            avg_culture = total_culture / (longueur * largeur) if (longueur * largeur) > 0 else 0
            
            # Score basé sur le boost potentiel
            if building.boost_100 > 0 and avg_culture >= building.boost_100:
                return 1000 + avg_culture
            elif building.boost_50 > 0 and avg_culture >= building.boost_50:
                return 500 + avg_culture
            elif building.boost_25 > 0 and avg_culture >= building.boost_25:
                return 250 + avg_culture
            else:
                return avg_culture
        
        elif building.type == "culturel":
            # Pour les culturels, privilégier les positions avec de l'espace autour
            space_score = 0
            for i in range(max(0, x - 2), min(self.terrain.width, x + longueur + 2)):
                for j in range(max(0, y - 2), min(self.terrain.height, y + largeur + 2)):
                    if i >= x and i < x + longueur and j >= y and j < y + largeur:
                        continue
                    if self.terrain.grid[j, i] == 1 and not self.terrain.occupied[j, i]:
                        space_score += 1
            
            return space_score
        
        return 1

def normalize_column_name(name):
    """Normalise les noms de colonnes"""
    if pd.isna(name):
        return ""
    
    name = str(name).strip().lower()
    
    # Remplacer les caractères accentués
    replacements = {
        'é': 'e', 'è': 'e', 'ê': 'e', 'ë': 'e',
        'à': 'a', 'â': 'a', 'ä': 'a',
        'î': 'i', 'ï': 'i',
        'ô': 'o', 'ö': 'o',
        'ù': 'u', 'û': 'u', 'ü': 'u',
        'ç': 'c'
    }
    
    for accented, unaccented in replacements.items():
        name = name.replace(accented, unaccented)
    
    # Enlever les caractères spéciaux
    name = re.sub(r'[^a-zA-Z0-9%]', '', name)
    
    return name

def read_input_file(file):
    """Lit le fichier Excel d'entrée"""
    try:
        xl = pd.ExcelFile(file)
        
        # Lire le terrain
        terrain_df = pd.read_excel(xl, sheet_name=0, header=None)
        terrain_grid = terrain_df.values.tolist()
        
        # Lire les bâtiments
        buildings_df = pd.read_excel(xl, sheet_name=1)
        
        # Normaliser les noms de colonnes
        buildings_df.columns = [normalize_column_name(col) for col in buildings_df.columns]
        
        return terrain_grid, buildings_df
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier: {e}")
        return None, None

def create_buildings_from_df(df):
    """Crée les objets Building à partir du DataFrame"""
    
    # Mapping des colonnes
    col_mapping = {
        'nom': ['nom', 'name'],
        'longueur': ['longueur', 'long', 'length'],
        'largeur': ['largeur', 'larg', 'width'],
        'quantite': ['quantite', 'quantité', 'quantity', 'qt', 'qte', 'nb', 'nombre'],
        'type': ['type'],
        'culture': ['culture', 'cult'],
        'rayonnement': ['rayonnement', 'rayon', 'radius'],
        'boost25': ['boost25%', 'boost25', '25%'],
        'boost50': ['boost50%', 'boost50', '50%'],
        'boost100': ['boost100%', 'boost100', '100%'],
        'production': ['production', 'prod']
    }
    
    def find_col(possible_names):
        for name in possible_names:
            for col in df.columns:
                if name in col:
                    return col
        return None
    
    # Trouver les colonnes
    nom_col = find_col(col_mapping['nom'])
    longueur_col = find_col(col_mapping['longueur'])
    largeur_col = find_col(col_mapping['largeur'])
    quantite_col = find_col(col_mapping['quantite'])
    type_col = find_col(col_mapping['type'])
    culture_col = find_col(col_mapping['culture'])
    rayonnement_col = find_col(col_mapping['rayonnement'])
    boost25_col = find_col(col_mapping['boost25'])
    boost50_col = find_col(col_mapping['boost50'])
    boost100_col = find_col(col_mapping['boost100'])
    production_col = find_col(col_mapping['production'])
    
    if not nom_col:
        st.error("Colonne 'Nom' non trouvée")
        # Afficher les colonnes disponibles pour debug
        st.write("Colonnes disponibles:", list(df.columns))
        return []
    
    buildings = []
    for idx, row in df.iterrows():
        try:
            # Récupérer la quantité
            quantite = 1
            if quantite_col:
                val = row[quantite_col]
                if pd.notna(val) and val != '':
                    try:
                        quantite = int(float(val))
                    except:
                        quantite = 1
            
            # Ignorer si quantité = 0
            if quantite == 0:
                continue
            
            # Récupérer le type
            type_bat = ""
            if type_col:
                val = row[type_col]
                if pd.notna(val) and val != '':
                    type_bat = str(val)
            
            # Récupérer la production
            production = ""
            if production_col:
                val = row[production_col]
                if pd.notna(val) and val != '':
                    production = str(val)
            
            building = Building(
                nom=row[nom_col],
                longueur=row[longueur_col] if longueur_col else 1,
                largeur=row[largeur_col] if largeur_col else 1,
                quantite=quantite,
                type_bat=type_bat,
                culture=row[culture_col] if culture_col else 0,
                rayonnement=row[rayonnement_col] if rayonnement_col else 0,
                boost_25=row[boost25_col] if boost25_col else 0,
                boost_50=row[boost50_col] if boost50_col else 0,
                boost_100=row[boost100_col] if boost100_col else 0,
                production=production
            )
            buildings.append(building)
        except Exception as e:
            st.warning(f"Erreur sur la ligne {idx+1}: {e}")
            continue
    
    return buildings

def create_output_excel(terrain, boost_results, total_culture, boost_counts, unplaced_buildings):
    """Crée le fichier Excel de sortie"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Feuille 1: Terrain avec bâtiments placés
        terrain_display = np.array(terrain.grid, dtype=object)
        
        # Créer une carte des bâtiments
        building_map = {}
        color_index = 1
        
        for building, x, y, orientation, longueur, largeur in terrain.buildings:
            if building.nom not in building_map:
                # Créer un code court
                words = building.nom.split()
                if len(words) > 1:
                    short_name = ''.join(w[0].upper() for w in words[:2])
                else:
                    short_name = building.nom[:3].upper()
                building_map[building.nom] = f"{short_name}_{color_index}"
                color_index += 1
            
            for i in range(longueur):
                for j in range(largeur):
                    terrain_display[y + j, x + i] = building_map[building.nom]
        
        terrain_df = pd.DataFrame(terrain_display)
        terrain_df.to_excel(writer, sheet_name='Terrain avec batiments', index=False, header=False)
        
        # Appliquer les couleurs
        worksheet = writer.sheets['Terrain avec batiments']
        
        # Définir les couleurs par type
        type_colors = {
            'culturel': 'FFE4B5',  # Beige
            'producteur': 'ADD8E6',  # Bleu clair
        }
        
        # Ajuster la largeur des colonnes
        for i, col in enumerate(worksheet.columns):
            worksheet.column_dimensions[col[0].column_letter].width = 5
        
        # Colorer les cellules
        for building, x, y, orientation, longueur, largeur in terrain.buildings:
            fill_color = type_colors.get(building.type, 'D3D3D3')
            fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
            
            for i in range(longueur):
                for j in range(largeur):
                    cell = worksheet.cell(row=y + j + 1, column=x + i + 1)
                    cell.fill = fill
                    cell.font = Font(bold=True, size=8)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Feuille 2: Légende
        legend_data = []
        for building_name, short_code in building_map.items():
            building = next(b[0] for b in terrain.buildings if b[0].nom == building_name)
            legend_data.append({
                "Code": short_code,
                "Nom": building_name,
                "Type": building.type,
                "Production": building.production if building.production else "-"
            })
        
        if legend_data:
            legend_df = pd.DataFrame(legend_data)
            legend_df.to_excel(writer, sheet_name='Legende', index=False)
        
        # Feuille 3: Résultats des boosts
        if boost_results:
            boost_df = pd.DataFrame(boost_results)
            boost_df.to_excel(writer, sheet_name='Boosts de production', index=False)
        
        # Feuille 4: Statistiques
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
        
        # Feuille 5: Positions
        positions_data = []
        for building, x, y, orientation, longueur, largeur in terrain.buildings:
            positions_data.append({
                "Nom": building.nom,
                "Type": building.type,
                "Production": building.production if building.production else "-",
                "Position X": x,
                "Position Y": y,
                "Orientation": orientation,
                "Longueur": longueur,
                "Largeur": largeur,
                "Culture": building.culture if building.culture > 0 else "-",
                "Rayonnement": building.rayonnement if building.rayonnement > 0 else "-"
            })
        
        if positions_data:
            positions_df = pd.DataFrame(positions_data)
            positions_df.to_excel(writer, sheet_name='Positions', index=False)
        
        # Feuille 6: Bâtiments non placés avec raisons
        if unplaced_buildings:
            unplaced_df = pd.DataFrame(unplaced_buildings)
            unplaced_df.to_excel(writer, sheet_name='Non places', index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Placeur de bâtiments optimisé", layout="wide")
    
    st.title("🏗️ Placeur de bâtiments optimisé")
    st.markdown("""
    Cette application optimise le placement de bâtiments sur un terrain en respectant l'ordre de priorité :
    1. **Guérison** 🏥
    2. **Nourriture** 🌾
    3. **Or** 👑
    4. **Productions spéciales** ✨
    """)
    
    with st.sidebar:
        st.header("📂 Chargement des données")
        uploaded_file = st.file_uploader(
            "Choisir le fichier Excel", 
            type=['xlsx', 'xls']
        )
        
        if uploaded_file:
            st.success("✅ Fichier chargé avec succès!")
    
    if uploaded_file:
        # Lire le fichier
        terrain_grid, buildings_df = read_input_file(uploaded_file)
        
        if terrain_grid is not None and buildings_df is not None:
            # Créer les objets
            terrain = Terrain(terrain_grid)
            buildings = create_buildings_from_df(buildings_df)
            
            if not buildings:
                st.error("❌ Aucun bâtiment n'a pu être créé. Vérifiez le format des colonnes.")
                return
            
            # Afficher les données
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("🗺️ Terrain")
                terrain_preview = pd.DataFrame(terrain_grid)
                st.dataframe(terrain_preview, use_container_width=True)
                
                total_cases = terrain.height * terrain.width
                free_cases = np.sum(terrain.grid == 1)
                st.caption(f"📊 Dimensions: {terrain.height} × {terrain.width} = {total_cases} cases")
                st.caption(f"✅ Cases libres: {free_cases}")
                st.caption(f"❌ Cases occupées: {total_cases - free_cases}")
            
            with col2:
                st.subheader("🏢 Bâtiments à placer")
                
                # Calculer l'espace nécessaire
                total_area_needed = 0
                for b in buildings:
                    if b.quantite > 0:
                        total_area_needed += b.longueur * b.largeur * b.quantite
                
                # Statistiques
                total_batiments = sum(b.quantite for b in buildings)
                total_culturels = sum(b.quantite for b in buildings if b.type == "culturel")
                total_producteurs = sum(b.quantite for b in buildings if b.type == "producteur" and b.production)
                
                st.metric("Total à placer", total_batiments)
                st.metric("Surface nécessaire", f"{total_area_needed} cases")
                st.metric("Surface disponible", free_cases)
                
                if total_area_needed > free_cases:
                    st.warning(f"⚠️ Surface insuffisante! Manque {total_area_needed - free_cases} cases")
                
                st.metric("Bâtiments culturels", total_culturels)
                st.metric("Bâtiments producteurs", total_producteurs)
                
                # Aperçu des bâtiments
                buildings_display = []
                for b in buildings:
                    if b.quantite > 0:
                        buildings_display.append({
                            "Nom": b.nom,
                            "Type": b.type if b.type else "-",
                            "Production": b.production if b.production else "-",
                            "Dimensions": f"{b.longueur}×{b.largeur}",
                            "Quantité": b.quantite,
                            "Culture": f"{b.culture:.0f}" if b.culture > 0 else "-",
                            "Rayon": b.rayonnement if b.rayonnement > 0 else "-"
                        })
                st.dataframe(pd.DataFrame(buildings_display), use_container_width=True)
            
            # Bouton d'optimisation
            if st.button("🚀 Lancer l'optimisation", type="primary"):
                with st.spinner("Optimisation en cours..."):
                    # Placer les bâtiments
                    placer = BuildingPlacer(terrain, buildings)
                    stats = placer.place_all()
                    
                    # Calculer les résultats
                    boost_results, total_culture, boost_counts = terrain.get_production_boosts()
                    
                    # Identifier les bâtiments non placés avec raisons
                    unplaced_buildings = []
                    for b in buildings:
                        if b.placed < b.quantite:
                            reasons = "; ".join(b.failed_attempts) if b.failed_attempts else "Aucun emplacement disponible"
                            unplaced_buildings.append({
                                "Nom": b.nom,
                                "Type": b.type if b.type else "-",
                                "Production": b.production if b.production else "-",
                                "À placer": b.quantite,
                                "Placés": b.placed,
                                "Raison": reasons[:200]  # Limiter la longueur
                            })
                    
                    # Afficher les résultats
                    st.subheader("📊 Résultats de l'optimisation")
                    
                    # Métriques
                    total_a_placer = sum(b.quantite for b in buildings)
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Bâtiments placés", f"{stats['total']}/{total_a_placer}")
                    with col2:
                        taux = (stats['total'] / total_a_placer * 100) if total_a_placer > 0 else 0
                        st.metric("Taux de placement", f"{taux:.1f}%")
                    with col3:
                        total_culture_sum = sum(total_culture.values())
                        st.metric("Culture totale", f"{total_culture_sum:.0f}")
                    
                    # Visualisation
                    st.subheader("🗺️ Terrain avec bâtiments placés")
                    
                    # Créer une visualisation
                    terrain_viz = np.array(terrain.grid, dtype=str)
                    building_symbols = {}
                    
                    # Symboles par type
                    type_symbols = {
                        'culturel': '🔵',
                        'producteur': '🔴'
                    }
                    
                    for building, x, y, orientation, longueur, largeur in terrain.buildings:
                        if building.nom not in building_symbols:
                            symbol = type_symbols.get(building.type, '⚪')
                            building_symbols[building.nom] = symbol
                        
                        for i in range(longueur):
                            for j in range(largeur):
                                terrain_viz[y + j, x + i] = building_symbols[building.nom]
                    
                    terrain_viz[terrain_viz == '1'] = '⬜'
                    terrain_viz[terrain_viz == '0'] = '⬛'
                    
                    st.dataframe(pd.DataFrame(terrain_viz), use_container_width=True)
                    
                    # Légende
                    if building_symbols:
                        st.subheader("🏷️ Légende")
                        legend_cols = st.columns(4)
                        for i, (building_name, symbol) in enumerate(building_symbols.items()):
                            with legend_cols[i % 4]:
                                st.markdown(f"{symbol} {building_name}")
                    
                    # Résultats des boosts
                    if boost_results:
                        st.subheader("📈 Détail des boosts")
                        
                        # Compter les boosts par type
                        boost_summary = defaultdict(lambda: defaultdict(int))
                        for r in boost_results:
                            boost_val = int(r['Boost'].replace('%', ''))
                            boost_summary[r['Production']][boost_val] += 1
                        
                        # Afficher le résumé
                        summary_data = []
                        for prod, boosts in boost_summary.items():
                            summary_data.append({
                                "Production": prod,
                                "0%": boosts.get(0, 0),
                                "25%": boosts.get(25, 0),
                                "50%": boosts.get(50, 0),
                                "100%": boosts.get(100, 0),
                                "Total": sum(boosts.values())
                            })
                        
                        st.dataframe(pd.DataFrame(summary_data), use_container_width=True)
                        
                        # Détail
                        with st.expander("Voir le détail par bâtiment"):
                            boost_df = pd.DataFrame(boost_results)
                            st.dataframe(boost_df, use_container_width=True)
                    
                    # Téléchargement
                    output_file = create_output_excel(
                        terrain, boost_results, total_culture, 
                        boost_counts, unplaced_buildings
                    )
                    
                    st.download_button(
                        label="📥 Télécharger les résultats (Excel)",
                        data=output_file,
                        file_name="resultats_placement.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    
    else:
        st.info("👆 Chargez un fichier Excel pour commencer")
        
        with st.expander("ℹ️ Format du fichier"):
            st.markdown("""
            **Onglet 1: Terrain**
            - Matrice de 0 (occupé) et 1 (libre)
            
            **Onglet 2: Bâtiments**
            - Nom, Longueur, Largeur, Quantité, Type, Culture, Rayonnement
            - Boost 25%, Boost 50%, Boost 100%, Production
            
            La colonne **Quantité** indique le nombre d'exemplaires de chaque bâtiment à placer.
            """)

if __name__ == "__main__":
    main()