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
        self.id = f"{nom}_{id(self)}"  # Identifiant unique
        
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
        self.building_type_map = np.full_like(self.grid, '', dtype=object)  # Pour le type de bâtiment
        
    def can_place(self, x, y, longueur, largeur):
        """Vérifie si un bâtiment peut être placé à la position (x,y)"""
        if x + longueur > self.width or y + largeur > self.height:
            return False
            
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
                return 5  # Priorité plus basse pour les productions spéciales
            return 4
        return 6  # Priorité la plus basse pour les bâtiments culturels
    
    def place_all(self):
        """Place tous les bâtiments en optimisant les boosts"""
        
        # Séparer les bâtiments par type
        cultural_buildings = [b for b in self.buildings if b.type == "culturel"]
        producer_buildings = [b for b in self.buildings if b.type == "producteur" and b.production]
        
        # Trier les bâtiments par priorité
        cultural_buildings.sort(key=lambda b: (-b.culture, -b.rayonnement))
        producer_buildings.sort(key=lambda b: (self.get_priority_score(b), -b.longueur * b.largeur))
        
        # Statistiques
        stats = {"culturels": 0, "producteurs": 0, "total": 0}
        total_culturels = sum(b.quantite for b in cultural_buildings)
        total_producteurs = sum(b.quantite for b in producer_buildings)
        
        # Placer les bâtiments culturels
        if cultural_buildings:
            st.write(f"🏛️ Placement des {total_culturels} bâtiments culturels...")
            progress_bar = st.progress(0)
            for i, building in enumerate(cultural_buildings):
                for _ in range(building.quantite - building.placed):
                    if self.try_place_building(building):
                        stats["culturels"] += 1
                        stats["total"] += 1
                progress_bar.progress((i + 1) / len(cultural_buildings))
        
        # Recalculer l'effet culturel
        self.terrain.calculate_cultural_effect()
        
        # Placer les producteurs
        if producer_buildings:
            st.write(f"🏭 Placement des {total_producteurs} bâtiments producteurs...")
            progress_bar = st.progress(0)
            for i, building in enumerate(producer_buildings):
                for _ in range(building.quantite - building.placed):
                    if self.try_place_building(building):
                        stats["producteurs"] += 1
                        stats["total"] += 1
                progress_bar.progress((i + 1) / len(producer_buildings))
        
        # Vérifier les bâtiments non placés
        unplaced = []
        for b in self.buildings:
            if b.placed < b.quantite:
                unplaced.append(f"{b.nom} ({b.placed}/{b.quantite})")
        
        if unplaced:
            st.warning(f"⚠️ Bâtiments non placés: {', '.join(unplaced)}")
        else:
            st.success(f"✅ Tous les {stats['total']} bâtiments ont été placés!")
        
        return stats
    
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
            # Pour les culturels, privilégier les positions centrales
            center_dist = abs(x - self.terrain.width/2) + abs(y - self.terrain.height/2)
            return -center_dist  # Plus on est proche du centre, mieux c'est
        
        return 0

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
        'quantite': ['quantite', 'quantité', 'quantity', 'qt'],
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
        return []
    
    buildings = []
    for idx, row in df.iterrows():
        try:
            # Ne garder que les bâtiments avec une quantité > 0
            quantite = row[quantite_col] if quantite_col else 1
            if pd.isna(quantite) or quantite == 0:
                continue
                
            building = Building(
                nom=row[nom_col],
                longueur=row[longueur_col] if longueur_col else 1,
                largeur=row[largeur_col] if largeur_col else 1,
                quantite=quantite,
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

def create_output_excel(terrain, boost_results, total_culture, boost_counts, unplaced_buildings):
    """Crée le fichier Excel de sortie avec couleurs par type"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Feuille 1: Terrain avec bâtiments placés
        terrain_display = np.array(terrain.grid, dtype=object)
        
        # Créer une carte des bâtiments avec leurs noms complets
        building_map = {}
        building_colors = {}
        color_index = 1
        
        # Définir les couleurs par type
        type_colors = {
            'culturel': 'FFE4B5',  # Beige
            'producteur': 'ADD8E6',  # Bleu clair
            'default': 'D3D3D3'  # Gris clair
        }
        
        for building, x, y, orientation, longueur, largeur in terrain.buildings:
            # Générer un ID court pour l'affichage
            if building.nom not in building_map:
                # Prendre les premières lettres significatives
                words = building.nom.split()
                if len(words) > 1:
                    short_name = ''.join(w[0].upper() for w in words[:2])
                else:
                    short_name = building.nom[:3].upper()
                building_map[building.nom] = f"{short_name}_{color_index}"
                building_colors[building.nom] = type_colors.get(building.type, type_colors['default'])
                color_index += 1
            
            # Placer le bâtiment
            for i in range(longueur):
                for j in range(largeur):
                    terrain_display[y + j, x + i] = building_map[building.nom]
        
        terrain_df = pd.DataFrame(terrain_display)
        terrain_df.to_excel(writer, sheet_name='Terrain avec batiments', index=False, header=False)
        
        # Appliquer les couleurs
        worksheet = writer.sheets['Terrain avec batiments']
        
        # Ajuster la largeur des colonnes
        for i, col in enumerate(worksheet.columns):
            worksheet.column_dimensions[col[0].column_letter].width = 5
        
        # Colorer les cellules selon le type
        for building, x, y, orientation, longueur, largeur in terrain.buildings:
            fill_color = type_colors.get(building.type, type_colors['default'])
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
        
        # Feuille 6: Bâtiments non placés
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
                
                # Statistiques
                total_batiments = sum(b.quantite for b in buildings)
                total_culturels = sum(b.quantite for b in buildings if b.type == "culturel")
                total_producteurs = sum(b.quantite for b in buildings if b.type == "producteur" and b.production)
                
                st.metric("Total à placer", total_batiments)
                st.metric("Bâtiments culturels", total_culturels)
                st.metric("Bâtiments producteurs", total_producteurs)
                
                # Aperçu des bâtiments
                buildings_display = []
                for b in buildings:
                    buildings_display.append({
                        "Nom": b.nom,
                        "Type": b.type,
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
                    
                    # Identifier les bâtiments non placés
                    unplaced_buildings = []
                    for b in buildings:
                        if b.placed < b.quantite:
                            unplaced_buildings.append({
                                "Nom": b.nom,
                                "Type": b.type,
                                "Production": b.production,
                                "À placer": b.quantite,
                                "Placés": b.placed
                            })
                    
                    # Afficher les résultats
                    st.subheader("📊 Résultats de l'optimisation")
                    
                    # Métriques
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Bâtiments placés", f"{stats['total']}/{total_batiments}")
                    with col2:
                        st.metric("Taux de placement", f"{stats['total']/total_batiments*100:.1f}%")
                    with col3:
                        total_culture_sum = sum(total_culture.values())
                        st.metric("Culture totale", f"{total_culture_sum:.0f}")
                    
                    # Visualisation
                    st.subheader("🗺️ Terrain avec bâtiments placés")
                    
                    # Créer une visualisation
                    terrain_viz = np.array(terrain.grid, dtype=str)
                    building_symbols = {}
                    colors = ['🟦', '🟥', '🟩', '🟨', '🟪', '🟫', '🟧']
                    
                    # Symboles par type
                    type_symbols = {
                        'culturel': '🔵',
                        'producteur': '🔴'
                    }
                    
                    color_idx = 0
                    for building, x, y, orientation, longueur, largeur in terrain.buildings:
                        if building.nom not in building_symbols:
                            symbol = type_symbols.get(building.type, '⬜')
                            building_symbols[building.nom] = symbol
                            color_idx += 1
                        
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
            """)

if __name__ == "__main__":
    main()