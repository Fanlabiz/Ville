import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import re
from collections import defaultdict

class Building:
    def __init__(self, nom, longueur, largeur, quantite, type_bat, culture, rayonnement, 
                 boost_25, boost_50, boost_100, production):
        self.nom = nom
        self.longueur = int(float(longueur)) if pd.notna(longueur) and str(longueur).strip() != '' else 1
        self.largeur = int(float(largeur)) if pd.notna(largeur) and str(largeur).strip() != '' else 1
        self.quantite = int(float(quantite)) if pd.notna(quantite) and str(quantite).strip() != '' else 1
        self.type = str(type_bat).lower() if pd.notna(type_bat) and str(type_bat).strip() != '' else ""
        self.culture = float(culture) if pd.notna(culture) and str(culture).strip() != '' else 0
        self.rayonnement = int(float(rayonnement)) if pd.notna(rayonnement) and str(rayonnement).strip() != '' else 0
        self.boost_25 = float(boost_25) if pd.notna(boost_25) and str(boost_25).strip() != '' else 0
        self.boost_50 = float(boost_50) if pd.notna(boost_50) and str(boost_50).strip() != '' else 0
        self.boost_100 = float(boost_100) if pd.notna(boost_100) and str(boost_100).strip() != '' else 0
        self.production = str(production) if pd.notna(production) and str(production).strip() != '' else ""
        
        self.placed = 0
        self.positions = []
        self.failed = False
        self.id = f"{nom}_{id(self)}"
        
    def get_dimensions(self, orientation='H'):
        if orientation == 'H':
            return self.longueur, self.largeur
        else:
            return self.largeur, self.longueur
    
    def get_area(self):
        return self.longueur * self.largeur

class Terrain:
    def __init__(self, grid):
        self.grid = np.array(grid)
        self.height, self.width = self.grid.shape
        self.occupied = np.zeros_like(self.grid, dtype=bool)
        self.buildings = []
        self.cultural_buildings = []
        
    def can_place(self, x, y, longueur, largeur):
        if x + longueur > self.width or y + largeur > self.height:
            return False
        for i in range(longueur):
            for j in range(largeur):
                if self.grid[y + j, x + i] == 0 or self.occupied[y + j, x + i]:
                    return False
        return True
    
    def place_building(self, building, x, y, orientation):
        longueur, largeur = building.get_dimensions(orientation)
        for i in range(longueur):
            for j in range(largeur):
                self.occupied[y + j, x + i] = True
        building.placed += 1
        building.positions.append((x, y, orientation))
        self.buildings.append((building, x, y, orientation, longueur, largeur))
        
        if building.type == "culturel" and building.culture > 0:
            self.cultural_buildings.append({
                'building': building,
                'x': x, 'y': y, 'orientation': orientation,
                'longueur': longueur, 'largeur': largeur,
                'rayonnement': building.rayonnement,
                'culture': building.culture,
                'id': building.id
            })
    
    def is_in_radiation_zone(self, cultural, x, y):
        """
        Vérifie si la case (x,y) est dans la zone de rayonnement du bâtiment culturel.
        Le rayonnement est une bande autour du bâtiment.
        """
        # Coordonnées du bâtiment culturel
        cx, cy = cultural['x'], cultural['y']
        cl, cL = cultural['longueur'], cultural['largeur']
        
        # La bande de rayonnement est à l'extérieur du bâtiment, jusqu'à rayonnement cases de distance
        # Une case est dans la zone si elle est à une distance <= rayonnement du bord du bâtiment
        
        # Distance minimale de la case au rectangle du bâtiment
        dx = 0
        if x < cx:
            dx = cx - x
        elif x >= cx + cl:
            dx = x - (cx + cl - 1)
        
        dy = 0
        if y < cy:
            dy = cy - y
        elif y >= cy + cL:
            dy = y - (cy + cL - 1)
        
        distance = max(dx, dy)
        
        # La case est dans la zone si elle est à une distance > 0 et <= rayonnement
        return 0 < distance <= cultural['rayonnement']
    
    def get_culture_for_position(self, x, y, longueur, largeur):
        """
        Calcule la culture reçue par un bâtiment à une position donnée.
        Chaque bâtiment culturel ne compte qu'une seule fois.
        """
        if not self.cultural_buildings:
            return 0
        
        # Ensemble des IDs des culturels qui touchent ce bâtiment
        affecting_ids = set()
        
        # Pour chaque bâtiment culturel
        for cb in self.cultural_buildings:
            # Vérifier si UNE case du bâtiment est dans la zone de rayonnement du culturel
            for i in range(longueur):
                for j in range(largeur):
                    px, py = x + i, y + j
                    if 0 <= px < self.width and 0 <= py < self.height:
                        if self.is_in_radiation_zone(cb, px, py):
                            affecting_ids.add(cb['id'])
                            break
                if cb['id'] in affecting_ids:
                    break
        
        # Calculer la culture totale (chaque culturel ne compte qu'une fois)
        total_culture = 0
        for id_ in affecting_ids:
            for cb in self.cultural_buildings:
                if cb['id'] == id_:
                    total_culture += cb['culture']
                    break
        
        return total_culture
    
    def get_all_positions(self, building):
        positions = []
        for orientation in ['H', 'V']:
            l, L = building.get_dimensions(orientation)
            if l > self.width or L > self.height:
                continue
            for y in range(self.height - L + 1):
                for x in range(self.width - l + 1):
                    if self.can_place(x, y, l, L):
                        positions.append((x, y, orientation))
        return positions
    
    def find_largest_free_area(self):
        """Trouve la plus grande zone contiguë de cases libres"""
        visited = np.zeros_like(self.grid, dtype=bool)
        largest_area = []
        
        for y in range(self.height):
            for x in range(self.width):
                if self.grid[y, x] == 1 and not self.occupied[y, x] and not visited[y, x]:
                    # BFS pour trouver la zone
                    area = []
                    queue = [(x, y)]
                    visited[y, x] = True
                    
                    while queue:
                        cx, cy = queue.pop(0)
                        area.append((cx, cy))
                        
                        # Voisins (4 directions)
                        for dx, dy in [(0,1), (1,0), (0,-1), (-1,0)]:
                            nx, ny = cx + dx, cy + dy
                            if (0 <= nx < self.width and 0 <= ny < self.height and 
                                self.grid[ny, nx] == 1 and not self.occupied[ny, nx] and 
                                not visited[ny, nx]):
                                visited[ny, nx] = True
                                queue.append((nx, ny))
                    
                    if len(area) > len(largest_area):
                        largest_area = area
        
        return largest_area
    
    def get_space_preservation_score(self, x, y, longueur, largeur, remaining_buildings):
        """
        Calcule un score pour préserver l'espace pour les bâtiments restants.
        Plus le score est élevé, mieux c'est.
        """
        # Simuler le placement temporaire
        temp_occupied = self.occupied.copy()
        for i in range(longueur):
            for j in range(largeur):
                if 0 <= y+j < self.height and 0 <= x+i < self.width:
                    temp_occupied[y+j, x+i] = True
        
        # Vérifier si on peut encore placer les plus grands bâtiments restants
        if remaining_buildings:
            # Trier par taille décroissante
            remaining_by_size = sorted(remaining_buildings, key=lambda b: -b.get_area())
            biggest = remaining_by_size[0]
            
            # Vérifier si le plus grand bâtiment peut encore être placé
            can_place_biggest = False
            for orientation in ['H', 'V']:
                bl, bL = biggest.get_dimensions(orientation)
                if bl > self.width or bL > self.height:
                    continue
                for by in range(self.height - bL + 1):
                    for bx in range(self.width - bl + 1):
                        # Vérifier si on peut placer sans utiliser temp_occupied
                        can = True
                        for i in range(bl):
                            for j in range(bL):
                                if self.grid[by+j, bx+i] == 0 or temp_occupied[by+j, bx+i]:
                                    can = False
                                    break
                            if not can:
                                break
                        if can:
                            can_place_biggest = True
                            break
                    if can_place_biggest:
                        break
                    break
            
            if not can_place_biggest:
                return -10000  # Pénalité énorme si on bloque le plus grand bâtiment
        
        # Bonus si on préserve de grandes zones contiguës
        # On va estimer la taille de la plus grande zone libre après placement
        # (simulation simplifiée)
        
        return 0  # Score neutre par défaut
    
    def get_production_boosts(self):
        """Calcule les boosts pour tous les bâtiments producteurs"""
        results = []
        total_culture_by_type = defaultdict(float)
        boost_counts = defaultdict(lambda: {0: 0, 25: 0, 50: 0, 100: 0})
        
        for building, x, y, orientation, longueur, largeur in self.buildings:
            if building.type == "producteur" and building.production:
                prod_type = building.production.strip()
                if not prod_type:
                    continue
                
                total_culture = self.get_culture_for_position(x, y, longueur, largeur)
                total_culture_by_type[prod_type] += total_culture
                
                boost = 0
                if building.boost_100 > 0 and total_culture >= building.boost_100:
                    boost = 100
                elif building.boost_50 > 0 and total_culture >= building.boost_50:
                    boost = 50
                elif building.boost_25 > 0 and total_culture >= building.boost_25:
                    boost = 25
                    
                boost_counts[prod_type][boost] += 1
                
                results.append({
                    "Nom": building.nom,
                    "Production": building.production,
                    "Culture reçue": round(total_culture, 2),
                    "Boost": f"{boost}%"
                })
        
        return results, dict(total_culture_by_type), dict(boost_counts)

def normalize_name(name):
    if pd.isna(name):
        return ""
    name = str(name).strip().lower()
    replacements = {'é':'e','è':'e','ê':'e','ë':'e','à':'a','â':'a','ä':'a',
                   'î':'i','ï':'i','ô':'o','ö':'o','ù':'u','û':'u','ü':'u','ç':'c'}
    for a, b in replacements.items():
        name = name.replace(a, b)
    name = re.sub(r'[^a-zA-Z0-9%]', '', name)
    return name

def read_file(file):
    try:
        xl = pd.ExcelFile(file)
        terrain = pd.read_excel(xl, sheet_name=0, header=None).values.tolist()
        batiments = pd.read_excel(xl, sheet_name=1)
        batiments.columns = [normalize_name(c) for c in batiments.columns]
        return terrain, batiments
    except Exception as e:
        st.error(f"Erreur: {e}")
        return None, None

def create_buildings(df):
    col_map = {
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
    
    def find(cols):
        for c in cols:
            for col in df.columns:
                if c in col:
                    return col
        return None
    
    nom = find(col_map['nom'])
    if not nom:
        st.error("Colonne 'Nom' introuvable")
        return []
    
    buildings = []
    for _, row in df.iterrows():
        try:
            q = 1
            quantite_col = find(col_map['quantite'])
            if quantite_col:
                val = row[quantite_col]
                if pd.notna(val) and str(val).strip() != '':
                    q = int(float(val))
            if q == 0:
                continue
            
            longueur_col = find(col_map['longueur'])
            largeur_col = find(col_map['largeur'])
            type_col = find(col_map['type'])
            culture_col = find(col_map['culture'])
            rayonnement_col = find(col_map['rayonnement'])
            boost25_col = find(col_map['boost25'])
            boost50_col = find(col_map['boost50'])
            boost100_col = find(col_map['boost100'])
            production_col = find(col_map['production'])
            
            b = Building(
                nom=row[nom],
                longueur=row[longueur_col] if longueur_col else 1,
                largeur=row[largeur_col] if largeur_col else 1,
                quantite=q,
                type_bat=row[type_col] if type_col else "",
                culture=row[culture_col] if culture_col else 0,
                rayonnement=row[rayonnement_col] if rayonnement_col else 0,
                boost_25=row[boost25_col] if boost25_col else 0,
                boost_50=row[boost50_col] if boost50_col else 0,
                boost_100=row[boost100_col] if boost100_col else 0,
                production=row[production_col] if production_col else ""
            )
            buildings.append(b)
        except Exception as e:
            st.warning(f"Erreur sur une ligne: {e}")
            continue
    return buildings

def create_output_excel(terrain, boosts, totals, counts, buildings_originaux):
    """Crée le fichier Excel de sortie"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Feuille 1: Terrain avec bâtiments
        terrain_display = np.array(terrain.grid, dtype=object)
        
        # Créer des codes courts pour les bâtiments
        building_map = {}
        color_index = 1
        for building, x, y, orientation, l, L in terrain.buildings:
            if building.nom not in building_map:
                words = building.nom.split()
                if len(words) > 1:
                    short = ''.join(w[0].upper() for w in words[:2])
                else:
                    short = building.nom[:3].upper()
                building_map[building.nom] = f"{short}_{color_index}"
                color_index += 1
            
            for i in range(l):
                for j in range(L):
                    terrain_display[y + j, x + i] = building_map[building.nom]
        
        terrain_df = pd.DataFrame(terrain_display)
        terrain_df.to_excel(writer, sheet_name='Terrain', index=False, header=False)
        
        # Ajuster la largeur des colonnes
        worksheet = writer.sheets['Terrain']
        for i, col in enumerate(worksheet.columns):
            worksheet.column_dimensions[col[0].column_letter].width = 5
        
        # Colorer les cellules
        type_colors = {'culturel': 'FFE4B5', 'producteur': 'ADD8E6'}
        for building, x, y, orientation, l, L in terrain.buildings:
            color = type_colors.get(building.type, 'D3D3D3')
            fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
            for i in range(l):
                for j in range(L):
                    cell = worksheet.cell(row=y + j + 1, column=x + i + 1)
                    cell.fill = fill
                    cell.font = Font(bold=True, size=8)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Feuille 2: Légende
        legend = []
        for nom, code in building_map.items():
            for b, _, _, _, _, _ in terrain.buildings:
                if b.nom == nom:
                    legend.append({
                        "Code": code,
                        "Nom": nom,
                        "Type": b.type,
                        "Production": b.production if b.production else "-"
                    })
                    break
        if legend:
            pd.DataFrame(legend).to_excel(writer, sheet_name='Legende', index=False)
        
        # Feuille 3: Boosts
        if boosts:
            pd.DataFrame(boosts).to_excel(writer, sheet_name='Boosts', index=False)
        
        # Feuille 4: Statistiques
        stats = []
        for prod, culture in totals.items():
            if prod in counts:
                stats.append({
                    "Production": prod,
                    "Culture totale": round(culture, 2),
                    "Boost 0%": counts[prod][0],
                    "Boost 25%": counts[prod][25],
                    "Boost 50%": counts[prod][50],
                    "Boost 100%": counts[prod][100]
                })
        if stats:
            pd.DataFrame(stats).to_excel(writer, sheet_name='Statistiques', index=False)
        
        # Feuille 5: Positions (avec rayonnement)
        positions = []
        for building, x, y, orientation, l, L in terrain.buildings:
            positions.append({
                "Nom": building.nom,
                "Type": building.type,
                "Production": building.production if building.production else "-",
                "X": x, "Y": y,
                "Orientation": orientation,
                "Longueur": l, "Largeur": L,
                "Culture": building.culture if building.culture > 0 else "-",
                "Rayonnement": building.rayonnement if building.rayonnement > 0 else "-"
            })
        if positions:
            pd.DataFrame(positions).to_excel(writer, sheet_name='Positions', index=False)
        
        # Feuille 6: Non placés
        non_places = []
        for b in buildings_originaux:
            if b.placed < b.quantite:
                non_places.append({
                    "Nom": b.nom,
                    "Type": b.type if b.type else "-",
                    "Production": b.production if b.production else "-",
                    "Quantité": b.quantite,
                    "Placés": b.placed,
                    "Restants": b.quantite - b.placed
                })
        
        non_places_df = pd.DataFrame(non_places) if non_places else pd.DataFrame(columns=["Nom", "Type", "Production", "Quantité", "Placés", "Restants"])
        non_places_df.to_excel(writer, sheet_name='Non places', index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Placeur de bâtiments", layout="wide")
    st.title("🏗️ Placeur de bâtiments - Préservation d'espace")
    
    with st.sidebar:
        st.header("📂 Chargement")
        file = st.file_uploader("Fichier Excel", type=['xlsx', 'xls'])
    
    if not file:
        st.info("Chargez un fichier Excel pour commencer")
        return
    
    terrain_grid, batiments_df = read_file(file)
    if terrain_grid is None or batiments_df is None:
        return
    
    terrain = Terrain(terrain_grid)
    buildings = create_buildings(batiments_df)
    
    if not buildings:
        st.error("Aucun bâtiment valide")
        return
    
    if st.button("🚀 Lancer l'optimisation", type="primary"):
        with st.spinner("Placement en cours..."):
            # Créer les exemplaires
            tous = []
            for b in buildings:
                for i in range(b.quantite):
                    new = Building(
                        b.nom, b.longueur, b.largeur, 1, 
                        b.type, b.culture, b.rayonnement,
                        b.boost_25, b.boost_50, b.boost_100, 
                        b.production
                    )
                    new.id = f"{b.nom}_{i}"
                    tous.append(new)
            
            # Séparer par type
            culturels = [b for b in tous if b.type == "culturel"]
            producteurs = [b for b in tous if b.type == "producteur" and b.production]
            
            # Trier par taille décroissante pour placer les plus grands en premier
            # (les grands bâtiments sont plus difficiles à placer)
            tous_par_taille = sorted(tous, key=lambda b: -b.get_area())
            
            # Placer tous les bâtiments en préservant l'espace
            terrain_final = terrain
            
            # Statistiques pour le suivi
            stats = {"total": len(tous), "places": 0, "non_places": 0}
            
            for i, b in enumerate(tous_par_taille):
                # Récupérer les bâtiments restants (ceux qui ne sont pas encore placés)
                remaining = [bb for bb in tous_par_taille[i+1:] if bb.placed == 0]
                
                meilleur_score = -1e9
                meilleure_pos = None
                
                positions = terrain_final.get_all_positions(b)
                
                for x, y, o in positions:
                    l, L = b.get_dimensions(o)
                    
                    # Score de base
                    score = 0
                    
                    # 1. Bonus pour la culture (si producteur)
                    if b.type == "producteur":
                        culture = terrain_final.get_culture_for_position(x, y, l, L)
                        score += culture * 10
                        
                        # Bonus pour atteindre les seuils
                        if b.boost_100 > 0 and culture >= b.boost_100:
                            score += 100000
                        elif b.boost_50 > 0 and culture >= b.boost_50:
                            score += 50000
                        elif b.boost_25 > 0 and culture >= b.boost_25:
                            score += 25000
                    
                    # 2. Pénalité si on fragmente l'espace
                    # On vérifie rapidement si le plus grand bâtiment restant peut encore être placé
                    if remaining:
                        biggest = max(remaining, key=lambda bb: bb.get_area())
                        can_place_biggest = False
                        
                        # Test rapide sur quelques positions
                        for orientation in ['H', 'V']:
                            bl, bL = biggest.get_dimensions(orientation)
                            if bl > terrain_final.width or bL > terrain_final.height:
                                continue
                            
                            # On ne teste que quelques positions pour ne pas ralentir
                            step = max(1, min(terrain_final.height // 5, terrain_final.width // 5))
                            for by in range(0, terrain_final.height - bL + 1, step):
                                for bx in range(0, terrain_final.width - bl + 1, step):
                                    # Simuler le placement
                                    can = True
                                    for ii in range(bl):
                                        for jj in range(bL):
                                            if (terrain_final.grid[by+jj, bx+ii] == 0 or 
                                                terrain_final.occupied[by+jj, bx+ii] or
                                                (bx+ii >= x and bx+ii < x+l and by+jj >= y and by+jj < y+L)):
                                                can = False
                                                break
                                        if not can:
                                            break
                                    if can:
                                        can_place_biggest = True
                                        break
                                if can_place_biggest:
                                    break
                            if can_place_biggest:
                                break
                        
                        if not can_place_biggest:
                            score -= 50000  # Pénalité si on bloque le plus grand bâtiment
                    
                    # 3. Bonus pour rester groupé (évite de disperser)
                    if terrain_final.buildings:
                        # Distance moyenne aux autres bâtiments
                        total_dist = 0
                        count = 0
                        center_x, center_y = x + l//2, y + L//2
                        for bb, bx, by, bo, bl, bL in terrain_final.buildings:
                            bcx, bcy = bx + bl//2, by + bL//2
                            dist = abs(center_x - bcx) + abs(center_y - bcy)
                            total_dist += dist
                            count += 1
                        if count > 0:
                            avg_dist = total_dist / count
                            score -= avg_dist * 10  # Pénalité pour être loin
                    
                    if score > meilleur_score:
                        meilleur_score = score
                        meilleure_pos = (x, y, o)
                
                if meilleure_pos:
                    x, y, o = meilleure_pos
                    terrain_final.place_building(b, x, y, o)
                    stats["places"] += 1
                else:
                    b.failed = True
                    stats["non_places"] += 1
            
            # Calculer les boosts
            boosts, totals, counts = terrain_final.get_production_boosts()
            
            # Mettre à jour les quantités placées dans les bâtiments originaux
            for original_b in buildings:
                original_b.placed = 0
                original_b.positions = []
                for placed_b in tous:
                    if placed_b.nom == original_b.nom and not placed_b.failed:
                        original_b.placed += 1
                        if placed_b.positions:
                            original_b.positions.append(placed_b.positions[0])
            
            # Créer et proposer le téléchargement du fichier Excel
            output_file = create_output_excel(terrain_final, boosts, totals, counts, buildings)
            
            st.success(f"✅ Placement terminé : {stats['places']} bâtiments placés, {stats['non_places']} non placés")
            st.download_button(
                label="📥 Télécharger les résultats (Excel)",
                data=output_file,
                file_name="resultats_placement.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()