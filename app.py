import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font
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
        self.id = f"{nom}_{id(self)}"
        
    def get_dimensions(self, orientation='H'):
        if orientation == 'H':
            return self.longueur, self.largeur
        else:
            return self.largeur, self.longueur

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
            center_x = x + longueur // 2
            center_y = y + largeur // 2
            self.cultural_buildings.append({
                'building': building,
                'x': x, 'y': y,
                'longueur': longueur, 'largeur': largeur,
                'center_x': center_x, 'center_y': center_y,
                'rayonnement': building.rayonnement,
                'culture': building.culture,
                'id': building.id
            })
    
    def is_in_radiation_zone(self, cultural, x, y):
        """Vérifie si (x,y) est dans la zone de rayonnement du culturel"""
        cx, cy = cultural['x'], cultural['y']
        cl, cL = cultural['longueur'], cultural['largeur']
        
        # Distance au bord du bâtiment
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
        
        return 0 < max(dx, dy) <= cultural['rayonnement']
    
    def get_culture_for_position(self, x, y, longueur, largeur):
        """Calcule la culture reçue à une position"""
        if not self.cultural_buildings:
            return 0
        
        affecting = set()
        for cb in self.cultural_buildings:
            for i in range(longueur):
                for j in range(largeur):
                    if self.is_in_radiation_zone(cb, x + i, y + j):
                        affecting.add(cb['id'])
                        break
                if cb['id'] in affecting:
                    break
        
        return sum(cb['culture'] for id_ in affecting for cb in self.cultural_buildings if cb['id'] == id_)
    
    def get_all_positions(self, building):
        """Toutes les positions possibles pour un bâtiment"""
        positions = []
        for o in ['H', 'V']:
            l, L = building.get_dimensions(o)
            if l <= self.width and L <= self.height:
                for y in range(self.height - L + 1):
                    for x in range(self.width - l + 1):
                        if self.can_place(x, y, l, L):
                            positions.append((x, y, o))
        return positions

def normalize_name(name):
    if pd.isna(name):
        return ""
    name = str(name).strip().lower()
    for a, b in {'é':'e','è':'e','ê':'e','à':'a','â':'a','ù':'u','û':'u','ç':'c'}.items():
        name = name.replace(a, b)
    return re.sub(r'[^a-zA-Z0-9%]', '', name)

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
        'quantite': ['quantite', 'quantité', 'quantity', 'qt', 'nb'],
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
        return []
    
    buildings = []
    for _, row in df.iterrows():
        try:
            q = 1
            if qc := find(col_map['quantite']):
                if pd.notna(row[qc]) and str(row[qc]).strip():
                    q = int(float(row[qc]))
            if q == 0:
                continue
            
            b = Building(
                nom=row[nom],
                longueur=row[find(col_map['longueur'])] if find(col_map['longueur']) else 1,
                largeur=row[find(col_map['largeur'])] if find(col_map['largeur']) else 1,
                quantite=q,
                type_bat=row[find(col_map['type'])] if find(col_map['type']) else "",
                culture=row[find(col_map['culture'])] if find(col_map['culture']) else 0,
                rayonnement=row[find(col_map['rayonnement'])] if find(col_map['rayonnement']) else 0,
                boost_25=row[find(col_map['boost25'])] if find(col_map['boost25']) else 0,
                boost_50=row[find(col_map['boost50'])] if find(col_map['boost50']) else 0,
                boost_100=row[find(col_map['boost100'])] if find(col_map['boost100']) else 0,
                production=row[find(col_map['production'])] if find(col_map['production']) else ""
            )
            buildings.append(b)
        except:
            continue
    return buildings

def create_output_excel(terrain, boosts, totals, counts, buildings):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Terrain
        disp = np.array(terrain.grid, dtype=object)
        mapping, idx = {}, 1
        for b, x, y, o, l, L in terrain.buildings:
            if b.nom not in mapping:
                short = ''.join(w[0].upper() for w in b.nom.split()[:2]) or b.nom[:3]
                mapping[b.nom] = f"{short}_{idx}"
                idx += 1
            for i in range(l):
                for j in range(L):
                    disp[y + j, x + i] = mapping[b.nom]
        
        pd.DataFrame(disp).to_excel(writer, sheet_name='Terrain', index=False, header=False)
        
        # Légende
        legend = []
        for nom, code in mapping.items():
            for b, _, _, _, _, _ in terrain.buildings:
                if b.nom == nom:
                    legend.append({"Code": code, "Nom": nom, "Type": b.type, "Production": b.production or "-"})
                    break
        if legend:
            pd.DataFrame(legend).to_excel(writer, sheet_name='Legende', index=False)
        
        # Boosts
        if boosts:
            pd.DataFrame(boosts).to_excel(writer, sheet_name='Boosts', index=False)
        
        # Stats
        stats = []
        for p, c in totals.items():
            if p in counts:
                stats.append({"Production": p, "Culture": round(c, 2), "0%": counts[p][0], "25%": counts[p][25], "50%": counts[p][50], "100%": counts[p][100]})
        if stats:
            pd.DataFrame(stats).to_excel(writer, sheet_name='Statistiques', index=False)
        
        # Positions
        pos = []
        for b, x, y, o, l, L in terrain.buildings:
            pos.append({"Nom": b.nom, "Type": b.type, "Production": b.production or "-", "X": x, "Y": y, "Orientation": o, "Longueur": l, "Largeur": L, "Culture": b.culture or "-", "Rayonnement": b.rayonnement or "-"})
        if pos:
            pd.DataFrame(pos).to_excel(writer, sheet_name='Positions', index=False)
        
        # Non placés
        np_ = []
        for b in buildings:
            if b.placed < b.quantite:
                np_.append({"Nom": b.nom, "Type": b.type or "-", "Production": b.production or "-", "Quantité": b.quantite, "Placés": b.placed, "Restants": b.quantite - b.placed})
        pd.DataFrame(np_ if np_ else []).to_excel(writer, sheet_name='Non places', index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Placeur", layout="wide")
    st.title("🏗️ Placeur de bâtiments")
    
    file = st.sidebar.file_uploader("Fichier Excel", type=['xlsx', 'xls'])
    if not file:
        st.info("Chargez un fichier Excel")
        return
    
    grid, df = read_file(file)
    if grid is None:
        return
    
    terrain = Terrain(grid)
    buildings = create_buildings(df)
    if not buildings:
        st.error("Aucun bâtiment")
        return
    
    if st.button("🚀 Lancer", type="primary"):
        with st.spinner("Placement..."):
            # Créer les exemplaires
            all_b = []
            for b in buildings:
                for i in range(b.quantite):
                    nb = Building(b.nom, b.longueur, b.largeur, 1, b.type, b.culture, b.rayonnement,
                                 b.boost_25, b.boost_50, b.boost_100, b.production)
                    nb.id = f"{b.nom}_{i}"
                    all_b.append(nb)
            
            # Placer les culturels en premier
            cult = [b for b in all_b if b.type == "culturel"]
            prod = [b for b in all_b if b.type == "producteur" and b.production]
            
            # Placement simple des culturels : on les met au centre
            for b in cult:
                pos = terrain.get_all_positions(b)
                if pos:
                    # Prendre la position la plus centrale
                    cx, cy = terrain.width // 2, terrain.height // 2
                    best = min(pos, key=lambda p: abs(p[0] + b.longueur//2 - cx) + abs(p[1] + b.largeur//2 - cy))
                    terrain.place_building(b, best[0], best[1], best[2])
            
            # Maintenant, placer les producteurs pour maximiser la culture
            # On va d'abord calculer la zone de rayonnement de chaque culturel
            zone_map = np.zeros((terrain.height, terrain.width), dtype=int)
            for y in range(terrain.height):
                for x in range(terrain.width):
                    if terrain.grid[y, x] == 1:
                        for cb in terrain.cultural_buildings:
                            if terrain.is_in_radiation_zone(cb, x, y):
                                zone_map[y, x] += 1
            
            # Afficher la carte
            st.write("Carte des zones (nombre de culturels qui touchent chaque case):")
            st.dataframe(pd.DataFrame(zone_map), use_container_width=True)
            
            # Placer les producteurs
            for b in prod:
                best_score = -1
                best_pos = None
                
                for x, y, o in terrain.get_all_positions(b):
                    l, L = b.get_dimensions(o)
                    
                    # Compter combien de culturels touchent cette position
                    count = 0
                    for i in range(l):
                        for j in range(L):
                            px, py = x + i, y + j
                            if 0 <= px < terrain.width and 0 <= py < terrain.height:
                                count = max(count, zone_map[py, px])
                    
                    # La culture reçue est la somme des plus gros culturels
                    if count > 0:
                        cultures = sorted([cb['culture'] for cb in terrain.cultural_buildings], reverse=True)
                        culture = sum(cultures[:count])
                    else:
                        culture = 0
                    
                    # Score = priorité au nombre de culturels
                    score = count * 10000 + culture
                    
                    if score > best_score:
                        best_score = score
                        best_pos = (x, y, o)
                
                if best_pos:
                    terrain.place_building(b, best_pos[0], best_pos[1], best_pos[2])
            
            # Calculer les résultats
            boosts, totals, counts = terrain.get_production_boosts()
            
            # Mettre à jour les quantités
            for ob in buildings:
                ob.placed = sum(1 for pb in all_b if pb.nom == ob.nom and pb.placed > 0)
            
            # Téléchargement
            out = create_output_excel(terrain, boosts, totals, counts, buildings)
            st.download_button("📥 Télécharger", out, "resultats.xlsx")

if __name__ == "__main__":
    main()