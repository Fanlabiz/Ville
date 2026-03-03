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
    
    def get_area(self):
        return self.longueur * self.largeur

class Terrain:
    def __init__(self, grid):
        self.grid = np.array(grid)
        self.height, self.width = self.grid.shape
        self.occupied = np.zeros_like(self.grid, dtype=bool)
        self.buildings = []
        self.cultural_buildings = []
        self.cultural_zone = np.zeros_like(self.grid, dtype=bool)  # Cases dans une zone de rayonnement
        
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
            self.update_cultural_zone(building, x, y, orientation)
    
    def update_cultural_zone(self, building, x, y, orientation):
        """Met à jour la carte des zones de rayonnement"""
        l, L = building.get_dimensions(orientation)
        rayon = building.rayonnement
        
        for dx in range(max(0, x - rayon), min(self.width, x + l + rayon)):
            for dy in range(max(0, y - rayon), min(self.height, y + L + rayon)):
                if self.is_in_radiation_zone_building(building, x, y, orientation, dx, dy):
                    self.cultural_zone[dy, dx] = True
    
    def is_in_radiation_zone_building(self, building, x, y, orientation, px, py):
        """Vérifie si (px,py) est dans la zone de rayonnement du bâtiment"""
        l, L = building.get_dimensions(orientation)
        
        dx = 0
        if px < x:
            dx = x - px
        elif px >= x + l:
            dx = px - (x + l - 1)
        
        dy = 0
        if py < y:
            dy = y - py
        elif py >= y + L:
            dy = py - (y + L - 1)
        
        return 0 < max(dx, dy) <= building.rayonnement
    
    def is_in_radiation_zone(self, cultural, x, y):
        cx, cy = cultural['x'], cultural['y']
        cl, cL = cultural['longueur'], cultural['largeur']
        
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
    
    def get_radiation_zone_size(self, building, x, y, orientation):
        """Calcule le nombre de cases de rayonnement qui restent dans le terrain"""
        l, L = building.get_dimensions(orientation)
        rayon = building.rayonnement
        
        count = 0
        for dx in range(max(0, x - rayon), min(self.width, x + l + rayon)):
            for dy in range(max(0, y - rayon), min(self.height, y + L + rayon)):
                if self.is_in_radiation_zone_building(building, x, y, orientation, dx, dy):
                    count += 1
        return count
    
    def get_culture_for_position(self, x, y, longueur, largeur):
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
        positions = []
        for o in ['H', 'V']:
            l, L = building.get_dimensions(o)
            if l <= self.width and L <= self.height:
                for y in range(self.height - L + 1):
                    for x in range(self.width - l + 1):
                        if self.can_place(x, y, l, L):
                            positions.append((x, y, o))
        return positions
    
    def get_positions_in_cultural_zone(self, building):
        """Retourne les positions où le bâtiment serait dans une zone de rayonnement"""
        positions = []
        for o in ['H', 'V']:
            l, L = building.get_dimensions(o)
            if l <= self.width and L <= self.height:
                for y in range(self.height - L + 1):
                    for x in range(self.width - l + 1):
                        if self.can_place(x, y, l, L):
                            # Vérifier si au moins une case est dans une zone culturelle
                            in_zone = False
                            for i in range(l):
                                for j in range(L):
                                    if self.cultural_zone[y + j, x + i]:
                                        in_zone = True
                                        break
                                if in_zone:
                                    break
                            if in_zone:
                                positions.append((x, y, o))
        return positions
    
    def get_production_boosts(self):
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
        disp = np.array(terrain.grid, dtype=object)
        mapping, idx = {}, 1
        for b, x, y, o, l, L in terrain.buildings:
            if b.nom not in mapping:
                words = b.nom.split()
                short = ''.join(w[0].upper() for w in words[:2]) if len(words) > 1 else b.nom[:3].upper()
                mapping[b.nom] = f"{short}_{idx}"
                idx += 1
            for i in range(l):
                for j in range(L):
                    disp[y + j, x + i] = mapping[b.nom]
        
        pd.DataFrame(disp).to_excel(writer, sheet_name='Terrain', index=False, header=False)
        
        ws = writer.sheets['Terrain']
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = 5
        
        colors = {'culturel': 'FFE4B5', 'producteur': 'ADD8E6'}
        for b, x, y, o, l, L in terrain.buildings:
            fill = PatternFill(start_color=colors.get(b.type, 'D3D3D3'), end_color=colors.get(b.type, 'D3D3D3'), fill_type='solid')
            for i in range(l):
                for j in range(L):
                    ws.cell(row=y + j + 1, column=x + i + 1).fill = fill
        
        legend = []
        for nom, code in mapping.items():
            for b, _, _, _, _, _ in terrain.buildings:
                if b.nom == nom:
                    legend.append({"Code": code, "Nom": nom, "Type": b.type, "Production": b.production or "-"})
                    break
        if legend:
            pd.DataFrame(legend).to_excel(writer, sheet_name='Legende', index=False)
        
        if boosts:
            pd.DataFrame(boosts).to_excel(writer, sheet_name='Boosts', index=False)
        
        stats = []
        for p, c in totals.items():
            if p in counts:
                stats.append({"Production": p, "Culture totale": round(c, 2), 
                             "0%": counts[p][0], "25%": counts[p][25], 
                             "50%": counts[p][50], "100%": counts[p][100]})
        if stats:
            pd.DataFrame(stats).to_excel(writer, sheet_name='Statistiques', index=False)
        
        pos = []
        for b, x, y, o, l, L in terrain.buildings:
            pos.append({"Nom": b.nom, "Type": b.type, "Production": b.production or "-",
                       "X": x, "Y": y, "Orientation": o, "Longueur": l, "Largeur": L,
                       "Culture": b.culture if b.culture > 0 else "-",
                       "Rayonnement": b.rayonnement if b.rayonnement > 0 else "-"})
        if pos:
            pd.DataFrame(pos).to_excel(writer, sheet_name='Positions', index=False)
        
        np_ = []
        for b in buildings:
            if b.placed < b.quantite:
                np_.append({"Nom": b.nom, "Type": b.type or "-", "Production": b.production or "-",
                           "Quantité": b.quantite, "Placés": b.placed, "Restants": b.quantite - b.placed})
        pd.DataFrame(np_).to_excel(writer, sheet_name='Non places', index=False)
    
    output.seek(0)
    return output

def main():
    st.set_page_config(page_title="Placeur", layout="wide")
    st.title("🏗️ Placeur de bâtiments - Algorithme alterné")
    
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
            
            # Séparer par type
            cult = [b for b in all_b if b.type == "culturel"]
            prod = [b for b in all_b if b.type == "producteur" and b.production]
            
            # Trier les culturels par taille décroissante
            cult.sort(key=lambda b: -b.get_area())
            
            # Trier les producteurs par priorité puis taille
            def priorite(b):
                p = b.production.lower()
                if 'guerison' in p: return 1
                if 'nourriture' in p: return 2
                if 'or' in p: return 3
                return 4
            
            prod.sort(key=lambda b: (priorite(b), -b.get_area()))
            
            # Statistiques
            stats = {"culturels_places": 0, "producteurs_places": 0, "total": 0}
            
            # NOUVEL ALGORITHME : Alterner culturels et producteurs
            st.write("🔄 Algorithme alterné : 1 culturel → des producteurs → 1 culturel → des producteurs...")
            
            while cult or prod:
                # ÉTAPE 1: Placer un culturel (le plus grand restant)
                if cult:
                    b = cult.pop(0)  # Prendre le plus grand
                    st.write(f"\n🏛️ Placement du culturel: {b.nom}")
                    
                    best_score = -1
                    best_pos = None
                    
                    positions = terrain.get_all_positions(b)
                    
                    for x, y, o in positions:
                        # Score = taille de la zone de rayonnement
                        rayon_score = terrain.get_radiation_zone_size(b, x, y, o)
                        
                        if rayon_score > best_score:
                            best_score = rayon_score
                            best_pos = (x, y, o)
                    
                    if best_pos:
                        terrain.place_building(b, best_pos[0], best_pos[1], best_pos[2])
                        stats["culturels_places"] += 1
                        stats["total"] += 1
                        rayon = terrain.get_radiation_zone_size(b, best_pos[0], best_pos[1], best_pos[2])
                        st.write(f"  ✅ à ({best_pos[0]},{best_pos[1]}) - zone: {rayon} cases")
                    else:
                        st.write(f"  ❌ Impossible de placer {b.nom}")
                        # Le remettre dans la liste pour plus tard ? Non, on abandonne
                
                # ÉTAPE 2: Placer des producteurs tant qu'il y a des zones libres
                if prod:
                    st.write(f"\n🏭 Placement des producteurs dans les zones de rayonnement...")
                    
                    places_dans_cycle = 0
                    while prod:
                        # Chercher des positions dans les zones de rayonnement
                        b = prod[0]  # Prendre le premier (plus prioritaire)
                        
                        positions_in_zone = terrain.get_positions_in_cultural_zone(b)
                        
                        if not positions_in_zone:
                            # Plus de positions dans les zones, passer au prochain cycle
                            st.write(f"  ⏸️ Plus de zones pour {b.nom} (cycle suivant)")
                            break
                        
                        # Choisir la meilleure position (maximiser la culture)
                        best_score = -1
                        best_pos = None
                        
                        for x, y, o in positions_in_zone:
                            l, L = b.get_dimensions(o)
                            culture = terrain.get_culture_for_position(x, y, l, L)
                            
                            score = culture
                            if b.boost_100 > 0 and culture >= b.boost_100:
                                score += 10000
                            elif b.boost_50 > 0 and culture >= b.boost_50:
                                score += 5000
                            elif b.boost_25 > 0 and culture >= b.boost_25:
                                score += 2500
                            
                            if score > best_score:
                                best_score = score
                                best_pos = (x, y, o)
                        
                        if best_pos:
                            terrain.place_building(b, best_pos[0], best_pos[1], best_pos[2])
                            prod.pop(0)  # Enlever de la liste
                            stats["producteurs_places"] += 1
                            stats["total"] += 1
                            places_dans_cycle += 1
                            
                            l, L = b.get_dimensions(best_pos[2])
                            culture = terrain.get_culture_for_position(best_pos[0], best_pos[1], l, L)
                            
                            boost_text = ""
                            if culture >= b.boost_100:
                                boost_text = "🎯 BOOST 100%!"
                            elif culture >= b.boost_50:
                                boost_text = "🎯 Boost 50%"
                            elif culture >= b.boost_25:
                                boost_text = "🎯 Boost 25%"
                            
                            st.write(f"  ✅ {b.nom} à ({best_pos[0]},{best_pos[1]}) → {culture:.0f} culture {boost_text}")
                        else:
                            # Ne devrait pas arriver
                            break
                    
                    if places_dans_cycle > 0:
                        st.write(f"  → {places_dans_cycle} producteurs placés dans ce cycle")
            
            # Calculer les résultats
            boosts, totals, counts = terrain.get_production_boosts()
            
            # Mettre à jour les quantités
            for ob in buildings:
                ob.placed = 0
                ob.positions = []
                for pb in all_b:
                    if pb.nom == ob.nom and pb.placed > 0:
                        ob.placed += 1
                        if pb.positions:
                            ob.positions.append(pb.positions[0])
            
            # Afficher le résumé final
            st.subheader("📊 Résumé final")
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Bâtiments placés", f"{stats['total']}/{len(all_b)}")
            with col2:
                st.metric("Culture totale", f"{sum(totals.values()):.0f}")
            with col3:
                total_boosts = sum(counts[p][100] for p in counts)
                st.metric("Boost 100%", total_boosts)
            
            st.write(f"🏛️ Culturels: {stats['culturels_places']}/{len([b for b in buildings if b.type == 'culturel' and b.quantite > 0])}")
            st.write(f"🏭 Producteurs: {stats['producteurs_places']}/{len([b for b in buildings if b.type == 'producteur' and b.production and b.quantite > 0])}")
            
            # Téléchargement
            out = create_output_excel(terrain, boosts, totals, counts, buildings)
            st.download_button("📥 Télécharger les résultats", out, "resultats.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    main()