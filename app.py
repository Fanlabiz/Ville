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
                'x': x, 'y': y, 'orientation': orientation,
                'longueur': longueur, 'largeur': largeur,
                'center_x': center_x, 'center_y': center_y,
                'rayonnement': building.rayonnement,
                'culture': building.culture,
                'id': building.id
            })
    
    def get_culture_for_position(self, x, y, longueur, largeur):
        if not self.cultural_buildings:
            return 0
        
        affecting_ids = set()
        for cb in self.cultural_buildings:
            for i in range(longueur):
                for j in range(largeur):
                    px, py = x + i, y + j
                    if 0 <= px < self.width and 0 <= py < self.height:
                        distance = max(abs(px - cb['center_x']), abs(py - cb['center_y']))
                        if distance <= cb['rayonnement']:
                            affecting_ids.add(cb['id'])
                            break
                if cb['id'] in affecting_ids:
                    break
        
        total = 0
        for id_ in affecting_ids:
            for cb in self.cultural_buildings:
                if cb['id'] == id_:
                    total += cb['culture']
                    break
        return total
    
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

def main():
    st.set_page_config(page_title="Placeur de bâtiments", layout="wide")
    st.title("🏗️ Placeur de bâtiments")
    
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
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("🗺️ Terrain")
        st.dataframe(pd.DataFrame(terrain_grid), use_container_width=True)
        st.caption(f"Dimensions: {terrain.height}x{terrain.width}")
        st.caption(f"Cases libres: {np.sum(terrain.grid == 1)}")
    with col2:
        st.subheader("🏢 Bâtiments")
        data = []
        for b in buildings:
            data.append({
                "Nom": b.nom, 
                "Type": b.type, 
                "Production": b.production if b.production else "-",
                "Dimensions": f"{b.longueur}x{b.largeur}", 
                "Quantité": b.quantite,
                "Culture": f"{b.culture:.0f}" if b.culture > 0 else "-"
            })
        st.dataframe(pd.DataFrame(data), use_container_width=True)
    
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
            
            # Trier
            def priorite_producteur(b):
                p = b.production.lower()
                if 'guerison' in p: return 1
                if 'nourriture' in p: return 2
                if 'or' in p: return 3
                return 4
            
            culturels.sort(key=lambda b: (-b.rayonnement, -b.culture))
            producteurs.sort(key=priorite_producteur)
            
            ordre = culturels + producteurs
            terrain_final = terrain
            progress = st.progress(0)
            output_text = []
            
            for i, b in enumerate(ordre):
                meilleur_score = -1
                meilleure_pos = None
                
                positions = terrain_final.get_all_positions(b)
                
                for x, y, o in positions:
                    if b.type == "culturel":
                        l, L = b.get_dimensions(o)
                        cx, cy = x + l//2, y + L//2
                        score = 0
                        for dx in range(max(0, cx - b.rayonnement), min(terrain_final.width, cx + b.rayonnement + 1)):
                            for dy in range(max(0, cy - b.rayonnement), min(terrain_final.height, cy + b.rayonnement + 1)):
                                if terrain_final.grid[dy, dx] == 1 and not terrain_final.occupied[dy, dx]:
                                    score += 1
                    else:
                        l, L = b.get_dimensions(o)
                        culture = terrain_final.get_culture_for_position(x, y, l, L)
                        score = culture
                        if b.boost_100 > 0 and culture >= b.boost_100:
                            score += 10000
                        elif b.boost_50 > 0 and culture >= b.boost_50:
                            score += 5000
                        elif b.boost_25 > 0 and culture >= b.boost_25:
                            score += 2500
                    
                    if score > meilleur_score:
                        meilleur_score = score
                        meilleure_pos = (x, y, o)
                
                if meilleure_pos:
                    x, y, o = meilleure_pos
                    terrain_final.place_building(b, x, y, o)
                    if b.type == "producteur":
                        l, L = b.get_dimensions(o)
                        c = terrain_final.get_culture_for_position(x, y, l, L)
                        msg = f"✅ {b.nom} à ({x},{y}) → {c:.0f} culture"
                        output_text.append(msg)
                    else:
                        msg = f"✅ {b.nom} à ({x},{y})"
                        output_text.append(msg)
                else:
                    msg = f"❌ Impossible de placer {b.nom}"
                    output_text.append(msg)
                
                progress.progress((i + 1) / len(ordre))
            
            # Afficher les messages de placement
            for msg in output_text:
                st.write(msg)
            
            # Calculer les boosts
            boosts, totals, counts = terrain_final.get_production_boosts()
            
            # Afficher les résultats
            st.subheader("📊 Résultats")
            col1, col2, col3 = st.columns(3)
            with col1:
                total_places = sum(b.placed for b in buildings)
                total_a_placer = sum(b.quantite for b in buildings)
                st.metric("Placés", f"{total_places}/{total_a_placer}")
            with col2:
                st.metric("Taux", f"{total_places/total_a_placer*100:.1f}%")
            with col3:
                st.metric("Culture totale", f"{sum(totals.values()):.0f}")
            
            # Visualisation
            st.subheader("🗺️ Terrain final")
            viz = np.array(terrain_final.grid, dtype=str)
            symbols = {}
            color_idx = 0
            colors = ['🔵', '🔴', '🟢', '🟡', '🟠', '🟣']
            
            for b, x, y, o, l, L in terrain_final.buildings:
                if b.nom not in symbols:
                    symbols[b.nom] = colors[color_idx % len(colors)]
                    color_idx += 1
                for i in range(l):
                    for j in range(L):
                        viz[y + j, x + i] = symbols[b.nom]
            
            viz[viz == '1'] = '⬜'
            viz[viz == '0'] = '⬛'
            st.dataframe(pd.DataFrame(viz), use_container_width=True)
            
            # Légende
            if symbols:
                st.subheader("🏷️ Légende")
                cols = st.columns(4)
                for i, (nom, sym) in enumerate(symbols.items()):
                    with cols[i % 4]:
                        st.write(f"{sym} {nom}")
            
            # Tableau des boosts
            if boosts:
                st.subheader("📈 Boosts par bâtiment")
                st.dataframe(pd.DataFrame(boosts), use_container_width=True)
                
                # Résumé par type
                st.subheader("📊 Résumé des boosts")
                summary = []
                for prod, c in totals.items():
                    if prod in counts:
                        summary.append({
                            "Production": prod,
                            "Culture totale": f"{c:.0f}",
                            "0%": counts[prod][0],
                            "25%": counts[prod][25],
                            "50%": counts[prod][50],
                            "100%": counts[prod][100]
                        })
                if summary:
                    st.dataframe(pd.DataFrame(summary), use_container_width=True)

if __name__ == "__main__":
    main()