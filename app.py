import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import io
from dataclasses import dataclass
from typing import List, Tuple, Optional, Dict
import copy
import time
import random

# Constantes pour les couleurs
COULEUR_CULTUREL = "FFD966"  # Orange
COULEUR_PRODUCTEUR = "92D050"  # Vert
COULEUR_NEUTRE = "D9D9D9"  # Gris
COULEUR_VIDE = "FFFFFF"  # Blanc

@dataclass
class Batiment:
    nom: str
    longueur: int
    largeur: int
    quantite: int
    type: str  # 'culturel', 'producteur', 'neutre'
    culture: float
    rayonnement: int
    boost_25: float
    boost_50: float
    boost_100: float
    production: str
    
    def __post_init__(self):
        if self.type not in ['culturel', 'producteur', 'neutre']:
            if pd.isna(self.culture) or self.culture == 0:
                self.type = 'neutre'
            elif self.type == 'culturel' or (not pd.isna(self.culture) and self.culture > 0):
                self.type = 'culturel'
            elif self.type == 'producteur' or (not pd.isna(self.production) and self.production != ''):
                self.type = 'producteur'

@dataclass
class Position:
    x: int
    y: int
    orientation: str  # 'H' ou 'V'

class PlacementBatimentsOptimise:
    def __init__(self, terrain: List[List[int]], batiments: List[Batiment]):
        self.terrain_original = terrain
        self.hauteur = len(terrain)
        self.largeur = len(terrain[0]) if terrain else 0
        self.terrain = [[0 if cell == 0 else -1 for cell in row] for row in terrain]
        self.batiments = batiments
        self.batiments_places = []
        self.batiments_non_places = []
        self.journal = []
        self.culture_recue = {}
        self.max_tentatives = 1000  # Limite de tentatives pour éviter les boucles infinies
        self.start_time = None
        self.max_duration = 30  # Durée maximale en secondes
        
    def log(self, message):
        self.journal.append(message)
        
    def case_est_libre(self, x: int, y: int) -> bool:
        if x < 0 or x >= self.largeur or y < 0 or y >= self.hauteur:
            return False
        return self.terrain[y][x] == -1
    
    def peut_placer_batiment(self, batiment: Batiment, pos: Position) -> bool:
        if pos.orientation == 'H':
            longueur, largeur = batiment.longueur, batiment.largeur
        else:
            longueur, largeur = batiment.largeur, batiment.longueur
            
        # Vérifier les limites
        if pos.x + longueur > self.largeur or pos.y + largeur > self.hauteur:
            return False
            
        # Vérifier que toutes les cases sont libres
        for dx in range(longueur):
            for dy in range(largeur):
                if not self.case_est_libre(pos.x + dx, pos.y + dy):
                    return False
        return True
    
    def placer_batiment(self, batiment: Batiment, pos: Position, index: int):
        if pos.orientation == 'H':
            longueur, largeur = batiment.longueur, batiment.largeur
        else:
            longueur, largeur = batiment.largeur, batiment.longueur
            
        # Marquer les cases
        for dx in range(longueur):
            for dy in range(largeur):
                self.terrain[pos.y + dy][pos.x + dx] = index
                
        self.batiments_places.append((batiment, pos, index))
        self.log(f"PLACÉ: {batiment.nom} à ({pos.x}, {pos.y}) orientation {pos.orientation}")
        
    def retirer_batiment(self, index: int):
        # Retirer du terrain
        for y in range(self.hauteur):
            for x in range(self.largeur):
                if self.terrain[y][x] == index:
                    self.terrain[y][x] = -1
                    
        # Retirer de la liste
        self.batiments_places = [bp for bp in self.batiments_places if bp[2] != index]
        self.log(f"RETIRÉ: batiment index {index}")
        
    def calculer_cases_libres(self) -> int:
        return sum(1 for y in range(self.hauteur) for x in range(self.largeur) if self.terrain[y][x] == -1)
    
    def trouver_positions_rapide(self, batiment: Batiment) -> List[Position]:
        """Version optimisée pour trouver rapidement des positions"""
        positions = []
        
        # Stratégie: chercher d'abord dans les zones prometteuses
        if batiment.type == 'neutre':
            # Pour les neutres, priorité aux bords
            for y in range(self.hauteur):
                for x in range(self.largeur):
                    # Vérifier si on est près d'un bord
                    if x == 0 or y == 0 or x + batiment.longueur >= self.largeur or y + batiment.largeur >= self.hauteur:
                        pos_h = Position(x, y, 'H')
                        if self.peut_placer_batiment(batiment, pos_h):
                            positions.append(pos_h)
                            
                    if batiment.longueur != batiment.largeur:
                        if x == 0 or y == 0 or x + batiment.largeur >= self.largeur or y + batiment.longueur >= self.hauteur:
                            pos_v = Position(x, y, 'V')
                            if self.peut_placer_batiment(batiment, pos_v):
                                positions.append(pos_v)
        else:
            # Pour les autres, recherche normale mais limitée
            echantillon = min(50, self.largeur * self.hauteur)
            for _ in range(echantillon):
                x = random.randint(0, max(0, self.largeur - 1))
                y = random.randint(0, max(0, self.hauteur - 1))
                
                pos_h = Position(x, y, 'H')
                if self.peut_placer_batiment(batiment, pos_h):
                    positions.append(pos_h)
                    
                if batiment.longueur != batiment.largeur:
                    pos_v = Position(x, y, 'V')
                    if self.peut_placer_batiment(batiment, pos_v):
                        positions.append(pos_v)
                        
                if len(positions) >= 10:  # Limite le nombre de positions à tester
                    break
        
        return positions[:20]  # Garde seulement les 20 meilleures positions
    
    def calculer_score_emplacement_rapide(self, batiment: Batiment, pos: Position) -> float:
        """Version simplifiée du calcul de score"""
        score = 0
        
        # Score basé sur la proximité des bords pour les neutres
        if batiment.type == 'neutre':
            if pos.orientation == 'H':
                dist_bord_x = min(pos.x, self.largeur - (pos.x + batiment.longueur))
                dist_bord_y = min(pos.y, self.hauteur - (pos.y + batiment.largeur))
            else:
                dist_bord_x = min(pos.x, self.largeur - (pos.x + batiment.largeur))
                dist_bord_y = min(pos.y, self.hauteur - (pos.y + batiment.longueur))
            
            score += (100 - dist_bord_x * 10) + (100 - dist_bord_y * 10)
        
        # Pénalité pour les positions qui fragmentent l'espace
        cases_voisines = 0
        if pos.orientation == 'H':
            for dx in [-1, batiment.longueur]:
                for dy in range(batiment.largeur):
                    if 0 <= pos.x + dx < self.largeur and 0 <= pos.y + dy < self.hauteur:
                        if self.terrain[pos.y + dy][pos.x + dx] != -1:
                            cases_voisines += 1
        else:
            for dx in range(batiment.largeur):
                for dy in [-1, batiment.longueur]:
                    if 0 <= pos.x + dx < self.largeur and 0 <= pos.y + dy < self.hauteur:
                        if self.terrain[pos.y + dy][pos.x + dx] != -1:
                            cases_voisines += 1
        
        score += cases_voisines * 5  # Légère préférence pour les placements adjacents
        
        return score
    
    def calculer_culture_recue_rapide(self):
        """Version optimisée du calcul de culture"""
        self.culture_recue = {}
        
        # Pré-calculer les zones de rayonnement des bâtiments culturels
        zones_rayonnement = []
        for bat_culturel, pos_c, index_c in self.batiments_places:
            if bat_culturel.type != 'culturel':
                continue
                
            if pos_c.orientation == 'H':
                x1 = pos_c.x - bat_culturel.rayonnement
                y1 = pos_c.y - bat_culturel.rayonnement
                x2 = pos_c.x + bat_culturel.longueur - 1 + bat_culturel.rayonnement
                y2 = pos_c.y + bat_culturel.largeur - 1 + bat_culturel.rayonnement
            else:
                x1 = pos_c.x - bat_culturel.rayonnement
                y1 = pos_c.y - bat_culturel.rayonnement
                x2 = pos_c.x + bat_culturel.largeur - 1 + bat_culturel.rayonnement
                y2 = pos_c.y + bat_culturel.longueur - 1 + bat_culturel.rayonnement
                
            zones_rayonnement.append((x1, y1, x2, y2, bat_culturel.culture))
        
        # Pour chaque producteur, calculer la culture
        for batiment, pos, index in self.batiments_places:
            if batiment.type != 'producteur':
                continue
                
            if pos.orientation == 'H':
                x1_p = pos.x
                y1_p = pos.y
                x2_p = pos.x + batiment.longueur - 1
                y2_p = pos.y + batiment.largeur - 1
            else:
                x1_p = pos.x
                y1_p = pos.y
                x2_p = pos.x + batiment.largeur - 1
                y2_p = pos.y + batiment.longueur - 1
            
            culture_totale = 0
            for x1_c, y1_c, x2_c, y2_c, culture in zones_rayonnement:
                if (x2_p >= x1_c and x1_p <= x2_c and y2_p >= y1_c and y1_p <= y2_c):
                    culture_totale += culture
                    
            self.culture_recue[index] = culture_totale
    
    def get_boost_niveau(self, culture_recue: float, batiment: Batiment) -> str:
        if culture_recue >= batiment.boost_100:
            return "100%"
        elif culture_recue >= batiment.boost_50:
            return "50%"
        elif culture_recue >= batiment.boost_25:
            return "25%"
        else:
            return "0%"
    
    def placer_tous_batiments_optimise(self):
        """Version optimisée du placement avec timeout"""
        self.start_time = time.time()
        self.log("DÉBUT DU PLACEMENT OPTIMISÉ")
        
        # Préparer la liste des bâtiments à placer
        tous_batiments = []
        for batiment in self.batiments:
            for _ in range(batiment.quantite):
                nouveau_bat = copy.deepcopy(batiment)
                nouveau_bat.quantite = 1
                tous_batiments.append(nouveau_bat)
        
        # Trier les bâtiments par taille (les plus grands d'abord)
        tous_batiments.sort(key=lambda b: -(b.longueur * b.largeur))
        
        # Séparer par type
        neutres = [b for b in tous_batiments if b.type == 'neutre']
        culturels = [b for b in tous_batiments if b.type == 'culturel']
        producteurs = [b for b in tous_batiments if b.type == 'producteur']
        
        # Stratégie de placement: placer les plus grands d'abord
        ordre_placement = neutres + culturels + producteurs
        
        tentatives = 0
        index_actuel = 0
        
        while index_actuel < len(ordre_placement) and tentatives < self.max_tentatives:
            # Vérifier le timeout
            if time.time() - self.start_time > self.max_duration:
                self.log("TIMEOUT: Arrêt de l'optimisation (dépassement du temps limite)")
                break
                
            batiment = ordre_placement[index_actuel]
            
            self.log(f"ÉVALUATION: {batiment.nom} (taille: {batiment.longueur}x{batiment.largeur})")
            
            # Trouver des positions
            positions = self.trouver_positions_rapide(batiment)
            
            if positions:
                # Trier par score
                positions_avec_score = [(p, self.calculer_score_emplacement_rapide(batiment, p)) 
                                      for p in positions]
                positions_avec_score.sort(key=lambda x: x[1], reverse=True)
                
                # Prendre la meilleure position
                meilleure_pos = positions_avec_score[0][0]
                self.placer_batiment(batiment, meilleure_pos, index_actuel)
                index_actuel += 1
            else:
                self.log(f"IMPOSSIBLE: {batiment.nom} - aucune position disponible")
                self.batiments_non_places.append(batiment)
                index_actuel += 1
            
            tentatives += 1
            
            # Afficher la progression
            if index_actuel % 5 == 0:
                self.log(f"Progression: {index_actuel}/{len(ordre_placement)} bâtiments traités")
        
        # Ajouter les bâtiments non traités à la liste des non placés
        if index_actuel < len(ordre_placement):
            self.batiments_non_places.extend(ordre_placement[index_actuel:])
        
        # Calculer la culture reçue
        self.calculer_culture_recue_rapide()
        
        self.log(f"FIN DU PLACEMENT - {len(self.batiments_places)} placés, {len(self.batiments_non_places)} non placés")
        return len(self.batiments_places) > 0

def lire_fichier_excel(file) -> Tuple[List[List[int]], List[Batiment]]:
    """Lit le fichier Excel et retourne le terrain et la liste des bâtiments"""
    xl = pd.ExcelFile(file)
    
    # Lire le terrain (premier onglet)
    df_terrain = pd.read_excel(xl, sheet_name=0, header=None)
    terrain = df_terrain.values.tolist()
    
    # Lire les bâtiments (deuxième onglet)
    df_batiments = pd.read_excel(xl, sheet_name=1)
    batiments = []
    
    for _, row in df_batiments.iterrows():
        # Déterminer le type
        type_bat = 'neutre'
        if not pd.isna(row.get('Culture', 0)) and row.get('Culture', 0) > 0:
            type_bat = 'culturel'
        elif not pd.isna(row.get('Production', '')) and row.get('Production', '') != '':
            type_bat = 'producteur'
            
        batiment = Batiment(
            nom=row['Nom'],
            longueur=int(row['Longueur']),
            largeur=int(row['Largeur']),
            quantite=int(row['Quantite']),
            type=type_bat,
            culture=float(row['Culture']) if not pd.isna(row.get('Culture', 0)) else 0,
            rayonnement=int(row['Rayonnement']) if not pd.isna(row.get('Rayonnement', 0)) else 0,
            boost_25=float(row['Boost 25%']) if not pd.isna(row.get('Boost 25%', 0)) else 0,
            boost_50=float(row['Boost 50%']) if not pd.isna(row.get('Boost 50%', 0)) else 0,
            boost_100=float(row['Boost 100%']) if not pd.isna(row.get('Boost 100%', 0)) else 0,
            production=row['Production'] if not pd.isna(row.get('Production', '')) else ''
        )
        batiments.append(batiment)
        
    return terrain, batiments

def generer_fichier_resultat(placement: PlacementBatimentsOptimise) -> bytes:
    """Génère un fichier Excel avec les résultats"""
    wb = Workbook()
    
    # Feuille Journal
    ws_journal = wb.active
    ws_journal.title = "Journal"
    ws_journal.append(["Journal des opérations"])
    for ligne in placement.journal:
        ws_journal.append([ligne])
    
    # Feuille Statistiques
    ws_stats = wb.create_sheet("Statistiques")
    ws_stats.append(["Statistiques de placement"])
    
    # Calcul des boosts
    boosts = {"0%": 0, "25%": 0, "50%": 0, "100%": 0}
    culture_par_type = {"guerison": 0, "nourriture": 0, "or": 0}
    culture_totale = 0
    
    for batiment, _, index in placement.batiments_places:
        if batiment.type == 'producteur':
            culture = placement.culture_recue.get(index, 0)
            culture_totale += culture
            boost = placement.get_boost_niveau(culture, batiment)
            boosts[boost] += 1
            
            # Catégoriser par type de production
            prod = str(batiment.production).lower()
            if 'guerison' in prod:
                culture_par_type["guerison"] += culture
            elif 'nourriture' in prod:
                culture_par_type["nourriture"] += culture
            elif 'or' in prod:
                culture_par_type["or"] += culture
    
    ws_stats.append(["Culture totale reçue:", culture_totale])
    ws_stats.append([])
    ws_stats.append(["Boosts atteints:"])
    for boost, count in boosts.items():
        ws_stats.append([f"Boost {boost}:", count, "bâtiments"])
    
    ws_stats.append([])
    ws_stats.append(["Culture par type:"])
    for type_culture, valeur in culture_par_type.items():
        ws_stats.append([f"Culture {type_culture}:", valeur])
    
    ws_stats.append([])
    ws_stats.append(["Cases non utilisées:", placement.calculer_cases_libres()])
    
    # Calcul des cases des bâtiments non placés
    cases_non_placees = sum(b.longueur * b.largeur for b in placement.batiments_non_places)
    ws_stats.append(["Cases des bâtiments non placés:", cases_non_placees])
    ws_stats.append(["Bâtiments placés:", len(placement.batiments_places)])
    ws_stats.append(["Bâtiments non placés:", len(placement.batiments_non_places)])
    
    # Feuille Terrain
    ws_terrain = wb.create_sheet("Terrain final")
    
    # Définir les styles
    fill_culturel = PatternFill(start_color=COULEUR_CULTUREL, end_color=COULEUR_CULTUREL, fill_type="solid")
    fill_producteur = PatternFill(start_color=COULEUR_PRODUCTEUR, end_color=COULEUR_PRODUCTEUR, fill_type="solid")
    fill_neutre = PatternFill(start_color=COULEUR_NEUTRE, end_color=COULEUR_NEUTRE, fill_type="solid")
    fill_vide = PatternFill(start_color=COULEUR_VIDE, end_color=COULEUR_VIDE, fill_type="solid")
    
    # Créer une matrice pour le terrain avec les noms des bâtiments
    terrain_noms = [["" for _ in range(placement.largeur)] for _ in range(placement.hauteur)]
    
    for batiment, pos, index in placement.batiments_places:
        if pos.orientation == 'H':
            longueur, largeur = batiment.longueur, batiment.largeur
        else:
            longueur, largeur = batiment.largeur, batiment.longueur
            
        for dx in range(longueur):
            for dy in range(largeur):
                x = pos.x + dx
                y = pos.y + dy
                if batiment.type == 'culturel':
                    couleur = "CULTUREL"
                elif batiment.type == 'producteur':
                    couleur = "PRODUCTEUR"
                else:
                    couleur = "NEUTRE"
                terrain_noms[y][x] = f"{batiment.nom}"
    
    # Écrire le terrain
    for y in range(placement.hauteur):
        row = []
        for x in range(placement.largeur):
            if placement.terrain[y][x] == -1:
                row.append("")
            else:
                row.append(terrain_noms[y][x])
        ws_terrain.append(row)
    
    # Appliquer les couleurs
    for y in range(placement.hauteur):
        for x in range(placement.largeur):
            cell = ws_terrain.cell(row=y+1, column=x+1)
            if placement.terrain[y][x] == -1:
                cell.fill = fill_vide
            else:
                # Trouver le type de bâtiment
                for batiment, pos, index in placement.batiments_places:
                    if pos.orientation == 'H':
                        if pos.x <= x < pos.x + batiment.longueur and pos.y <= y < pos.y + batiment.largeur:
                            if batiment.type == 'culturel':
                                cell.fill = fill_culturel
                            elif batiment.type == 'producteur':
                                cell.fill = fill_producteur
                            else:
                                cell.fill = fill_neutre
                            break
                    else:
                        if pos.x <= x < pos.x + batiment.largeur and pos.y <= y < pos.y + batiment.longueur:
                            if batiment.type == 'culturel':
                                cell.fill = fill_culturel
                            elif batiment.type == 'producteur':
                                cell.fill = fill_producteur
                            else:
                                cell.fill = fill_neutre
                            break
    
    # Ajuster la largeur des colonnes
    for column in ws_terrain.columns:
        max_length = 0
        for cell in column:
            try:
                if cell.value and len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 30)
        ws_terrain.column_dimensions[column[0].column_letter].width = adjusted_width
    
    # Feuille Bâtiments non placés
    ws_non_places = wb.create_sheet("Non placés")
    ws_non_places.append(["Bâtiments non placés"])
    if placement.batiments_non_places:
        for batiment in placement.batiments_non_places:
            ws_non_places.append([batiment.nom, f"L:{batiment.longueur}", f"l:{batiment.largeur}", batiment.type])
    else:
        ws_non_places.append(["Tous les bâtiments ont été placés"])
    
    # Sauvegarder dans un buffer
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output.getvalue()

def main():
    st.set_page_config(page_title="Placement de Bâtiments", layout="wide")
    
    st.title("🏗️ Optimiseur de Placement de Bâtiments")
    st.markdown("""
    Cet outil permet d'optimiser le placement de bâtiments sur un terrain 
    en maximisant les boosts de production.
    """)
    
    with st.sidebar:
        st.header("📁 Chargement des données")
        uploaded_file = st.file_uploader(
            "Choisir le fichier Excel",
            type=['xlsx', 'xls'],
            help="Le fichier doit contenir deux onglets: terrain (0/1) et bâtiments"
        )
        
        if uploaded_file:
            st.success("Fichier chargé avec succès!")
            
            # Options d'optimisation
            st.header("⚙️ Options d'optimisation")
            timeout = st.slider("Temps maximum (secondes)", 5, 60, 30)
            max_tentatives = st.slider("Nombre maximum de tentatives", 100, 5000, 1000)
            
            if st.button("🚀 Lancer l'optimisation", type="primary"):
                st.session_state['run_optimization'] = True
                st.session_state['timeout'] = timeout
                st.session_state['max_tentatives'] = max_tentatives
    
    if uploaded_file and st.session_state.get('run_optimization', False):
        with st.spinner("Optimisation en cours..."):
            try:
                # Lire le fichier
                terrain, batiments = lire_fichier_excel(uploaded_file)
                
                # Afficher les informations
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Taille du terrain", f"{len(terrain)} × {len(terrain[0])}")
                with col2:
                    cases_libres = sum(1 for row in terrain for cell in row if cell == 1)
                    st.metric("Cases libres", cases_libres)
                with col3:
                    total_batiments = sum(b.quantite for b in batiments)
                    st.metric("Total bâtiments", total_batiments)
                
                # Barre de progression
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Créer et exécuter le placement optimisé
                placement = PlacementBatimentsOptimise(terrain, batiments)
                placement.max_duration = st.session_state['timeout']
                placement.max_tentatives = st.session_state['max_tentatives']
                
                # Simulation de progression
                for i in range(10):
                    time.sleep(0.1)
                    progress_bar.progress((i + 1) * 10)
                    status_text.text(f"Étape {i+1}/10: Recherche de placements...")
                
                succes = placement.placer_tous_batiments_optimise()
                progress_bar.progress(100)
                status_text.text("Optimisation terminée!")
                
                # Afficher les résultats
                if succes:
                    st.success("✅ Optimisation terminée avec succès!")
                else:
                    st.warning("⚠️ Optimisation terminée avec des limitations")
                
                # Onglets de résultats
                tab1, tab2, tab3, tab4 = st.tabs(["📊 Statistiques", "📜 Journal", "🗺️ Visualisation", "📦 Non placés"])
                
                with tab1:
                    st.subheader("Statistiques de placement")
                    
                    # Calcul des stats
                    cases_utilisees = sum(1 for y in range(placement.hauteur) 
                                         for x in range(placement.largeur) 
                                         if placement.terrain[y][x] != -1)
                    cases_libres_final = placement.calculer_cases_libres()
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Bâtiments placés", len(placement.batiments_places))
                    with col2:
                        st.metric("Bâtiments non placés", len(placement.batiments_non_places))
                    with col3:
                        st.metric("Taux de remplissage", 
                                 f"{cases_utilisees/(placement.hauteur*placement.largeur)*100:.1f}%")
                    
                    # Culture et boosts
                    st.subheader("🎯 Analyse des boosts")
                    
                    boosts_data = []
                    culture_totale = 0
                    for batiment, _, index in placement.batiments_places:
                        if batiment.type == 'producteur':
                            culture = placement.culture_recue.get(index, 0)
                            culture_totale += culture
                            boost = placement.get_boost_niveau(culture, batiment)
                            boosts_data.append({
                                "Bâtiment": batiment.nom,
                                "Culture reçue": f"{culture:.1f}",
                                "Boost": boost
                            })
                    
                    if boosts_data:
                        st.dataframe(pd.DataFrame(boosts_data), use_container_width=True)
                        st.metric("Culture totale reçue", f"{culture_totale:.1f}")
                    else:
                        st.info("Aucun bâtiment producteur placé")
                
                with tab2:
                    st.subheader("Journal des opérations")
                    # Afficher les dernières 50 lignes du journal
                    for ligne in placement.journal[-50:]:
                        st.text(ligne)
                    
                    if len(placement.journal) > 50:
                        st.caption(f"Affichage des 50 dernières lignes sur {len(placement.journal)}")
                
                with tab3:
                    st.subheader("Visualisation du terrain")
                    
                    # Créer une représentation avec des couleurs
                    html_table = "<table style='border-collapse: collapse;'>"
                    for y in range(placement.hauteur):
                        html_table += "<tr>"
                        for x in range(placement.largeur):
                            if placement.terrain[y][x] == -1:
                                couleur = "#FFFFFF"
                                texte = "⬜"
                            else:
                                # Trouver le type
                                trouve = False
                                for batiment, pos, index in placement.batiments_places:
                                    if pos.orientation == 'H':
                                        if pos.x <= x < pos.x + batiment.longueur and pos.y <= y < pos.y + batiment.largeur:
                                            if batiment.type == 'culturel':
                                                couleur = "#FFD966"
                                                texte = "🟧"
                                            elif batiment.type == 'producteur':
                                                couleur = "#92D050"
                                                texte = "🟩"
                                            else:
                                                couleur = "#D9D9D9"
                                                texte = "⬛"
                                            trouve = True
                                            break
                                    else:
                                        if pos.x <= x < pos.x + batiment.largeur and pos.y <= y < pos.y + batiment.longueur:
                                            if batiment.type == 'culturel':
                                                couleur = "#FFD966"
                                                texte = "🟧"
                                            elif batiment.type == 'producteur':
                                                couleur = "#92D050"
                                                texte = "🟩"
                                            else:
                                                couleur = "#D9D9D9"
                                                texte = "⬛"
                                            trouve = True
                                            break
                                if not trouve:
                                    couleur = "#FF0000"
                                    texte = "❌"
                            
                            html_table += f"<td style='background-color: {couleur}; width: 30px; height: 30px; text-align: center; border: 1px solid #ccc;'>{texte}</td>"
                        html_table += "</tr>"
                    html_table += "</table>"
                    
                    st.markdown(html_table, unsafe_allow_html=True)
                    
                    st.markdown("""
                    **Légende:**
                    - 🟧 : Bâtiment culturel
                    - 🟩 : Bâtiment producteur
                    - ⬛ : Bâtiment neutre
                    - ⬜ : Case libre
                    """)
                
                with tab4:
                    st.subheader("Bâtiments non placés")
                    if placement.batiments_non_places:
                        data = []
                        for b in placement.batiments_non_places:
                            data.append({
                                "Nom": b.nom,
                                "Type": b.type,
                                "Dimensions": f"{b.longueur}×{b.largeur}",
                                "Culture": b.culture if b.culture > 0 else "-"
                            })
                        st.dataframe(pd.DataFrame(data), use_container_width=True)
                        
                        cases_non_placees = sum(b.longueur * b.largeur for b in placement.batiments_non_places)
                        st.warning(f"⚠️ {cases_non_placees} cases non utilisées à cause des bâtiments non placés")
                    else:
                        st.success("✅ Tous les bâtiments ont été placés!")
                
                # Téléchargement du fichier résultat
                st.divider()
                output_data = generer_fichier_resultat(placement)
                st.download_button(
                    label="📥 Télécharger le fichier Excel des résultats",
                    data=output_data,
                    file_name="resultat_placement.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                
            except Exception as e:
                st.error(f"Erreur lors de l'optimisation: {str(e)}")
                st.exception(e)

if __name__ == "__main__":
    main()