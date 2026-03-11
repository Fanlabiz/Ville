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

class PlacementBatiments:
    def __init__(self, terrain: List[List[int]], batiments: List[Batiment]):
        self.terrain_original = terrain
        self.hauteur = len(terrain)
        self.largeur = len(terrain[0]) if terrain else 0
        self.terrain = [[0 if cell == 0 else -1 for cell in row] for row in terrain]
        self.batiments = batiments
        self.batiments_places = []
        self.batiments_non_places = []
        self.journal = []
        self.culture_recue = {}  # index batiment -> culture recue
        
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
    
    def trouver_plus_grand_batiment_restant(self, batiments_restants: List[Batiment]) -> int:
        if not batiments_restants:
            return 0
        return max(max(b.longueur, b.largeur) for b in batiments_restants)
    
    def assez_de_place(self, batiments_restants: List[Batiment]) -> bool:
        cases_libres = self.calculer_cases_libres()
        cases_necessaires = sum(b.longueur * b.largeur * b.quantite for b in batiments_restants)
        return cases_libres >= cases_necessaires
    
    def trouver_positions_possibles(self, batiment: Batiment) -> List[Position]:
        positions = []
        # Orientation horizontale
        for y in range(self.hauteur):
            for x in range(self.largeur):
                pos_h = Position(x, y, 'H')
                if self.peut_placer_batiment(batiment, pos_h):
                    positions.append(pos_h)
                    
        # Orientation verticale si différente
        if batiment.longueur != batiment.largeur:
            for y in range(self.hauteur):
                for x in range(self.largeur):
                    pos_v = Position(x, y, 'V')
                    if self.peut_placer_batiment(batiment, pos_v):
                        positions.append(pos_v)
        return positions
    
    def calculer_score_emplacement(self, batiment: Batiment, pos: Position, 
                                  batiments_restants: List[Batiment]) -> float:
        score = 0
        
        # Priorité: bords du terrain pour les neutres
        if batiment.type == 'neutre':
            if pos.orientation == 'H':
                if pos.x == 0 or pos.x + batiment.longueur >= self.largeur:
                    score += 100
                if pos.y == 0 or pos.y + batiment.largeur >= self.hauteur:
                    score += 50
            else:
                if pos.x == 0 or pos.x + batiment.largeur >= self.largeur:
                    score += 100
                if pos.y == 0 or pos.y + batiment.longueur >= self.hauteur:
                    score += 50
                    
        # Éviter de bloquer des espaces
        cases_libres_apres = self.calculer_cases_libres() - (batiment.longueur * batiment.largeur)
        if batiments_restants:
            plus_grand = self.trouver_plus_grand_batiment_restant(batiments_restants)
            # Vérifier si on peut encore placer le plus grand bâtiment
            if cases_libres_apres < plus_grand * plus_grand:
                score -= 200
                
        return score
    
    def calculer_culture_recue(self):
        """Calcule la culture reçue par chaque bâtiment producteur"""
        self.culture_recue = {}
        
        # Pour chaque bâtiment placé
        for i, (batiment, pos, index) in enumerate(self.batiments_places):
            if batiment.type != 'producteur':
                continue
                
            culture_totale = 0
            
            # Vérifier les bâtiments culturels dans le rayonnement
            for j, (bat_culturel, pos_c, index_c) in enumerate(self.batiments_places):
                if bat_culturel.type != 'culturel':
                    continue
                    
                # Déterminer les cases du bâtiment culturel
                if pos_c.orientation == 'H':
                    x1_c, y1_c = pos_c.x, pos_c.y
                    x2_c = x1_c + bat_culturel.longueur - 1
                    y2_c = y1_c + bat_culturel.largeur - 1
                else:
                    x1_c, y1_c = pos_c.x, pos_c.y
                    x2_c = x1_c + bat_culturel.largeur - 1
                    y2_c = y1_c + bat_culturel.longueur - 1
                
                # Étendre la zone de rayonnement
                x1_c -= bat_culturel.rayonnement
                y1_c -= bat_culturel.rayonnement
                x2_c += bat_culturel.rayonnement
                y2_c += bat_culturel.rayonnement
                
                # Déterminer les cases du bâtiment producteur
                if pos.orientation == 'H':
                    x1_p, y1_p = pos.x, pos.y
                    x2_p = x1_p + batiment.longueur - 1
                    y2_p = y1_p + batiment.largeur - 1
                else:
                    x1_p, y1_p = pos.x, pos.y
                    x2_p = x1_p + batiment.largeur - 1
                    y2_p = y1_p + batiment.longueur - 1
                
                # Vérifier l'intersection
                if (x2_p >= x1_c and x1_p <= x2_c and 
                    y2_p >= y1_c and y1_p <= y2_c):
                    culture_totale += bat_culturel.culture
                    
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
    
    def placer_tous_batiments(self):
        self.log("DÉBUT DU PLACEMENT")
        
        # Séparer les bâtiments par type
        batiments_neutres = [b for b in self.batiments if b.type == 'neutre' and b.quantite > 0]
        batiments_culturels = [b for b in self.batiments if b.type == 'culturel' and b.quantite > 0]
        batiments_producteurs = [b for b in self.batiments if b.type == 'producteur' and b.quantite > 0]
        
        # Étendre les listes selon la quantité
        tous_batiments = []
        for b in batiments_neutres:
            tous_batiments.extend([copy.deepcopy(b) for _ in range(b.quantite)])
        for b in batiments_culturels:
            tous_batiments.extend([copy.deepcopy(b) for _ in range(b.quantite)])
        for b in batiments_producteurs:
            tous_batiments.extend([copy.deepcopy(b) for _ in range(b.quantite)])
            
        # Ordonner: neutres d'abord, puis alterner culturels et producteurs
        batiments_ordre = []
        batiments_ordre.extend(tous_batiments[:len(batiments_neutres)])
        
        i_culturel = len(batiments_neutres)
        i_producteur = len(batiments_neutres) + len(batiments_culturels)
        
        while i_culturel < len(batiments_neutres) + len(batiments_culturels) or \
              i_producteur < len(tous_batiments):
            if i_culturel < len(batiments_neutres) + len(batiments_culturels):
                batiments_ordre.append(tous_batiments[i_culturel])
                i_culturel += 1
            if i_producteur < len(tous_batiments):
                batiments_ordre.append(tous_batiments[i_producteur])
                i_producteur += 1
        
        index_actuel = 0
        pile_placements = []
        
        while index_actuel < len(batiments_ordre):
            batiment = batiments_ordre[index_actuel]
            batiments_restants = batiments_ordre[index_actuel + 1:]
            
            self.log(f"ÉVALUATION: {batiment.nom}")
            
            positions = self.trouver_positions_possibles(batiment)
            
            if not positions:
                self.log(f"IMPOSSIBLE: {batiment.nom} - aucune position disponible")
                # Backtracking
                if pile_placements:
                    dernier_index, dernier_positions = pile_placements.pop()
                    self.retirer_batiment(dernier_index)
                    index_actuel = dernier_index
                    
                    # Essayer une autre position
                    if dernier_positions:
                        nouvelle_pos = dernier_positions.pop(0)
                        batiment_prec = batiments_ordre[dernier_index]
                        self.placer_batiment(batiment_prec, nouvelle_pos, dernier_index)
                        pile_placements.append((dernier_index, dernier_positions))
                        index_actuel += 1
                else:
                    self.batiments_non_places.append(batiment)
                    index_actuel += 1
                continue
            
            # Trier les positions par score
            positions_avec_score = [(p, self.calculer_score_emplacement(batiment, p, batiments_restants)) 
                                   for p in positions]
            positions_avec_score.sort(key=lambda x: x[1], reverse=True)
            positions_triees = [p for p, _ in positions_avec_score]
            
            # Essayer la meilleure position
            meilleure_pos = positions_triees[0]
            self.placer_batiment(batiment, meilleure_pos, index_actuel)
            pile_placements.append((index_actuel, positions_triees[1:]))
            index_actuel += 1
        
        # Calculer la culture reçue
        self.calculer_culture_recue()
        self.log("FIN DU PLACEMENT")

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

def generer_fichier_resultat(placement: PlacementBatiments) -> bytes:
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
                terrain_noms[y][x] = f"{batiment.nom} ({couleur})"
    
    # Écrire le terrain
    for y in range(placement.hauteur):
        row = []
        for x in range(placement.largeur):
            if placement.terrain[y][x] == -1:
                row.append("LIBRE")
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
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 30)
        ws_terrain.column_dimensions[column_letter].width = adjusted_width
    
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
            if st.button("🚀 Lancer l'optimisation", type="primary"):
                st.session_state['run_optimization'] = True
    
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
                    st.metric("Types de bâtiments", len(batiments))
                
                # Créer et exécuter le placement
                placement = PlacementBatiments(terrain, batiments)
                placement.placer_tous_batiments()
                
                # Afficher les résultats
                st.success("✅ Optimisation terminée!")
                
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
                
                with tab2:
                    st.subheader("Journal des opérations")
                    for ligne in placement.journal:
                        st.text(ligne)
                
                with tab3:
                    st.subheader("Visualisation du terrain")
                    
                    # Créer une représentation textuelle
                    terrain_viz = []
                    for y in range(placement.hauteur):
                        row = []
                        for x in range(placement.largeur):
                            if placement.terrain[y][x] == -1:
                                row.append("⬜")  # Libre
                            else:
                                # Trouver le type
                                trouve = False
                                for batiment, pos, index in placement.batiments_places:
                                    if pos.orientation == 'H':
                                        if pos.x <= x < pos.x + batiment.longueur and pos.y <= y < pos.y + batiment.largeur:
                                            if batiment.type == 'culturel':
                                                row.append("🟧")  # Orange
                                            elif batiment.type == 'producteur':
                                                row.append("🟩")  # Vert
                                            else:
                                                row.append("⬛")  # Gris
                                            trouve = True
                                            break
                                    else:
                                        if pos.x <= x < pos.x + batiment.largeur and pos.y <= y < pos.y + batiment.longueur:
                                            if batiment.type == 'culturel':
                                                row.append("🟧")
                                            elif batiment.type == 'producteur':
                                                row.append("🟩")
                                            else:
                                                row.append("⬛")
                                            trouve = True
                                            break
                                if not trouve:
                                    row.append("❌")  # Erreur
                        terrain_viz.append(" ".join(row))
                    
                    st.code("\n".join(terrain_viz))
                    
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