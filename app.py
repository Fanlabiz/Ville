import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import io
from copy import deepcopy
import random

# Configuration de la page Streamlit
st.set_page_config(page_title="Placement de bâtiments", layout="wide")
st.title("🏗️ Optimiseur de placement de bâtiments")

# Initialisation des variables de session
if 'journal' not in st.session_state:
    st.session_state.journal = []
if 'log_count' not in st.session_state:
    st.session_state.log_count = 0
if 'solution' not in st.session_state:
    st.session_state.solution = None

class Batiment:
    def __init__(self, nom, longueur, largeur, quantite, type_bat, culture, rayonnement, 
                 boost_25, boost_50, boost_100, production):
        self.nom = nom
        self.longueur = int(longueur)
        self.largeur = int(largeur)
        self.quantite = int(quantite)
        self.type = type_bat
        self.culture = float(culture) if culture and culture != '' else 0
        self.rayonnement = int(rayonnement) if rayonnement and rayonnement != '' else 0
        self.boost_25 = float(boost_25) if boost_25 and boost_25 != '' else 0
        self.boost_50 = float(boost_50) if boost_50 and boost_50 != '' else 0
        self.boost_100 = float(boost_100) if boost_100 and boost_100 != '' else 0
        self.production = production if production and production != '' else "Aucune"
        self.placed = False
        self.position = None
        self.orientation = None
        self.culture_recue = 0
        self.instance_id = id(self)  # Pour différencier les instances
        
    def surface(self):
        return self.longueur * self.largeur
    
    def __str__(self):
        return f"{self.nom} ({self.longueur}x{self.largeur})"

class Terrain:
    def __init__(self, grille):
        self.grille = np.array(grille, dtype=int)
        self.hauteur, self.largeur = self.grille.shape
        self.batiments_places = []
        
    def case_libre(self, x, y):
        return 0 <= x < self.hauteur and 0 <= y < self.largeur and self.grille[x, y] == 1
    
    def est_sur_bord(self, x, y, longueur, largeur, orientation):
        """Vérifie si un bâtiment est placé sur un bord"""
        if orientation == 'H':
            return (x == 0 or x + longueur - 1 == self.hauteur - 1 or 
                    y == 0 or y + largeur - 1 == self.largeur - 1)
        else:
            return (x == 0 or x + largeur - 1 == self.hauteur - 1 or 
                    y == 0 or y + longueur - 1 == self.largeur - 1)
    
    def trouver_tous_emplacements(self, batiment):
        """Trouve tous les emplacements possibles pour un bâtiment"""
        emplacements = []
        
        for x in range(self.hauteur):
            for y in range(self.largeur):
                # Orientation horizontale
                if y + batiment.largeur <= self.largeur and x + batiment.longueur <= self.hauteur:
                    valide = True
                    for i in range(batiment.longueur):
                        for j in range(batiment.largeur):
                            if not self.case_libre(x + i, y + j):
                                valide = False
                                break
                        if not valide:
                            break
                    if valide:
                        emplacements.append((x, y, 'H'))
                
                # Orientation verticale (si différent)
                if batiment.longueur != batiment.largeur:
                    if y + batiment.longueur <= self.largeur and x + batiment.largeur <= self.hauteur:
                        valide = True
                        for i in range(batiment.largeur):
                            for j in range(batiment.longueur):
                                if not self.case_libre(x + i, y + j):
                                    valide = False
                                    break
                            if not valide:
                                break
                        if valide:
                            emplacements.append((x, y, 'V'))
        
        return emplacements
    
    def peut_placer(self, batiment, x, y, orientation):
        """Vérifie si on peut placer un bâtiment à la position (x,y) avec l'orientation donnée"""
        if orientation == 'H':
            if x + batiment.longueur > self.hauteur or y + batiment.largeur > self.largeur:
                return False
            for i in range(batiment.longueur):
                for j in range(batiment.largeur):
                    if not self.case_libre(x + i, y + j):
                        return False
        else:  # Vertical
            if x + batiment.largeur > self.hauteur or y + batiment.longueur > self.largeur:
                return False
            for i in range(batiment.largeur):
                for j in range(batiment.longueur):
                    if not self.case_libre(x + i, y + j):
                        return False
        return True
    
    def placer_batiment(self, batiment, x, y, orientation):
        """Place un bâtiment sur le terrain"""
        if orientation == 'H':
            for i in range(batiment.longueur):
                for j in range(batiment.largeur):
                    self.grille[x + i, y + j] = 2  # Marqué comme occupé par un bâtiment
        else:
            for i in range(batiment.largeur):
                for j in range(batiment.longueur):
                    self.grille[x + i, y + j] = 2
        
        batiment.placed = True
        batiment.position = (x, y)
        batiment.orientation = orientation
        self.batiments_places.append(batiment)
        
    def enlever_batiment(self, batiment):
        """Enlève un bâtiment du terrain"""
        if batiment.position:
            x, y = batiment.position
            if batiment.orientation == 'H':
                for i in range(batiment.longueur):
                    for j in range(batiment.largeur):
                        self.grille[x + i, y + j] = 1
            else:
                for i in range(batiment.largeur):
                    for j in range(batiment.longueur):
                        self.grille[x + i, y + j] = 1
            
            self.batiments_places.remove(batiment)
            batiment.placed = False
            batiment.position = None
            batiment.orientation = None
    
    def calculer_culture_recue(self):
        """Calcule la culture reçue par tous les bâtiments producteurs"""
        producteurs = [b for b in self.batiments_places if b.type == 'Producteur']
        culturels = [b for b in self.batiments_places if b.type == 'Culturel']
        
        # Réinitialiser la culture reçue
        for prod in producteurs:
            prod.culture_recue = 0
        
        # Pour chaque culturel, ajouter sa culture aux producteurs dans son rayonnement
        for cult in culturels:
            if cult.placed:
                x, y = cult.position
                rayon = cult.rayonnement
                
                # Définir la zone de rayonnement (inclut le bâtiment lui-même et la bande autour)
                x_min = max(0, x - rayon)
                x_max = min(self.hauteur, x + cult.longueur + rayon)
                y_min = max(0, y - rayon)
                y_max = min(self.largeur, y + cult.largeur + rayon)
                
                for prod in producteurs:
                    if prod.placed:
                        px, py = prod.position
                        # Vérifier chaque case du producteur
                        if prod.orientation == 'H':
                            for i in range(prod.longueur):
                                for j in range(prod.largeur):
                                    case_x, case_y = px + i, py + j
                                    if (x_min <= case_x < x_max and 
                                        y_min <= case_y < y_max):
                                        prod.culture_recue += cult.culture
                                        break  # Une seule case suffit pour le rayonnement
                                else:
                                    continue
                                break
                        else:
                            for i in range(prod.largeur):
                                for j in range(prod.longueur):
                                    case_x, case_y = px + i, py + j
                                    if (x_min <= case_x < x_max and 
                                        y_min <= case_y < y_max):
                                        prod.culture_recue += cult.culture
                                        break
                                else:
                                    continue
                                break
    
    def get_boost_niveau(self, culture_recue, batiment):
        """Détermine le niveau de boost atteint"""
        if culture_recue >= batiment.boost_100:
            return 100
        elif culture_recue >= batiment.boost_50:
            return 50
        elif culture_recue >= batiment.boost_25:
            return 25
        else:
            return 0
    
    def surface_libre_totale(self):
        """Calcule la surface libre totale"""
        return np.sum(self.grille == 1)

def ajouter_log(message):
    """Ajoute un message au journal"""
    st.session_state.journal.append(message)
    st.session_state.log_count += 1
    if st.session_state.log_count > 1000:
        raise Exception("Limite de 1000 entrées dans le journal atteinte")

def priorite_bords(terrain, batiment):
    """Trouve un emplacement sur les bords si possible"""
    emplacements = terrain.trouver_tous_emplacements(batiment)
    
    # Priorité 1: Emplacements sur les bords
    bords = []
    interieur = []
    
    for emp in emplacements:
        x, y, orientation = emp
        if terrain.est_sur_bord(x, y, batiment.longueur, batiment.largeur, orientation):
            bords.append(emp)
        else:
            interieur.append(emp)
    
    # Retourner d'abord les bords, puis l'intérieur
    return bords + interieur

def calculer_score_placement(terrain, batiment, x, y, orientation, batiments_culturels_places, producteurs_a_venir):
    """Calcule un score pour un emplacement basé sur le potentiel de boost futur"""
    score = 0
    
    # Simuler le placement
    terrain_simule = deepcopy(terrain)
    terrain_simule.placer_batiment(deepcopy(batiment), x, y, orientation)
    
    # Calculer la culture actuelle
    terrain_simule.calculer_culture_recue()
    
    # Score basé sur la culture reçue par ce bâtiment (si producteur)
    if batiment.type == 'Producteur':
        for prod in terrain_simule.batiments_places:
            if prod.instance_id == batiment.instance_id:
                score += prod.culture_recue * 10
    
    # Score basé sur le rayonnement pour les futurs producteurs
    if batiment.type == 'Culturel':
        # Compter combien de cases libres sont dans le rayonnement
        x_min = max(0, x - batiment.rayonnement)
        x_max = min(terrain.hauteur, x + batiment.longueur + batiment.rayonnement)
        y_min = max(0, y - batiment.rayonnement)
        y_max = min(terrain.largeur, y + batiment.largeur + batiment.rayonnement)
        
        cases_dans_rayon = 0
        for i in range(x_min, x_max):
            for j in range(y_min, y_max):
                if terrain_simule.grille[i, j] == 1:  # Case libre
                    cases_dans_rayon += 1
        
        score += cases_dans_rayon * batiment.culture
    
    return score

def placer_batiments(terrain, batiments):
    """Algorithme principal de placement amélioré"""
    
    # Créer toutes les instances de bâtiments
    tous_batiments = []
    for b in batiments:
        for i in range(b.quantite):
            nouveau_bat = deepcopy(b)
            nouveau_bat.instance_id = id(nouveau_bat)
            tous_batiments.append(nouveau_bat)
    
    # Séparer par type
    neutres = [b for b in tous_batiments if b.type == 'Neutre']
    culturels = [b for b in tous_batiments if b.type == 'Culturel']
    producteurs = [b for b in tous_batiments if b.type == 'Producteur']
    
    # Trier par surface décroissante
    neutres.sort(key=lambda b: b.surface(), reverse=True)
    culturels.sort(key=lambda b: b.surface(), reverse=True)
    producteurs.sort(key=lambda b: b.surface(), reverse=True)
    
    # Journal du placement
    ajouter_log(f"Début du placement - Terrain: {terrain.hauteur}x{terrain.largeur}")
    ajouter_log(f"Bâtiments à placer: {len(neutres)} neutres, {len(culturels)} culturels, {len(producteurs)} producteurs")
    
    # Étape 1: Placer les neutres sur les bords
    neutres_non_places = neutres.copy()
    neutres_places = []
    
    for batiment in neutres:
        ajouter_log(f"Recherche d'emplacement pour neutre: {batiment.nom}")
        
        # Chercher d'abord sur les bords
        emplacements = priorite_bords(terrain, batiment)
        
        place = False
        for x, y, orientation in emplacements:
            if terrain.peut_placer(batiment, x, y, orientation):
                terrain.placer_batiment(batiment, x, y, orientation)
                neutres_places.append(batiment)
                neutres_non_places.remove(batiment)
                ajouter_log(f"  ✅ Placé sur bord: {batiment.nom} à ({x},{y}) orientation {orientation}")
                place = True
                break
        
        if not place:
            ajouter_log(f"  ❌ Impossible de placer {batiment.nom} sur un bord")
    
    ajouter_log(f"Neutres placés: {len(neutres_places)}/{len(neutres)}")
    
    # Étape 2: Placer le reste en alternance
    tous_reste = culturels + producteurs
    tous_reste.sort(key=lambda b: b.surface(), reverse=True)
    
    # Backtracking avec mémoire
    max_essais = 1000
    essai = 0
    
    while tous_reste and essai < max_essais:
        essai += 1
        batiment = tous_reste[0]
        
        ajouter_log(f"Essai {essai}: Tentative placement {batiment.nom} ({batiment.type})")
        
        # Trouver tous les emplacements possibles
        emplacements = terrain.trouver_tous_emplacements(batiment)
        
        if not emplacements:
            ajouter_log(f"  ❌ Aucun emplacement pour {batiment.nom}")
            # Backtracking
            if len(tous_reste) > 1:
                # Essayer de déplacer le précédent
                precedent = tous_reste[1] if len(tous_reste) > 1 else None
                if precedent and precedent.placed:
                    terrain.enlever_batiment(precedent)
                    tous_reste.insert(0, precedent)
                    ajouter_log(f"  ↩️ Retrait de {precedent.nom} pour réessayer")
                    continue
        
        # Trier les emplacements par score
        emplacements_avec_score = []
        for x, y, orientation in emplacements:
            score = calculer_score_placement(terrain, batiment, x, y, orientation, 
                                           [c for c in culturels if c.placed],
                                           [p for p in producteurs if not p.placed])
            emplacements_avec_score.append((score, x, y, orientation))
        
        emplacements_avec_score.sort(reverse=True)  # Meilleur score d'abord
        
        # Essayer les emplacements
        place = False
        for score, x, y, orientation in emplacements_avec_score:
            if terrain.peut_placer(batiment, x, y, orientation):
                terrain.placer_batiment(batiment, x, y, orientation)
                tous_reste.pop(0)
                ajouter_log(f"  ✅ Placé: {batiment.nom} à ({x},{y}) orientation {orientation} (score: {score:.1f})")
                place = True
                break
        
        if not place:
            ajouter_log(f"  ❌ Échec placement {batiment.nom}")
            # Backtracking
            if len(tous_reste) > 1:
                tous_reste = tous_reste[1:] + [tous_reste[0]]  # Rotation
            else:
                break
    
    # Étape 3: Essayer de placer les neutres restants à l'intérieur
    for batiment in neutres_non_places:
        ajouter_log(f"Tentative placement neutre restant: {batiment.nom}")
        emplacements = terrain.trouver_tous_emplacements(batiment)
        
        for x, y, orientation in emplacements:
            if terrain.peut_placer(batiment, x, y, orientation):
                terrain.placer_batiment(batiment, x, y, orientation)
                ajouter_log(f"  ✅ Placé à l'intérieur: {batiment.nom} à ({x},{y})")
                neutres_non_places.remove(batiment)
                break
    
    # Calculer la culture finale
    terrain.calculer_culture_recue()
    
    return terrain, neutres_non_places + tous_reste

def calculer_stats(terrain):
    """Calcule les statistiques finales"""
    producteurs = [b for b in terrain.batiments_places if b.type == 'Producteur']
    
    culture_totale = sum(b.culture_recue for b in producteurs)
    
    # Compter par type de production
    culture_guerison = sum(b.culture_recue for b in producteurs if b.production == 'Guerison')
    culture_nourriture = sum(b.culture_recue for b in producteurs if b.production == 'Nourriture')
    culture_or = sum(b.culture_recue for b in producteurs if b.production == 'Or')
    
    # Compter les boosts atteints
    boosts = {25: 0, 50: 0, 100: 0}
    boost_details = []
    
    for prod in producteurs:
        niveau = terrain.get_boost_niveau(prod.culture_recue, prod)
        if niveau >= 25:
            boosts[25] += 1
        if niveau >= 50:
            boosts[50] += 1
        if niveau >= 100:
            boosts[100] += 1
        
        boost_details.append({
            'nom': prod.nom,
            'culture_recue': prod.culture_recue,
            'boost': niveau
        })
    
    return {
        'culture_totale': culture_totale,
        'culture_guerison': culture_guerison,
        'culture_nourriture': culture_nourriture,
        'culture_or': culture_or,
        'boosts': boosts,
        'boost_details': boost_details
    }

def creer_excel_resultat(terrain, tous_batiments, stats, journal, batiments_non_places):
    """Crée un fichier Excel avec les résultats"""
    wb = Workbook()
    
    # Feuille Terrain Final
    ws_terrain = wb.active
    ws_terrain.title = "Terrain Final"
    
    # Créer une grille avec les noms des bâtiments
    grille_texte = []
    for i in range(terrain.hauteur):
        ligne = []
        for j in range(terrain.largeur):
            if terrain.grille[i, j] == 0:
                ligne.append("0")
            elif terrain.grille[i, j] == 1:
                ligne.append("1")
            else:
                # Trouver le bâtiment à cette position
                batiment_trouve = None
                for b in terrain.batiments_places:
                    if b.position:
                        x, y = b.position
                        if b.orientation == 'H':
                            if (x <= i < x + b.longueur and 
                                y <= j < y + b.largeur):
                                batiment_trouve = b
                                break
                        else:
                            if (x <= i < x + b.largeur and 
                                y <= j < y + b.longueur):
                                batiment_trouve = b
                                break
                
                if batiment_trouve:
                    ligne.append(f"{batiment_trouve.nom}")
                else:
                    ligne.append("B")
        grille_texte.append(ligne)
    
    df_terrain = pd.DataFrame(grille_texte)
    for r in dataframe_to_rows(df_terrain, index=False, header=False):
        ws_terrain.append(r)
    
    # Appliquer les couleurs
    orange_fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
    green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
    gray_fill = PatternFill(start_color='808080', end_color='808080', fill_type='solid')
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    white_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    
    for i in range(1, terrain.hauteur + 1):
        for j in range(1, terrain.largeur + 1):
            cell = ws_terrain.cell(row=i, column=j)
            if terrain.grille[i-1, j-1] == 0:
                cell.fill = red_fill
            elif terrain.grille[i-1, j-1] == 1:
                cell.fill = white_fill
            else:
                # Vérifier si c'est un bâtiment
                for b in terrain.batiments_places:
                    if b.position:
                        x, y = b.position
                        if b.orientation == 'H':
                            if (x <= i-1 < x + b.longueur and 
                                y <= j-1 < y + b.largeur):
                                if b.type == 'Culturel':
                                    cell.fill = orange_fill
                                elif b.type == 'Producteur':
                                    cell.fill = green_fill
                                else:
                                    cell.fill = gray_fill
                                break
                        else:
                            if (x <= i-1 < x + b.largeur and 
                                y <= j-1 < y + b.longueur):
                                if b.type == 'Culturel':
                                    cell.fill = orange_fill
                                elif b.type == 'Producteur':
                                    cell.fill = green_fill
                                else:
                                    cell.fill = gray_fill
                                break
    
    # Ajuster la largeur des colonnes
    for col in ws_terrain.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 20)
        ws_terrain.column_dimensions[column].width = adjusted_width
    
    # Feuille Statistiques
    ws_stats = wb.create_sheet("Statistiques")
    stats_data = [
        ["Statistique", "Valeur"],
        ["Culture totale reçue", f"{stats['culture_totale']:.2f}"],
        ["Culture Guérison", f"{stats['culture_guerison']:.2f}"],
        ["Culture Nourriture", f"{stats['culture_nourriture']:.2f}"],
        ["Culture Or", f"{stats['culture_or']:.2f}"],
        ["Bâtiments avec boost ≥25%", stats['boosts'][25]],
        ["Bâtiments avec boost ≥50%", stats['boosts'][50]],
        ["Bâtiments avec boost 100%", stats['boosts'][100]],
        ["Cases non utilisées", np.sum(terrain.grille == 1)],
        ["Surface totale des bâtiments non placés", 
         sum(b.surface() for b in batiments_non_places)],
        ["Bâtiments placés", len(terrain.batiments_places)],
        ["Bâtiments non placés", len(batiments_non_places)],
    ]
    
    for row in stats_data:
        ws_stats.append(row)
    
    # Détails des boosts par bâtiment
    ws_stats.append([])
    ws_stats.append(["Détails des boosts par producteur"])
    ws_stats.append(["Nom", "Culture reçue", "Niveau boost"])
    for detail in stats['boost_details']:
        ws_stats.append([detail['nom'], f"{detail['culture_recue']:.2f}", f"{detail['boost']}%"])
    
    # Feuille Journal
    ws_journal = wb.create_sheet("Journal")
    ws_journal.append(["#", "Message"])
    for i, entry in enumerate(journal, 1):
        ws_journal.append([i, entry])
    
    # Ajuster la largeur de la colonne journal
    ws_journal.column_dimensions['B'].width = 100
    
    # Feuille Bâtiments non placés
    ws_non_places = wb.create_sheet("Bâtiments non placés")
    ws_non_places.append(["Nom", "Type", "Longueur", "Largeur", "Surface", "Quantité totale"])
    
    # Compter par type
    non_places_count = {}
    for b in batiments_non_places:
        key = (b.nom, b.type, b.longueur, b.largeur)
        if key not in non_places_count:
            non_places_count[key] = 0
        non_places_count[key] += 1
    
    for (nom, type_bat, long, larg), count in non_places_count.items():
        ws_non_places.append([nom, type_bat, long, larg, long*larg, count])
    
    # Feuille Récapitulatif placement
    ws_recap = wb.create_sheet("Récapitulatif placement")
    ws_recap.append(["Nom", "Type", "Position X", "Position Y", "Orientation", "Culture reçue", "Boost"])
    
    for b in terrain.batiments_places:
        if b.position:
            x, y = b.position
            boost = terrain.get_boost_niveau(b.culture_recue, b) if b.type == 'Producteur' else 'N/A'
            ws_recap.append([
                b.nom, b.type, x, y, b.orientation,
                f"{b.culture_recue:.2f}" if b.type == 'Producteur' else '-',
                f"{boost}%" if boost != 'N/A' else '-'
            ])
    
    return wb

# Interface Streamlit
st.sidebar.header("📂 Chargement des données")

uploaded_file = st.sidebar.file_uploader(
    "Choisissez votre fichier Excel",
    type=['xlsx', 'xls'],
    help="Le fichier doit contenir un onglet 'Terrain' et un onglet 'Batiments'"
)

if uploaded_file:
    try:
        # Lecture du fichier Excel
        df_terrain = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        df_batiments = pd.read_excel(uploaded_file, sheet_name=1)
        
        # Nettoyer les noms de colonnes
        df_batiments.columns = ['Nom', 'Longueur', 'Largeur', 'Quantite', 'Type', 
                                'Culture', 'Rayonnement', 'Boost_25', 'Boost_50', 
                                'Boost_100', 'Production']
        
        st.success("✅ Fichier chargé avec succès!")
        
        # Afficher un aperçu
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Aperçu du terrain")
            st.dataframe(df_terrain.head(10))
            st.write(f"Dimensions: {df_terrain.shape[0]} lignes × {df_terrain.shape[1]} colonnes")
            st.write(f"Cases libres (1): {np.sum(df_terrain.values == 1)}")
            st.write(f"Cases occupées (0): {np.sum(df_terrain.values == 0)}")
        
        with col2:
            st.subheader("Aperçu des bâtiments")
            st.dataframe(df_batiments.head(10))
            
            # Statistiques sur les bâtiments
            total_batiments = df_batiments['Quantite'].sum()
            st.write(f"Total bâtiments à placer: {total_batiments}")
            
            type_counts = df_batiments.groupby('Type')['Quantite'].sum()
            for type_bat, count in type_counts.items():
                st.write(f"- {type_bat}: {count}")
        
        # Bouton pour lancer l'optimisation
        if st.button("🚀 Lancer l'optimisation", type="primary"):
            with st.spinner("Optimisation en cours... (cela peut prendre quelques instants)"):
                try:
                    # Réinitialiser le journal
                    st.session_state.journal = []
                    st.session_state.log_count = 0
                    
                    # Créer le terrain
                    terrain = Terrain(df_terrain.values)
                    
                    # Créer les bâtiments
                    batiments = []
                    for _, row in df_batiments.iterrows():
                        batiment = Batiment(
                            row['Nom'], row['Longueur'], row['Largeur'], row['Quantite'],
                            row['Type'], row['Culture'], row['Rayonnement'], 
                            row['Boost_25'], row['Boost_50'], row['Boost_100'], 
                            row['Production'] if pd.notna(row['Production']) else ""
                        )
                        batiments.append(batiment)
                    
                    ajouter_log("=== DÉBUT DE L'OPTIMISATION ===")
                    
                    # Lancer l'algorithme
                    terrain_optimise, non_places = placer_batiments(deepcopy(terrain), batiments)
                    
                    # Calculer les statistiques
                    stats = calculer_stats(terrain_optimise)
                    
                    ajouter_log(f"=== OPTIMISATION TERMINÉE ===")
                    ajouter_log(f"Bâtiments placés: {len(terrain_optimise.batiments_places)}")
                    ajouter_log(f"Bâtiments non placés: {len(non_places)}")
                    ajouter_log(f"Cultures totales: {stats['culture_totale']:.2f}")
                    
                    # Sauvegarder la solution
                    st.session_state.solution = {
                        'terrain': terrain_optimise,
                        'batiments': batiments,
                        'stats': stats,
                        'journal': st.session_state.journal,
                        'batiments_non_places': non_places
                    }
                    
                    st.success("✅ Optimisation terminée!")
                    
                except Exception as e:
                    st.error(f"Erreur pendant l'optimisation: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
        
        # Afficher les résultats si disponibles
        if st.session_state.solution:
            solution = st.session_state.solution
            
            st.header("📊 Résultats de l'optimisation")
            
            # Métriques principales
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Culture totale", f"{solution['stats']['culture_totale']:.1f}")
            with col2:
                st.metric("Boost ≥25%", solution['stats']['boosts'][25])
            with col3:
                st.metric("Boost ≥50%", solution['stats']['boosts'][50])
            with col4:
                st.metric("Boost 100%", solution['stats']['boosts'][100])
            
            # Deuxième ligne de métriques
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Bâtiments placés", len(solution['terrain'].batiments_places))
            with col2:
                st.metric("Bâtiments non placés", len(solution['batiments_non_places']))
            with col3:
                st.metric("Cases libres restantes", np.sum(solution['terrain'].grille == 1))
            
            # Visualisation du terrain
            st.subheader("🗺️ Visualisation du terrain")
            
            # Créer une matrice de couleurs pour l'affichage
            terrain_mat = solution['terrain'].grille.copy()
            legend_text = []
            
            # Colorier les bâtiments
            for b in solution['terrain'].batiments_places:
                if b.position:
                    x, y = b.position
                    if b.orientation == 'H':
                        for i in range(b.longueur):
                            for j in range(b.largeur):
                                if b.type == 'Culturel':
                                    terrain_mat[x + i, y + j] = 3
                                elif b.type == 'Producteur':
                                    terrain_mat[x + i, y + j] = 4
                                else:
                                    terrain_mat[x + i, y + j] = 5
                    else:
                        for i in range(b.largeur):
                            for j in range(b.longueur):
                                if b.type == 'Culturel':
                                    terrain_mat[x + i, y + j] = 3
                                elif b.type == 'Producteur':
                                    terrain_mat[x + i, y + j] = 4
                                else:
                                    terrain_mat[x + i, y + j] = 5
            
            # Afficher avec st.dataframe et coloration conditionnelle
            df_visuel = pd.DataFrame(terrain_mat)
            
            def colorier_cases(val):
                if val == 0:
                    return 'background-color: #ffcccc; color: black'  # Rouge clair (occupé)
                elif val == 1:
                    return 'background-color: #ffffff; color: black'  # Blanc (libre)
                elif val == 3:
                    return 'background-color: #ffb347; color: black'  # Orange (culturel)
                elif val == 4:
                    return 'background-color: #90ee90; color: black'  # Vert clair (producteur)
                elif val == 5:
                    return 'background-color: #d3d3d3; color: black'  # Gris (neutre)
                return ''
            
            styled_df = df_visuel.style.map(colorier_cases)
            st.dataframe(styled_df, height=500, use_container_width=True)
            
            # Légende
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.markdown("🟥 **Occupé**")
            with col2:
                st.markdown("⬜ **Libre**")
            with col3:
                st.markdown("🟧 **Culturel**")
            with col4:
                st.markdown("🟩 **Producteur**")
            with col5:
                st.markdown("⬜ **Neutre**")
            
            # Détails des boosts
            with st.expander("📈 Détails des boosts par producteur"):
                boost_df = pd.DataFrame(solution['stats']['boost_details'])
                if not boost_df.empty:
                    st.dataframe(boost_df)
            
            # Journal
            with st.expander("📋 Voir le journal d'exécution"):
                for i, entry in enumerate(solution['journal']):
                    st.text(f"{i+1:3d}: {entry}")
            
            # Bâtiments non placés
            if solution['batiments_non_places']:
                st.subheader("⚠️ Bâtiments non placés")
                
                # Compter par type
                non_places_count = {}
                for b in solution['batiments_non_places']:
                    key = (b.nom, b.type, b.longueur, b.largeur)
                    if key not in non_places_count:
                        non_places_count[key] = 0
                    non_places_count[key] += 1
                
                non_places_data = []
                for (nom, type_bat, long, larg), count in non_places_count.items():
                    non_places_data.append({
                        'Nom': nom,
                        'Type': type_bat,
                        'Dimensions': f"{long}x{larg}",
                        'Surface': long*larg,
                        'Quantité': count
                    })
                
                st.dataframe(pd.DataFrame(non_places_data))
                st.write(f"Surface totale non placée: {sum(b.surface() for b in solution['batiments_non_places'])} cases")
            
            # Bouton de téléchargement
            st.subheader("💾 Télécharger les résultats")
            
            # Créer le fichier Excel
            wb = creer_excel_resultat(
                solution['terrain'],
                solution['batiments'],
                solution['stats'],
                solution['journal'],
                solution['batiments_non_places']
            )
            
            # Sauvegarder dans un buffer
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            
            st.download_button(
                label="📥 Télécharger le rapport Excel complet",
                data=excel_buffer,
                file_name="resultats_placement.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
    except Exception as e:
        st.error(f"Erreur de lecture du fichier: {str(e)}")
        st.info("Assurez-vous que votre fichier Excel a le format correct.")
        import traceback
        st.code(traceback.format_exc())

else:
    st.info("👈 Veuillez charger un fichier Excel pour commencer")
    
    # Afficher un exemple de format attendu
    with st.expander("📝 Format du fichier attendu"):
        st.markdown("""
        ### Onglet 1: Terrain
        - Matrice de 0 et 1
        - 1 = case libre
        - 0 = case occupée
        - Les lignes représentent les rangées du terrain
        
        ### Onglet 2: Bâtiments (avec en-têtes)
        | Nom | Longueur | Largeur | Quantite | Type | Culture | Rayonnement | Boost 25% | Boost 50% | Boost 100% | Production |
        |-----|----------|---------|----------|------|---------|-------------|-----------|-----------|------------|------------|
        | Maison | 2 | 2 | 3 | Neutre | 0 | 0 | 0 | 0 | 0 | |
        | Temple | 3 | 2 | 1 | Culturel | 5 | 2 | 0 | 0 | 0 | |
        | Ferme | 4 | 3 | 2 | Producteur | 0 | 0 | 10 | 20 | 30 | Nourriture |
        
        **Types possibles**: Culturel, Producteur, Neutre
        **Production**: Guerison, Nourriture, Or, ou vide
        """)

# Pied de page
st.markdown("---")
st.markdown("🎯 Optimiseur de placement v2.0 - Compatible iPad")
