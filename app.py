import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import io
from copy import deepcopy
import math

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
    
    def peut_placer(self, batiment, x, y, orientation):
        """Vérifie si on peut placer un bâtiment à la position (x,y) avec l'orientation donnée"""
        if orientation == 'H':  # Horizontal
            if y + batiment.largeur > self.largeur:
                return False
            for i in range(batiment.longueur):
                for j in range(batiment.largeur):
                    if not self.case_libre(x + i, y + j):
                        return False
        else:  # Vertical
            if x + batiment.largeur > self.hauteur:
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
    
    def calculer_culture_recue(self, batiments_culturels):
        """Calcule la culture reçue par tous les bâtiments producteurs"""
        producteurs = [b for b in self.batiments_places if b.type == 'Producteur']
        
        for prod in producteurs:
            prod.culture_recue = 0
            
        for cult in batiments_culturels:
            if cult.placed:
                x, y = cult.position
                rayon = cult.rayonnement
                
                # Définir la zone de rayonnement
                x_min = max(0, x - rayon)
                x_max = min(self.hauteur, x + cult.longueur + rayon)
                y_min = max(0, y - rayon)
                y_max = min(self.largeur, y + cult.largeur + rayon)
                
                for prod in producteurs:
                    if prod.placed:
                        px, py = prod.position
                        # Vérifier si le producteur est dans la zone
                        if (px >= x_min and px < x_max and 
                            py >= y_min and py < y_max):
                            prod.culture_recue += cult.culture
    
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

def ajouter_log(message):
    """Ajoute un message au journal"""
    st.session_state.journal.append(message)
    st.session_state.log_count += 1
    if st.session_state.log_count > 1000:
        raise Exception("Limite de 1000 entrées dans le journal atteinte")

def trouver_emplacement(terrain, batiment, positions_essayees):
    """Trouve un emplacement valide pour un bâtiment"""
    for x in range(terrain.hauteur):
        for y in range(terrain.largeur):
            for orientation in ['H', 'V']:
                if (x, y, orientation) not in positions_essayees:
                    if terrain.peut_placer(batiment, x, y, orientation):
                        return (x, y, orientation)
    return None

def assez_de_place_restante(terrain, batiments_restants):
    """Vérifie s'il reste assez de place pour les plus grands bâtiments"""
    if not batiments_restants:
        return True
    
    # Calculer la surface totale disponible
    surface_disponible = np.sum(terrain.grille == 1)
    
    # Trier les bâtiments restants par surface décroissante
    batiments_tries = sorted(batiments_restants, key=lambda b: b.surface(), reverse=True)
    
    # Vérifier si on peut placer le plus grand
    plus_grand = batiments_tries[0]
    
    # Chercher un emplacement pour le plus grand bâtiment
    for x in range(terrain.hauteur):
        for y in range(terrain.largeur):
            for orientation in ['H', 'V']:
                if terrain.peut_placer(plus_grand, x, y, orientation):
                    return True
    
    return False

def placer_batiments(terrain, batiments):
    """Algorithme principal de placement"""
    
    # Séparer les bâtiments par type
    neutres = [b for b in batiments if b.type == 'Neutre']
    culturels = [b for b in batiments if b.type == 'Culturel']
    producteurs = [b for b in batiments if b.type == 'Producteur']
    
    # Trier par surface décroissante
    neutres.sort(key=lambda b: b.surface(), reverse=True)
    culturels.sort(key=lambda b: b.surface(), reverse=True)
    producteurs.sort(key=lambda b: b.surface(), reverse=True)
    
    # Liste de tous les bâtiments à placer (avec répétitions selon quantite)
    tous_batiments = []
    for b in neutres:
        for _ in range(b.quantite):
            tous_batiments.append(deepcopy(b))
    for b in culturels:
        for _ in range(b.quantite):
            tous_batiments.append(deepcopy(b))
    for b in producteurs:
        for _ in range(b.quantite):
            tous_batiments.append(deepcopy(b))
    
    # Trier par ordre de placement
    # D'abord les neutres (pour les bords), puis alternance des autres
    neutres_a_placer = [b for b in tous_batiments if b.type == 'Neutre']
    autres = [b for b in tous_batiments if b.type != 'Neutre']
    
    ordre_placement = neutres_a_placer.copy()
    
    # Ajouter en alternance culturels et producteurs
    i, j = 0, 0
    while i < len(culturels) or j < len(producteurs):
        if i < len(culturels):
            ordre_placement.append(culturels[i])
            i += 1
        if j < len(producteurs):
            ordre_placement.append(producteurs[j])
            j += 1
    
    pile_placement = []
    positions_essayees = {}
    
    for batiment in ordre_placement:
        if batiment.nom not in positions_essayees:
            positions_essayees[batiment.nom] = set()
        
        ajouter_log(f"Évaluation du bâtiment: {batiment.nom}")
        
        place = False
        while not place:
            # Chercher un emplacement
            emplacement = trouver_emplacement(terrain, batiment, positions_essayees[batiment.nom])
            
            if emplacement:
                x, y, orientation = emplacement
                positions_essayees[batiment.nom].add((x, y, orientation))
                
                # Vérifier s'il reste assez de place pour les autres
                terrain.placer_batiment(batiment, x, y, orientation)
                
                batiments_restants = [b for b in ordre_placement if not b.placed and b != batiment]
                
                if assez_de_place_restante(terrain, batiments_restants):
                    ajouter_log(f"Bâtiment placé: {batiment.nom} à ({x},{y}) orientation {orientation}")
                    pile_placement.append(batiment)
                    place = True
                else:
                    terrain.enlever_batiment(batiment)
                    ajouter_log(f"Place insuffisante après placement, retrait de {batiment.nom}")
            else:
                # Backtracking
                if pile_placement:
                    dernier = pile_placement.pop()
                    terrain.enlever_batiment(dernier)
                    ajouter_log(f"Retrait du bâtiment: {dernier.nom} pour réessayer")
                else:
                    ajouter_log(f"Impossible de placer {batiment.nom}")
                    break
    
    return terrain

def calculer_stats(terrain, batiments_culturels):
    """Calcule les statistiques finales"""
    terrain.calculer_culture_recue(batiments_culturels)
    
    producteurs = [b for b in terrain.batiments_places if b.type == 'Producteur']
    
    culture_totale = sum(b.culture_recue for b in producteurs)
    
    # Compter par type de production
    culture_guerison = sum(b.culture_recue for b in producteurs if b.production == 'Guerison')
    culture_nourriture = sum(b.culture_recue for b in producteurs if b.production == 'Nourriture')
    culture_or = sum(b.culture_recue for b in producteurs if b.production == 'Or')
    
    # Compter les boosts atteints
    boosts = {25: 0, 50: 0, 100: 0}
    for prod in producteurs:
        niveau = terrain.get_boost_niveau(prod.culture_recue, prod)
        if niveau >= 25:
            boosts[25] += 1
        if niveau >= 50:
            boosts[50] += 1
        if niveau >= 100:
            boosts[100] += 1
    
    return {
        'culture_totale': culture_totale,
        'culture_guerison': culture_guerison,
        'culture_nourriture': culture_nourriture,
        'culture_or': culture_or,
        'boosts': boosts
    }

def creer_excel_resultat(terrain, batiments, stats, journal, batiments_non_places):
    """Crée un fichier Excel avec les résultats"""
    wb = Workbook()
    
    # Feuille Terrain Final
    ws_terrain = wb.active
    ws_terrain.title = "Terrain Final"
    
    # Créer une grille colorée
    grille_couleurs = []
    for i in range(terrain.hauteur):
        ligne = []
        for j in range(terrain.largeur):
            case = terrain.grille[i, j]
            if case == 0:
                ligne.append("0 (Occupé)")
            elif case == 1:
                ligne.append("1 (Libre)")
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
                    ligne.append(f"{batiment_trouve.nom} ({batiment_trouve.type})")
                else:
                    ligne.append("Bâtiment")
        grille_couleurs.append(ligne)
    
    df_terrain = pd.DataFrame(grille_couleurs)
    for r in dataframe_to_rows(df_terrain, index=False, header=False):
        ws_terrain.append(r)
    
    # Appliquer les couleurs
    orange_fill = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')
    green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
    gray_fill = PatternFill(start_color='808080', end_color='808080', fill_type='solid')
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    
    for i in range(1, terrain.hauteur + 1):
        for j in range(1, terrain.largeur + 1):
            cell = ws_terrain.cell(row=i, column=j)
            if terrain.grille[i-1, j-1] == 0:
                cell.fill = red_fill
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
    
    # Feuille Statistiques
    ws_stats = wb.create_sheet("Statistiques")
    stats_data = [
        ["Statistique", "Valeur"],
        ["Culture totale reçue", stats['culture_totale']],
        ["Culture Guérison", stats['culture_guerison']],
        ["Culture Nourriture", stats['culture_nourriture']],
        ["Culture Or", stats['culture_or']],
        ["Nombre de bâtiments avec boost 25%", stats['boosts'][25]],
        ["Nombre de bâtiments avec boost 50%", stats['boosts'][50]],
        ["Nombre de bâtiments avec boost 100%", stats['boosts'][100]],
        ["Cases non utilisées", np.sum(terrain.grille == 1)],
        ["Surface totale des bâtiments non placés", 
         sum(b.surface() for b in batiments_non_places)],
    ]
    
    for row in stats_data:
        ws_stats.append(row)
    
    # Feuille Journal
    ws_journal = wb.create_sheet("Journal")
    ws_journal.append(["Entrées du journal"])
    for entry in journal:
        ws_journal.append([entry])
    
    # Feuille Bâtiments non placés
    ws_non_places = wb.create_sheet("Bâtiments non placés")
    ws_non_places.append(["Nom", "Type", "Longueur", "Largeur", "Surface"])
    for b in batiments_non_places:
        ws_non_places.append([b.nom, b.type, b.longueur, b.largeur, b.surface()])
    
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
        
        st.success("✅ Fichier chargé avec succès!")
        
        # Afficher un aperçu
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Aperçu du terrain")
            st.dataframe(df_terrain.head(10))
            st.write(f"Dimensions: {df_terrain.shape[0]}x{df_terrain.shape[1]}")
        
        with col2:
            st.subheader("Aperçu des bâtiments")
            st.dataframe(df_batiments.head(10))
            st.write(f"Nombre de types: {len(df_batiments)}")
        
        # Bouton pour lancer l'optimisation
        if st.button("🚀 Lancer l'optimisation", type="primary"):
            with st.spinner("Optimisation en cours..."):
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
                            row.iloc[0], row.iloc[1], row.iloc[2], row.iloc[3],
                            row.iloc[4], row.iloc[5], row.iloc[6], row.iloc[7],
                            row.iloc[8], row.iloc[9], row.iloc[10] if len(row) > 10 else ""
                        )
                        batiments.append(batiment)
                    
                    ajouter_log("Début de l'optimisation")
                    
                    # Lancer l'algorithme
                    terrain_optimise = placer_batiments(deepcopy(terrain), batiments)
                    
                    # Calculer les statistiques
                    batiments_culturels = [b for b in batiments if b.type == 'Culturel']
                    stats = calculer_stats(terrain_optimise, batiments_culturels)
                    
                    # Identifier les bâtiments non placés
                    tous_noms = []
                    for b in batiments:
                        for _ in range(b.quantite):
                            tous_noms.append(b.nom)
                    
                    noms_places = [b.nom for b in terrain_optimise.batiments_places]
                    batiments_non_places = []
                    for nom in tous_noms:
                        if nom not in noms_places:
                            # Trouver le bâtiment correspondant
                            for b in batiments:
                                if b.nom == nom:
                                    batiments_non_places.append(b)
                                    break
                    
                    # Sauvegarder la solution
                    st.session_state.solution = {
                        'terrain': terrain_optimise,
                        'batiments': batiments,
                        'stats': stats,
                        'journal': st.session_state.journal,
                        'batiments_non_places': batiments_non_places
                    }
                    
                    st.success("✅ Optimisation terminée!")
                    
                except Exception as e:
                    st.error(f"Erreur pendant l'optimisation: {str(e)}")
        
        # Afficher les résultats si disponibles
        if st.session_state.solution:
            solution = st.session_state.solution
            
            st.header("📊 Résultats de l'optimisation")
            
            # Statistiques
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Culture totale", f"{solution['stats']['culture_totale']:.1f}")
            with col2:
                st.metric("Boost 25%", solution['stats']['boosts'][25])
            with col3:
                st.metric("Boost 50%", solution['stats']['boosts'][50])
            with col4:
                st.metric("Boost 100%", solution['stats']['boosts'][100])
            
            # Visualisation simplifiée du terrain
            st.subheader("🗺️ Visualisation du terrain")
            
            # Créer une matrice de couleurs pour l'affichage
            terrain_mat = solution['terrain'].grille.copy()
            
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
                    return 'background-color: #ffcccc'  # Rouge clair (occupé)
                elif val == 1:
                    return 'background-color: #ffffff'  # Blanc (libre)
                elif val == 3:
                    return 'background-color: #ffb347'  # Orange (culturel)
                elif val == 4:
                    return 'background-color: #90ee90'  # Vert clair (producteur)
                elif val == 5:
                    return 'background-color: #d3d3d3'  # Gris (neutre)
                return ''
            
            styled_df = df_visuel.style.map(colorier_cases)
            st.dataframe(styled_df, height=400)
            
            # Journal
            with st.expander("📋 Voir le journal d'exécution"):
                for entry in solution['journal'][-50:]:  # Afficher les 50 dernières entrées
                    st.text(entry)
            
            # Bâtiments non placés
            if solution['batiments_non_places']:
                st.subheader("⚠️ Bâtiments non placés")
                non_places_data = []
                for b in solution['batiments_non_places']:
                    non_places_data.append({
                        'Nom': b.nom,
                        'Type': b.type,
                        'Surface': b.surface()
                    })
                st.dataframe(pd.DataFrame(non_places_data))
            
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
                label="📥 Télécharger le rapport Excel",
                data=excel_buffer,
                file_name="resultats_placement.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    except Exception as e:
        st.error(f"Erreur de lecture du fichier: {str(e)}")
        st.info("Assurez-vous que votre fichier Excel a le format correct.")

else:
    st.info("👈 Veuillez charger un fichier Excel pour commencer")
    
    # Afficher un exemple de format attendu
    with st.expander("📝 Format du fichier attendu"):
        st.markdown("""
        ### Onglet 1: Terrain
        - Matrice de 0 et 1
        - 1 = case libre
        - 0 = case occupée
        
        ### Onglet 2: Bâtiments
        Colonnes dans l'ordre:
        1. Nom
        2. Longueur
        3. Largeur
        4. Quantité
        5. Type (Culturel/Producteur/Neutre)
        6. Culture
        7. Rayonnement
        8. Boost 25%
        9. Boost 50%
        10. Boost 100%
        11. Production (optionnel)
        """)

# Pied de page
st.markdown("---")
st.markdown("🎯 Optimiseur de placement - Compatible iPad")