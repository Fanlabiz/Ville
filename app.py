import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
import io
from copy import deepcopy

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
if 'optimisation_terminee' not in st.session_state:
    st.session_state.optimisation_terminee = False
if 'limite_atteinte' not in st.session_state:
    st.session_state.limite_atteinte = False

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
        self.id = id(self)
        
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
    
    def peut_placer(self, batiment, x, y, orientation):
        """Vérifie si on peut placer un bâtiment"""
        if orientation == 'H':
            if x + batiment.longueur > self.hauteur or y + batiment.largeur > self.largeur:
                return False
            for i in range(batiment.longueur):
                for j in range(batiment.largeur):
                    if not self.case_libre(x + i, y + j):
                        return False
        else:
            if x + batiment.largeur > self.hauteur or y + batiment.longueur > self.largeur:
                return False
            for i in range(batiment.largeur):
                for j in range(batiment.longueur):
                    if not self.case_libre(x + i, y + j):
                        return False
        return True
    
    def placer_batiment(self, batiment, x, y, orientation):
        """Place un bâtiment"""
        if orientation == 'H':
            for i in range(batiment.longueur):
                for j in range(batiment.largeur):
                    self.grille[x + i, y + j] = 2
        else:
            for i in range(batiment.largeur):
                for j in range(batiment.longueur):
                    self.grille[x + i, y + j] = 2
        
        batiment.placed = True
        batiment.position = (x, y)
        batiment.orientation = orientation
        self.batiments_places.append(batiment)
        
    def enlever_batiment(self, batiment):
        """Enlève un bâtiment"""
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
            
            if batiment in self.batiments_places:
                self.batiments_places.remove(batiment)
            batiment.placed = False
            batiment.position = None
            batiment.orientation = None
    
    def calculer_culture_recue(self):
        """Calcule la culture reçue"""
        producteurs = [b for b in self.batiments_places if b.type == 'Producteur']
        culturels = [b for b in self.batiments_places if b.type == 'Culturel']
        
        for prod in producteurs:
            prod.culture_recue = 0
        
        for cult in culturels:
            if cult.placed:
                x, y = cult.position
                rayon = cult.rayonnement
                
                x_min = max(0, x - rayon)
                x_max = min(self.hauteur, x + cult.longueur + rayon)
                y_min = max(0, y - rayon)
                y_max = min(self.largeur, y + cult.largeur + rayon)
                
                for prod in producteurs:
                    if prod.placed:
                        px, py = prod.position
                        # Vérifier si le producteur est dans la zone
                        if prod.orientation == 'H':
                            for i in range(prod.longueur):
                                for j in range(prod.largeur):
                                    if (x_min <= px + i < x_max and 
                                        y_min <= py + j < y_max):
                                        prod.culture_recue += cult.culture
                                        break
                                else:
                                    continue
                                break
                        else:
                            for i in range(prod.largeur):
                                for j in range(prod.longueur):
                                    if (x_min <= px + i < x_max and 
                                        y_min <= py + j < y_max):
                                        prod.culture_recue += cult.culture
                                        break
                                else:
                                    continue
                                break
    
    def get_boost_niveau(self, culture_recue, batiment):
        """Détermine le niveau de boost"""
        if culture_recue >= batiment.boost_100:
            return 100
        elif culture_recue >= batiment.boost_50:
            return 50
        elif culture_recue >= batiment.boost_25:
            return 25
        return 0

def ajouter_log(message):
    """Ajoute un message au journal et vérifie la limite"""
    st.session_state.journal.append(message)
    st.session_state.log_count += 1
    
    # Si on atteint exactement 1000, on lève l'exception
    if st.session_state.log_count >= 1000:
        st.session_state.limite_atteinte = True
        raise Exception("LIMITE_1000_ATTEINTE")

def placer_batiments_avec_backtracking(terrain, batiments):
    """Algorithme avec backtracking qui s'arrête à 1000 entrées"""
    
    # Créer toutes les instances
    tous_batiments = []
    for b in batiments:
        for _ in range(b.quantite):
            nouveau = deepcopy(b)
            nouveau.id = id(nouveau)
            tous_batiments.append(nouveau)
    
    # Séparer par type
    neutres = [b for b in tous_batiments if b.type == 'Neutre']
    culturels = [b for b in tous_batiments if b.type == 'Culturel']
    producteurs = [b for b in tous_batiments if b.type == 'Producteur']
    
    # Trier par surface (du plus grand au plus petit)
    neutres.sort(key=lambda b: b.surface(), reverse=True)
    culturels.sort(key=lambda b: b.surface(), reverse=True)
    producteurs.sort(key=lambda b: b.surface(), reverse=True)
    
    ajouter_log(f"Début placement - Terrain: {terrain.hauteur}x{terrain.largeur}")
    ajouter_log(f"À placer: {len(neutres)} neutres, {len(culturels)} culturels, {len(producteurs)} producteurs")
    
    # Fonction pour trouver tous les emplacements possibles
    def trouver_tous_emplacements(batiment):
        emplacements = []
        for x in range(terrain.hauteur):
            for y in range(terrain.largeur):
                for orientation in ['H', 'V']:
                    if terrain.peut_placer(batiment, x, y, orientation):
                        # Calculer un score pour cet emplacement
                        score = 0
                        if terrain.est_sur_bord(x, y, batiment.longueur, batiment.largeur, orientation):
                            score += 100  # Priorité aux bords
                        emplacements.append((score, x, y, orientation))
        
        # Trier par score décroissant
        emplacements.sort(reverse=True)
        return [(x, y, orientation) for _, x, y, orientation in emplacements]
    
    # Phase 1: Placer les neutres sur les bords
    neutres_places = []
    neutres_non_places = []
    
    for batiment in neutres:
        try:
            ajouter_log(f"Recherche neutre: {batiment.nom}")
            emplacements = trouver_tous_emplacements(batiment)
            
            place = False
            for x, y, orientation in emplacements:
                if terrain.peut_placer(batiment, x, y, orientation):
                    terrain.placer_batiment(batiment, x, y, orientation)
                    neutres_places.append(batiment)
                    ajouter_log(f"  ✅ Neutre placé: {batiment.nom} à ({x},{y}) orientation {orientation}")
                    place = True
                    break
            
            if not place:
                neutres_non_places.append(batiment)
                ajouter_log(f"  ❌ Impossible de placer neutre: {batiment.nom}")
        except Exception as e:
            if str(e) == "LIMITE_1000_ATTEINTE":
                # On remonte l'exception mais avec les résultats partiels
                raise
            else:
                raise
    
    # Phase 2: Placer en alternance culturels et producteurs
    tous_reste = culturels + producteurs
    tous_reste.sort(key=lambda b: b.surface(), reverse=True)
    
    # Backtracking simple
    pile_placement = []
    index = 0
    max_essais = 500
    essai = 0
    
    while index < len(tous_reste) and essai < max_essais:
        try:
            essai += 1
            batiment = tous_reste[index]
            
            ajouter_log(f"Essai placement {batiment.nom} ({batiment.type})")
            
            emplacements = trouver_tous_emplacements(batiment)
            
            if emplacements:
                x, y, orientation = emplacements[0]
                terrain.placer_batiment(batiment, x, y, orientation)
                pile_placement.append(batiment)
                ajouter_log(f"  ✅ Placé: {batiment.nom} à ({x},{y})")
                index += 1
            else:
                ajouter_log(f"  ❌ Aucun emplacement, backtracking...")
                if pile_placement:
                    dernier = pile_placement.pop()
                    terrain.enlever_batiment(dernier)
                    ajouter_log(f"  ↩️ Retrait de {dernier.nom}")
                    tous_reste.insert(index, dernier)
                else:
                    ajouter_log(f"  ⚠️ Impossible de progresser")
                    break
        except Exception as e:
            if str(e) == "LIMITE_1000_ATTEINTE":
                raise
            else:
                raise
    
    # Phase 3: Essayer de placer les neutres restants
    for batiment in neutres_non_places[:]:
        try:
            ajouter_log(f"Tentative neutre restant: {batiment.nom}")
            emplacements = trouver_tous_emplacements(batiment)
            
            for x, y, orientation in emplacements:
                if terrain.peut_placer(batiment, x, y, orientation):
                    terrain.placer_batiment(batiment, x, y, orientation)
                    neutres_places.append(batiment)
                    neutres_non_places.remove(batiment)
                    ajouter_log(f"  ✅ Neutre placé à l'intérieur: {batiment.nom}")
                    break
        except Exception as e:
            if str(e) == "LIMITE_1000_ATTEINTE":
                raise
            else:
                raise
    
    # Calculer la culture finale
    terrain.calculer_culture_recue()
    
    # Bilan
    tous_non_places = neutres_non_places + [b for b in culturels if not b.placed] + [b for b in producteurs if not b.placed]
    
    ajouter_log(f"=== BILAN FINAL ===")
    ajouter_log(f"Neutres placés: {len(neutres_places)}/{len(neutres)}")
    ajouter_log(f"Culturels placés: {len([b for b in culturels if b.placed])}/{len(culturels)}")
    ajouter_log(f"Producteurs placés: {len([b for b in producteurs if b.placed])}/{len(producteurs)}")
    ajouter_log(f"Total placés: {len(terrain.batiments_places)}")
    ajouter_log(f"Total non placés: {len(tous_non_places)}")
    
    return terrain, tous_non_places

def calculer_stats(terrain):
    """Calcule les statistiques"""
    producteurs = [b for b in terrain.batiments_places if b.type == 'Producteur']
    
    culture_totale = sum(b.culture_recue for b in producteurs)
    
    culture_guerison = sum(b.culture_recue for b in producteurs if b.production == 'Guerison')
    culture_nourriture = sum(b.culture_recue for b in producteurs if b.production == 'Nourriture')
    culture_or = sum(b.culture_recue for b in producteurs if b.production == 'Or')
    
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
            'culture_recue': round(prod.culture_recue, 2),
            'boost': f"{niveau}%"
        })
    
    return {
        'culture_totale': round(culture_totale, 2),
        'culture_guerison': round(culture_guerison, 2),
        'culture_nourriture': round(culture_nourriture, 2),
        'culture_or': round(culture_or, 2),
        'boosts': boosts,
        'boost_details': boost_details
    }

def creer_excel_resultat(terrain, tous_batiments, stats, journal, batiments_non_places, limite_atteinte):
    """Crée le fichier Excel de résultats"""
    wb = Workbook()
    
    # Feuille Terrain
    ws_terrain = wb.active
    ws_terrain.title = "Terrain Final"
    
    # Créer la grille de texte
    grille_texte = []
    for i in range(terrain.hauteur):
        ligne = []
        for j in range(terrain.largeur):
            if terrain.grille[i, j] == 0:
                ligne.append("0")
            elif terrain.grille[i, j] == 1:
                ligne.append("1")
            else:
                batiment_trouve = None
                for b in terrain.batiments_places:
                    if b.position:
                        x, y = b.position
                        if b.orientation == 'H':
                            if x <= i < x + b.longueur and y <= j < y + b.largeur:
                                batiment_trouve = b
                                break
                        else:
                            if x <= i < x + b.largeur and y <= j < y + b.longueur:
                                batiment_trouve = b
                                break
                
                if batiment_trouve:
                    ligne.append(f"{batiment_trouve.nom}")
                else:
                    ligne.append("B")
        grille_texte.append(ligne)
    
    # Écrire dans Excel
    for r, ligne in enumerate(grille_texte, 1):
        for c, valeur in enumerate(ligne, 1):
            ws_terrain.cell(row=r, column=c, value=valeur)
    
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
    
    # Ajouter un avertissement si limite atteinte
    if limite_atteinte:
        ws_stats.append(["⚠️ ATTENTION", "Limite de 1000 entrées journal atteinte - Résultats PARTIELS"])
        ws_stats.append(["Date", str(pd.Timestamp.now())])
        ws_stats.append([])
    
    stats_data = [
        ["Statistique", "Valeur"],
        ["Culture totale reçue", stats['culture_totale']],
        ["Culture Guérison", stats['culture_guerison']],
        ["Culture Nourriture", stats['culture_nourriture']],
        ["Culture Or", stats['culture_or']],
        ["Bâtiments avec boost ≥25%", stats['boosts'][25]],
        ["Bâtiments avec boost ≥50%", stats['boosts'][50]],
        ["Bâtiments avec boost 100%", stats['boosts'][100]],
        ["Cases non utilisées", int(np.sum(terrain.grille == 1))],
        ["Bâtiments placés", len(terrain.batiments_places)],
        ["Bâtiments non placés", len(batiments_non_places)],
    ]
    
    for row in stats_data:
        ws_stats.append(row)
    
    # Détails des boosts
    if stats['boost_details']:
        ws_stats.append([])
        ws_stats.append(["Détails des boosts par producteur"])
        ws_stats.append(["Nom", "Culture reçue", "Boost"])
        for detail in stats['boost_details']:
            ws_stats.append([detail['nom'], detail['culture_recue'], detail['boost']])
    
    # Feuille Journal (COMPLET - jusqu'à 1000 entrées)
    ws_journal = wb.create_sheet("Journal")
    ws_journal.append(["#", "Message"])
    for i, entry in enumerate(journal, 1):
        ws_journal.append([i, entry])
    
    # Ajuster la largeur
    ws_journal.column_dimensions['B'].width = 100
    
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
        # Lecture du fichier
        df_terrain = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        df_batiments = pd.read_excel(uploaded_file, sheet_name=1)
        
        # Nettoyer les noms de colonnes
        df_batiments.columns = ['Nom', 'Longueur', 'Largeur', 'Quantite', 'Type', 
                                'Culture', 'Rayonnement', 'Boost_25', 'Boost_50', 
                                'Boost_100', 'Production']
        
        st.success("✅ Fichier chargé avec succès!")
        
        # Aperçu
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Aperçu du terrain")
            st.dataframe(df_terrain.head(10))
            st.write(f"Dimensions: {df_terrain.shape[0]}×{df_terrain.shape[1]}")
            st.write(f"Cases libres: {np.sum(df_terrain.values == 1)}")
        
        with col2:
            st.subheader("Aperçu des bâtiments")
            st.dataframe(df_batiments.head(10))
            total_batiments = df_batiments['Quantite'].sum()
            st.write(f"Total à placer: {total_batiments}")
        
        # Bouton de lancement
        if st.sidebar.button("🚀 Lancer l'optimisation", type="primary", use_container_width=True):
            with st.spinner("Optimisation en cours..."):
                # Réinitialiser
                st.session_state.journal = []
                st.session_state.log_count = 0
                st.session_state.limite_atteinte = False
                st.session_state.optimisation_terminee = False
                st.session_state.solution = None
                
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
                
                terrain_optimise = None
                non_places = []
                
                try:
                    # Lancer l'algorithme
                    terrain_optimise, non_places = placer_batiments_avec_backtracking(deepcopy(terrain), batiments)
                    
                    # Si on arrive ici, c'est que la limite n'a pas été atteinte
                    st.session_state.optimisation_terminee = True
                    
                except Exception as e:
                    if str(e) == "LIMITE_1000_ATTEINTE":
                        st.session_state.limite_atteinte = True
                        st.session_state.optimisation_terminee = True
                        st.warning("⚠️ Limite de 1000 entrées journal atteinte - Résultats partiels disponibles")
                    else:
                        st.error(f"Erreur inattendue: {str(e)}")
                        import traceback
                        st.code(traceback.format_exc())
                        st.stop()
                
                # Calculer les stats (que l'algorithme soit terminé ou interrompu)
                if terrain_optimise:
                    stats = calculer_stats(terrain_optimise)
                    
                    # Sauvegarder la solution
                    st.session_state.solution = {
                        'terrain': terrain_optimise,
                        'batiments': batiments,
                        'stats': stats,
                        'journal': st.session_state.journal,
                        'batiments_non_places': non_places,
                        'limite_atteinte': st.session_state.limite_atteinte
                    }
                    
                    if not st.session_state.limite_atteinte:
                        st.success("✅ Optimisation terminée avec succès!")
        
        # Afficher les résultats (TOUJOURS en dehors du bloc try/except)
        if st.session_state.solution and st.session_state.optimisation_terminee:
            solution = st.session_state.solution
            
            st.header("📊 Résultats" + (" (PARTIELS)" if solution['limite_atteinte'] else ""))
            
            if solution['limite_atteinte']:
                st.warning("⚠️ Ces résultats sont partiels car la limite de 1000 entrées journal a été atteinte")
                st.info(f"📝 Journal: {len(solution['journal'])} entrées sur 1000")
            
            # Métriques
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Culture totale", solution['stats']['culture_totale'])
            with col2:
                st.metric("Boost ≥25%", solution['stats']['boosts'][25])
            with col3:
                st.metric("Boost ≥50%", solution['stats']['boosts'][50])
            with col4:
                st.metric("Boost 100%", solution['stats']['boosts'][100])
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Placés", len(solution['terrain'].batiments_places))
            with col2:
                st.metric("Non placés", len(solution['batiments_non_places']))
            with col3:
                st.metric("Cases libres", int(np.sum(solution['terrain'].grille == 1)))
            
            # Visualisation simplifiée
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
            
            df_visuel = pd.DataFrame(terrain_mat)
            
            def colorier_cases(val):
                if val == 0:
                    return 'background-color: #ffcccc'  # Rouge clair
                elif val == 1:
                    return 'background-color: #ffffff'  # Blanc
                elif val == 3:
                    return 'background-color: #ffb347'  # Orange
                elif val == 4:
                    return 'background-color: #90ee90'  # Vert clair
                elif val == 5:
                    return 'background-color: #d3d3d3'  # Gris
                return ''
            
            styled_df = df_visuel.style.map(colorier_cases)
            st.dataframe(styled_df, height=400, use_container_width=True)
            
            # Légende
            cols = st.columns(5)
            cols[0].markdown("🟥 **Occupé**")
            cols[1].markdown("⬜ **Libre**")
            cols[2].markdown("🟧 **Culturel**")
            cols[3].markdown("🟩 **Producteur**")
            cols[4].markdown("⬜ **Neutre**")
            
            # Détails des boosts
            with st.expander("📈 Détails des boosts par producteur"):
                if solution['stats']['boost_details']:
                    st.dataframe(pd.DataFrame(solution['stats']['boost_details']))
            
            # Journal
            with st.expander(f"📋 Journal d'exécution ({len(solution['journal'])} entrées)"):
                if solution['limite_atteinte']:
                    st.warning("⚠️ Limite de 1000 entrées atteinte - Journal complet mais optimisation interrompue")
                
                # Afficher les 50 dernières entrées
                for entry in solution['journal'][-50:]:
                    st.text(entry)
                
                if len(solution['journal']) > 50:
                    st.info(f"... et {len(solution['journal']) - 50} entrées précédentes (voir fichier Excel)")
            
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
            
            # BOUTON DE TÉLÉCHARGEMENT - TOUJOURS PRÉSENT
            st.subheader("💾 Télécharger les résultats")
            
            wb = creer_excel_resultat(
                solution['terrain'],
                solution['batiments'],
                solution['stats'],
                solution['journal'],
                solution['batiments_non_places'],
                solution['limite_atteinte']
            )
            
            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            
            # Nom du fichier avec indication si partiel
            filename = "resultats_placement"
            if solution['limite_atteinte']:
                filename += "_partiels"
            filename += ".xlsx"
            
            st.download_button(
                label="📥 Télécharger le rapport Excel" + (" (PARTIEL)" if solution['limite_atteinte'] else ""),
                data=excel_buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    except Exception as e:
        st.error(f"Erreur de lecture du fichier: {str(e)}")
        import traceback
        st.code(traceback.format_exc())

else:
    st.info("👈 Chargez un fichier Excel pour commencer")
    
    with st.expander("📝 Format du fichier attendu"):
        st.markdown("""
        ### Onglet 1: Terrain
        - Matrice de 0 et 1
        - 1 = case libre
        - 0 = case occupée
        
        ### Onglet 2: Bâtiments (avec en-têtes)
        | Nom | Longueur | Largeur | Quantite | Type | Culture | Rayonnement | Boost 25% | Boost 50% | Boost 100% | Production |
        |-----|----------|---------|----------|------|---------|-------------|-----------|-----------|------------|------------|
        
        **Types possibles**: Culturel, Producteur, Neutre
        **Production**: Guerison, Nourriture, Or, ou vide
        """)

st.markdown("---")
st.markdown("🔒 Optimiseur de placement v3.2 - Arrêt à 1000 entrées avec résultats partiels")