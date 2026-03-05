import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import plotly.graph_objects as go
from collections import defaultdict
import time

# Configuration de la page
st.set_page_config(
    page_title="Placement de Bâtiments",
    page_icon="🏗️",
    layout="wide"
)

##############################
# CLASSE POUR LE BACKTRACKING
##############################
class PlacementBacktracking:
    def __init__(self, terrain, batiments):
        self.terrain_original = terrain.copy()
        self.hauteur, self.largeur = terrain.shape
        self.batiments = batiments
        
        # Statistiques
        self.stats = {
            'backtracks': 0,
            'placements_essayes': 0,
            'temps_execution': 0
        }
        
        # Préparer la liste des bâtiments à placer (avec quantités)
        self.batiments_a_placer = []
        for b in batiments:
            for _ in range(int(b['quantite'])):
                # Créer une copie pour éviter les modifications
                bat = b.copy()
                # Ajouter un champ pour suivre le boost
                bat['boost'] = '0%'
                bat['culture_recue'] = 0
                self.batiments_a_placer.append(bat)
        
        # Trier par taille (du plus grand au plus petit) pour optimiser le backtracking
        self.batiments_a_placer.sort(
            key=lambda x: (-(x['longueur'] * x['largeur']), 
                          x['type'] != 'culturel')  # Priorité aux culturels
        )
        
        # Initialiser le terrain de travail
        self.terrain = terrain.astype(object).copy()
        self.carte_rayonnement = np.zeros((self.hauteur, self.largeur))
        
        # Meilleure solution trouvée
        self.meilleure_solution = {
            'terrain': None,
            'batiments_places': [],
            'carte_rayonnement': None,
            'score': -1,
            'nb_places': 0
        }
        
        # Cache pour les placements possibles
        self.cache_placements = {}
    
    def est_case_libre(self, i, j):
        """Vérifie si une case est libre"""
        return self.terrain[i, j] == 1
    
    def trouver_infos_originales(self, nom):
        """Trouve les infos originales d'un bâtiment"""
        for b in self.batiments:
            if b['nom'] == nom:
                return b.get('rayonnement', 0), b.get('culture', 0)
        return 0, 0
    
    def calculer_zone_rayonnement(self, x, y, longueur, largeur, rayonnement, valeur_culture):
        """Calcule la zone de rayonnement autour d'un bâtiment culturel"""
        if rayonnement <= 0 or valeur_culture <= 0:
            return
        
        x_min = max(0, x - rayonnement)
        x_max = min(self.hauteur, x + longueur + rayonnement)
        y_min = max(0, y - rayonnement)
        y_max = min(self.largeur, y + largeur + rayonnement)
        
        for i in range(x_min, x_max):
            for j in range(y_min, y_max):
                if (i < x or i >= x + longueur or j < y or j >= y + largeur):
                    self.carte_rayonnement[i, j] += valeur_culture
    
    def calculer_culture_pour_position(self, x, y, longueur, largeur):
        """Calcule la culture reçue par un bâtiment à une position donnée"""
        culture = 0
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                culture += self.carte_rayonnement[i, j]
        return culture
    
    def calculer_boost(self, batiment, culture_recue):
        """Calcule le boost en fonction de la culture reçue"""
        if batiment['type'] != 'producteur' or culture_recue == 0:
            return '0%'
        
        boost_25 = batiment.get('boost_25', 0)
        boost_50 = batiment.get('boost_50', 0)
        boost_100 = batiment.get('boost_100', 0)
        
        if culture_recue >= boost_100 and boost_100 > 0:
            return '100%'
        elif culture_recue >= boost_50 and boost_50 > 0:
            return '50%'
        elif culture_recue >= boost_25 and boost_25 > 0:
            return '25%'
        else:
            return '0%'
    
    def trouver_positions_possibles(self, batiment):
        """Trouve toutes les positions possibles pour un bâtiment"""
        # Clé pour le cache
        cache_key = f"{batiment['nom']}_{batiment['longueur']}_{batiment['largeur']}"
        
        if cache_key in self.cache_placements:
            # Filtrer les positions qui sont encore libres
            positions = []
            for pos in self.cache_placements[cache_key]:
                x, y, longueur, largeur, orientation = pos
                if self.peut_placer(x, y, longueur, largeur):
                    positions.append(pos)
            return positions
        
        positions = []
        orientations = [(batiment['longueur'], batiment['largeur'], 'horizontal')]
        if batiment['longueur'] != batiment['largeur']:
            orientations.append((batiment['largeur'], batiment['longueur'], 'vertical'))
        
        for longueur, largeur, orientation in orientations:
            for x in range(self.hauteur - longueur + 1):
                for y in range(self.largeur - largeur + 1):
                    if self.peut_placer(x, y, longueur, largeur):
                        positions.append((x, y, longueur, largeur, orientation))
        
        # Mettre en cache
        self.cache_placements[cache_key] = positions
        return positions
    
    def peut_placer(self, x, y, longueur, largeur):
        """Vérifie si on peut placer un rectangle à cette position"""
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                if not self.est_case_libre(i, j):
                    return False
        return True
    
    def placer_batiment(self, batiment, x, y, longueur, largeur, orientation):
        """Place un bâtiment et retourne les infos"""
        # Calculer la culture reçue
        culture_recue = self.calculer_culture_pour_position(x, y, longueur, largeur)
        
        # Calculer le boost
        boost = self.calculer_boost(batiment, culture_recue)
        
        # Marquer les cases comme occupées
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                self.terrain[i, j] = f"{batiment['nom'][:3]}_{i}_{j}"
        
        # Ajouter le rayonnement si c'est un culturel
        if batiment['type'] == 'culturel':
            rayonnement, culture = self.trouver_infos_originales(batiment['nom'])
            self.calculer_zone_rayonnement(x, y, longueur, largeur, rayonnement, culture)
        
        # Créer l'enregistrement du bâtiment
        batiment_place = {
            'nom': batiment['nom'],
            'type': batiment['type'],
            'x': x,
            'y': y,
            'longueur': longueur,
            'largeur': largeur,
            'orientation': orientation,
            'culture_recue': culture_recue,
            'boost': boost,
            'production': batiment.get('production', ''),
            'boost_25': batiment.get('boost_25', 0),
            'boost_50': batiment.get('boost_50', 0),
            'boost_100': batiment.get('boost_100', 0)
        }
        
        return batiment_place
    
    def retirer_batiment(self, batiment_place):
        """Retire un bâtiment du terrain"""
        x, y = batiment_place['x'], batiment_place['y']
        longueur, largeur = batiment_place['longueur'], batiment_place['largeur']
        
        # Remettre les cases à 1
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                self.terrain[i, j] = 1
        
        # Recalculer complètement la carte de rayonnement
        self.carte_rayonnement = np.zeros((self.hauteur, self.largeur))
        for b in self.batiments_places_actuels:
            if b['type'] == 'culturel':
                rayonnement, culture = self.trouver_infos_originales(b['nom'])
                self.calculer_zone_rayonnement(
                    b['x'], b['y'], b['longueur'], b['largeur'],
                    rayonnement, culture
                )
    
    def calculer_score_solution(self, batiments_places):
        """Calcule un score pour une solution"""
        score = 0
        
        # Bonus pour chaque bâtiment placé
        score += len(batiments_places) * 10000
        
        # Bonus pour la culture reçue
        for b in batiments_places:
            if b['type'] == 'producteur':
                score += b['culture_recue']
                if b['boost'] == '100%':
                    score += 1000
                elif b['boost'] == '50%':
                    score += 500
                elif b['boost'] == '25%':
                    score += 100
        
        return score
    
    def backtrack(self, index, batiments_places):
        """Fonction récursive de backtracking"""
        self.stats['backtracks'] += 1
        
        # Si on a placé tous les bâtiments
        if index >= len(self.batiments_a_placer):
            score = self.calculer_score_solution(batiments_places)
            if score > self.meilleure_solution['score']:
                self.meilleure_solution['score'] = score
                self.meilleure_solution['nb_places'] = len(batiments_places)
                self.meilleure_solution['batiments_places'] = batiments_places.copy()
                self.meilleure_solution['terrain'] = self.terrain.copy()
                self.meilleure_solution['carte_rayonnement'] = self.carte_rayonnement.copy()
                
                # Afficher la progression
                if self.stats['backtracks'] % 1000 == 0:
                    st.write(f"  Meilleure solution: {len(batiments_places)} bâtiments placés (exploration: {self.stats['backtracks']})")
            return True
        
        batiment = self.batiments_a_placer[index]
        
        # Sauvegarder l'état actuel des bâtiments placés
        self.batiments_places_actuels = batiments_places
        
        # Trouver toutes les positions possibles
        positions = self.trouver_positions_possibles(batiment)
        self.stats['placements_essayes'] += len(positions)
        
        # Trier les positions par score potentiel
        positions_scorees = []
        for x, y, longueur, largeur, orientation in positions:
            culture = self.calculer_culture_pour_position(x, y, longueur, largeur)
            
            if batiment['type'] == 'culturel':
                # Pour les culturels, favoriser les positions centrales
                score_pos = -abs(x - self.hauteur/2) - abs(y - self.largeur/2)
            else:
                # Pour les producteurs, favoriser la culture
                score_pos = culture * 1000
            
            positions_scorees.append((score_pos, x, y, longueur, largeur, orientation))
        
        # Trier par score décroissant
        positions_scorees.sort(reverse=True)
        
        solution_trouvee = False
        
        for _, x, y, longueur, largeur, orientation in positions_scorees:
            # Placer le bâtiment
            batiment_place = self.placer_batiment(batiment, x, y, longueur, largeur, orientation)
            batiment_place['index'] = index
            
            # Appel récursif
            if self.backtrack(index + 1, batiments_places + [batiment_place]):
                solution_trouvee = True
                # Optionnel: décommenter pour s'arrêter à la première solution
                # return True
            
            # Retirer le bâtiment
            self.retirer_batiment(batiment_place)
        
        return solution_trouvee
    
    def executer_placement(self, temps_max=60):
        """Exécute l'algorithme de backtracking avec timeout"""
        start_time = time.time()
        
        st.write("### Démarrage du backtracking...")
        st.write(f"Nombre de bâtiments à placer: {len(self.batiments_a_placer)}")
        
        # Lancer le backtracking
        self.backtrack(0, [])
        
        self.stats['temps_execution'] = time.time() - start_time
        
        # Appliquer la meilleure solution trouvée
        if self.meilleure_solution['terrain'] is not None:
            self.terrain = self.meilleure_solution['terrain']
            self.carte_rayonnement = self.meilleure_solution['carte_rayonnement']
            self.batiments_places = self.meilleure_solution['batiments_places']
        else:
            self.batiments_places = []
        
        # Déterminer les bâtiments non placés
        indices_places = {b['index'] for b in self.batiments_places}
        batiments_non_places = [
            self.batiments_a_placer[i] 
            for i in range(len(self.batiments_a_placer)) 
            if i not in indices_places
        ]
        
        return self.terrain, self.batiments_places, batiments_non_places
    
    def calculer_statistiques(self):
        """Calcule les statistiques de production"""
        stats = defaultdict(lambda: {
            'total_culture': 0,
            'boost_25': 0,
            'boost_50': 0,
            'boost_100': 0,
            'nb_batiments': 0
        })
        
        for batiment in self.batiments_places:
            if batiment['type'] == 'producteur':
                prod_type = batiment['production'] if batiment['production'] else 'Rien'
                
                stats[prod_type]['total_culture'] += batiment['culture_recue']
                stats[prod_type]['nb_batiments'] += 1
                
                if batiment['boost'] == '25%':
                    stats[prod_type]['boost_25'] += 1
                elif batiment['boost'] == '50%':
                    stats[prod_type]['boost_50'] += 1
                elif batiment['boost'] == '100%':
                    stats[prod_type]['boost_100'] += 1
        
        return dict(stats)


##############################
# FONCTIONS DE GESTION EXCEL (inchangées)
##############################
def lire_fichier_excel(uploaded_file):
    """
    Lit le fichier Excel uploadé et extrait les données du terrain et des bâtiments
    """
    # Lire tous les onglets du fichier Excel
    xls = pd.ExcelFile(uploaded_file)
    
    # Le premier onglet contient le terrain
    df_terrain = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)
    terrain = df_terrain.values.astype(int)
    
    # Le second onglet contient les bâtiments
    df_batiments = pd.read_excel(xls, sheet_name=xls.sheet_names[1])
    
    # Normaliser les noms de colonnes
    df_batiments.columns = df_batiments.columns.str.strip().str.replace(' ', '').str.lower()
    
    # Dictionnaire de correspondance des noms de colonnes possibles
    mapping_colonnes = {
        'nom': ['nom', 'name', 'nome'],
        'longueur': ['longueur', 'length', 'long'],
        'largeur': ['largeur', 'width'],
        'quantite': ['quantite', 'quantity', 'qté', 'qte', 'qt', 'quantité'],
        'type': ['type', 'tipo'],
        'culture': ['culture', 'cult'],
        'rayonnement': ['rayonnement', 'range', 'rayon'],
        'boost25%': ['boost25%', 'boost25', '25%boost', 'boost25pourcent'],
        'boost50%': ['boost50%', 'boost50', '50%boost', 'boost50pourcent'],
        'boost100%': ['boost100%', 'boost100', '100%boost', 'boost100pourcent'],
        'production': ['production', 'prod']
    }
    
    # Fonction pour trouver la colonne correspondante
    def trouver_colonne(noms_possibles):
        for nom in noms_possibles:
            if nom in df_batiments.columns:
                return nom
        return None
    
    # Récupérer les noms de colonnes réels
    colonne_nom = trouver_colonne(mapping_colonnes['nom'])
    colonne_longueur = trouver_colonne(mapping_colonnes['longueur'])
    colonne_largeur = trouver_colonne(mapping_colonnes['largeur'])
    colonne_quantite = trouver_colonne(mapping_colonnes['quantite'])
    colonne_type = trouver_colonne(mapping_colonnes['type'])
    colonne_culture = trouver_colonne(mapping_colonnes['culture'])
    colonne_rayonnement = trouver_colonne(mapping_colonnes['rayonnement'])
    colonne_boost25 = trouver_colonne(mapping_colonnes['boost25%'])
    colonne_boost50 = trouver_colonne(mapping_colonnes['boost50%'])
    colonne_boost100 = trouver_colonne(mapping_colonnes['boost100%'])
    colonne_production = trouver_colonne(mapping_colonnes['production'])
    
    # Vérifier que les colonnes essentielles sont trouvées
    colonnes_manquantes = []
    if not colonne_quantite:
        colonnes_manquantes.append('Quantite')
    if not colonne_nom:
        colonnes_manquantes.append('Nom')
    if not colonne_longueur:
        colonnes_manquantes.append('Longueur')
    if not colonne_largeur:
        colonnes_manquantes.append('Largeur')
    
    if colonnes_manquantes:
        st.error(f"Colonnes manquantes dans le fichier: {', '.join(colonnes_manquantes)}")
        st.info("Les colonnes trouvées sont: " + ', '.join(df_batiments.columns))
        return None, None
    
    batiments = []
    for _, row in df_batiments.iterrows():
        try:
            batiment = {
                'nom': str(row[colonne_nom]),
                'longueur': int(float(row[colonne_longueur])),
                'largeur': int(float(row[colonne_largeur])),
                'quantite': int(float(row[colonne_quantite])),
                'type': str(row[colonne_type]).lower(),
                'culture': float(row[colonne_culture]) if colonne_culture and pd.notna(row[colonne_culture]) else 0,
                'rayonnement': int(float(row[colonne_rayonnement])) if colonne_rayonnement and pd.notna(row[colonne_rayonnement]) else 0,
                'boost_25': float(row[colonne_boost25]) if colonne_boost25 and pd.notna(row[colonne_boost25]) else 0,
                'boost_50': float(row[colonne_boost50]) if colonne_boost50 and pd.notna(row[colonne_boost50]) else 0,
                'boost_100': float(row[colonne_boost100]) if colonne_boost100 and pd.notna(row[colonne_boost100]) else 0,
                'production': str(row[colonne_production]) if colonne_production and pd.notna(row[colonne_production]) else ''
            }
            batiments.append(batiment)
        except Exception as e:
            st.warning(f"Erreur lors de la lecture d'une ligne: {e}")
            continue
    
    return terrain, batiments

def generer_fichier_resultat(terrain_original, terrain_place, batiments_places, stats_culture, tous_les_batiments, batiments_non_places_list):
    """
    Génère un fichier Excel avec les résultats
    """
    output = io.BytesIO()
    
    # Créer un workbook avec openpyxl directement
    wb = Workbook()
    
    # Supprimer l'onglet par défaut
    wb.remove(wb.active)
    
    # Définir les styles de couleur
    fill_vert = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')
    font_vert = Font(color='006100')
    
    fill_orange = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid')
    font_orange = Font(color='9C6500')
    
    fill_gris = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
    font_gris = Font(color='333333')
    
    # Onglet 1: Terrain original
    ws1 = wb.create_sheet('Terrain_Original')
    for row in dataframe_to_rows(pd.DataFrame(terrain_original), index=False, header=False):
        ws1.append(row)
    
    # Onglet 2: Terrain avec bâtiments placés
    ws2 = wb.create_sheet('Terrain_Place')
    terrain_affichage = []
    for i in range(terrain_place.shape[0]):
        ligne = []
        for j in range(terrain_place.shape[1]):
            valeur = terrain_place[i, j]
            if isinstance(valeur, (int, np.integer)):
                if valeur == 1:
                    ligne.append('LIBRE')
                elif valeur == 0:
                    ligne.append('OCCUPE')
                else:
                    ligne.append(str(valeur))
            else:
                ligne.append(str(valeur))
        terrain_affichage.append(ligne)
    
    df_terrain_place = pd.DataFrame(terrain_affichage)
    for row in dataframe_to_rows(df_terrain_place, index=False, header=False):
        ws2.append(row)
    
    # Appliquer les couleurs à l'onglet Terrain_Place
    for i in range(df_terrain_place.shape[0]):
        for j in range(df_terrain_place.shape[1]):
            cellule = str(df_terrain_place.iat[i, j])
            if cellule not in ['LIBRE', 'OCCUPE'] and '_' in cellule:
                # C'est un bâtiment, chercher son type
                nom_bat = cellule.split('_')[0]
                for bat in batiments_places:
                    if bat['nom'][:3] == nom_bat or (len(bat['nom']) >= 3 and bat['nom'][:3] == nom_bat):
                        cell = ws2.cell(row=i+1, column=j+1)
                        if bat['type'] == 'producteur':
                            cell.fill = fill_vert
                            cell.font = font_vert
                        elif bat['type'] == 'culturel':
                            cell.fill = fill_orange
                            cell.font = font_orange
                        else:
                            cell.fill = fill_gris
                            cell.font = font_gris
                        break
    
    # Onglet 3: Liste des bâtiments placés
    ws3 = wb.create_sheet('Placements')
    data_placement = []
    for bat in batiments_places:
        data_placement.append({
            'Nom': bat['nom'],
            'Type': bat['type'],
            'Position_X': bat['x'],
            'Position_Y': bat['y'],
            'Orientation': bat['orientation'],
            'Culture_recue': bat['culture_recue'],
            'Boost_atteint': bat['boost']
        })
    
    df_placement = pd.DataFrame(data_placement)
    
    # Écrire les en-têtes
    for col_idx, col_name in enumerate(df_placement.columns, 1):
        ws3.cell(row=1, column=col_idx, value=col_name)
    
    # Écrire les données
    for row_idx, row in df_placement.iterrows():
        for col_idx, value in enumerate(row, 1):
            ws3.cell(row=row_idx+2, column=col_idx, value=value)
    
    # Appliquer les couleurs à l'onglet Placements
    type_col_idx = None
    for col_idx, col_name in enumerate(df_placement.columns, 1):
        if col_name == 'Type':
            type_col_idx = col_idx
            break
    
    if type_col_idx is not None:
        for row_idx, row in df_placement.iterrows():
            cell = ws3.cell(row=row_idx+2, column=type_col_idx)
            if row['Type'] == 'producteur':
                cell.fill = fill_vert
                cell.font = font_vert
            elif row['Type'] == 'culturel':
                cell.fill = fill_orange
                cell.font = font_orange
            else:
                cell.fill = fill_gris
                cell.font = font_gris
    
    # Onglet 4: Bâtiments non placés
    quantites_placees = {}
    for bat in batiments_places:
        if bat['nom'] not in quantites_placees:
            quantites_placees[bat['nom']] = 0
        quantites_placees[bat['nom']] += 1
    
    # Calculer les statistiques pour les bâtiments non placés
    cases_libres_restantes = np.sum(terrain_place == 1)
    total_cases_batiments_non_places = 0
    total_cases_batiments_places = 0
    
    # Calculer le nombre de cases occupées par les bâtiments placés
    for bat in batiments_places:
        total_cases_batiments_places += bat['longueur'] * bat['largeur']
    
    # Compter les bâtiments non placés
    non_places_comptes = defaultdict(int)
    for bat in batiments_non_places_list:
        non_places_comptes[bat['nom']] += 1
    
    non_places = []
    for bat in tous_les_batiments:
        nom = bat['nom']
        quantite_demandee = int(bat['quantite'])
        quantite_placee = quantites_placees.get(nom, 0)
        quantite_non_placee = non_places_comptes.get(nom, 0)
        
        if quantite_non_placee > 0:
            cases_batiment = bat['longueur'] * bat['largeur'] * quantite_non_placee
            total_cases_batiments_non_places += cases_batiment
            
            non_places.append({
                'Nom': nom,
                'Type': bat['type'],
                'Longueur': bat['longueur'],
                'Largeur': bat['largeur'],
                'Quantite_demandee': quantite_demandee,
                'Quantite_placee': quantite_placee,
                'Reste_a_placer': quantite_non_placee,
                'Culture': bat['culture'] if bat['type'] == 'culturel' else 'N/A',
                'Rayonnement': bat['rayonnement'] if bat['type'] == 'culturel' else 'N/A',
                'Production': bat['production'] if bat['type'] == 'producteur' else 'N/A',
                'Cases_necessaires': cases_batiment
            })
    
    # Ajouter une ligne de total pour les bâtiments non placés
    if non_places:
        non_places.append({
            'Nom': 'TOTAL',
            'Type': '',
            'Longueur': '',
            'Largeur': '',
            'Quantite_demandee': '',
            'Quantite_placee': '',
            'Reste_a_placer': sum(b['Reste_a_placer'] for b in non_places),
            'Culture': '',
            'Rayonnement': '',
            'Production': '',
            'Cases_necessaires': total_cases_batiments_non_places
        })
    
    if non_places:
        ws4 = wb.create_sheet('Non_Places')
        df_non_places = pd.DataFrame(non_places)
        
        # Écrire les en-têtes
        for col_idx, col_name in enumerate(df_non_places.columns, 1):
            ws4.cell(row=1, column=col_idx, value=col_name)
        
        # Écrire les données
        for row_idx, row in df_non_places.iterrows():
            for col_idx, value in enumerate(row, 1):
                ws4.cell(row=row_idx+2, column=col_idx, value=value)
        
        # Appliquer les couleurs à l'onglet Non_Places
        type_non_places_idx = None
        for col_idx, col_name in enumerate(df_non_places.columns, 1):
            if col_name == 'Type':
                type_non_places_idx = col_idx
                break
        
        if type_non_places_idx is not None:
            for row_idx, row in df_non_places.iterrows():
                if row['Nom'] != 'TOTAL':
                    cell = ws4.cell(row=row_idx+2, column=type_non_places_idx)
                    if row['Type'] == 'producteur':
                        cell.fill = fill_vert
                        cell.font = font_vert
                    elif row['Type'] == 'culturel':
                        cell.fill = fill_orange
                        cell.font = font_orange
                    else:
                        cell.fill = fill_gris
                        cell.font = font_gris
    
    # Onglet 5: Statistiques
    ws5 = wb.create_sheet('Statistiques')
    
    stats_data = []
    for prod, stats in stats_culture.items():
        stats_data.append({
            'Type_Production': prod,
            'Culture_Total_Recue': stats['total_culture'],
            'Boost_25_atteint': stats['boost_25'],
            'Boost_50_atteint': stats['boost_50'],
            'Boost_100_atteint': stats['boost_100'],
            'Nombre_batiments': stats['nb_batiments']
        })
    
    if stats_data:
        df_stats = pd.DataFrame(stats_data)
        
        # Écrire les en-têtes
        for col_idx, col_name in enumerate(df_stats.columns, 1):
            ws5.cell(row=1, column=col_idx, value=col_name)
        
        # Écrire les données
        for row_idx, row in df_stats.iterrows():
            for col_idx, value in enumerate(row, 1):
                ws5.cell(row=row_idx+2, column=col_idx, value=value)
    
    # Onglet 6: Résumé
    ws6 = wb.create_sheet('Resume')
    
    resume_data = [
        ['Description', 'Valeur'],
        ['Cases libres initiales', np.sum(terrain_original == 1)],
        ['Cases occupées initiales', np.sum(terrain_original == 0)],
        ['Cases libres restantes', cases_libres_restantes],
        ['Cases occupées par des bâtiments', total_cases_batiments_places],
        [''],
        ['Bâtiments placés', len(batiments_places)],
        ['Bâtiments non placés', sum(b['Reste_a_placer'] for b in non_places if b['Nom'] != 'TOTAL') if non_places else 0],
        [''],
        ['Cases nécessaires pour les bâtiments non placés', total_cases_batiments_non_places],
        ['Suffisamment de cases libres ?', 'OUI' if cases_libres_restantes >= total_cases_batiments_non_places else 'NON'],
    ]
    
    # Ajouter les totaux de culture par type
    if stats_culture:
        resume_data.extend([
            [''],
            ['TOTAL Culture par type de production'],
            ['Type', 'Culture totale recue']
        ])
        
        for prod, stats in stats_culture.items():
            resume_data.append([prod, stats['total_culture']])
    
    df_resume = pd.DataFrame(resume_data)
    for row in dataframe_to_rows(df_resume, index=False, header=False):
        ws6.append(row)
    
    # Sauvegarder le workbook
    wb.save(output)
    output.seek(0)
    
    return output


##############################
# INTERFACE STREAMLIT
##############################

# Titre
st.title("🏗️ Optimiseur de Placement de Bâtiments (Backtracking)")
st.markdown("---")

# Sidebar pour les instructions
with st.sidebar:
    st.header("📋 Instructions")
    st.markdown("""
    1. Préparez votre fichier Excel avec deux onglets:
        - **Onglet 1**: Terrain (matrice de 0 et 1)
        - **Onglet 2**: Bâtiments avec colonnes:
            - Nom, Longueur, Largeur, Quantité, Type
            - Culture, Rayonnement, Boost 25%, Boost 50%, Boost 100%, Production
    
    2. Uploadez le fichier ci-dessous
    
    3. Lancez l'optimisation
    
    4. Téléchargez les résultats
    """)
    
    st.markdown("---")
    st.header("📁 Upload du fichier")
    uploaded_file = st.file_uploader(
        "Choisissez votre fichier Excel",
        type=['xlsx', 'xls'],
        help="Format accepté: .xlsx, .xls"
    )

# Zone principale
if uploaded_file is not None:
    try:
        # Lecture du fichier
        with st.spinner("Lecture du fichier en cours..."):
            terrain, batiments = lire_fichier_excel(uploaded_file)
        
        if terrain is None or batiments is None:
            st.stop()
        
        # Affichage des données lues
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📊 Terrain original")
            st.dataframe(pd.DataFrame(terrain), use_container_width=True, height=400)
            
            # Statistiques du terrain
            cases_libres = np.sum(terrain == 1)
            cases_occupees = np.sum(terrain == 0)
            st.info(f"📌 Cases libres: {cases_libres} | Cases occupées: {cases_occupees}")
        
        with col2:
            st.subheader("🏢 Bâtiments à placer")
            df_batiments = pd.DataFrame(batiments)
            st.dataframe(df_batiments, use_container_width=True, height=400)
            
            # Résumé des bâtiments
            total_batiments = sum(int(b['quantite']) for b in batiments)
            total_cases_necessaires = sum(int(b['quantite']) * b['longueur'] * b['largeur'] for b in batiments)
            st.info(f"📦 Total de bâtiments à placer: {total_batiments}")
            st.info(f"📐 Cases totales nécessaires: {total_cases_necessaires}")
        
        st.markdown("---")
        
        # Bouton pour lancer l'optimisation
        if st.button("🚀 Lancer l'optimisation", type="primary", use_container_width=True):
            with st.spinner("Optimisation en cours... Cela peut prendre plusieurs minutes."):
                # Création de l'instance de placement
                placement = PlacementBacktracking(terrain, batiments)
                
                # Exécution de l'algorithme
                terrain_place, batiments_places, batiments_non_places = placement.executer_placement()
                
                # Calcul des statistiques
                stats = placement.calculer_statistiques()
                
                # Sauvegarde dans la session
                st.session_state['terrain_place'] = terrain_place
                st.session_state['batiments_places'] = batiments_places
                st.session_state['stats'] = stats
                st.session_state['terrain_original'] = terrain
                st.session_state['batiments_complets'] = batiments
                st.session_state['batiments_non_places'] = batiments_non_places
                
                st.success(f"✅ Optimisation terminée avec succès! {len(batiments_places)} bâtiments placés, {len(batiments_non_places)} non placés.")
                
                # Afficher les statistiques de l'algorithme
                st.info(f"📊 Statistiques de l'algorithme: {placement.stats['backtracks']} backtracks, {placement.stats['placements_essayes']} placements essayés")
        
        # Affichage des résultats si disponibles
        if 'terrain_place' in st.session_state:
            st.markdown("---")
            st.header("📈 Résultats de l'optimisation")
            
            # Statistiques globales
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Bâtiments placés", len(st.session_state['batiments_places']))
            with col2:
                st.metric("Bâtiments non placés", len(st.session_state['batiments_non_places']))
            with col3:
                cases_libres = np.sum(st.session_state['terrain_place'] == 1)
                st.metric("Cases libres restantes", cases_libres)
            with col4:
                total = len(st.session_state['batiments_places']) + len(st.session_state['batiments_non_places'])
                st.metric("Taux de placement", f"{len(st.session_state['batiments_places'])/total*100:.1f}%" if total > 0 else "0%")
            
            # Statistiques de production
            st.subheader("📊 Statistiques de production")
            if st.session_state['stats']:
                stats_df = pd.DataFrame([
                    {
                        'Production': k,
                        'Culture totale': v['total_culture'],
                        'Boost 25%': v['boost_25'],
                        'Boost 50%': v['boost_50'],
                        'Boost 100%': v['boost_100'],
                        'Nb bâtiments': v['nb_batiments']
                    }
                    for k, v in st.session_state['stats'].items()
                ])
                st.dataframe(stats_df, use_container_width=True)
            
            # Visualisation du terrain
            st.subheader("🗺️ Visualisation du terrain")
            
            # Créer une matrice pour la visualisation
            vis_terrain = np.zeros_like(st.session_state['terrain_original'], dtype=float)
            for bat in st.session_state['batiments_places']:
                valeur = 1 if bat['type'] == 'culturel' else 2
                for i in range(bat['x'], bat['x'] + bat['longueur']):
                    for j in range(bat['y'], bat['y'] + bat['largeur']):
                        vis_terrain[i, j] = valeur
            
            # Ajouter les cases libres/occupées originales
            for i in range(vis_terrain.shape[0]):
                for j in range(vis_terrain.shape[1]):
                    if vis_terrain[i, j] == 0:
                        vis_terrain[i, j] = 3 if st.session_state['terrain_original'][i, j] == 1 else 4
            
            # Création de la figure Plotly
            fig = go.Figure(data=go.Heatmap(
                z=vis_terrain,
                colorscale=[
                    [0, 'lightblue'],   # Bâtiment culturel
                    [0.33, 'lightgreen'], # Bâtiment producteur
                    [0.66, 'white'],      # Case libre
                    [1, 'lightgray']      # Case occupée
                ],
                showscale=False,
                text=[[str(st.session_state['terrain_place'][i, j])[:10] + '...' 
                       if len(str(st.session_state['terrain_place'][i, j])) > 10 
                       else str(st.session_state['terrain_place'][i, j]) 
                       for j in range(vis_terrain.shape[1])] 
                      for i in range(vis_terrain.shape[0])],
                hoverinfo='text'
            ))
            
            fig.update_layout(
                title="Carte des placements",
                xaxis_title="Colonnes",
                yaxis_title="Lignes",
                height=600
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Liste des placements
            with st.expander("📋 Voir le détail des placements"):
                placements_df = pd.DataFrame(st.session_state['batiments_places'])
                st.dataframe(placements_df, use_container_width=True)
            
            # Bâtiments non placés
            if st.session_state['batiments_non_places']:
                st.warning("⚠️ Certains bâtiments n'ont pas pu être placés!")
                
                # Compter par type
                non_places_count = defaultdict(int)
                for bat in st.session_state['batiments_non_places']:
                    non_places_count[bat['nom']] += 1
                
                non_places_df = pd.DataFrame([
                    {'Nom': nom, 'Quantité non placée': count}
                    for nom, count in non_places_count.items()
                ])
                st.dataframe(non_places_df, use_container_width=True)
                
                # Afficher le résumé des cases
                cases_libres_restantes = np.sum(st.session_state['terrain_place'] == 1)
                cases_necessaires = sum(
                    bat['longueur'] * bat['largeur'] 
                    for bat in st.session_state['batiments_non_places']
                )
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Cases libres restantes", cases_libres_restantes)
                with col2:
                    st.metric("Cases nécessaires", cases_necessaires)
                with col3:
                    st.metric("Suffisant", "✅ OUI" if cases_libres_restantes >= cases_necessaires else "❌ NON")
            
            # Bouton de téléchargement
            st.markdown("---")
            
            # Génération du fichier de résultats
            output_file = generer_fichier_resultat(
                st.session_state['terrain_original'],
                st.session_state['terrain_place'],
                st.session_state['batiments_places'],
                st.session_state['stats'],
                st.session_state['batiments_complets'],
                st.session_state['batiments_non_places']
            )
            
            st.download_button(
                label="📥 Télécharger les résultats (Excel)",
                data=output_file,
                file_name="resultats_placement.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement: {str(e)}")
        st.exception(e)

else:
    # Message d'accueil
    st.info("👈 Veuillez uploader un fichier Excel pour commencer")
    
    # Exemple de structure
    with st.expander("📝 Voir un exemple de structure de fichier"):
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Onglet Terrain")
            exemple_terrain = pd.DataFrame([
                [1, 1, 1, 0, 1],
                [1, 0, 1, 1, 1],
                [1, 1, 1, 0, 1],
                [0, 1, 1, 1, 1]
            ])
            st.dataframe(exemple_terrain)
        
        with col2:
            st.subheader("Onglet Bâtiments")
            exemple_batiments = pd.DataFrame([
                ['Maison', 2, 2, 1, 'culturel', 10, 1, 5, 10, 20, ''],
                ['Ferme', 3, 2, 2, 'producteur', 0, 0, 10, 20, 30, 'Nourriture'],
                ['Atelier', 2, 1, 1, 'producteur', 0, 0, 5, 15, 25, 'Or']
            ], columns=['Nom', 'Longueur', 'Largeur', 'Quantité', 'Type', 
                       'Culture', 'Rayonnement', 'Boost 25%', 'Boost 50%', 'Boost 100%', 'Production'])
            st.dataframe(exemple_batiments)

# Footer
st.markdown("---")
st.markdown("🚀 Application avec backtracking pour exploration systématique")