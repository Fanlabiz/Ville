import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import plotly.graph_objects as go
from collections import defaultdict
import random
from copy import deepcopy

# Configuration de la page
st.set_page_config(
    page_title="Placement de Bâtiments",
    page_icon="🏗️",
    layout="wide"
)

##############################
# GESTIONNAIRE DE PLACEMENT
##############################
class PlacementManager:
    def __init__(self, terrain, batiments):
        self.terrain_original = terrain.copy()
        self.hauteur, self.largeur = terrain.shape
        self.batiments = batiments
        
        # Statistiques
        self.stats = {
            'phase1_placements': 0,
            'phase2_deplacements': 0
        }
        
        # Initialiser
        self.reset()
    
    def reset(self):
        """Réinitialise l'état"""
        self.terrain = self.terrain_original.astype(object).copy()
        self.batiments_places = []
        self.carte_culture = np.zeros((self.hauteur, self.largeur))
        self.non_places = []
    
    ##############################
    # PHASE 1 : PLACEMENT COMPACT
    ##############################
    
    def est_libre(self, x, y, h, l):
        """Vérifie si une zone est libre"""
        if x + h > self.hauteur or y + l > self.largeur:
            return False
        for i in range(x, x + h):
            for j in range(y, y + l):
                if self.terrain[i, j] != 1:
                    return False
        return True
    
    def trouver_tous_placements(self, batiment):
        """Trouve tous les placements possibles pour un bâtiment"""
        placements = []
        
        orientations = [(batiment['longueur'], batiment['largeur'], 'horizontal')]
        if batiment['longueur'] != batiment['largeur']:
            orientations.append((batiment['largeur'], batiment['longueur'], 'vertical'))
        
        for h, l, orientation in orientations:
            for x in range(self.hauteur - h + 1):
                for y in range(self.largeur - l + 1):
                    if self.est_libre(x, y, h, l):
                        # Calculer un score de compacité (favoriser les bords)
                        score_compacite = x + y  # Proche du bord = meilleur score
                        placements.append((x, y, h, l, orientation, score_compacite))
        
        return placements
    
    def placer_batiment_phase1(self, batiment, x, y, h, l, orientation):
        """Place un bâtiment pendant la phase 1 (sans boost)"""
        # Marquer les cases
        for i in range(x, x + h):
            for j in range(y, y + l):
                self.terrain[i, j] = f"{batiment['nom'][:3]}_{i}_{j}"
        
        # Ajouter aux places
        self.batiments_places.append({
            'nom': batiment['nom'],
            'type': batiment['type'],
            'x': x,
            'y': y,
            'longueur': h,
            'largeur': l,
            'orientation': orientation,
            'culture_recue': 0,
            'boost': '0%',
            'production': batiment.get('production', '')
        })
        
        self.stats['phase1_placements'] += 1
    
    def phase1_placement_compact(self):
        """Phase 1 : Placer tous les bâtiments de manière compacte"""
        st.write("### Phase 1 : Placement compact")
        
        # Préparer la liste des bâtiments à placer
        a_placer = []
        for b in self.batiments:
            for _ in range(int(b['quantite'])):
                a_placer.append(b.copy())
        
        # Trier par taille décroissante (grands d'abord)
        a_placer.sort(key=lambda x: -(x['longueur'] * x['largeur']))
        
        total = len(a_placer)
        progress_bar = st.progress(0)
        
        for i, batiment in enumerate(a_placer):
            progress_bar.progress((i + 1) / total)
            
            placements = self.trouver_tous_placements(batiment)
            
            if placements:
                # Prendre le placement avec le meilleur score de compacité
                placements.sort(key=lambda p: p[5])  # Trier par score_compacite
                x, y, h, l, orientation, _ = placements[0]
                self.placer_batiment_phase1(batiment, x, y, h, l, orientation)
            else:
                # Garder pour la phase 2
                self.non_places.append(batiment)
        
        progress_bar.empty()
        st.write(f"✅ {self.stats['phase1_placements']} bâtiments placés")
        st.write(f"❌ {len(self.non_places)} bâtiments non placés")
    
    ##############################
    # PHASE 2 : OPTIMISATION DES BOOSTS
    ##############################
    
    def mettre_a_jour_carte_culture(self):
        """Met à jour la carte de culture (rayonnement des bâtiments culturels)"""
        self.carte_culture = np.zeros((self.hauteur, self.largeur))
        
        for bat in self.batiments_places:
            if bat['type'] == 'culturel':
                # Chercher les infos originales
                for b_orig in self.batiments:
                    if b_orig['nom'] == bat['nom']:
                        rayonnement = b_orig.get('rayonnement', 0)
                        culture = b_orig.get('culture', 0)
                        
                        if rayonnement > 0 and culture > 0:
                            x_min = max(0, bat['x'] - rayonnement)
                            x_max = min(self.hauteur, bat['x'] + bat['longueur'] + rayonnement)
                            y_min = max(0, bat['y'] - rayonnement)
                            y_max = min(self.largeur, bat['y'] + bat['largeur'] + rayonnement)
                            
                            for i in range(x_min, x_max):
                                for j in range(y_min, y_max):
                                    if (i < bat['x'] or i >= bat['x'] + bat['longueur'] or 
                                        j < bat['y'] or j >= bat['y'] + bat['largeur']):
                                        self.carte_culture[i, j] += culture
                        break
    
    def calculer_culture_batiment(self, bat):
        """Calcule la culture reçue par un bâtiment"""
        culture = 0
        for i in range(bat['x'], bat['x'] + bat['longueur']):
            for j in range(bat['y'], bat['y'] + bat['largeur']):
                culture += self.carte_culture[i, j]
        return culture
    
    def calculer_boost(self, batiment, culture):
        """Calcule le boost en fonction de la culture"""
        if batiment['type'] != 'producteur' or culture == 0:
            return '0%'
        
        # Chercher les seuils
        for b_orig in self.batiments:
            if b_orig['nom'] == batiment['nom']:
                boost_25 = b_orig.get('boost_25', 0)
                boost_50 = b_orig.get('boost_50', 0)
                boost_100 = b_orig.get('boost_100', 0)
                
                if culture >= boost_100 and boost_100 > 0:
                    return '100%'
                elif culture >= boost_50 and boost_50 > 0:
                    return '50%'
                elif culture >= boost_25 and boost_25 > 0:
                    return '25%'
                break
        
        return '0%'
    
    def calculer_score_placement(self, batiment, x, y, h, l):
        """Calcule le score potentiel d'un placement"""
        if not self.est_libre(x, y, h, l):
            return -1
        
        # Culture potentielle
        culture = 0
        for i in range(x, x + h):
            for j in range(y, y + l):
                culture += self.carte_culture[i, j]
        
        # Boost potentiel
        boost = self.calculer_boost(batiment, culture)
        multiplicateur = {
            '100%': 4,
            '50%': 2,
            '25%': 1.2,
            '0%': 0.8
        }[boost]
        
        return culture * multiplicateur
    
    def retirer_batiment(self, index):
        """Retire un bâtiment du terrain"""
        bat = self.batiments_places.pop(index)
        
        # Effacer du terrain
        for i in range(bat['x'], bat['x'] + bat['longueur']):
            for j in range(bat['y'], bat['y'] + bat['largeur']):
                self.terrain[i, j] = 1
    
    def trouver_meilleur_deplacement(self, batiment, ancien_x, ancien_y, ancien_h, ancien_l):
        """Trouve le meilleur nouvel emplacement pour un bâtiment"""
        meilleur_score = -1
        meilleur_placement = None
        
        orientations = [(batiment['longueur'], batiment['largeur'], 'horizontal')]
        if batiment['longueur'] != batiment['largeur']:
            orientations.append((batiment['largeur'], batiment['longueur'], 'vertical'))
        
        # Chercher dans un rayon limité autour de l'ancienne position
        rayon = 5
        for dx in range(-rayon, rayon + 1):
            for dy in range(-rayon, rayon + 1):
                for h, l, orientation in orientations:
                    x = ancien_x + dx
                    y = ancien_y + dy
                    
                    if x < 0 or y < 0 or x + h > self.hauteur or y + l > self.largeur:
                        continue
                    
                    score = self.calculer_score_placement(batiment, x, y, h, l)
                    if score > meilleur_score:
                        meilleur_score = score
                        meilleur_placement = (x, y, h, l, orientation)
        
        return meilleur_placement, meilleur_score
    
    def phase2_optimisation_boosts(self):
        """Phase 2 : Optimiser les boosts en déplaçant les bâtiments"""
        st.write("### Phase 2 : Optimisation des boosts")
        
        # Mettre à jour la carte de culture
        self.mettre_a_jour_carte_culture()
        
        # Identifier les producteurs avec mauvais boost
        producteurs = []
        for idx, bat in enumerate(self.batiments_places):
            if bat['type'] == 'producteur':
                culture = self.calculer_culture_batiment(bat)
                boost = self.calculer_boost(bat, culture)
                if boost != '100%':  # Chercher à améliorer
                    producteurs.append((idx, bat, culture, boost))
        
        # Trier par priorité (d'abord ceux sans boost)
        producteurs.sort(key=lambda x: (
            0 if x[3] == '0%' else 1 if x[3] == '25%' else 2 if x[3] == '50%' else 3,
            -x[2]  # Culture actuelle (décroissant)
        ))
        
        total = len(producteurs)
        progress_bar = st.progress(0)
        
        for i, (idx, bat, culture_actuelle, boost_actuel) in enumerate(producteurs):
            progress_bar.progress((i + 1) / total)
            
            # Sauvegarder l'état
            ancien_x, ancien_y = bat['x'], bat['y']
            ancien_h, ancien_l = bat['longueur'], bat['largeur']
            
            # Retirer temporairement
            self.retirer_batiment(idx)
            
            # Trouver meilleur nouvel emplacement
            meilleur_placement, meilleur_score = self.trouver_meilleur_deplacement(
                bat, ancien_x, ancien_y, ancien_h, ancien_l
            )
            
            if meilleur_placement:
                x, y, h, l, orientation = meilleur_placement
                
                # Calculer le nouveau boost potentiel
                culture_potentielle = 0
                for ci in range(x, x + h):
                    for cj in range(y, y + l):
                        culture_potentielle += self.carte_culture[ci, cj]
                
                nouveau_boost = self.calculer_boost(bat, culture_potentielle)
                
                # Décider si on déplace
                ordre_boost = {'0%': 0, '25%': 1, '50%': 2, '100%': 3}
                if ordre_boost[nouveau_boost] > ordre_boost[boost_actuel]:
                    # Meilleur boost → on déplace
                    self.placer_batiment_phase1(bat, x, y, h, l, orientation)
                    self.stats['phase2_deplacements'] += 1
                    
                    # Mettre à jour la carte de culture
                    if bat['type'] == 'culturel':
                        self.mettre_a_jour_carte_culture()
                else:
                    # Pas mieux → on remet à l'ancien endroit
                    self.placer_batiment_phase1(bat, ancien_x, ancien_y, ancien_h, ancien_l, bat['orientation'])
            else:
                # Pas de placement possible → on remet
                self.placer_batiment_phase1(bat, ancien_x, ancien_y, ancien_h, ancien_l, bat['orientation'])
        
        progress_bar.empty()
        st.write(f"🔄 {self.stats['phase2_deplacements']} bâtiments déplacés")
    
    ##############################
    # PHASE 3 : PLACEMENT FINAL
    ##############################
    
    def phase3_placement_final(self):
        """Phase 3 : Placer les bâtiments restants"""
        if not self.non_places:
            return
        
        st.write("### Phase 3 : Placement final")
        
        # Trier par taille (petits d'abord)
        self.non_places.sort(key=lambda x: x['longueur'] * x['largeur'])
        
        progress_bar = st.progress(0)
        total = len(self.non_places)
        places = 0
        
        for i, batiment in enumerate(self.non_places):
            progress_bar.progress((i + 1) / total)
            
            placements = self.trouver_tous_placements(batiment)
            
            if placements:
                # Prendre le meilleur placement (maximiser culture)
                meilleur_placement = None
                meilleure_culture = -1
                
                for x, y, h, l, orientation, _ in placements:
                    culture = 0
                    for ci in range(x, x + h):
                        for cj in range(y, y + l):
                            culture += self.carte_culture[ci, cj]
                    
                    if culture > meilleure_culture:
                        meilleure_culture = culture
                        meilleur_placement = (x, y, h, l, orientation)
                
                if meilleur_placement:
                    x, y, h, l, orientation = meilleur_placement
                    self.placer_batiment_phase1(batiment, x, y, h, l, orientation)
                    places += 1
        
        progress_bar.empty()
        st.write(f"➕ {places} bâtiments supplémentaires placés")
        
        # Mettre à jour la liste des non placés
        self.non_places = [b for i, b in enumerate(self.non_places) 
                          if i >= places]
    
    ##############################
    # CALCUL DES STATISTIQUES FINALES
    ##############################
    
    def calculer_statistiques_finales(self):
        """Calcule les statistiques avec les boosts"""
        self.mettre_a_jour_carte_culture()
        
        stats = defaultdict(lambda: {
            'total_culture': 0,
            'boost_25': 0,
            'boost_50': 0,
            'boost_100': 0,
            'nb_batiments': 0
        })
        
        for bat in self.batiments_places:
            if bat['type'] == 'producteur':
                culture = self.calculer_culture_batiment(bat)
                boost = self.calculer_boost(bat, culture)
                
                prod_type = bat['production'] if bat['production'] else 'Rien'
                
                stats[prod_type]['total_culture'] += culture
                stats[prod_type]['nb_batiments'] += 1
                
                if boost == '25%':
                    stats[prod_type]['boost_25'] += 1
                elif boost == '50%':
                    stats[prod_type]['boost_50'] += 1
                elif boost == '100%':
                    stats[prod_type]['boost_100'] += 1
                
                # Mettre à jour le boost dans l'objet
                bat['culture_recue'] = culture
                bat['boost'] = boost
        
        return dict(stats)
    
    ##############################
    # EXÉCUTION PRINCIPALE
    ##############################
    
    def executer(self):
        """Exécute les trois phases"""
        self.reset()
        
        # Phase 1 : Placement compact
        self.phase1_placement_compact()
        
        # Phase 2 : Optimisation des boosts
        self.phase2_optimisation_boosts()
        
        # Phase 3 : Placement final
        self.phase3_placement_final()
        
        # Statistiques finales
        stats = self.calculer_statistiques_finales()
        
        return self.terrain, self.batiments_places, self.non_places, stats


##############################
# FONCTIONS DE GESTION EXCEL
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
    
    # Créer un workbook avec openpyxl
    wb = Workbook()
    wb.remove(wb.active)
    
    # Styles
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
    
    # Appliquer les couleurs
    for i in range(df_terrain_place.shape[0]):
        for j in range(df_terrain_place.shape[1]):
            cellule = str(df_terrain_place.iat[i, j])
            if cellule not in ['LIBRE', 'OCCUPE'] and '_' in cellule:
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
    
    # Onglet 3: Placements
    ws3 = wb.create_sheet('Placements')
    data_placement = []
    for bat in batiments_places:
        data_placement.append({
            'Nom': bat['nom'],
            'Type': bat['type'],
            'Position_X': bat['x'],
            'Position_Y': bat['y'],
            'Orientation': bat['orientation'],
            'Culture_recue': bat.get('culture_recue', 0),
            'Boost_atteint': bat.get('boost', '0%')
        })
    
    df_placement = pd.DataFrame(data_placement)
    for col_idx, col_name in enumerate(df_placement.columns, 1):
        ws3.cell(row=1, column=col_idx, value=col_name)
    for row_idx, row in df_placement.iterrows():
        for col_idx, value in enumerate(row, 1):
            ws3.cell(row=row_idx+2, column=col_idx, value=value)
    
    # Onglet 4: Non placés
    quantites_placees = defaultdict(int)
    for bat in batiments_places:
        quantites_placees[bat['nom']] += 1
    
    cases_libres_restantes = np.sum(terrain_place == 1)
    total_cases_batiments_non_places = 0
    total_cases_batiments_places = sum(b['longueur'] * b['largeur'] for b in batiments_places)
    
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
        
        ws4 = wb.create_sheet('Non_Places')
        df_non_places = pd.DataFrame(non_places)
        for col_idx, col_name in enumerate(df_non_places.columns, 1):
            ws4.cell(row=1, column=col_idx, value=col_name)
        for row_idx, row in df_non_places.iterrows():
            for col_idx, value in enumerate(row, 1):
                ws4.cell(row=row_idx+2, column=col_idx, value=value)
    
    # Onglet 5: Statistiques
    ws5 = wb.create_sheet('Statistiques')
    stats_data = []
    for prod, s in stats_culture.items():
        stats_data.append({
            'Type_Production': prod,
            'Culture_Total_Recue': s['total_culture'],
            'Boost_25_atteint': s['boost_25'],
            'Boost_50_atteint': s['boost_50'],
            'Boost_100_atteint': s['boost_100'],
            'Nombre_batiments': s['nb_batiments']
        })
    
    if stats_data:
        df_stats = pd.DataFrame(stats_data)
        for col_idx, col_name in enumerate(df_stats.columns, 1):
            ws5.cell(row=1, column=col_idx, value=col_name)
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
    
    if stats_culture:
        resume_data.extend([
            [''],
            ['TOTAL Culture par type de production'],
            ['Type', 'Culture totale recue']
        ])
        for prod, s in stats_culture.items():
            resume_data.append([prod, s['total_culture']])
    
    df_resume = pd.DataFrame(resume_data)
    for row in dataframe_to_rows(df_resume, index=False, header=False):
        ws6.append(row)
    
    wb.save(output)
    output.seek(0)
    return output


##############################
# INTERFACE STREAMLIT
##############################

st.title("🏗️ Optimiseur de Placement de Bâtiments (2 Phases)")
st.markdown("---")

with st.sidebar:
    st.header("📋 Instructions")
    st.markdown("""
    1. Préparez votre fichier Excel avec deux onglets
    2. Uploadez le fichier
    3. Lancez l'optimisation
    4. Téléchargez les résultats
    """)
    st.markdown("---")
    st.header("📁 Upload du fichier")
    uploaded_file = st.file_uploader(
        "Choisissez votre fichier Excel",
        type=['xlsx', 'xls']
    )

if uploaded_file is not None:
    try:
        with st.spinner("Lecture du fichier en cours..."):
            terrain, batiments = lire_fichier_excel(uploaded_file)
        
        if terrain is None or batiments is None:
            st.stop()
        
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("📊 Terrain original")
            st.dataframe(pd.DataFrame(terrain), use_container_width=True, height=400)
            cases_libres = np.sum(terrain == 1)
            cases_occupees = np.sum(terrain == 0)
            st.info(f"📌 Cases libres: {cases_libres} | Cases occupées: {cases_occupees}")
        
        with col2:
            st.subheader("🏢 Bâtiments à placer")
            df_batiments = pd.DataFrame(batiments)
            st.dataframe(df_batiments, use_container_width=True, height=400)
            total_batiments = sum(int(b['quantite']) for b in batiments)
            total_cases_necessaires = sum(int(b['quantite']) * b['longueur'] * b['largeur'] for b in batiments)
            st.info(f"📦 Total de bâtiments: {total_batiments}")
            st.info(f"📐 Cases nécessaires: {total_cases_necessaires}")
        
        st.markdown("---")
        
        if st.button("🚀 Lancer l'optimisation", type="primary", use_container_width=True):
            with st.spinner("Optimisation en cours..."):
                manager = PlacementManager(terrain, batiments)
                terrain_place, batiments_places, non_places, stats = manager.executer()
                
                st.session_state['terrain_place'] = terrain_place
                st.session_state['batiments_places'] = batiments_places
                st.session_state['stats'] = stats
                st.session_state['terrain_original'] = terrain
                st.session_state['batiments_complets'] = batiments
                st.session_state['batiments_non_places'] = non_places
                
                st.success(f"✅ Terminé! {len(batiments_places)} placés, {len(non_places)} non placés")
        
        if 'terrain_place' in st.session_state:
            st.markdown("---")
            st.header("📈 Résultats")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Placés", len(st.session_state['batiments_places']))
            with col2:
                st.metric("Non placés", len(st.session_state['batiments_non_places']))
            with col3:
                cases_libres = np.sum(st.session_state['terrain_place'] == 1)
                st.metric("Cases libres", cases_libres)
            with col4:
                total = len(st.session_state['batiments_places']) + len(st.session_state['batiments_non_places'])
                st.metric("Taux", f"{len(st.session_state['batiments_places'])/total*100:.1f}%" if total > 0 else "0%")
            
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
            
            st.subheader("🗺️ Visualisation")
            vis_terrain = np.zeros_like(st.session_state['terrain_original'], dtype=float)
            for bat in st.session_state['batiments_places']:
                valeur = 1 if bat['type'] == 'culturel' else 2
                for i in range(bat['x'], bat['x'] + bat['longueur']):
                    for j in range(bat['y'], bat['y'] + bat['largeur']):
                        vis_terrain[i, j] = valeur
            
            for i in range(vis_terrain.shape[0]):
                for j in range(vis_terrain.shape[1]):
                    if vis_terrain[i, j] == 0:
                        vis_terrain[i, j] = 3 if st.session_state['terrain_original'][i, j] == 1 else 4
            
            fig = go.Figure(data=go.Heatmap(
                z=vis_terrain,
                colorscale=[
                    [0, 'lightblue'], [0.33, 'lightgreen'],
                    [0.66, 'white'], [1, 'lightgray']
                ],
                showscale=False,
                text=[[str(st.session_state['terrain_place'][i, j])[:10] + '...' 
                       if len(str(st.session_state['terrain_place'][i, j])) > 10 
                       else str(st.session_state['terrain_place'][i, j]) 
                       for j in range(vis_terrain.shape[1])] 
                      for i in range(vis_terrain.shape[0])],
                hoverinfo='text'
            ))
            fig.update_layout(height=600)
            st.plotly_chart(fig, use_container_width=True)
            
            if st.session_state['batiments_non_places']:
                st.warning("⚠️ Bâtiments non placés!")
                non_places_count = defaultdict(int)
                for bat in st.session_state['batiments_non_places']:
                    non_places_count[bat['nom']] += 1
                st.dataframe(pd.DataFrame([
                    {'Nom': nom, 'Quantité': count}
                    for nom, count in non_places_count.items()
                ]))
            
            st.markdown("---")
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
        st.error(f"❌ Erreur: {str(e)}")
        st.exception(e)

else:
    st.info("👈 Veuillez uploader un fichier Excel")
    with st.expander("📝 Exemple de structure"):
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Terrain")
            st.dataframe(pd.DataFrame([
                [1,1,1,0,1],
                [1,0,1,1,1],
                [1,1,1,0,1],
                [0,1,1,1,1]
            ]))
        with col2:
            st.subheader("Bâtiments")
            st.dataframe(pd.DataFrame([
                ['Maison',2,2,1,'culturel',10,1,5,10,20,''],
                ['Ferme',3,2,2,'producteur',0,0,10,20,30,'Nourriture'],
            ], columns=['Nom','Longueur','Largeur','Quantité','Type',
                       'Culture','Rayonnement','Boost25%','Boost50%','Boost100%','Production']))

st.markdown("---")
st.markdown("🚀 **Stratégie en 3 phases :** Placement compact → Optimisation des boosts → Placement final")