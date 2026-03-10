import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import plotly.graph_objects as go
from collections import defaultdict
from copy import deepcopy

# Configuration de la page
st.set_page_config(
    page_title="Placement de Bâtiments",
    page_icon="🏗️",
    layout="wide"
)

##############################
# GESTIONNAIRE DE PLACEMENT AVEC BACKTRACKING INTELLIGENT
##############################
class PlacementManager:
    def __init__(self, terrain, batiments):
        self.terrain_original = terrain.copy()
        self.hauteur, self.largeur = terrain.shape
        self.batiments = batiments
        
        # Statistiques
        self.stats = {
            'placements': 0,
            'retours_arriere': 0,
            'verifications_espace': 0
        }
        
        # Initialiser
        self.reset()
    
    def reset(self):
        """Réinitialise l'état"""
        self.terrain = self.terrain_original.astype(object).copy()
        self.batiments_places = []
        self.carte_culture = np.zeros((self.hauteur, self.largeur))
        self.historique_placements = []  # Pour le backtracking
    
    ##############################
    # FONCTIONS UTILITAIRES
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
                        # Calculer un score (proche des bords pour compacité)
                        score = -(x + y)  # Plus proche du bord = meilleur score
                        placements.append((x, y, h, l, orientation, score))
        
        # Trier par score (meilleur d'abord)
        placements.sort(key=lambda p: p[5])
        return placements
    
    def trouver_plus_grand_batiment_restant(self, batiments_restants):
        """Trouve les dimensions du plus grand bâtiment restant (en surface)"""
        if not batiments_restants:
            return 0, 0
        
        # Trouver la plus grande surface
        max_surface = 0
        max_h, max_l = 0, 0
        
        for bat in batiments_restants:
            surface = bat['longueur'] * bat['largeur']
            if surface > max_surface:
                max_surface = surface
                max_h, max_l = bat['longueur'], bat['largeur']
        
        return max_h, max_l
    
    def verifier_espace_pour_plus_grand(self, batiments_restants):
        """
        Vérifie s'il reste assez d'espace pour placer le plus grand bâtiment restant
        Retourne True si possible, False sinon
        """
        if not batiments_restants:
            return True
        
        self.stats['verifications_espace'] += 1
        
        # Dimensions du plus grand bâtiment restant
        max_h, max_l = self.trouver_plus_grand_batiment_restant(batiments_restants)
        
        # Chercher un espace qui peut contenir ce bâtiment
        for orientation in [(max_h, max_l), (max_l, max_h)]:
            h, l = orientation
            for x in range(self.hauteur - h + 1):
                for y in range(self.largeur - l + 1):
                    if self.est_libre(x, y, h, l):
                        return True  # Trouvé un espace
        
        return False  # Pas d'espace trouvé
    
    def calculer_culture_pour_position(self, x, y, h, l):
        """Calcule la culture reçue par un bâtiment à une position"""
        culture = 0
        for i in range(x, x + h):
            for j in range(y, y + l):
                culture += self.carte_culture[i, j]
        return culture
    
    def ajouter_rayonnement_culturel(self, batiment, x, y, h, l):
        """Ajoute le rayonnement d'un bâtiment culturel"""
        # Chercher les infos originales
        for b_orig in self.batiments:
            if b_orig['nom'] == batiment['nom']:
                rayonnement = b_orig.get('rayonnement', 0)
                culture = b_orig.get('culture', 0)
                
                if rayonnement > 0 and culture > 0:
                    x_min = max(0, x - rayonnement)
                    x_max = min(self.hauteur, x + h + rayonnement)
                    y_min = max(0, y - rayonnement)
                    y_max = min(self.largeur, y + l + rayonnement)
                    
                    for i in range(x_min, x_max):
                        for j in range(y_min, y_max):
                            if (i < x or i >= x + h or j < y or j >= y + l):
                                self.carte_culture[i, j] += culture
                break
    
    def placer_batiment(self, batiment, x, y, h, l, orientation):
        """Place un bâtiment sur le terrain"""
        # Marquer les cases
        for i in range(x, x + h):
            for j in range(y, y + l):
                self.terrain[i, j] = f"{batiment['nom'][:3]}_{i}_{j}"
        
        # Calculer la culture reçue
        culture_recue = self.calculer_culture_pour_position(x, y, h, l)
        
        # Ajouter le rayonnement si culturel
        if batiment['type'] == 'culturel':
            self.ajouter_rayonnement_culturel(batiment, x, y, h, l)
        
        # Créer l'objet bâtiment placé
        bat_place = {
            'nom': batiment['nom'],
            'type': batiment['type'],
            'x': x,
            'y': y,
            'longueur': h,
            'largeur': l,
            'orientation': orientation,
            'culture_recue': culture_recue,
            'production': batiment.get('production', '')
        }
        
        # Calculer le boost plus tard (à la fin)
        bat_place['boost'] = '0%'
        
        self.batiments_places.append(bat_place)
        self.historique_placements.append(bat_place)
        self.stats['placements'] += 1
        
        return bat_place
    
    def retirer_dernier_batiment(self):
        """Retire le dernier bâtiment placé (backtracking)"""
        if not self.batiments_places:
            return None
        
        bat = self.batiments_places.pop()
        self.historique_placements.pop()
        
        # Effacer du terrain
        for i in range(bat['x'], bat['x'] + bat['longueur']):
            for j in range(bat['y'], bat['y'] + bat['largeur']):
                self.terrain[i, j] = 1
        
        # Recalculer toute la carte de culture
        self.carte_culture = np.zeros((self.hauteur, self.largeur))
        for b in self.batiments_places:
            if b['type'] == 'culturel':
                self.ajouter_rayonnement_culturel(b, b['x'], b['y'], b['longueur'], b['largeur'])
        
        self.stats['retours_arriere'] += 1
        return bat
    
    ##############################
    # ORDRE DE PLACEMENT
    ##############################
    
    def preparer_ordre_placement(self):
        """
        Prépare l'ordre de placement des bâtiments:
        1. D'abord tous les neutres
        2. Ensuite alternance culturels/producteurs
        """
        # Compter les quantités
        tous_les_batiments = []
        for b in self.batiments:
            for _ in range(int(b['quantite'])):
                tous_les_batiments.append(b.copy())
        
        # Séparer par type
        neutres = [b for b in tous_les_batiments if b['type'] not in ['culturel', 'producteur']]
        culturels = [b for b in tous_les_batiments if b['type'] == 'culturel']
        producteurs = [b for b in tous_les_batiments if b['type'] == 'producteur']
        
        # Trier par taille décroissante pour chaque groupe
        neutres.sort(key=lambda x: -(x['longueur'] * x['largeur']))
        culturels.sort(key=lambda x: -(x['longueur'] * x['largeur']))
        producteurs.sort(key=lambda x: -(x['longueur'] * x['largeur']))
        
        # Construire l'ordre final
        ordre = []
        
        # 1. Ajouter tous les neutres
        ordre.extend(neutres)
        
        # 2. Alterner culturels et producteurs
        i, j = 0, 0
        while i < len(culturels) or j < len(producteurs):
            if i < len(culturels):
                ordre.append(culturels[i])
                i += 1
            if j < len(producteurs):
                ordre.append(producteurs[j])
                j += 1
        
        return ordre
    
    ##############################
    # ALGORITHME PRINCIPAL
    ##############################
    
    def executer_placement(self):
        """Exécute l'algorithme de placement avec backtracking"""
        st.write("### Algorithme de placement")
        
        # Préparer l'ordre de placement
        ordre_placement = self.preparer_ordre_placement()
        total = len(ordre_placement)
        
        st.write(f"📋 Ordre de placement: {total} bâtiments")
        st.write(f"   - Neutres d'abord, puis alternance culturels/producteurs")
        
        # Barre de progression
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        index = 0
        pile_retour = []  # Pour backtracking
        
        while index < len(ordre_placement):
            batiment = ordre_placement[index]
            status_text.text(f"Placement {index + 1}/{total}: {batiment['nom']}...")
            
            # Bâtiments restants (après celui-ci)
            restants = ordre_placement[index + 1:]
            
            # Trouver tous les placements possibles
            placements = self.trouver_tous_placements(batiment)
            
            placement_reussi = False
            
            for placement in placements:
                x, y, h, l, orientation, _ = placement
                
                # Placer temporairement
                self.placer_batiment(batiment, x, y, h, l, orientation)
                
                # Vérifier s'il reste assez d'espace pour le plus grand restant
                if self.verifier_espace_pour_plus_grand(restants):
                    placement_reussi = True
                    pile_retour.append(index)  # Sauvegarder pour backtracking
                    break  # Placement valide
                else:
                    # Annuler ce placement
                    self.retirer_dernier_batiment()
            
            if placement_reussi:
                # Passer au suivant
                index += 1
                progress_bar.progress(index / total)
            else:
                # Aucun placement valide trouvé → backtracking
                if not pile_retour:
                    st.warning("❌ Impossible de continuer - backtracking impossible")
                    break
                
                # Revenir au dernier placement réussi
                dernier_index = pile_retour.pop()
                while index > dernier_index:
                    self.retirer_dernier_batiment()
                    index -= 1
                
                # Réessayer avec le bâtiment courant
                st.write(f"↩️ Backtracking - retour au bâtiment {index}")
        
        progress_bar.empty()
        status_text.empty()
        
        st.write(f"✅ Placement terminé: {len(self.batiments_places)} bâtiments placés")
        st.write(f"📊 Statistiques: {self.stats['placements']} placements, {self.stats['retours_arriere']} retours arrière, {self.stats['verifications_espace']} vérifications d'espace")
        
        # Identifier les non placés
        noms_places = {b['nom'] for b in self.batiments_places}
        non_places = []
        
        for bat in self.batiments:
            count_place = sum(1 for b in self.batiments_places if b['nom'] == bat['nom'])
            for _ in range(int(bat['quantite']) - count_place):
                non_places.append(bat)
        
        return self.terrain, self.batiments_places, non_places
    
    ##############################
    # CALCUL DES BOOSTS FINAUX
    ##############################
    
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
    
    def calculer_statistiques_finales(self):
        """Calcule les statistiques avec les boosts"""
        stats = defaultdict(lambda: {
            'total_culture': 0,
            'boost_25': 0,
            'boost_50': 0,
            'boost_100': 0,
            'nb_batiments': 0
        })
        
        for bat in self.batiments_places:
            if bat['type'] == 'producteur':
                culture = bat['culture_recue']
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
                
                # Mettre à jour le boost
                bat['boost'] = boost
        
        return dict(stats)


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

st.title("🏗️ Optimiseur de Placement avec Backtracking Intelligent")
st.markdown("---")

with st.sidebar:
    st.header("📋 Stratégie")
    st.markdown("""
    **Ordre de placement :**
    1. Tous les bâtiments neutres d'abord
    2. Alternance culturels/producteurs
    
    **Règle de placement :**
    - Après chaque placement, vérifier qu'il reste assez d'espace pour le plus grand bâtiment restant
    - Si impossible → backtracking
    
    **Avantages :**
    - Évite les blocages prématurés
    - Optimise l'utilisation de l'espace
    - Garantit qu'on ne bloque pas les grands bâtiments
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
                terrain_place, batiments_places, non_places = manager.executer_placement()
                stats = manager.calculer_statistiques_finales()
                
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
                ['Rocher',1,1,1,'neutre',0,0,0,0,0,'']
            ], columns=['Nom','Longueur','Largeur','Quantité','Type',
                       'Culture','Rayonnement','Boost25%','Boost50%','Boost100%','Production']))

st.markdown("---")
st.markdown("🚀 **Algorithme avec backtracking intelligent** - Garantit qu'on ne bloque jamais le plus grand bâtiment restant")