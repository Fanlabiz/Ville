import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows
import io
import plotly.graph_objects as go
from copy import deepcopy
import random
from collections import deque

# Configuration de la page
st.set_page_config(
    page_title="Placement de Bâtiments",
    page_icon="🏗️",
    layout="wide"
)

##############################
# CLASSE DE PLACEMENT AVEC RECHERCHE TABOU
##############################
class PlacementBatiments:
    def __init__(self, terrain, batiments):
        self.terrain_original = terrain.copy()
        self.terrain = terrain.astype(object).copy()
        self.hauteur, self.largeur = terrain.shape
        self.batiments = batiments
        self.batiments_places = []
        self.carte_rayonnement = np.zeros((self.hauteur, self.largeur))
        self.production_priority = {'Guerison': 1, 'Nourriture': 2, 'Or': 3, '': 4}
        
        # Dictionnaire pour suivre les quantités placées
        self.quantites_placees = {}
        
        # Paramètres de la recherche tabou
        self.taille_liste_tabou = 50
        self.nb_iterations_max = 500
        self.nb_iterations_sans_amelioration_max = 100
        
        # Statistiques
        self.stats_placement = {
            'tentatives': 0,
            'placements_reussis': 0,
            'contraintes_bande_ignorees': 0,
            'iterations_tabou': 0
        }
    
    def est_case_libre(self, i, j):
        """Vérifie si une case est libre (valeur 1)"""
        return self.terrain[i, j] == 1
    
    def est_case_occupee_par_batiment(self, i, j):
        """Vérifie si une case est occupée par un bâtiment (string)"""
        return isinstance(self.terrain[i, j], str)
    
    def trouver_infos_originales(self, nom):
        """
        Trouve le rayonnement et la culture originaux d'un bâtiment
        """
        for b in self.batiments:
            if b['nom'] == nom:
                return b['rayonnement'], b['culture']
        return 0, 0
    
    def trouver_seuils_boost_originaux(self, nom):
        """
        Trouve les seuils de boost originaux d'un bâtiment
        """
        for b in self.batiments:
            if b['nom'] == nom:
                return b.get('boost_25', 0), b.get('boost_50', 0), b.get('boost_100', 0)
        return 0, 0, 0
    
    def calculer_zone_rayonnement(self, x, y, longueur, largeur, rayonnement, valeur_culture):
        """Calcule la zone de rayonnement autour d'un bâtiment culturel"""
        zone = []
        x_min = max(0, x - rayonnement)
        x_max = min(self.hauteur, x + longueur + rayonnement)
        y_min = max(0, y - rayonnement)
        y_max = min(self.largeur, y + largeur + rayonnement)
        
        for i in range(x_min, x_max):
            for j in range(y_min, y_max):
                if (i < x or i >= x + longueur or j < y or j >= y + largeur):
                    if self.est_case_libre(i, j) or self.est_case_occupee_par_batiment(i, j):
                        zone.append((i, j))
                        self.carte_rayonnement[i, j] += valeur_culture
        
        return zone
    
    def peut_placer_batiment(self, x, y, longueur, largeur):
        """Vérifie si un bâtiment peut être placé à la position (x,y)"""
        if x + longueur > self.hauteur or y + largeur > self.largeur:
            return False
        
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                if not self.est_case_libre(i, j):
                    return False
        return True
    
    def verifier_bande_1(self, x, y, longueur, largeur, strict=True):
        """Vérifie si le placement crée des bandes de largeur 1"""
        if not strict:
            return True
        
        if x == 0 or x + longueur == self.hauteur or y == 0 or y + largeur == self.largeur:
            if x > 0:
                bande_valide = False
                for j in range(y, y + largeur):
                    if x - 1 >= 0 and (not self.est_case_libre(x-1, j) or j == 0 or j == self.largeur-1):
                        bande_valide = True
                        break
                if not bande_valide:
                    return False
            
            if x + longueur < self.hauteur:
                bande_valide = False
                for j in range(y, y + largeur):
                    if x + longueur < self.hauteur and (not self.est_case_libre(x+longueur, j) or j == 0 or j == self.largeur-1):
                        bande_valide = True
                        break
                if not bande_valide:
                    return False
            
            if y > 0:
                bande_valide = False
                for i in range(x, x + longueur):
                    if y - 1 >= 0 and (not self.est_case_libre(i, y-1) or i == 0 or i == self.hauteur-1):
                        bande_valide = True
                        break
                if not bande_valide:
                    return False
            
            if y + largeur < self.largeur:
                bande_valide = False
                for i in range(x, x + longueur):
                    if y + largeur < self.largeur and (not self.est_case_libre(i, y+largeur) or i == 0 or i == self.hauteur-1):
                        bande_valide = True
                        break
                if not bande_valide:
                    return False
        
        return True
    
    def calculer_culture_pour_position(self, x, y, longueur, largeur):
        """Calcule la culture totale reçue par un bâtiment à une position donnée"""
        culture_totale = 0
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                culture_totale += self.carte_rayonnement[i, j]
        return culture_totale
    
    def trouver_meilleur_placement_culturel(self, batiment):
        """Trouve le meilleur placement pour un bâtiment culturel"""
        meilleur_score = -1
        meilleur_placement = None
        
        orientations = [(batiment['longueur'], batiment['largeur'])]
        if batiment['longueur'] != batiment['largeur']:
            orientations.append((batiment['largeur'], batiment['longueur']))
        
        for longueur, largeur in orientations:
            for i in range(self.hauteur - longueur + 1):
                for j in range(self.largeur - largeur + 1):
                    if self.peut_placer_batiment(i, j, longueur, largeur):
                        if not self.verifier_bande_1(i, j, longueur, largeur, strict=True):
                            continue
                        
                        zone_temp = []
                        for ii in range(max(0, i - batiment['rayonnement']), 
                                       min(self.hauteur, i + longueur + batiment['rayonnement'])):
                            for jj in range(max(0, j - batiment['rayonnement']), 
                                          min(self.largeur, j + largeur + batiment['rayonnement'])):
                                if (ii < i or ii >= i + longueur or jj < j or jj >= j + largeur):
                                    if self.est_case_libre(ii, jj) or self.est_case_occupee_par_batiment(ii, jj):
                                        zone_temp.append((ii, jj))
                        
                        producteurs_dans_zone = 0
                        for bat_place in self.batiments_places:
                            if bat_place['type'] == 'producteur':
                                for bi in range(bat_place['x'], bat_place['x'] + bat_place['longueur']):
                                    for bj in range(bat_place['y'], bat_place['y'] + bat_place['largeur']):
                                        if (bi, bj) in zone_temp:
                                            producteurs_dans_zone += 1
                                            break
                        
                        score = len(zone_temp) + producteurs_dans_zone * 100
                        
                        if score > meilleur_score:
                            meilleur_score = score
                            meilleur_placement = (i, j, longueur, largeur)
        
        if meilleur_placement is None:
            for longueur, largeur in orientations:
                for i in range(self.hauteur - longueur + 1):
                    for j in range(self.largeur - largeur + 1):
                        if self.peut_placer_batiment(i, j, longueur, largeur):
                            self.stats_placement['contraintes_bande_ignorees'] += 1
                            
                            zone_temp = []
                            for ii in range(max(0, i - batiment['rayonnement']), 
                                           min(self.hauteur, i + longueur + batiment['rayonnement'])):
                                for jj in range(max(0, j - batiment['rayonnement']), 
                                              min(self.largeur, j + largeur + batiment['rayonnement'])):
                                    if (ii < i or ii >= i + longueur or jj < j or jj >= j + largeur):
                                        if self.est_case_libre(ii, jj) or self.est_case_occupee_par_batiment(ii, jj):
                                            zone_temp.append((ii, jj))
                            
                            producteurs_dans_zone = 0
                            for bat_place in self.batiments_places:
                                if bat_place['type'] == 'producteur':
                                    for bi in range(bat_place['x'], bat_place['x'] + bat_place['longueur']):
                                        for bj in range(bat_place['y'], bat_place['y'] + bat_place['largeur']):
                                            if (bi, bj) in zone_temp:
                                                producteurs_dans_zone += 1
                                                break
                            
                            score = len(zone_temp) + producteurs_dans_zone * 100
                            
                            if score > meilleur_score:
                                meilleur_score = score
                                meilleur_placement = (i, j, longueur, largeur)
        
        return meilleur_placement
    
    def trouver_meilleur_placement_producteur(self, batiment):
        """Trouve le meilleur placement pour un bâtiment producteur"""
        meilleur_score = -1
        meilleur_placement = None
        meilleure_culture = -1
        
        orientations = [(batiment['longueur'], batiment['largeur'])]
        if batiment['longueur'] != batiment['largeur']:
            orientations.append((batiment['largeur'], batiment['longueur']))
        
        for longueur, largeur in orientations:
            for i in range(self.hauteur - longueur + 1):
                for j in range(self.largeur - largeur + 1):
                    if self.peut_placer_batiment(i, j, longueur, largeur):
                        if not self.verifier_bande_1(i, j, longueur, largeur, strict=True):
                            continue
                        
                        culture = self.calculer_culture_pour_position(i, j, longueur, largeur)
                        score = culture * 1000 + (self.hauteur - i) * 10 + (self.largeur - j)
                        
                        if culture > meilleure_culture or (culture == meilleure_culture and score > meilleur_score):
                            meilleure_culture = culture
                            meilleur_score = score
                            meilleur_placement = (i, j, longueur, largeur, culture)
        
        if meilleur_placement is None:
            for longueur, largeur in orientations:
                for i in range(self.hauteur - longueur + 1):
                    for j in range(self.largeur - largeur + 1):
                        if self.peut_placer_batiment(i, j, longueur, largeur):
                            self.stats_placement['contraintes_bande_ignorees'] += 1
                            
                            culture = self.calculer_culture_pour_position(i, j, longueur, largeur)
                            score = culture * 1000 + (self.hauteur - i) * 10 + (self.largeur - j)
                            
                            if culture > meilleure_culture or (culture == meilleure_culture and score > meilleur_score):
                                meilleure_culture = culture
                                meilleur_score = score
                                meilleur_placement = (i, j, longueur, largeur, culture)
        
        return meilleur_placement
    
    def placer_batiment(self, batiment, x, y, longueur, largeur, culture_recue=0):
        """
        Place un bâtiment sur le terrain
        """
        nom_batiment = batiment['nom']
        
        # Vérifier que nous n'avons pas déjà atteint la quantité maximale
        quantite_demandee = 0
        for b in self.batiments:
            if b['nom'] == nom_batiment:
                quantite_demandee = int(b['quantite'])
                break
        
        quantite_actuelle = self.quantites_placees.get(nom_batiment, 0)
        if quantite_actuelle >= quantite_demandee:
            return None  # Ne pas placer si on a déjà atteint la quantité
        
        # Placer le bâtiment
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                self.terrain[i, j] = f"{nom_batiment[:3]}_{i}_{j}"
        
        # Si c'est un bâtiment culturel, calculer sa zone de rayonnement
        if batiment['type'] == 'culturel':
            # Récupérer les infos originales du bâtiment
            rayonnement, culture = self.trouver_infos_originales(batiment['nom'])
            self.calculer_zone_rayonnement(x, y, longueur, largeur, rayonnement, culture)
        
        # Déterminer le boost atteint
        boost = '0%'
        if batiment['type'] == 'producteur' and culture_recue > 0:
            # Récupérer les seuils de boost depuis le dictionnaire batiment ou depuis les infos originales
            boost_25 = batiment.get('boost_25', 0)
            boost_50 = batiment.get('boost_50', 0)
            boost_100 = batiment.get('boost_100', 0)
            
            # Si les valeurs sont 0, essayer de les récupérer depuis les infos originales
            if boost_25 == 0 and boost_50 == 0 and boost_100 == 0:
                boost_25, boost_50, boost_100 = self.trouver_seuils_boost_originaux(batiment['nom'])
            
            if culture_recue >= boost_100 and boost_100 > 0:
                boost = '100%'
            elif culture_recue >= boost_50 and boost_50 > 0:
                boost = '50%'
            elif culture_recue >= boost_25 and boost_25 > 0:
                boost = '25%'
        
        # Enregistrer le bâtiment placé avec toutes les informations nécessaires
        batiment_place = {
            'nom': batiment['nom'],
            'type': batiment['type'],
            'x': x,
            'y': y,
            'longueur': longueur,
            'largeur': largeur,
            'orientation': 'horizontal' if longueur == batiment['longueur'] else 'vertical',
            'culture_recue': culture_recue,
            'boost': boost,
            'production': batiment['production'],
            # Garder les seuils de boost pour les calculs futurs
            'boost_25': batiment.get('boost_25', 0),
            'boost_50': batiment.get('boost_50', 0),
            'boost_100': batiment.get('boost_100', 0)
        }
        self.batiments_places.append(batiment_place)
        self.quantites_placees[nom_batiment] = quantite_actuelle + 1
        
        return batiment_place
    
    def retirer_batiment(self, index):
        """Retire un bâtiment du terrain (pour la recherche tabou)"""
        if index < 0 or index >= len(self.batiments_places):
            return None
        
        batiment = self.batiments_places.pop(index)
        
        # Mettre à jour les quantités placées
        nom = batiment['nom']
        self.quantites_placees[nom] = self.quantites_placees.get(nom, 1) - 1
        if self.quantites_placees[nom] <= 0:
            del self.quantites_placees[nom]
        
        # Effacer le bâtiment du terrain
        for i in range(batiment['x'], batiment['x'] + batiment['longueur']):
            for j in range(batiment['y'], batiment['y'] + batiment['largeur']):
                self.terrain[i, j] = 1
        
        # Recalculer la carte de rayonnement
        self.carte_rayonnement = np.zeros((self.hauteur, self.largeur))
        for b in self.batiments_places:
            if b['type'] == 'culturel':
                rayonnement, culture = self.trouver_infos_originales(b['nom'])
                self.calculer_zone_rayonnement(b['x'], b['y'], b['longueur'], b['largeur'],
                                              rayonnement, culture)
        
        return batiment
    
    def verifier_configuration_valide(self):
        """
        Vérifie que la configuration actuelle est valide (pas de doublons et quantités respectées)
        """
        # Vérifier les doublons de position
        positions = {}
        for bat in self.batiments_places:
            for i in range(bat['x'], bat['x'] + bat['longueur']):
                for j in range(bat['y'], bat['y'] + bat['largeur']):
                    pos = (i, j)
                    if pos in positions:
                        return False
                    positions[pos] = True
        
        # Vérifier les quantités
        quantites = {}
        for bat in self.batiments_places:
            nom = bat['nom']
            if nom not in quantites:
                quantites[nom] = 0
            quantites[nom] += 1
        
        for bat in self.batiments:
            nom = bat['nom']
            quantite_demandee = int(bat['quantite'])
            quantite_placee = quantites.get(nom, 0)
            if quantite_placee > quantite_demandee:
                return False
        
        return True
    
    def calculer_score_configuration(self):
        """
        Calcule un score pour la configuration actuelle
        Plus le score est élevé, meilleure est la configuration
        """
        score = 0
        
        # Pénalité pour les bâtiments non placés
        batiments_non_places = self.compter_batiments_non_places()
        score -= batiments_non_places * 1000
        
        # Bonus pour la culture reçue
        for batiment in self.batiments_places:
            if batiment['type'] == 'producteur':
                # Plus la culture reçue est élevée, mieux c'est
                score += batiment['culture_recue']
                
                # Bonus supplémentaire pour les boosts élevés
                if batiment['boost'] == '100%':
                    score += 500
                elif batiment['boost'] == '50%':
                    score += 200
                elif batiment['boost'] == '25%':
                    score += 50
        
        return score
    
    def compter_batiments_non_places(self):
        """Compte le nombre de bâtiments non placés"""
        total_non_places = 0
        for bat in self.batiments:
            nom = bat['nom']
            quantite_demandee = int(bat['quantite'])
            quantite_placee = self.quantites_placees.get(nom, 0)
            total_non_places += max(0, quantite_demandee - quantite_placee)
        
        return total_non_places
    
    def generer_voisinage(self, liste_tabou):
        """
        Génère les configurations voisines en déplaçant un bâtiment
        """
        voisins = []
        
        # Voisinage 1: Déplacer un bâtiment
        for i, batiment in enumerate(self.batiments_places):
            # Sauvegarder l'état
            x_old, y_old = batiment['x'], batiment['y']
            longueur, largeur = batiment['longueur'], batiment['largeur']
            
            # Retirer le bâtiment temporairement
            self.retirer_batiment(i)
            
            # Chercher de nouvelles positions
            for dx in range(-3, 4):
                for dy in range(-3, 4):
                    if dx == 0 and dy == 0:
                        continue
                    
                    x_new = x_old + dx
                    y_new = y_old + dy
                    
                    if x_new < 0 or y_new < 0 or x_new + longueur > self.hauteur or y_new + largeur > self.largeur:
                        continue
                    
                    if self.peut_placer_batiment(x_new, y_new, longueur, largeur):
                        # Vérifier si ce mouvement est tabou
                        mouvement = (batiment['nom'], x_old, y_old, x_new, y_new)
                        if mouvement in liste_tabou:
                            continue
                        
                        # Placer le bâtiment à la nouvelle position
                        culture = self.calculer_culture_pour_position(x_new, y_new, longueur, largeur)
                        resultat = self.placer_batiment(batiment, x_new, y_new, longueur, largeur, culture)
                        
                        if resultat is not None and self.verifier_configuration_valide():
                            # Calculer le score
                            score = self.calculer_score_configuration()
                            voisins.append((score, deepcopy(self), mouvement))
                        
                        # Revenir à l'état précédent
                        self.retirer_batiment(-1)  # Retire le dernier ajouté
            
            # Remettre le bâtiment à sa place
            self.placer_batiment(batiment, x_old, y_old, longueur, largeur, batiment['culture_recue'])
        
        # Trier les voisins par score décroissant
        voisins.sort(key=lambda x: -x[0])
        
        return voisins[:10]
    
    def executer_placement_initial(self):
        """
        Exécute le placement initial (sans recherche tabou)
        """
        # Séparer les bâtiments par type
        culturels = []
        producteurs = []
        
        for b in self.batiments:
            for _ in range(int(b['quantite'])):
                if b['type'] == 'culturel':
                    culturels.append(b.copy())
                else:
                    producteurs.append(b.copy())
        
        # Trier les culturels par taille (du plus grand au plus petit)
        culturels.sort(key=lambda x: (x['longueur'] * x['largeur'], x['longueur']), reverse=True)
        
        # Trier les producteurs par priorité de production
        producteurs.sort(key=lambda x: (self.production_priority.get(x['production'], 4), 
                                      -x['longueur'] * x['largeur']))
        
        st.write("### Phase 1: Placement initial")
        index_producteur = 0
        iteration = 0
        max_iterations = 200
        
        while (culturels or index_producteur < len(producteurs)) and iteration < max_iterations:
            iteration += 1
            
            if culturels:
                batiment = culturels.pop(0)
                placement = self.trouver_meilleur_placement_culturel(batiment)
                if placement:
                    x, y, longueur, largeur = placement
                    self.placer_batiment(batiment, x, y, longueur, largeur)
                    self.stats_placement['placements_reussis'] += 1
                else:
                    culturels.append(batiment)
            
            places_dans_cette_iteration = 0
            tentative_producteurs = 0
            while (index_producteur < len(producteurs) and 
                   places_dans_cette_iteration < 3 and 
                   tentative_producteurs < len(producteurs) - index_producteur):
                
                batiment = producteurs[index_producteur]
                placement = self.trouver_meilleur_placement_producteur(batiment)
                
                if placement:
                    x, y, longueur, largeur, culture = placement
                    self.placer_batiment(batiment, x, y, longueur, largeur, culture)
                    index_producteur += 1
                    places_dans_cette_iteration += 1
                    self.stats_placement['placements_reussis'] += 1
                else:
                    index_producteur += 1
                    tentative_producteurs += 1
        
        return self.terrain, self.batiments_places
    
    def executer_recherche_tabou(self):
        """
        Exécute la recherche tabou pour améliorer le placement
        """
        st.write("### Phase 2: Recherche tabou")
        
        # Initialisation
        meilleure_config = deepcopy(self)
        meilleur_score = self.calculer_score_configuration()
        configuration_courante = self
        
        liste_tabou = deque(maxlen=self.taille_liste_tabou)
        
        iterations_sans_amelioration = 0
        barre_progression = st.progress(0)
        
        for iteration in range(self.nb_iterations_max):
            self.stats_placement['iterations_tabou'] += 1
            
            # Mettre à jour la barre de progression
            barre_progression.progress((iteration + 1) / self.nb_iterations_max)
            
            # Générer les voisins
            voisins = configuration_courante.generer_voisinage(liste_tabou)
            
            if not voisins:
                break
            
            # Prendre le meilleur voisin non tabou (ou avec aspiration)
            meilleur_voisin = None
            meilleur_score_voisin = -float('inf')
            
            for score_voisin, config_voisine, mouvement in voisins:
                if mouvement not in liste_tabou or score_voisin > meilleur_score:
                    if score_voisin > meilleur_score_voisin:
                        meilleur_score_voisin = score_voisin
                        meilleur_voisin = (config_voisine, mouvement)
                    break
            
            if meilleur_voisin is None:
                break
            
            # Mettre à jour la configuration courante
            configuration_courante, mouvement = meilleur_voisin
            liste_tabou.append(mouvement)
            
            # Mettre à jour la meilleure configuration
            if meilleur_score_voisin > meilleur_score:
                meilleur_score = meilleur_score_voisin
                meilleure_config = configuration_courante
                iterations_sans_amelioration = 0
                st.write(f"  Iteration {iteration+1}: Nouveau meilleur score = {meilleur_score}")
            else:
                iterations_sans_amelioration += 1
            
            # Critère d'arrêt
            if iterations_sans_amelioration >= self.nb_iterations_sans_amelioration_max:
                st.write(f"  Arrêt après {iteration+1} itérations (pas d'amélioration)")
                break
        
        barre_progression.empty()
        
        # Mettre à jour avec la meilleure configuration trouvée
        self.terrain = meilleure_config.terrain
        self.batiments_places = meilleure_config.batiments_places
        self.carte_rayonnement = meilleure_config.carte_rayonnement
        self.quantites_placees = meilleure_config.quantites_placees
        
        return self.terrain, self.batiments_places
    
    def executer_placement(self):
        """
        Exécute l'algorithme complet avec recherche tabou
        """
        # Phase 1: Placement initial
        self.executer_placement_initial()
        
        # Afficher les statistiques initiales
        non_places_init = self.compter_batiments_non_places()
        st.write(f"Bâtiments non placés après phase 1: {non_places_init}")
        
        # Phase 2: Recherche tabou
        self.executer_recherche_tabou()
        
        # Statistiques finales
        non_places_final = self.compter_batiments_non_places()
        st.write(f"Bâtiments non placés après recherche tabou: {non_places_final}")
        
        st.write("### Statistiques de placement")
        st.write(f"Placements réussis: {self.stats_placement['placements_reussis']}")
        st.write(f"Contraintes de bande ignorées: {self.stats_placement['contraintes_bande_ignorees']}")
        st.write(f"Itérations de recherche tabou: {self.stats_placement['iterations_tabou']}")
        
        return self.terrain, self.batiments_places
    
    def calculer_statistiques(self):
        """Calcule les statistiques finales"""
        stats = {}
        
        for batiment in self.batiments_places:
            if batiment['type'] == 'producteur':
                prod_type = batiment['production'] if batiment['production'] else 'Rien'
                
                if prod_type not in stats:
                    stats[prod_type] = {
                        'total_culture': 0,
                        'boost_25': 0,
                        'boost_50': 0,
                        'boost_100': 0,
                        'nb_batiments': 0
                    }
                
                stats[prod_type]['total_culture'] += batiment['culture_recue']
                stats[prod_type]['nb_batiments'] += 1
                
                if batiment['boost'] == '25%':
                    stats[prod_type]['boost_25'] += 1
                elif batiment['boost'] == '50%':
                    stats[prod_type]['boost_50'] += 1
                elif batiment['boost'] == '100%':
                    stats[prod_type]['boost_100'] += 1
        
        return stats


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
    
    # Afficher les noms de colonnes pour déboguer
    st.write("Colonnes trouvées dans le fichier:", list(df_batiments.columns))
    
    # Normaliser les noms de colonnes (enlever les espaces, gérer les accents, etc.)
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

def generer_fichier_resultat(terrain_original, terrain_place, batiments_places, stats_culture, tous_les_batiments):
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
    
    non_places = []
    for bat in tous_les_batiments:
        nom = bat['nom']
        quantite_demandee = int(bat['quantite'])
        quantite_placee = quantites_placees.get(nom, 0)
        
        if quantite_placee < quantite_demandee:
            reste = quantite_demandee - quantite_placee
            cases_batiment = bat['longueur'] * bat['largeur'] * reste
            total_cases_batiments_non_places += cases_batiment
            
            non_places.append({
                'Nom': nom,
                'Type': bat['type'],
                'Longueur': bat['longueur'],
                'Largeur': bat['largeur'],
                'Quantite_demandee': quantite_demandee,
                'Quantite_placee': quantite_placee,
                'Reste_a_placer': reste,
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
    
    # Calculer les statistiques spécifiques pour Guérison, Nourriture et Or
    stats_guerison = stats_culture.get('Guerison', {'total_culture': 0, 'boost_25': 0, 'boost_50': 0, 'boost_100': 0, 'nb_batiments': 0})
    stats_nourriture = stats_culture.get('Nourriture', {'total_culture': 0, 'boost_25': 0, 'boost_50': 0, 'boost_100': 0, 'nb_batiments': 0})
    stats_or = stats_culture.get('Or', {'total_culture': 0, 'boost_25': 0, 'boost_50': 0, 'boost_100': 0, 'nb_batiments': 0})
    
    stats_data = [
        {
            'Type_Production': 'Guerison',
            'Culture_Total_Recue': stats_guerison['total_culture'],
            'Boost_25_atteint': stats_guerison['boost_25'],
            'Boost_50_atteint': stats_guerison['boost_50'],
            'Boost_100_atteint': stats_guerison['boost_100'],
            'Nombre_batiments': stats_guerison['nb_batiments']
        },
        {
            'Type_Production': 'Nourriture',
            'Culture_Total_Recue': stats_nourriture['total_culture'],
            'Boost_25_atteint': stats_nourriture['boost_25'],
            'Boost_50_atteint': stats_nourriture['boost_50'],
            'Boost_100_atteint': stats_nourriture['boost_100'],
            'Nombre_batiments': stats_nourriture['nb_batiments']
        },
        {
            'Type_Production': 'Or',
            'Culture_Total_Recue': stats_or['total_culture'],
            'Boost_25_atteint': stats_or['boost_25'],
            'Boost_50_atteint': stats_or['boost_50'],
            'Boost_100_atteint': stats_or['boost_100'],
            'Nombre_batiments': stats_or['nb_batiments']
        }
    ]
    
    # Ajouter les autres types de production
    for prod, stats in stats_culture.items():
        if prod not in ['Guerison', 'Nourriture', 'Or']:
            stats_data.append({
                'Type_Production': prod,
                'Culture_Total_Recue': stats['total_culture'],
                'Boost_25_atteint': stats['boost_25'],
                'Boost_50_atteint': stats['boost_50'],
                'Boost_100_atteint': stats['boost_100'],
                'Nombre_batiments': stats['nb_batiments']
            })
    
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
        ['Cases restantes inutilisables', np.sum(terrain_original == 0) - (np.sum(terrain_place == 0) - total_cases_batiments_places)],
        [''],
        ['Bâtiments placés', len(batiments_places)],
        ['Bâtiments non placés', sum(b['Reste_a_placer'] for b in non_places if b['Nom'] != 'TOTAL') if non_places else 0],
        [''],
        ['Cases nécessaires pour les bâtiments non placés', total_cases_batiments_non_places],
        ['Suffisamment de cases libres ?', 'OUI' if cases_libres_restantes >= total_cases_batiments_non_places else 'NON'],
    ]
    
    # Ajouter les totaux de culture par type
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
st.title("🏗️ Optimiseur de Placement de Bâtiments (Recherche Tabou Améliorée)")
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
            st.dataframe(pd.DataFrame(terrain), use_container_width=True)
            
            # Statistiques du terrain
            cases_libres = np.sum(terrain == 1)
            cases_occupees = np.sum(terrain == 0)
            st.info(f"📌 Cases libres: {cases_libres} | Cases occupées: {cases_occupees}")
        
        with col2:
            st.subheader("🏢 Bâtiments à placer")
            df_batiments = pd.DataFrame(batiments)
            st.dataframe(df_batiments, use_container_width=True)
            
            # Résumé des bâtiments
            total_batiments = sum(int(b['quantite']) for b in batiments)
            total_cases_necessaires = sum(int(b['quantite']) * b['longueur'] * b['largeur'] for b in batiments)
            st.info(f"📦 Total de bâtiments à placer: {total_batiments}")
            st.info(f"📐 Cases totales nécessaires: {total_cases_necessaires}")
        
        st.markdown("---")
        
        # Bouton pour lancer l'optimisation
        if st.button("🚀 Lancer l'optimisation", type="primary", use_container_width=True):
            with st.spinner("Optimisation en cours... Cela peut prendre quelques instants."):
                # Création de l'instance de placement
                placement = PlacementBatiments(terrain, batiments)
                
                # Exécution de l'algorithme
                terrain_place, batiments_places = placement.executer_placement()
                
                # Calcul des statistiques
                stats = placement.calculer_statistiques()
                
                # Sauvegarde dans la session
                st.session_state['terrain_place'] = terrain_place
                st.session_state['batiments_places'] = batiments_places
                st.session_state['stats'] = stats
                st.session_state['terrain_original'] = terrain
                st.session_state['batiments_complets'] = batiments
                
                st.success("✅ Optimisation terminée avec succès!")
        
        # Affichage des résultats si disponibles
        if 'terrain_place' in st.session_state:
            st.markdown("---")
            st.header("📈 Résultats de l'optimisation")
            
            # Statistiques
            st.subheader("📊 Statistiques de production")
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
            st.subheader("📋 Détail des placements")
            placements_df = pd.DataFrame(st.session_state['batiments_places'])
            st.dataframe(placements_df, use_container_width=True)
            
            # Vérifier s'il reste des bâtiments non placés
            quantites_placees = {}
            for bat in st.session_state['batiments_places']:
                if bat['nom'] not in quantites_placees:
                    quantites_placees[bat['nom']] = 0
                quantites_placees[bat['nom']] += 1
            
            non_places = []
            cases_necessaires_total = 0
            for bat in st.session_state['batiments_complets']:
                nom = bat['nom']
                quantite_demandee = int(bat['quantite'])
                quantite_placee = quantites_placees.get(nom, 0)
                
                if quantite_placee < quantite_demandee:
                    reste = quantite_demandee - quantite_placee
                    cases_necessaires = bat['longueur'] * bat['largeur'] * reste
                    cases_necessaires_total += cases_necessaires
                    non_places.append({
                        'Nom': nom,
                        'Reste à placer': reste,
                        'Cases nécessaires': cases_necessaires
                    })
            
            if non_places:
                st.warning("⚠️ Certains bâtiments n'ont pas pu être placés!")
                st.dataframe(pd.DataFrame(non_places))
                
                # Afficher le résumé des cases
                cases_libres_restantes = np.sum(st.session_state['terrain_place'] == 1)
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Cases libres restantes", cases_libres_restantes)
                with col2:
                    st.metric("Cases nécessaires", cases_necessaires_total)
                with col3:
                    st.metric("Suffisant", "✅ OUI" if cases_libres_restantes >= cases_necessaires_total else "❌ NON")
            
            # Bouton de téléchargement
            st.markdown("---")
            
            # Génération du fichier de résultats
            output_file = generer_fichier_resultat(
                st.session_state['terrain_original'],
                st.session_state['terrain_place'],
                st.session_state['batiments_places'],
                st.session_state['stats'],
                st.session_state['batiments_complets']
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
st.markdown("🚀 Application développée avec l'algorithme de recherche tabou amélioré (avec gestion stricte des quantités)")