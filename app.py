import pandas as pd
import numpy as np
from typing import List, Dict, Tuple, Optional
import streamlit as st
from io import BytesIO
import openpyxl
from copy import deepcopy

class Batiment:
    def __init__(self, nom, longueur, largeur, quantite, type_batiment, culture, rayonnement, 
                 boost_25, boost_50, boost_100, production):
        self.nom = nom
        self.longueur = int(longueur)
        self.largeur = int(largeur)
        self.quantite = int(quantite)
        self.type = type_batiment.lower()
        self.culture = float(culture) if culture else 0
        self.rayonnement = int(rayonnement) if rayonnement else 0
        self.boost_25 = float(boost_25) if boost_25 else float('inf')
        self.boost_50 = float(boost_50) if boost_50 else float('inf')
        self.boost_100 = float(boost_100) if boost_100 else float('inf')
        self.production = production if production else ""
        
    def get_dimensions(self, orientation='horizontal'):
        if orientation == 'horizontal':
            return self.longueur, self.largeur
        else:
            return self.largeur, self.longueur

class PlacementBatiments:
    def __init__(self, terrain_df, batiments_df):
        self.terrain = terrain_df.values
        self.hauteur, self.largeur = self.terrain.shape
        self.batiments_originaux = self._charger_batiments(batiments_df)
        self.batiments_a_placer = self._creer_liste_batiments()
        self.batiments_places = []
        self.batiments_non_places = []
        self.culture_totale = 0
        self.culture_par_type = {'guerison': 0, 'nourriture': 0, 'or': 0}
        self.seuils_atteints = {'25%': 0, '50%': 0, '100%': 0}
        
    def _charger_batiments(self, batiments_df):
        batiments = []
        for _, row in batiments_df.iterrows():
            batiment = Batiment(
                row['Nom'], row['Longueur'], row['Largeur'], row['Quantite'],
                row['Type'], row['Culture'], row['Rayonnement'],
                row['Boost 25%'], row['Boost 50%'], row['Boost 100%'], row['Production']
            )
            batiments.append(batiment)
        return batiments
    
    def _creer_liste_batiments(self):
        liste = []
        for batiment in self.batiments_originaux:
            for _ in range(batiment.quantite):
                liste.append(batiment)
        return liste
    
    def case_est_libre(self, x, y):
        if 0 <= x < self.hauteur and 0 <= y < self.largeur:
            return self.terrain[x, y] == 1
        return False
    
    def peut_placer_batiment(self, batiment, x, y, orientation):
        longueur, largeur = batiment.get_dimensions(orientation)
        
        if x + longueur > self.hauteur or y + largeur > self.largeur:
            return False
            
        for i in range(longueur):
            for j in range(largeur):
                if not self.case_est_libre(x + i, y + j):
                    return False
        return True
    
    def placer_batiment(self, batiment, x, y, orientation):
        longueur, largeur = batiment.get_dimensions(orientation)
        
        for i in range(longueur):
            for j in range(largeur):
                self.terrain[x + i, y + j] = -1  # -1 pour indiquer un bâtiment placé
                
        self.batiments_places.append({
            'nom': batiment.nom,
            'x': x,
            'y': y,
            'orientation': orientation,
            'longueur': longueur,
            'largeur': largeur,
            'type': batiment.type,
            'culture': batiment.culture,
            'rayonnement': batiment.rayonnement,
            'production': batiment.production
        })
    
    def trouver_emplacement(self, batiment):
        """Trouve un emplacement valide pour un bâtiment"""
        orientations = ['horizontal', 'vertical']
        
        for orientation in orientations:
            longueur, largeur = batiment.get_dimensions(orientation)
            for x in range(self.hauteur - longueur + 1):
                for y in range(self.largeur - largeur + 1):
                    if self.peut_placer_batiment(batiment, x, y, orientation):
                        return x, y, orientation
        return None, None, None
    
    def calculer_culture_recue(self, batiment_place):
        """Calcule la culture reçue par un bâtiment producteur"""
        culture_recue = 0
        x, y = batiment_place['x'], batiment_place['y']
        longueur, largeur = batiment_place['longueur'], batiment_place['largeur']
        
        for batiment_culturel in self.batiments_places:
            if batiment_culturel['type'] == 'culturel':
                cx, cy = batiment_culturel['x'], batiment_culturel['y']
                cl, clargeur = batiment_culturel['longueur'], batiment_culturel['largeur']
                rayon = batiment_culturel['rayonnement']
                
                # Vérifier si le bâtiment producteur est dans la zone de rayonnement
                for i in range(longueur):
                    for j in range(largeur):
                        bat_x, bat_y = x + i, y + j
                        
                        # Vérifier les cases autour du bâtiment culturel
                        for ci in range(max(0, cx - rayon), min(self.hauteur, cx + cl + rayon)):
                            for cj in range(max(0, cy - rayon), min(self.largeur, cy + clargeur + rayon)):
                                if (bat_x == ci and bat_y == cj):
                                    culture_recue += batiment_culturel['culture']
                                    break
                            else:
                                continue
                            break
                        else:
                            continue
                        break
                    else:
                        continue
                    break
                    
        return culture_recue
    
    def calculer_boost(self, culture_recue):
        """Détermine le niveau de boost atteint"""
        if culture_recue >= self.seuils['100%']:
            return 100
        elif culture_recue >= self.seuils['50%']:
            return 50
        elif culture_recue >= self.seuils['25%']:
            return 25
        else:
            return 0
    
    def trouver_plus_grand_batiment_restant(self):
        """Trouve les dimensions du plus grand bâtiment non placé"""
        if not self.batiments_a_placer:
            return 0, 0
        
        max_area = 0
        max_dim = (0, 0)
        for batiment in self.batiments_a_placer:
            area = batiment.longueur * batiment.largeur
            if area > max_area:
                max_area = area
                max_dim = (max(batiment.longueur, batiment.largeur), 
                          min(batiment.longueur, batiment.largeur))
        return max_dim
    
    def reste_assez_de_place(self, batiment_a_placer):
        """Vérifie s'il reste assez de place pour le plus grand bâtiment restant"""
        plus_grand_l, plus_grand_largeur = self.trouver_plus_grand_batiment_restant()
        if plus_grand_l == 0:
            return True
            
        # Compter les cases libres contiguës
        cases_libres = np.sum(self.terrain == 1)
        
        # Vérification simplifiée : assez de cases libres en nombre
        if cases_libres < plus_grand_l * plus_grand_largeur:
            return False
            
        return True
    
    def placer_batiments_ordre(self):
        """Place les bâtiments selon l'ordre spécifié"""
        
        # Séparer les bâtiments par type
        batiments_culturels = [b for b in self.batiments_a_placer if b.type == 'culturel']
        batiments_producteurs = [b for b in self.batiments_a_placer if b.type == 'producteur']
        batiments_neutres = [b for b in self.batiments_a_placer if b.type not in ['culturel', 'producteur']]
        
        # Placer d'abord les bâtiments neutres sur les bords
        for batiment in batiments_neutres:
            self._placer_batiment_avec_backtracking(batiment, bord_prioritaire=True)
        
        # Placer en alternance culturels et producteurs
        i, j = 0, 0
        while i < len(batiments_culturels) or j < len(batiments_producteurs):
            if i < len(batiments_culturels):
                self._placer_batiment_avec_backtracking(batiments_culturels[i])
                i += 1
            
            if j < len(batiments_producteurs):
                self._placer_batiment_avec_backtracking(batiments_producteurs[j])
                j += 1
    
    def _placer_batiment_avec_backtracking(self, batiment, bord_prioritaire=False):
        """Place un bâtiment avec backtracking si nécessaire"""
        
        # Essayer de placer le bâtiment
        x, y, orientation = self.trouver_emplacement_optimise(batiment, bord_prioritaire)
        
        if x is not None:
            # Vérifier qu'il reste assez de place pour les futurs bâtiments
            self.placer_batiment(batiment, x, y, orientation)
            
            # Sauvegarder l'état avant de continuer
            terrain_sauvegarde = self.terrain.copy()
            
            if not self.reste_assez_de_place(batiment):
                # Backtracking : enlever le dernier bâtiment et essayer autre chose
                self.terrain = terrain_sauvegarde
                self.batiments_places.pop()
                
                # Essayer un autre emplacement
                return self._placer_batiment_avec_backtracking(batiment, bord_prioritaire)
            else:
                # Enlever le bâtiment de la liste à placer
                self.batiments_a_placer.remove(batiment)
                return True
        else:
            # Ajouter aux non placés
            self.batiments_non_places.append(batiment)
            return False
    
    def trouver_emplacement_optimise(self, batiment, bord_prioritaire=False):
        """Trouve un emplacement optimisé pour le bâtiment"""
        orientations = ['horizontal', 'vertical']
        meilleur_emplacement = None
        meilleur_score = -1
        
        for orientation in orientations:
            longueur, largeur = batiment.get_dimensions(orientation)
            
            # Déterminer la plage de recherche
            x_range = range(self.hauteur - longueur + 1)
            y_range = range(self.largeur - largeur + 1)
            
            for x in x_range:
                for y in y_range:
                    if self.peut_placer_batiment(batiment, x, y, orientation):
                        score = 0
                        
                        # Favoriser les bords si demandé
                        if bord_prioritaire:
                            if x == 0 or y == 0 or x + longueur == self.hauteur or y + largeur == self.largeur:
                                score += 100
                        
                        # Calculer un score basé sur l'emplacement
                        # (on peut ajouter d'autres heuristiques ici)
                        
                        if score > meilleur_score:
                            meilleur_score = score
                            meilleur_emplacement = (x, y, orientation)
        
        if meilleur_emplacement:
            return meilleur_emplacement
        else:
            return None, None, None
    
    def calculer_resultats(self):
        """Calcule tous les résultats finaux"""
        culture_par_producteur = {}
        
        for batiment in self.batiments_places:
            if batiment['type'] == 'producteur':
                culture_recue = self.calculer_culture_recue(batiment)
                culture_par_producteur[batiment['nom']] = culture_recue
                self.culture_totale += culture_recue
                
                # Catégoriser par type de production
                if 'guerison' in batiment['production'].lower():
                    self.culture_par_type['guerison'] += culture_recue
                elif 'nourriture' in batiment['production'].lower():
                    self.culture_par_type['nourriture'] += culture_recue
                elif 'or' in batiment['production'].lower():
                    self.culture_par_type['or'] += culture_recue
        
        return culture_par_producteur
    
    def generer_terrain_visuel(self):
        """Génère une représentation visuelle du terrain"""
        terrain_visuel = np.full((self.hauteur, self.largeur), '.', dtype='<U10')
        
        # Marquer les cases occupées initialement (0)
        for i in range(self.hauteur):
            for j in range(self.largeur):
                if self.terrain[i, j] == 0:
                    terrain_visuel[i, j] = 'X'
        
        # Marquer les bâtiments placés
        for batiment in self.batiments_places:
            x, y = batiment['x'], batiment['y']
            longueur, largeur = batiment['longueur'], batiment['largeur']
            
            for i in range(longueur):
                for j in range(largeur):
                    if i == 0 and j == 0:
                        terrain_visuel[x + i, y + j] = batiment['nom'][:3]
                    else:
                        terrain_visuel[x + i, y + j] = '■'
        
        return terrain_visuel

def main():
    st.set_page_config(page_title="Placement de Bâtiments", layout="wide")
    
    st.title("🏗️ Optimiseur de Placement de Bâtiments")
    st.markdown("---")
    
    # Sidebar pour le téléchargement
    with st.sidebar:
        st.header("📂 Fichier d'entrée")
        uploaded_file = st.file_uploader(
            "Choisir un fichier Excel", 
            type=['xlsx', 'xls'],
            help="Le fichier doit contenir deux onglets : 'Terrain' et 'Batiments'"
        )
        
        st.markdown("---")
        st.header("📊 Résultats")
        st.info("Les résultats seront affichés après le placement")
    
    if uploaded_file is not None:
        try:
            # Lire les données
            df_terrain = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            df_batiments = pd.read_excel(uploaded_file, sheet_name=1)
            
            # Afficher un aperçu des données
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📋 Aperçu du Terrain")
                st.dataframe(df_terrain.head(10), use_container_width=True)
                
            with col2:
                st.subheader("🏛️ Aperçu des Bâtiments")
                st.dataframe(df_batiments.head(10), use_container_width=True)
            
            # Bouton pour lancer le placement
            if st.button("🚀 Lancer le placement", type="primary"):
                with st.spinner("Placement des bâtiments en cours..."):
                    # Initialiser le placement
                    placement = PlacementBatiments(df_terrain, df_batiments)
                    
                    # Lancer l'algorithme de placement
                    placement.placer_batiments_ordre()
                    
                    # Calculer les résultats
                    culture_par_producteur = placement.calculer_resultats()
                    
                    # Afficher les résultats
                    st.markdown("---")
                    st.header("✅ Résultats du Placement")
                    
                    # Métriques principales
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("Bâtiments placés", len(placement.batiments_places))
                    with col2:
                        st.metric("Bâtiments non placés", len(placement.batiments_non_places))
                    with col3:
                        cases_utilisees = np.sum(placement.terrain == -1)
                        st.metric("Cases utilisées", int(cases_utilisees))
                    with col4:
                        cases_non_utilisees = np.sum(placement.terrain == 1)
                        st.metric("Cases libres restantes", int(cases_non_utilisees))
                    
                    # Culture totale
                    st.subheader("📈 Statistiques de Culture")
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("Culture totale", f"{placement.culture_totale:.2f}")
                    with col2:
                        st.metric("Culture Guérison", f"{placement.culture_par_type['guerison']:.2f}")
                    with col3:
                        st.metric("Culture Nourriture", f"{placement.culture_par_type['nourriture']:.2f}")
                    with col4:
                        st.metric("Culture Or", f"{placement.culture_par_type['or']:.2f}")
                    
                    # Visualisation du terrain
                    st.subheader("🗺️ Visualisation du Terrain")
                    terrain_visuel = placement.generer_terrain_visuel()
                    
                    # Convertir en DataFrame pour l'affichage
                    df_visuel = pd.DataFrame(terrain_visuel)
                    st.dataframe(df_visuel, use_container_width=True, height=400)
                    
                    # Liste des bâtiments placés
                    st.subheader("📋 Détail des bâtiments placés")
                    df_places = pd.DataFrame(placement.batiments_places)
                    if not df_places.empty:
                        st.dataframe(df_places, use_container_width=True)
                    
                    # Bâtiments non placés
                    if placement.batiments_non_places:
                        st.subheader("⚠️ Bâtiments non placés")
                        non_places_data = []
                        cases_non_placees = 0
                        
                        for batiment in placement.batiments_non_places:
                            non_places_data.append({
                                'Nom': batiment.nom,
                                'Type': batiment.type,
                                'Dimensions': f"{batiment.longueur}x{batiment.largeur}"
                            })
                            cases_non_placees += batiment.longueur * batiment.largeur
                        
                        st.dataframe(pd.DataFrame(non_places_data), use_container_width=True)
                        st.info(f"📦 Total de cases nécessaires pour les bâtiments non placés : {cases_non_placees}")
                    
                    # Générer le fichier de résultats
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # Feuille résumé
                        resume_data = {
                            'Métrique': [
                                'Bâtiments placés',
                                'Bâtiments non placés',
                                'Cases utilisées',
                                'Cases libres restantes',
                                'Culture totale',
                                'Culture Guérison',
                                'Culture Nourriture',
                                'Culture Or'
                            ],
                            'Valeur': [
                                len(placement.batiments_places),
                                len(placement.batiments_non_places),
                                int(np.sum(placement.terrain == -1)),
                                int(np.sum(placement.terrain == 1)),
                                placement.culture_totale,
                                placement.culture_par_type['guerison'],
                                placement.culture_par_type['nourriture'],
                                placement.culture_par_type['or']
                            ]
                        }
                        pd.DataFrame(resume_data).to_excel(writer, sheet_name='Résumé', index=False)
                        
                        # Feuille bâtiments placés
                        if not df_places.empty:
                            df_places.to_excel(writer, sheet_name='Bâtiments placés', index=False)
                        
                        # Feuille terrain final
                        df_terrain_final = pd.DataFrame(placement.terrain)
                        df_terrain_final.to_excel(writer, sheet_name='Terrain final', index=False, header=False)
                        
                        # Feuille visualisation
                        df_visuel.to_excel(writer, sheet_name='Visualisation', index=False, header=False)
                        
                        # Bâtiments non placés
                        if non_places_data:
                            pd.DataFrame(non_places_data).to_excel(writer, sheet_name='Non placés', index=False)
                    
                    output.seek(0)
                    
                    # Bouton de téléchargement
                    st.download_button(
                        label="📥 Télécharger les résultats (Excel)",
                        data=output,
                        file_name="resultats_placement_batiments.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement : {str(e)}")
            st.exception(e)

if __name__ == "__main__":
    main()