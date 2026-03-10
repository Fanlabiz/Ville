import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side
import io
from dataclasses import dataclass
from typing import List, Tuple, Optional
import copy

@dataclass
class Batiment:
    nom: str
    longueur: int
    largeur: int
    quantite: int
    type: str  # 'culturel' ou 'producteur'
    culture: float
    rayonnement: int
    boost_25: float
    boost_50: float
    boost_100: float
    production: str
    
@dataclass
class BatimentPlace:
    batiment: Batiment
    x: int
    y: int
    orientation: str  # 'H' ou 'V'
    culture_recue: float = 0

class PlacementBatiments:
    def __init__(self, terrain_df, batiments_df):
        self.terrain = terrain_df.values
        self.hauteur, self.largeur = self.terrain.shape
        self.terrain_occupe = np.zeros((self.hauteur, self.largeur), dtype=bool)
        self.terrain_occupe[self.terrain == 0] = True  # Cases occupées initialement
        
        # Normaliser les noms de colonnes pour éviter les problèmes de casse/accents
        self.batiments_df = self._normaliser_colonnes(batiments_df)
        self.batiments = self._charger_batiments()
        self.batiments_places = []
        self.batiments_non_places = []
    
    def _normaliser_colonnes(self, df):
        """Normalise les noms de colonnes pour gérer différentes variations"""
        df = df.copy()
        
        # Dictionnaire de correspondance des noms de colonnes
        mapping = {
            'nom': 'nom',
            'name': 'nom',
            'longueur': 'longueur',
            'length': 'longueur',
            'largeur': 'largeur',
            'width': 'largeur',
            'quantite': 'quantite',
            'quantity': 'quantite',
            'type': 'type',
            'culture': 'culture',
            'rayonnement': 'rayonnement',
            'radiation': 'rayonnement',
            'boost 25%': 'boost_25',
            'boost 25': 'boost_25',
            'boost25': 'boost_25',
            'boost 50%': 'boost_50',
            'boost 50': 'boost_50',
            'boost50': 'boost_50',
            'boost 100%': 'boost_100',
            'boost 100': 'boost_100',
            'boost100': 'boost_100',
            'production': 'production'
        }
        
        # Nettoyer et mapper les noms de colonnes
        nouvelles_colonnes = {}
        for col in df.columns:
            col_clean = str(col).strip().lower().replace('é', 'e').replace('è', 'e').replace('ê', 'e').replace('à', 'a').replace('â', 'a').replace('î', 'i').replace('ô', 'o').replace('û', 'u')
            
            # Chercher une correspondance
            trouve = False
            for key, value in mapping.items():
                if key in col_clean or col_clean in key:
                    nouvelles_colonnes[col] = value
                    trouve = True
                    break
            
            if not trouve:
                nouvelles_colonnes[col] = col_clean
        
        df.rename(columns=nouvelles_colonnes, inplace=True)
        return df
    
    def _charger_batiments(self):
        batiments = []
        df = self.batiments_df
        
        # Vérifier que toutes les colonnes nécessaires existent
        colonnes_requises = ['nom', 'longueur', 'largeur', 'quantite', 'type']
        for col in colonnes_requises:
            if col not in df.columns:
                st.error(f"Colonne manquante dans le fichier : {col}")
                st.write("Colonnes trouvées :", list(df.columns))
                return []
        
        for idx, row in df.iterrows():
            try:
                quantite = int(row['quantite']) if pd.notna(row['quantite']) else 1
                
                for i in range(quantite):
                    batiments.append(Batiment(
                        nom=str(row['nom']) if pd.notna(row['nom']) else f"Batiment_{idx}",
                        longueur=int(row['longueur']) if pd.notna(row['longueur']) else 1,
                        largeur=int(row['largeur']) if pd.notna(row['largeur']) else 1,
                        quantite=1,
                        type=str(row['type']).lower() if pd.notna(row['type']) else "neutre",
                        culture=float(row['culture']) if 'culture' in df.columns and pd.notna(row['culture']) else 0,
                        rayonnement=int(row['rayonnement']) if 'rayonnement' in df.columns and pd.notna(row['rayonnement']) else 0,
                        boost_25=float(row['boost_25']) if 'boost_25' in df.columns and pd.notna(row['boost_25']) else 0,
                        boost_50=float(row['boost_50']) if 'boost_50' in df.columns and pd.notna(row['boost_50']) else 0,
                        boost_100=float(row['boost_100']) if 'boost_100' in df.columns and pd.notna(row['boost_100']) else 0,
                        production=str(row['production']) if 'production' in df.columns and pd.notna(row['production']) else ""
                    ))
            except Exception as e:
                st.warning(f"Erreur lors du chargement de la ligne {idx}: {e}")
                continue
                
        return batiments
    
    def _trouver_plus_grand_batiment_restant(self, batiments_restants):
        if not batiments_restants:
            return 0, 0
        max_surface = 0
        max_long, max_larg = 0, 0
        for b in batiments_restants:
            surface = b.longueur * b.largeur
            if surface > max_surface:
                max_surface = surface
                max_long = b.longueur
                max_larg = b.largeur
        return max_long, max_larg
    
    def _verifier_placement_possible(self, x, y, longueur, largeur, orientation):
        if orientation == 'H':
            l, L = longueur, largeur
        else:
            l, L = largeur, longueur
            
        if x + l > self.hauteur or y + L > self.largeur:
            return False
        
        # Vérifier que toutes les cases sont libres
        for i in range(l):
            for j in range(L):
                if self.terrain_occupe[x + i][y + j]:
                    return False
        return True
    
    def _placer_batiment(self, batiment, x, y, orientation):
        if orientation == 'H':
            l, L = batiment.longueur, batiment.largeur
        else:
            l, L = batiment.largeur, batiment.longueur
            
        for i in range(l):
            for j in range(L):
                self.terrain_occupe[x + i][y + j] = True
        
        batiment_place = BatimentPlace(batiment, x, y, orientation)
        self.batiments_places.append(batiment_place)
        return batiment_place
    
    def _calculer_culture_recue(self, batiment_place, batiments_culturels):
        if batiment_place.batiment.type != 'producteur':
            return 0
        
        culture_totale = 0
        x1, y1 = batiment_place.x, batiment_place.y
        if batiment_place.orientation == 'H':
            x2 = x1 + batiment_place.batiment.longueur
            y2 = y1 + batiment_place.batiment.largeur
        else:
            x2 = x1 + batiment_place.batiment.largeur
            y2 = y1 + batiment_place.batiment.longueur
        
        for culturel in batiments_culturels:
            # Zone de rayonnement du culturel
            cx1 = max(0, culturel.x - culturel.batiment.rayonnement)
            cy1 = max(0, culturel.y - culturel.batiment.rayonnement)
            if culturel.orientation == 'H':
                cx2 = min(self.hauteur, culturel.x + culturel.batiment.longueur + culturel.batiment.rayonnement)
                cy2 = min(self.largeur, culturel.y + culturel.batiment.largeur + culturel.batiment.rayonnement)
            else:
                cx2 = min(self.hauteur, culturel.x + culturel.batiment.largeur + culturel.batiment.rayonnement)
                cy2 = min(self.largeur, culturel.y + culturel.batiment.longueur + culturel.batiment.rayonnement)
            
            # Vérifier si les zones se chevauchent
            if not (x2 <= cx1 or x1 >= cx2 or y2 <= cy1 or y1 >= cy2):
                culture_totale += culturel.batiment.culture
        
        return culture_totale
    
    def _calculer_boost(self, culture_recue, batiment):
        if culture_recue >= batiment.boost_100:
            return 100
        elif culture_recue >= batiment.boost_50:
            return 50
        elif culture_recue >= batiment.boost_25:
            return 25
        return 0
    
    def placer_tous_batiments(self):
        if not self.batiments:
            st.error("Aucun bâtiment à placer")
            return [], []
            
        # Séparer les bâtiments par type
        neutres = [b for b in self.batiments if b.type not in ['culturel', 'producteur']]
        culturels = [b for b in self.batiments if b.type == 'culturel']
        producteurs = [b for b in self.batiments if b.type == 'producteur']
        
        tous_batiments = neutres + culturels + producteurs
        batiments_restants = tous_batiments.copy()
        self.batiments_places = []
        self.terrain_occupe = np.zeros((self.hauteur, self.largeur), dtype=bool)
        self.terrain_occupe[self.terrain == 0] = True
        
        # Placer les bâtiments dans l'ordre
        progression = st.progress(0)
        for i, batiment in enumerate(tous_batiments):
            progression.progress((i + 1) / len(tous_batiments))
            
            place = False
            meilleur_placement = None
            meilleur_score = -1
            
            # Essayer toutes les positions et orientations possibles
            for x in range(self.hauteur):
                for y in range(self.largeur):
                    for orientation in ['H', 'V']:
                        if self._verifier_placement_possible(x, y, batiment.longueur, batiment.largeur, orientation):
                            # Vérifier qu'il reste assez de place pour le plus grand bâtiment restant
                            max_long, max_larg = self._trouver_plus_grand_batiment_restant(batiments_restants[i+1:])
                            
                            # Sauvegarder l'état actuel
                            terrain_sauve = self.terrain_occupe.copy()
                            
                            # Placer temporairement
                            self._placer_batiment(batiment, x, y, orientation)
                            
                            # Vérifier s'il reste assez de place
                            assez_place = True
                            if max_long > 0 and max_larg > 0:
                                place_trouvee = False
                                for bx in range(self.hauteur - max_long + 1):
                                    for by in range(self.largeur - max_larg + 1):
                                        if self._verifier_placement_possible(bx, by, max_long, max_larg, 'H'):
                                            place_trouvee = True
                                            break
                                    if place_trouvee:
                                        break
                                assez_place = place_trouvee
                            
                            # Restaurer l'état
                            self.terrain_occupe = terrain_sauve
                            self.batiments_places.pop()
                            
                            if assez_place:
                                # Calculer le score pour ce placement
                                score = 0
                                if batiment.type == 'producteur':
                                    # Pour les producteurs, on veut maximiser la culture reçue
                                    batiment_temp = BatimentPlace(batiment, x, y, orientation)
                                    culture = self._calculer_culture_recue(batiment_temp, 
                                        [p for p in self.batiments_places if p.batiment.type == 'culturel'])
                                    score = culture
                                
                                if score > meilleur_score:
                                    meilleur_score = score
                                    meilleur_placement = (x, y, orientation)
            
            if meilleur_placement:
                x, y, orientation = meilleur_placement
                batiment_place = self._placer_batiment(batiment, x, y, orientation)
                
                # Calculer la culture reçue pour les producteurs
                if batiment.type == 'producteur':
                    culturels_places = [p for p in self.batiments_places if p.batiment.type == 'culturel']
                    batiment_place.culture_recue = self._calculer_culture_recue(batiment_place, culturels_places)
                
                batiments_restants.remove(batiment)
            else:
                self.batiments_non_places.append(batiment)
        
        return self.batiments_places, self.batiments_non_places
    
    def calculer_statistiques(self):
        stats = {
            'culture_totale': 0,
            'culture_guerison': 0,
            'culture_nourriture': 0,
            'culture_or': 0,
            'boosts': {'25%': 0, '50%': 0, '100%': 0},
            'cases_non_utilisees': np.sum(~self.terrain_occupe),
            'surface_non_placee': sum(b.longueur * b.largeur for b in self.batiments_non_places)
        }
        
        for bp in self.batiments_places:
            if bp.batiment.type == 'producteur':
                stats['culture_totale'] += bp.culture_recue
                
                # Catégoriser par type de production
                prod = str(bp.batiment.production).lower()
                if 'guerison' in prod or 'guérison' in prod:
                    stats['culture_guerison'] += bp.culture_recue
                elif 'nourriture' in prod:
                    stats['culture_nourriture'] += bp.culture_recue
                elif 'or' in prod:
                    stats['culture_or'] += bp.culture_recue
                
                # Calculer le boost atteint
                boost = self._calculer_boost(bp.culture_recue, bp.batiment)
                if boost == 25:
                    stats['boosts']['25%'] += 1
                elif boost == 50:
                    stats['boosts']['50%'] += 1
                elif boost == 100:
                    stats['boosts']['100%'] += 1
        
        return stats
    
    def generer_visualisation(self):
        # Créer une matrice de visualisation
        vis = np.full((self.hauteur, self.largeur), '.', dtype=object)
        
        # Marquer les cases occupées initialement
        for i in range(self.hauteur):
            for j in range(self.largeur):
                if self.terrain[i][j] == 0:
                    vis[i][j] = '█'  # Case occupée
        
        # Placer les bâtiments
        lettres = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
        for i, bp in enumerate(self.batiments_places):
            if bp.orientation == 'H':
                l, L = bp.batiment.longueur, bp.batiment.largeur
            else:
                l, L = bp.batiment.largeur, bp.batiment.longueur
            
            # Choisir un symbole pour ce bâtiment
            if i < len(lettres):
                symbole = lettres[i]
            else:
                symbole = str(i)
            
            for x in range(l):
                for y in range(L):
                    vis[bp.x + x][bp.y + y] = symbole
        
        return vis

def main():
    st.set_page_config(page_title="Placement de Bâtiments", page_icon="🏗️", layout="wide")
    
    st.title("🏗️ Placement Optimisé de Bâtiments")
    
    with st.expander("📋 Instructions", expanded=True):
        st.markdown("""
        ### Format du fichier Excel attendu :
        
        **Onglet 1 - Terrain** :
        - Matrice de 0 (cases occupées) et 1 (cases libres)
        - Pas d'en-tête, que des nombres
        
        **Onglet 2 - Bâtiments** :
        - Colonnes : Nom, Longueur, Largeur, Quantité, Type, Culture, Rayonnement, Boost 25%, Boost 50%, Boost 100%, Production
        - Les noms de colonnes peuvent varier (accents, majuscules...)
        - Types possibles : "culturel", "producteur", ou autre pour neutre
        """)
    
    uploaded_file = st.file_uploader("📁 Choisissez votre fichier Excel", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # Lire les deux onglets
            terrain_df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            
            # Essayer de lire le second onglet avec différentes options
            try:
                batiments_df = pd.read_excel(uploaded_file, sheet_name=1)
            except:
                # Si sheet_name=1 ne fonctionne pas, essayer par index
                xl = pd.ExcelFile(uploaded_file)
                batiments_df = pd.read_excel(uploaded_file, sheet_name=xl.sheet_names[1])
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("🗺️ Aperçu du terrain")
                st.dataframe(terrain_df, height=300)
                st.caption(f"Dimensions : {terrain_df.shape[0]}×{terrain_df.shape[1]}")
            
            with col2:
                st.subheader("🏢 Aperçu des bâtiments")
                st.dataframe(batiments_df, height=300)
                st.caption(f"Nombre de types : {len(batiments_df)}")
            
            if st.button("🚀 Lancer le placement optimal", type="primary"):
                with st.spinner("Calcul du placement en cours..."):
                    # Créer et exécuter le placement
                    placement = PlacementBatiments(terrain_df, batiments_df)
                    
                    if not placement.batiments:
                        st.error("Aucun bâtiment valide trouvé dans le fichier")
                        st.stop()
                    
                    places, non_places = placement.placer_tous_batiments()
                    stats = placement.calculer_statistiques()
                    
                    # Afficher les résultats
                    st.subheader("📊 Résultats du placement")
                    
                    # Métriques principales
                    col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                    with col_m1:
                        st.metric("Bâtiments placés", len(places))
                    with col_m2:
                        st.metric("Bâtiments non placés", len(non_places))
                    with col_m3:
                        st.metric("Cases libres restantes", stats['cases_non_utilisees'])
                    with col_m4:
                        st.metric("Culture totale", f"{stats['culture_totale']:.0f}")
                    
                    # Culture par type
                    st.subheader("💫 Culture reçue par type")
                    col_c1, col_c2, col_c3 = st.columns(3)
                    with col_c1:
                        st.metric("Guérison", f"{stats['culture_guerison']:.0f}")
                    with col_c2:
                        st.metric("Nourriture", f"{stats['culture_nourriture']:.0f}")
                    with col_c3:
                        st.metric("Or", f"{stats['culture_or']:.0f}")
                    
                    # Boosts
                    st.subheader("⚡ Boosts atteints")
                    col_b1, col_b2, col_b3 = st.columns(3)
                    with col_b1:
                        st.metric("Boost 25%", stats['boosts']['25%'])
                    with col_b2:
                        st.metric("Boost 50%", stats['boosts']['50%'])
                    with col_b3:
                        st.metric("Boost 100%", stats['boosts']['100%'])
                    
                    # Visualisation du terrain
                    st.subheader("🗺️ Visualisation du placement")
                    vis = placement.generer_visualisation()
                    vis_df = pd.DataFrame(vis)
                    
                    # Afficher la légende
                    st.markdown("**Légende :**")
                    col_leg = st.columns(4)
                    with col_leg[0]:
                        st.markdown("█ : Case occupée")
                    with col_leg[1]:
                        st.markdown(". : Case libre")
                    with col_leg[2]:
                        st.markdown("Lettres : Bâtiments placés")
                    
                    st.dataframe(vis_df, height=400, use_container_width=True)
                    
                    # Liste des bâtiments non placés
                    if non_places:
                        st.subheader("⚠️ Bâtiments non placés")
                        non_places_data = []
                        for b in non_places:
                            non_places_data.append({
                                'Nom': b.nom,
                                'Type': b.type,
                                'Dimensions': f"{b.longueur}×{b.largeur}"
                            })
                        st.dataframe(pd.DataFrame(non_places_data))
                    
                    # Générer le fichier de résultats
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        # Feuille des statistiques
                        stats_df = pd.DataFrame([{
                            'Culture totale': stats['culture_totale'],
                            'Culture guérison': stats['culture_guerison'],
                            'Culture nourriture': stats['culture_nourriture'],
                            'Culture or': stats['culture_or'],
                            'Boost 25%': stats['boosts']['25%'],
                            'Boost 50%': stats['boosts']['50%'],
                            'Boost 100%': stats['boosts']['100%'],
                            'Bâtiments placés': len(places),
                            'Bâtiments non placés': len(non_places),
                            'Cases non utilisées': stats['cases_non_utilisees'],
                            'Surface non placée': stats['surface_non_placee']
                        }])
                        stats_df.to_excel(writer, sheet_name='Statistiques', index=False)
                        
                        # Feuille du placement
                        placement_data = []
                        for i, bp in enumerate(places):
                            placement_data.append({
                                'ID': i,
                                'Nom': bp.batiment.nom,
                                'Type': bp.batiment.type,
                                'Position X': bp.x,
                                'Position Y': bp.y,
                                'Orientation': bp.orientation,
                                'Culture reçue': bp.culture_recue,
                                'Boost (%)': placement._calculer_boost(bp.culture_recue, bp.batiment)
                            })
                        placement_df = pd.DataFrame(placement_data)
                        placement_df.to_excel(writer, sheet_name='Placement', index=False)
                        
                        # Feuille du terrain visualisé
                        vis_df.to_excel(writer, sheet_name='Terrain_place', header=False, index=False)
                        
                        # Bâtiments non placés
                        if non_places:
                            non_places_df = pd.DataFrame([{
                                'Nom': b.nom,
                                'Type': b.type,
                                'Longueur': b.longueur,
                                'Largeur': b.largeur
                            } for b in non_places])
                            non_places_df.to_excel(writer, sheet_name='Non_places', index=False)
                    
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 Télécharger les résultats Excel",
                        data=output,
                        file_name="resultats_placement.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
        
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement du fichier : {str(e)}")
            st.exception(e)
            
            # Aide au débogage
            st.info("💡 Astuce : Vérifiez que votre fichier Excel a bien deux onglets avec les bonnes colonnes")

if __name__ == "__main__":
    main()