import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side
import io
from dataclasses import dataclass
from typing import List, Tuple, Optional
import copy

# Configuration de la page Streamlit
st.set_page_config(page_title="Placeur de Bâtiments", layout="wide")

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
    
@dataclass
class BatimentPlace:
    batiment: Batiment
    x: int
    y: int
    orientation: str  # 'H' ou 'V'

class PlaceurBatiments:
    def __init__(self, terrain_df, batiments_df):
        self.terrain = terrain_df.values
        self.hauteur, self.largeur = self.terrain.shape
        self.batiments = self._charger_batiments(batiments_df)
        self.batiments_places = []
        self.cases_utilisees = np.zeros((self.hauteur, self.largeur), dtype=bool)
        self.journal = []
        self.culture_recue_par_batiment = {}  # Dictionnaire pour stocker la culture reçue par chaque bâtiment producteur
        self.culture_totale = 0
        self.seuils_atteints = {'25%': 0, '50%': 0, '100%': 0}
        
    def _charger_batiments(self, df):
        batiments = []
        for _, row in df.iterrows():
            type_bat = 'neutre'
            if pd.notna(row.get('Culture')) and row['Culture'] > 0:
                type_bat = 'culturel'
            elif pd.notna(row.get('Production')) and row['Production'] != '':
                type_bat = 'producteur'
                
            batiment = Batiment(
                nom=row['Nom'],
                longueur=int(row['Longueur']),
                largeur=int(row['Largeur']),
                quantite=int(row['Quantite']),
                type=type_bat,
                culture=float(row['Culture']) if pd.notna(row.get('Culture')) else 0,
                rayonnement=int(row['Rayonnement']) if pd.notna(row.get('Rayonnement')) else 0,
                boost_25=float(row['Boost 25%']) if pd.notna(row.get('Boost 25%')) else 0,
                boost_50=float(row['Boost 50%']) if pd.notna(row.get('Boost 50%')) else 0,
                boost_100=float(row['Boost 100%']) if pd.notna(row.get('Boost 100%')) else 0,
                production=row.get('Production', '') if pd.notna(row.get('Production')) else ''
            )
            # Ajouter plusieurs instances selon la quantité
            for i in range(batiment.quantite):
                batiments.append(copy.deepcopy(batiment))
        return batiments
    
    def _journaliser(self, message):
        if len(self.journal) < 1000:
            self.journal.append(message)
            return True
        return False
    
    def _case_est_libre(self, x, y):
        return (0 <= x < self.hauteur and 0 <= y < self.largeur and 
                self.terrain[x, y] == 1 and not self.cases_utilisees[x, y])
    
    def _peut_placer(self, batiment, x, y, orientation):
        if orientation == 'H':
            longueur, largeur = batiment.longueur, batiment.largeur
        else:
            longueur, largeur = batiment.largeur, batiment.longueur
            
        if x + longueur > self.hauteur or y + largeur > self.largeur:
            return False
            
        for i in range(longueur):
            for j in range(largeur):
                if not self._case_est_libre(x + i, y + j):
                    return False
        return True
    
    def _placer_batiment(self, batiment, x, y, orientation):
        if orientation == 'H':
            longueur, largeur = batiment.longueur, batiment.largeur
        else:
            longueur, largeur = batiment.largeur, batiment.longueur
            
        for i in range(longueur):
            for j in range(largeur):
                self.cases_utilisees[x + i, y + j] = True
                
        batiment_place = BatimentPlace(batiment, x, y, orientation)
        self.batiments_places.append(batiment_place)
        return batiment_place
    
    def _enlever_batiment(self, batiment_place):
        if batiment_place.orientation == 'H':
            longueur, largeur = batiment_place.batiment.longueur, batiment_place.batiment.largeur
        else:
            longueur, largeur = batiment_place.batiment.largeur, batiment_place.batiment.longueur
            
        for i in range(longueur):
            for j in range(largeur):
                self.cases_utilisees[batiment_place.x + i, batiment_place.y + j] = False
                
        self.batiments_places.remove(batiment_place)
    
    def _calculer_culture_recue(self, batiment_producteur, batiment_place):
        """Calcule la culture reçue par un bâtiment producteur depuis tous les bâtiments culturels"""
        culture_recue = 0
        
        for b_place in self.batiments_places:
            if b_place.batiment.type == 'culturel':
                # Déterminer la zone de rayonnement du bâtiment culturel
                if b_place.orientation == 'H':
                    longueur_c, largeur_c = b_place.batiment.longueur, b_place.batiment.largeur
                else:
                    longueur_c, largeur_c = b_place.batiment.largeur, b_place.batiment.longueur
                
                rayon = b_place.batiment.rayonnement
                
                # Vérifier si le bâtiment producteur est dans la zone de rayonnement
                for i in range(batiment_place.x, batiment_place.x + 
                               (batiment_place.batiment.longueur if batiment_place.orientation == 'H' 
                                else batiment_place.batiment.largeur)):
                    for j in range(batiment_place.y, batiment_place.y + 
                                   (batiment_place.batiment.largeur if batiment_place.orientation == 'H' 
                                    else batiment_place.batiment.longueur)):
                        # Vérifier si cette case est dans le rayonnement du bâtiment culturel
                        for ci in range(max(0, b_place.x - rayon), 
                                       min(self.hauteur, b_place.x + longueur_c + rayon)):
                            for cj in range(max(0, b_place.y - rayon), 
                                           min(self.largeur, b_place.y + largeur_c + rayon)):
                                if (b_place.x <= ci < b_place.x + longueur_c and 
                                    b_place.y <= cj < b_place.y + largeur_c):
                                    # Case du bâtiment culturel lui-même
                                    continue
                                if i == ci and j == cj:
                                    culture_recue += b_place.batiment.culture
                                    break
                            else:
                                continue
                            break
        return culture_recue
    
    def _mettre_a_jour_culture(self):
        """Met à jour la culture reçue par tous les bâtiments producteurs"""
        culture_par_batiment = {}
        culture_totale = 0
        seuils = {'25%': 0, '50%': 0, '100%': 0}
        
        for b_place in self.batiments_places:
            if b_place.batiment.type == 'producteur':
                culture = self._calculer_culture_recue(b_place.batiment, b_place)
                culture_par_batiment[b_place.batiment.nom] = culture
                culture_totale += culture
                
                # Déterminer les seuils atteints
                if culture >= b_place.batiment.boost_100:
                    seuils['100%'] += 1
                elif culture >= b_place.batiment.boost_50:
                    seuils['50%'] += 1
                elif culture >= b_place.batiment.boost_25:
                    seuils['25%'] += 1
                    
        self.culture_recue_par_batiment = culture_par_batiment
        self.culture_totale = culture_totale
        self.seuils_atteints = seuils
    
    def _trouver_plus_grand_batiment_restant(self, batiments_restants):
        max_taille = 0
        for b in batiments_restants:
            taille = b.longueur * b.largeur
            if taille > max_taille:
                max_taille = taille
        return max_taille
    
    def _verifier_place_disponible(self, batiment, batiments_restants):
        """Vérifie s'il reste assez de place pour le plus grand bâtiment restant"""
        if not batiments_restants:
            return True
            
        plus_grand = self._trouver_plus_grand_batiment_restant(batiments_restants)
        cases_libres = np.sum(self.terrain == 1) - np.sum(self.cases_utilisees)
        
        return cases_libres >= plus_grand
    
    def _trouver_emplacement(self, batiment, est_premier_passage=False):
        """Trouve un emplacement pour le bâtiment"""
        # Ordre de recherche : d'abord les bords si c'est un bâtiment neutre
        if batiment.type == 'neutre' and est_premier_passage:
            # Chercher sur les bords
            positions = []
            # Bord haut
            for y in range(self.largeur):
                positions.append((0, y))
            # Bord bas
            for y in range(self.largeur):
                positions.append((self.hauteur - batiment.longueur, y))
            # Bord gauche
            for x in range(self.hauteur):
                positions.append((x, 0))
            # Bord droit
            for x in range(self.hauteur):
                positions.append((x, self.largeur - batiment.largeur))
        else:
            # Chercher partout
            positions = [(x, y) for x in range(self.hauteur) for y in range(self.largeur)]
        
        # Tester les orientations
        orientations = ['H', 'V']
        for orientation in orientations:
            for x, y in positions:
                if self._peut_placer(batiment, x, y, orientation):
                    return (x, y, orientation)
        return None
    
    def placer_batiments(self):
        # Séparer les bâtiments par type
        neutres = [b for b in self.batiments if b.type == 'neutre']
        culturels = [b for b in self.batiments if b.type == 'culturel']
        producteurs = [b for b in self.batiments if b.type == 'producteur']
        
        batiments_non_places = self.batiments.copy()
        pile_placement = []
        
        # Phase 1: Placer les neutres
        self._journaliser("Phase 1: Placement des bâtiments neutres")
        for batiment in neutres:
            self._journaliser(f"Évaluation du bâtiment neutre: {batiment.nom}")
            emplacement = self._trouver_emplacement(batiment, est_premier_passage=True)
            
            if emplacement:
                x, y, orientation = emplacement
                self._placer_batiment(batiment, x, y, orientation)
                batiments_non_places.remove(batiment)
                pile_placement.append((batiment, x, y, orientation))
                self._journaliser(f"Placement du bâtiment neutre {batiment.nom} à ({x},{y}) orientation {orientation}")
            else:
                self._journaliser(f"Impossible de placer le bâtiment neutre {batiment.nom}")
        
        # Phase 2: Alterner culturels et producteurs
        self._journaliser("Phase 2: Placement des bâtiments culturels et producteurs")
        index_culturel = 0
        index_producteur = 0
        essais = 0
        max_essais = 1000
        
        while (index_culturel < len(culturels) or index_producteur < len(producteurs)) and essais < max_essais:
            essais += 1
            
            # Alterner: culturel d'abord
            if index_culturel < len(culturels):
                batiment = culturels[index_culturel]
                self._journaliser(f"Évaluation du bâtiment culturel: {batiment.nom}")
                
                emplacement = self._trouver_emplacement(batiment)
                if emplacement and self._verifier_place_disponible(batiment, culturels[index_culturel+1:] + producteurs[index_producteur:]):
                    x, y, orientation = emplacement
                    self._placer_batiment(batiment, x, y, orientation)
                    batiments_non_places.remove(batiment)
                    pile_placement.append((batiment, x, y, orientation))
                    index_culturel += 1
                    self._journaliser(f"Placement du bâtiment culturel {batiment.nom} à ({x},{y}) orientation {orientation}")
                else:
                    # Backtracking
                    if pile_placement:
                        dernier = pile_placement.pop()
                        self._enlever_batiment(BatimentPlace(dernier[0], dernier[1], dernier[2], dernier[3]))
                        batiments_non_places.append(dernier[0])
                        self._journaliser(f"Retrait du bâtiment {dernier[0].nom}")
                        
                        # Réajuster les index
                        if dernier[0].type == 'culturel':
                            index_culturel -= 1
                        elif dernier[0].type == 'producteur':
                            index_producteur -= 1
            
            # Ensuite producteur
            if index_producteur < len(producteurs):
                batiment = producteurs[index_producteur]
                self._journaliser(f"Évaluation du bâtiment producteur: {batiment.nom}")
                
                emplacement = self._trouver_emplacement(batiment)
                if emplacement and self._verifier_place_disponible(batiment, culturels[index_culturel:] + producteurs[index_producteur+1:]):
                    x, y, orientation = emplacement
                    self._placer_batiment(batiment, x, y, orientation)
                    batiments_non_places.remove(batiment)
                    pile_placement.append((batiment, x, y, orientation))
                    index_producteur += 1
                    self._journaliser(f"Placement du bâtiment producteur {batiment.nom} à ({x},{y}) orientation {orientation}")
                    
                    # Mettre à jour la culture après placement
                    self._mettre_a_jour_culture()
                else:
                    # Backtracking
                    if pile_placement:
                        dernier = pile_placement.pop()
                        self._enlever_batiment(BatimentPlace(dernier[0], dernier[1], dernier[2], dernier[3]))
                        batiments_non_places.append(dernier[0])
                        self._journaliser(f"Retrait du bâtiment {dernier[0].nom}")
                        
                        if dernier[0].type == 'culturel':
                            index_culturel -= 1
                        elif dernier[0].type == 'producteur':
                            index_producteur -= 1
        
        return batiments_non_places
    
    def generer_resultats(self, batiments_non_places):
        # Calcul des cases non utilisées
        cases_libres_initiales = np.sum(self.terrain == 1)
        cases_utilisees = np.sum(self.cases_utilisees)
        cases_non_utilisees = cases_libres_initiales - cases_utilisees
        
        # Calcul des cases des bâtiments non placés
        cases_batiments_non_places = sum(b.longueur * b.largeur for b in batiments_non_places)
        
        # Création de la visualisation du terrain
        terrain_visuel = np.full((self.hauteur, self.largeur), '□', dtype=object)
        
        # Marquer les cases occupées initialement (0)
        for i in range(self.hauteur):
            for j in range(self.largeur):
                if self.terrain[i, j] == 0:
                    terrain_visuel[i, j] = '■'  # Case occupée
        
        # Marquer les bâtiments placés
        for b_place in self.batiments_places:
            if b_place.orientation == 'H':
                longueur, largeur = b_place.batiment.longueur, b_place.batiment.largeur
            else:
                longueur, largeur = b_place.batiment.largeur, b_place.batiment.longueur
                
            type_code = {
                'culturel': 'C',
                'producteur': 'P',
                'neutre': 'N'
            }.get(b_place.batiment.type, '?')
            
            for i in range(longueur):
                for j in range(largeur):
                    terrain_visuel[b_place.x + i, b_place.y + j] = type_code
        
        return {
            'journal': self.journal,
            'culture_totale': self.culture_totale,
            'seuils_atteints': self.seuils_atteints,
            'terrain_visuel': terrain_visuel,
            'batiments_non_places': batiments_non_places,
            'cases_non_utilisees': cases_non_utilisees,
            'cases_batiments_non_places': cases_batiments_non_places,
            'batiments_places': self.batiments_places
        }

def main():
    st.title("🏗️ Placeur de Bâtiments")
    st.markdown("---")
    
    # Sidebar pour le chargement du fichier
    with st.sidebar:
        st.header("📁 Chargement du fichier")
        uploaded_file = st.file_uploader(
            "Choisir le fichier Excel", 
            type=['xlsx', 'xls'],
            help="Le fichier doit contenir deux onglets: 'Terrain' et 'Batiments'"
        )
        
        if uploaded_file:
            st.success("Fichier chargé avec succès!")
    
    # Zone principale
    if uploaded_file:
        try:
            # Lecture du fichier Excel
            df_terrain = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            df_batiments = pd.read_excel(uploaded_file, sheet_name=1)
            
            # Aperçu des données
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📊 Aperçu du terrain")
                st.dataframe(df_terrain, height=200)
                st.caption(f"Dimensions: {df_terrain.shape[0]} lignes × {df_terrain.shape[1]} colonnes")
                
            with col2:
                st.subheader("🏢 Aperçu des bâtiments")
                st.dataframe(df_batiments, height=200)
                st.caption(f"Nombre de types de bâtiments: {len(df_batiments)}")
            
            # Bouton de lancement
            if st.button("🚀 Lancer le placement", type="primary", use_container_width=True):
                with st.spinner("Placement des bâtiments en cours..."):
                    # Initialisation et exécution du placement
                    placeur = PlaceurBatiments(df_terrain, df_batiments)
                    batiments_non_places = placeur.placer_batiments()
                    resultats = placeur.generer_resultats(batiments_non_places)
                
                # Affichage des résultats
                st.markdown("---")
                st.header("📈 Résultats")
                
                # Métriques principales
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Culture totale", f"{resultats['culture_totale']:.1f}")
                with col2:
                    st.metric("Boost 25%", resultats['seuils_atteints']['25%'])
                with col3:
                    st.metric("Boost 50%", resultats['seuils_atteints']['50%'])
                with col4:
                    st.metric("Boost 100%", resultats['seuils_atteints']['100%'])
                
                # Terrain visuel
                st.subheader("🗺️ Carte du terrain")
                st.caption("Légende: □ Libre | ■ Occupé | C Culturel | P Producteur | N Neutre")
                
                # Convertir en DataFrame pour l'affichage
                df_visuel = pd.DataFrame(resultats['terrain_visuel'])
                st.dataframe(df_visuel, height=400, use_container_width=True)
                
                # Statistiques des bâtiments non placés
                st.subheader("📦 Bâtiments non placés")
                if resultats['batiments_non_places']:
                    non_places_data = []
                    for b in resultats['batiments_non_places']:
                        non_places_data.append({
                            'Nom': b.nom,
                            'Type': b.type,
                            'Taille': f"{b.longueur}×{b.largeur}",
                            'Surface': b.longueur * b.largeur
                        })
                    st.dataframe(pd.DataFrame(non_places_data))
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Cases non utilisées", resultats['cases_non_utilisees'])
                    with col2:
                        st.metric("Cases des bâtiments non placés", resultats['cases_batiments_non_places'])
                else:
                    st.success("✅ Tous les bâtiments ont été placés!")
                    st.metric("Cases non utilisées", resultats['cases_non_utilisees'])
                
                # Journal des opérations
                with st.expander("📋 Journal des opérations"):
                    for i, entry in enumerate(resultats['journal'], 1):
                        st.text(f"{i:3d}. {entry}")
                    if len(resultats['journal']) >= 1000:
                        st.warning("⚠️ Limite de 1000 entrées atteinte dans le journal")
                
                # Export des résultats
                st.markdown("---")
                st.subheader("💾 Export des résultats")
                
                # Création du fichier Excel de résultats
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    # Feuille des résultats
                    results_data = {
                        'Métrique': ['Culture totale', 'Boost 25%', 'Boost 50%', 'Boost 100%', 
                                    'Cases non utilisées', 'Cases bat. non placés'],
                        'Valeur': [resultats['culture_totale'], 
                                  resultats['seuils_atteints']['25%'],
                                  resultats['seuils_atteints']['50%'],
                                  resultats['seuils_atteints']['100%'],
                                  resultats['cases_non_utilisees'],
                                  resultats['cases_batiments_non_places']]
                    }
                    df_results = pd.DataFrame(results_data)
                    df_results.to_excel(writer, sheet_name='Résultats', index=False)
                    
                    # Feuille du terrain avec placements
                    df_terrain_placements = pd.DataFrame(resultats['terrain_visuel'])
                    df_terrain_placements.to_excel(writer, sheet_name='Terrain_Placements', index=False, header=False)
                    
                    # Feuille des bâtiments placés
                    places_data = []
                    for bp in resultats['batiments_places']:
                        places_data.append({
                            'Nom': bp.batiment.nom,
                            'Type': bp.batiment.type,
                            'Position X': bp.x,
                            'Position Y': bp.y,
                            'Orientation': bp.orientation,
                            'Culture reçue': resultats['culture_recue_par_batiment'].get(bp.batiment.nom, 0)
                        })
                    if places_data:
                        df_places = pd.DataFrame(places_data)
                        df_places.to_excel(writer, sheet_name='Batiments_Places', index=False)
                    
                    # Feuille des bâtiments non placés
                    if resultats['batiments_non_places']:
                        non_places_data = []
                        for b in resultats['batiments_non_places']:
                            non_places_data.append({
                                'Nom': b.nom,
                                'Type': b.type,
                                'Longueur': b.longueur,
                                'Largeur': b.largeur,
                                'Surface': b.longueur * b.largeur
                            })
                        df_non_places = pd.DataFrame(non_places_data)
                        df_non_places.to_excel(writer, sheet_name='Batiments_Non_Places', index=False)
                    
                    # Feuille du journal
                    journal_data = {'Étape': range(1, len(resultats['journal']) + 1),
                                   'Action': resultats['journal']}
                    df_journal = pd.DataFrame(journal_data)
                    df_journal.to_excel(writer, sheet_name='Journal', index=False)
                
                output.seek(0)
                
                st.download_button(
                    label="📥 Télécharger les résultats (Excel)",
                    data=output,
                    file_name="resultats_placement_batiments.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement du fichier: {str(e)}")
            st.exception(e)
    else:
        # Message d'accueil
        st.info("👈 Veuillez charger un fichier Excel dans le panneau latéral")
        
        with st.expander("📝 Format du fichier attendu"):
            st.markdown("""
            ### Structure du fichier Excel
            
            **Onglet 1: Terrain**
            - Matrice de 0 et 1
            - 1 = case libre
            - 0 = case occupée
            
            **Onglet 2: Bâtiments**
            - Colonnes: Nom, Longueur, Largeur, Quantite, Type, Culture, Rayonnement, Boost 25%, Boost 50%, Boost 100%, Production
            - Type: 'culturel' ou 'producteur' (ou vide pour neutre)
            - Culture: quantité de culture produite (pour bâtiments culturels)
            - Rayonnement: portée du boost culturel
            - Production: ce que produit le bâtiment (pour producteurs)
            """)

if __name__ == "__main__":
    main()