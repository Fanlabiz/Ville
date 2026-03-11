import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from io import BytesIO
from dataclasses import dataclass
from typing import List, Tuple, Optional, Dict, Set
import copy
import openpyxl
import traceback

# ============================================================================
# CLASSES DE BASE
# ============================================================================

@dataclass
class Batiment:
    """Représente un bâtiment à placer"""
    nom: str
    longueur: int
    largeur: int
    quantite: int
    type: str  # 'culturel', 'producteur', 'neutre'
    culture: float = 0
    rayonnement: int = 0
    boost_25: float = 0
    boost_50: float = 0
    boost_100: float = 0
    production: str = ""
    
    def __post_init__(self):
        # Normaliser le type
        if self.type:
            type_lower = self.type.lower().strip()
            if type_lower in ['culturel', 'culture', 'culturelle']:
                self.type = 'culturel'
            elif type_lower in ['producteur', 'productrice', 'producteurbis']:
                self.type = 'producteur'
            else:
                self.type = 'neutre'
        else:
            self.type = 'neutre'
    
    def get_dimensions(self, orientation: str) -> Tuple[int, int]:
        """Retourne (longueur, largeur) selon l'orientation"""
        if orientation == 'horizontal':
            return self.longueur, self.largeur
        else:  # vertical
            return self.largeur, self.longueur


class Terrain:
    """Gère le terrain et les placements"""
    
    def __init__(self, matrice):
        self.matrice = np.array(matrice, dtype=int)
        self.hauteur, self.largeur = self.matrice.shape
        self.occupation = np.zeros_like(self.matrice, dtype=object)
        self.culture_map = np.zeros_like(self.matrice, dtype=float)
        
    def est_libre(self, x: int, y: int, longueur: int, largeur: int) -> bool:
        """Vérifie si un emplacement est libre pour placer un bâtiment"""
        if x + longueur > self.hauteur or y + largeur > self.largeur:
            return False
        
        # Vérifier que toutes les cases sont libres (1 dans la matrice d'origine et non occupées)
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                if self.matrice[i, j] != 1 or self.occupation[i, j] != 0:
                    return False
        return True
    
    def placer_batiment(self, batiment: Batiment, x: int, y: int, orientation: str):
        """Place un bâtiment sur le terrain"""
        longueur, largeur = batiment.get_dimensions(orientation)
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                self.occupation[i, j] = batiment
    
    def enlever_batiment(self, batiment: Batiment, x: int, y: int, orientation: str):
        """Enlève un bâtiment du terrain"""
        longueur, largeur = batiment.get_dimensions(orientation)
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                self.occupation[i, j] = 0
    
    def calculer_culture_recue(self, batiments_culturels: List[Tuple[Batiment, int, int, str]]):
        """Calcule la culture reçue par chaque case à partir des bâtiments culturels"""
        self.culture_map = np.zeros_like(self.matrice, dtype=float)
        
        for batiment, x, y, orientation in batiments_culturels:
            if batiment.type == 'culturel' and batiment.rayonnement > 0:
                longueur, largeur = batiment.get_dimensions(orientation)
                
                # Définir la zone de rayonnement
                x_min = max(0, x - batiment.rayonnement)
                x_max = min(self.hauteur, x + longueur + batiment.rayonnement)
                y_min = max(0, y - batiment.rayonnement)
                y_max = min(self.largeur, y + largeur + batiment.rayonnement)
                
                # Ajouter la culture dans la zone
                self.culture_map[x_min:x_max, y_min:y_max] += batiment.culture
    
    def get_culture_pour_batiment(self, batiment: Batiment, x: int, y: int, orientation: str) -> float:
        """Calcule la culture totale reçue par un bâtiment producteur"""
        if batiment.type != 'producteur':
            return 0
            
        longueur, largeur = batiment.get_dimensions(orientation)
        culture_totale = 0
        
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                culture_totale += self.culture_map[i, j]
        
        return culture_totale
    
    def get_boost_niveau(self, culture: float, batiment: Batiment) -> str:
        """Détermine le niveau de boost en fonction de la culture reçue"""
        if culture >= batiment.boost_100:
            return "100%"
        elif culture >= batiment.boost_50:
            return "50%"
        elif culture >= batiment.boost_25:
            return "25%"
        else:
            return "0%"


class LogManager:
    """Gère le journal des opérations avec limite de 1000 entrées"""
    
    def __init__(self, max_entries=1000):
        self.entries = []
        self.max_entries = max_entries
        self.arret = False
        
    def ajouter(self, message: str):
        if len(self.entries) < self.max_entries:
            self.entries.append(message)
        else:
            self.arret = True
            
    def get_logs(self) -> List[str]:
        return self.entries


# ============================================================================
# ALGORITHME DE PLACEMENT
# ============================================================================

class PlacementBatiments:
    """Algorithme principal de placement"""
    
    def __init__(self, terrain_data: List[List[int]], batiments_data: List[Dict]):
        self.terrain = Terrain(terrain_data)
        self.batiments = self._creer_batiments(batiments_data)
        self.logger = LogManager()
        self.batiments_places = []  # (batiment, x, y, orientation)
        self.batiments_non_places = []
        self.historique_placements = {}  # Pour éviter de replacer au même endroit
        
    def _creer_batiments(self, data: List[Dict]) -> List[Batiment]:
        """Crée la liste des bâtiments à placer (en tenant compte de la quantité)"""
        batiments = []
        
        for i, row in enumerate(data):
            try:
                nom = row.get('Nom', f'Batiment_{i}')
                if pd.isna(nom):
                    nom = f'Batiment_{i}'
                
                # Gérer les valeurs manquantes avec des valeurs par défaut
                try:
                    longueur = int(row.get('Longueur', 1)) if not pd.isna(row.get('Longueur')) else 1
                except:
                    longueur = 1
                    
                try:
                    largeur = int(row.get('Largeur', 1)) if not pd.isna(row.get('Largeur')) else 1
                except:
                    largeur = 1
                    
                try:
                    quantite = int(row.get('Quantite', 1)) if not pd.isna(row.get('Quantite')) else 1
                except:
                    quantite = 1
                
                type_bat = row.get('Type', '')
                if pd.isna(type_bat) or type_bat == '' or type_bat is None:
                    type_bat = 'neutre'
                
                # Valeurs numériques avec gestion des NaN
                culture = 0
                if not pd.isna(row.get('Culture')):
                    try:
                        culture = float(row.get('Culture', 0))
                    except:
                        culture = 0
                
                rayonnement = 0
                if not pd.isna(row.get('Rayonnement')):
                    try:
                        rayonnement = int(row.get('Rayonnement', 0))
                    except:
                        rayonnement = 0
                
                boost_25 = 0
                if not pd.isna(row.get('Boost 25%')):
                    try:
                        boost_25 = float(row.get('Boost 25%', 0))
                    except:
                        boost_25 = 0
                
                boost_50 = 0
                if not pd.isna(row.get('Boost 50%')):
                    try:
                        boost_50 = float(row.get('Boost 50%', 0))
                    except:
                        boost_50 = 0
                
                boost_100 = 0
                if not pd.isna(row.get('Boost 100%')):
                    try:
                        boost_100 = float(row.get('Boost 100%', 0))
                    except:
                        boost_100 = 0
                
                production = row.get('Production', '')
                if pd.isna(production):
                    production = ''
                    
                batiment = Batiment(
                    nom=str(nom),
                    longueur=longueur,
                    largeur=largeur,
                    quantite=quantite,
                    type=str(type_bat),
                    culture=culture,
                    rayonnement=rayonnement,
                    boost_25=boost_25,
                    boost_50=boost_50,
                    boost_100=boost_100,
                    production=str(production)
                )
                
                # Ajouter la quantité spécifiée
                for _ in range(batiment.quantite):
                    batiments.append(copy.deepcopy(batiment))
                    
            except Exception as e:
                st.error(f"Erreur lors de la création du bâtiment à la ligne {i}: {str(e)}")
                continue
                
        return batiments
    
    def _get_plus_grand_batiment_restant(self) -> int:
        """Retourne la surface du plus grand bâtiment non placé"""
        if not self.batiments_non_places:
            return 0
            
        max_surface = 0
        for batiment in self.batiments_non_places:
            surface = batiment.longueur * batiment.largeur
            if surface > max_surface:
                max_surface = surface
        return max_surface
    
    def _trouver_emplacements(self, batiment: Batiment) -> List[Tuple[int, int, str]]:
        """Trouve tous les emplacements possibles pour un bâtiment"""
        emplacements = []
        
        # Essayer les deux orientations
        for orientation in ['horizontal', 'vertical']:
            longueur, largeur = batiment.get_dimensions(orientation)
            
            for x in range(self.terrain.hauteur):
                for y in range(self.terrain.largeur):
                    if self.terrain.est_libre(x, y, longueur, largeur):
                        # Vérifier si on n'a pas déjà essayé cet emplacement
                        key = (batiment.nom, x, y, orientation)
                        if key not in self.historique_placements.get(batiment.nom, set()):
                            emplacements.append((x, y, orientation))
        
        return emplacements
    
    def _verifier_espace_suffisant(self) -> bool:
        """Vérifie s'il reste assez de place pour le plus grand bâtiment restant"""
        if not self.batiments_non_places:
            return True
            
        plus_grand = self._get_plus_grand_batiment_restant()
        cases_libres = 0
        for i in range(self.terrain.hauteur):
            for j in range(self.terrain.largeur):
                if self.terrain.matrice[i, j] == 1 and self.terrain.occupation[i, j] == 0:
                    cases_libres += 1
        
        return cases_libres >= plus_grand
    
    def _calculer_score_placement(self, batiment: Batiment, x: int, y: int, orientation: str) -> float:
        """Calcule un score pour un placement (pour prioriser les bons emplacements)"""
        score = 0
        
        # Favoriser les bords pour les bâtiments neutres
        if batiment.type == 'neutre':
            longueur, largeur = batiment.get_dimensions(orientation)
            if x == 0 or y == 0 or x + longueur == self.terrain.hauteur or y + largeur == self.terrain.largeur:
                score += 100
        
        # Pour les producteurs, favoriser les zones avec culture
        if batiment.type == 'producteur' and hasattr(self.terrain, 'culture_map'):
            culture = self.terrain.get_culture_pour_batiment(batiment, x, y, orientation)
            score += culture
        
        return score
    
    def placer_batiment(self, batiment: Batiment) -> bool:
        """Tente de placer un bâtiment"""
        self.logger.ajouter(f"Évaluation du placement pour {batiment.nom}")
        
        # Trouver tous les emplacements possibles
        emplacements = self._trouver_emplacements(batiment)
        
        if not emplacements:
            self.logger.ajouter(f"Aucun emplacement trouvé pour {batiment.nom}")
            return False
        
        # Trier les emplacements par score
        emplacements_avec_score = []
        for x, y, orientation in emplacements:
            score = self._calculer_score_placement(batiment, x, y, orientation)
            emplacements_avec_score.append((score, x, y, orientation))
        
        emplacements_avec_score.sort(reverse=True)
        
        # Essayer chaque emplacement
        for score, x, y, orientation in emplacements_avec_score:
            # Vérifier l'espace restant
            self.terrain.placer_batiment(batiment, x, y, orientation)
            
            if self._verifier_espace_suffisant():
                # Placement réussi
                self.batiments_places.append((batiment, x, y, orientation))
                key = (batiment.nom, x, y, orientation)
                if batiment.nom not in self.historique_placements:
                    self.historique_placements[batiment.nom] = set()
                self.historique_placements[batiment.nom].add(key)
                
                self.logger.ajouter(f"✓ {batiment.nom} placé en ({x}, {y}) orientation {orientation}")
                return True
            else:
                # Pas assez de place, retirer le bâtiment
                self.terrain.enlever_batiment(batiment, x, y, orientation)
                self.logger.ajouter(f"✗ Placement de {batiment.nom} en ({x}, {y}) annulé - pas assez d'espace")
        
        return False
    
    def executer_placement(self):
        """Exécute l'algorithme de placement"""
        self.batiments_non_places = self.batiments.copy()
        
        # Compter les types pour debug
        types_count = {}
        for b in self.batiments_non_places:
            types_count[b.type] = types_count.get(b.type, 0) + 1
        st.write(f"DEBUG - Répartition après normalisation: {types_count}")
        
        # Séparer les bâtiments par type
        neutres = [b for b in self.batiments_non_places if b.type == 'neutre']
        culturels = [b for b in self.batiments_non_places if b.type == 'culturel']
        producteurs = [b for b in self.batiments_non_places if b.type == 'producteur']
        
        st.write(f"DEBUG - Neutres: {len(neutres)}, Culturels: {len(culturels)}, Producteurs: {len(producteurs)}")
        
        # Étape 1: Placer les neutres sur les bords
        self.logger.ajouter("=== PHASE 1: Placement des bâtiments neutres sur les bords ===")
        neutres_a_retirer = []
        for batiment in neutres:
            if self.placer_batiment(batiment):
                neutres_a_retirer.append(batiment)
        
        for batiment in neutres_a_retirer:
            self.batiments_non_places.remove(batiment)
        
        # Étape 2: Alterner culturels et producteurs
        self.logger.ajouter("=== PHASE 2: Placement alterné culturels/producteurs ===")
        
        iteration = 0
        max_iterations = 1000  # Pour éviter les boucles infinies
        
        while (culturels or producteurs) and iteration < max_iterations and not self.logger.arret:
            iteration += 1
            
            # Placer un culturel
            if culturels:
                batiment = culturels[0]
                if self.placer_batiment(batiment):
                    culturels.pop(0)
                    self.batiments_non_places.remove(batiment)
                    
                    # Mettre à jour la carte de culture
                    batiments_culturels = [(b, x, y, o) for b, x, y, o in self.batiments_places if b.type == 'culturel']
                    self.terrain.calculer_culture_recue(batiments_culturels)
            
            # Placer un producteur
            if producteurs and not self.logger.arret:
                batiment = producteurs[0]
                if self.placer_batiment(batiment):
                    producteurs.pop(0)
                    self.batiments_non_places.remove(batiment)
        
        self.logger.ajouter(f"=== PLACEMENT TERMINÉ ===")
        self.logger.ajouter(f"Bâtiments placés: {len(self.batiments_places)}")
        self.logger.ajouter(f"Bâtiments non placés: {len(self.batiments_non_places)}")
    
    def calculer_statistiques(self) -> Dict:
        """Calcule les statistiques de placement"""
        stats = {
            'culture_totale': 0,
            'boosts': {'25%': 0, '50%': 0, '100%': 0, '0%': 0},
            'culture_par_type': {},
            'cases_non_utilisees': 0,
            'surface_non_placee': sum(b.longueur * b.largeur for b in self.batiments_non_places)
        }
        
        # Compter les cases non utilisées
        for i in range(self.terrain.hauteur):
            for j in range(self.terrain.largeur):
                if self.terrain.matrice[i, j] == 1 and self.terrain.occupation[i, j] == 0:
                    stats['cases_non_utilisees'] += 1
        
        # Calculer la culture reçue par les producteurs
        for batiment, x, y, orientation in self.batiments_places:
            if batiment.type == 'producteur':
                culture = self.terrain.get_culture_pour_batiment(batiment, x, y, orientation)
                stats['culture_totale'] += culture
                
                # Déterminer le boost
                boost = self.terrain.get_boost_niveau(culture, batiment)
                stats['boosts'][boost] = stats['boosts'].get(boost, 0) + 1
                
                # Culture par type de production
                if batiment.production and not pd.isna(batiment.production) and batiment.production != '':
                    prod_type = str(batiment.production)
                    if prod_type not in stats['culture_par_type']:
                        stats['culture_par_type'][prod_type] = 0
                    stats['culture_par_type'][prod_type] += culture
        
        return stats


# ============================================================================
# FONCTIONS POUR L'EXPORT EXCEL
# ============================================================================

def create_output_excel(placement, stats, logs):
    """Crée le fichier Excel de résultats"""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Journal
        df_logs = pd.DataFrame({'Journal': logs})
        df_logs.to_excel(writer, sheet_name='Journal', index=False)
        
        # 2. Statistiques
        stats_data = {
            'Métrique': [
                'Culture totale reçue',
                'Cases non utilisées',
                'Surface des bâtiments non placés',
                'Nombre de bâtiments non placés'
            ],
            'Valeur': [
                stats['culture_totale'],
                stats['cases_non_utilisees'],
                stats['surface_non_placee'],
                len(placement.batiments_non_places)
            ]
        }
        df_stats = pd.DataFrame(stats_data)
        df_stats.to_excel(writer, sheet_name='Statistiques', index=False)
        
        # 3. Détail des boosts
        boosts_data = {
            'Niveau de boost': list(stats['boosts'].keys()),
            'Nombre de bâtiments': list(stats['boosts'].values())
        }
        df_boosts = pd.DataFrame(boosts_data)
        df_boosts.to_excel(writer, sheet_name='Boosts', index=False)
        
        # 4. Culture par type de production
        if stats['culture_par_type']:
            culture_type_data = {
                'Type de production': list(stats['culture_par_type'].keys()),
                'Culture reçue': list(stats['culture_par_type'].values())
            }
            df_culture_type = pd.DataFrame(culture_type_data)
            df_culture_type.to_excel(writer, sheet_name='Culture par type', index=False)
        
        # 5. Bâtiments non placés
        if placement.batiments_non_places:
            non_places_data = []
            for b in placement.batiments_non_places:
                non_places_data.append({
                    'Nom': b.nom,
                    'Type': b.type,
                    'Longueur': b.longueur,
                    'Largeur': b.largeur,
                    'Surface': b.longueur * b.largeur
                })
            df_non_places = pd.DataFrame(non_places_data)
            df_non_places.to_excel(writer, sheet_name='Non placés', index=False)
        
        # 6. Bâtiments placés avec leurs boosts
        places_data = []
        for batiment, x, y, orientation in placement.batiments_places:
            if batiment.type == 'producteur':
                culture = placement.terrain.get_culture_pour_batiment(batiment, x, y, orientation)
                boost = placement.terrain.get_boost_niveau(culture, batiment)
            else:
                culture = 0
                boost = '-'
                
            places_data.append({
                'Nom': batiment.nom,
                'Type': batiment.type,
                'Position X': x,
                'Position Y': y,
                'Orientation': orientation,
                'Culture reçue': culture if batiment.type == 'producteur' else '-',
                'Boost': boost if batiment.type == 'producteur' else '-'
            })
        
        if places_data:
            df_places = pd.DataFrame(places_data)
            df_places.to_excel(writer, sheet_name='Bâtiments placés', index=False)
    
    output.seek(0)
    return output


# ============================================================================
# INTERFACE STREAMLIT
# ============================================================================

def main():
    """Fonction principale de l'application Streamlit"""
    
    st.set_page_config(page_title="Placement de Bâtiments", layout="wide")
    
    st.title("🏗️ Optimiseur de Placement de Bâtiments")
    st.markdown("---")

    # Sidebar pour le chargement du fichier
    with st.sidebar:
        st.header("📁 Chargement des données")
        uploaded_file = st.file_uploader(
            "Choisir le fichier Excel", 
            type=['xlsx', 'xls'],
            help="Le fichier doit contenir deux onglets: 'Terrain' et 'Batiments'"
        )
        
        if uploaded_file:
            st.success("✅ Fichier chargé avec succès!")
        
        st.markdown("---")
        st.markdown("### ℹ️ Instructions")
        st.markdown("""
        1. Préparez un fichier Excel avec 2 onglets
        2. Onglet 1: **Terrain** (matrice de 0/1)
        3. Onglet 2: **Batiments** (liste des bâtiments)
        4. Lancez le placement
        5. Téléchargez les résultats
        """)

    # Traitement principal
    if uploaded_file:
        try:
            # Lecture des données
            excel_file = pd.ExcelFile(uploaded_file)
            
            # Vérifier le nombre d'onglets
            if len(excel_file.sheet_names) < 2:
                st.error("Le fichier doit contenir au moins 2 onglets")
                return
            
            df_terrain = pd.read_excel(uploaded_file, sheet_name=0, header=None)
            df_batiments = pd.read_excel(uploaded_file, sheet_name=1)
            
            st.subheader("📊 Aperçu des données")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.write("**Terrain** (extrait):")
                st.dataframe(df_terrain.head(), use_container_width=True)
                st.write(f"Dimensions: {df_terrain.shape[0]} lignes x {df_terrain.shape[1]} colonnes")
                st.write(f"Cases libres (1): {np.sum(df_terrain.values == 1)}")
                st.write(f"Cases occupées (0): {np.sum(df_terrain.values == 0)}")
            
            with col2:
                st.write("**Bâtiments** (extrait):")
                st.dataframe(df_batiments.head(), use_container_width=True)
                st.write(f"Colonnes: {list(df_batiments.columns)}")
                st.write(f"Nombre de types de bâtiments: {len(df_batiments)}")
                
                # Calculer le nombre total de bâtiments à placer
                total_batiments = 0
                if 'Quantite' in df_batiments.columns:
                    total_batiments = df_batiments['Quantite'].sum()
                    st.write(f"Total de bâtiments à placer: {int(total_batiments)}")
            
            # Bouton pour lancer le placement
            st.markdown("---")
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                bouton_lancer = st.button("🚀 LANCER LE PLACEMENT", type="primary", use_container_width=True)
            
            if bouton_lancer:
                with st.spinner("Placement en cours... (cela peut prendre quelques instants)"):
                    # Convertir les données
                    terrain_data = df_terrain.values.tolist()
                    batiments_data = df_batiments.to_dict('records')
                    
                    # Exécuter le placement
                    placement = PlacementBatiments(terrain_data, batiments_data)
                    placement.executer_placement()
                    
                    # Calculer les statistiques
                    stats = placement.calculer_statistiques()
                    
                    # Afficher les résultats
                    st.markdown("---")
                    st.subheader("✅ Résultats du placement")
                    
                    # Métriques principales
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        st.metric("Bâtiments placés", len(placement.batiments_places))
                    
                    with col2:
                        st.metric("Bâtiments non placés", len(placement.batiments_non_places))
                    
                    with col3:
                        st.metric("Culture totale", f"{stats['culture_totale']:.1f}")
                    
                    with col4:
                        st.metric("Cases libres restantes", stats['cases_non_utilisees'])
                    
                    # Statistiques détaillées
                    st.markdown("### 📊 Détail des boosts")
                    boosts_cols = st.columns(4)
                    boost_labels = ["0%", "25%", "50%", "100%"]
                    for i, label in enumerate(boost_labels):
                        with boosts_cols[i]:
                            st.metric(f"Boost {label}", stats['boosts'].get(label, 0))
                    
                    # Visualisation
                    st.markdown("### 🗺️ Carte du terrain")
                    
                    # Créer la visualisation
                    fig, ax = plt.subplots(figsize=(min(15, placement.terrain.largeur * 0.5), 
                                                   min(10, placement.terrain.hauteur * 0.5)))
                    
                    # Créer une matrice de couleurs
                    color_map = np.zeros((placement.terrain.hauteur, placement.terrain.largeur, 3))
                    
                    # Cases libres non occupées (blanc)
                    for i in range(placement.terrain.hauteur):
                        for j in range(placement.terrain.largeur):
                            if placement.terrain.matrice[i, j] == 1 and placement.terrain.occupation[i, j] == 0:
                                color_map[i, j] = [1, 1, 1]
                            elif placement.terrain.matrice[i, j] == 0:
                                color_map[i, j] = [0, 0, 0]
                    
                    # Colorier les bâtiments placés
                    colors = {
                        'culturel': [1, 0.5, 0],  # orange
                        'producteur': [0, 1, 0],   # vert
                        'neutre': [0.5, 0.5, 0.5]  # gris
                    }
                    
                    for batiment, x, y, orientation in placement.batiments_places:
                        longueur, largeur = batiment.get_dimensions(orientation)
                        color = colors.get(batiment.type, [0.5, 0.5, 0.5])
                        
                        for i in range(x, x + longueur):
                            for j in range(y, y + largeur):
                                if i < placement.terrain.hauteur and j < placement.terrain.largeur:
                                    color_map[i, j] = color
                        
                        # Ajouter le nom du bâtiment (si assez de place)
                        if longueur > 0 and largeur > 0:
                            ax.text(j - largeur/2 + 0.5, i - longueur/2 + 0.5, batiment.nom, 
                                   ha='center', va='center', fontsize=6, fontweight='bold',
                                   color='black', bbox=dict(facecolor='white', alpha=0.7, edgecolor='none', pad=1))
                    
                    ax.imshow(color_map, interpolation='nearest', origin='upper')
                    ax.set_title('Placement des bâtiments', fontsize=14, fontweight='bold')
                    ax.grid(True, alpha=0.3, color='black', linewidth=0.5)
                    
                    # Ajouter une légende
                    legend_elements = [
                        patches.Patch(color='orange', label='Culturel'),
                        patches.Patch(color='green', label='Producteur'),
                        patches.Patch(color='gray', label='Neutre'),
                        patches.Patch(color='black', label='Case occupée'),
                        patches.Patch(color='white', label='Case libre')
                    ]
                    ax.legend(handles=legend_elements, loc='upper right', fontsize=8)
                    
                    st.pyplot(fig)
                    plt.close(fig)
                    
                    # Journal
                    with st.expander("📋 Journal des opérations", expanded=False):
                        logs_text = "\n".join(placement.logger.entries[-100:])  # Afficher les 100 dernières
                        st.text_area("Journal", logs_text, height=300)
                    
                    # Téléchargement des résultats
                    st.markdown("---")
                    st.subheader("📥 Télécharger les résultats")
                    
                    output_excel = create_output_excel(placement, stats, placement.logger.entries)
                    
                    st.download_button(
                        label="📊 Télécharger le rapport Excel complet",
                        data=output_excel,
                        file_name="resultats_placement.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement: {str(e)}")
            st.write("Détails de l'erreur:")
            st.code(traceback.format_exc())
            st.info("Vérifiez que votre fichier Excel a le bon format")

    else:
        # Afficher un exemple quand aucun fichier n'est chargé
        st.info("👈 Veuillez charger un fichier Excel pour commencer")
        
        with st.expander("📝 Structure du fichier attendue", expanded=True):
            st.markdown("""
            ### Format du fichier Excel attendu
            
            **Onglet 1: Terrain** (sans en-tête)
            - Matrice de 0 et 1
            - 1 = case libre
            - 0 = case occupée
            
            **Onglet 2: Bâtiments** (avec en-têtes)
            - **Nom**: nom du bâtiment
            - **Longueur**: longueur en cases
            - **Largeur**: largeur en cases
            - **Quantite**: nombre d'exemplaires
            - **Type**: 'culturel', 'producteur' ou vide pour neutre
            - **Culture**: quantité de culture produite
            - **Rayonnement**: zone d'influence en cases
            - **Boost 25%**: seuil pour boost 25%
            - **Boost 50%**: seuil pour boost 50%
            - **Boost 100%**: seuil pour boost 100%
            - **Production**: type de production (ex: 'nourriture', 'or', etc.)
            """)


# ============================================================================
# POINT D'ENTRÉE
# ============================================================================

if __name__ == "__main__":
    main()