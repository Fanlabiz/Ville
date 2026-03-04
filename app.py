import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
import io
import plotly.graph_objects as go
from copy import deepcopy

# Configuration de la page
st.set_page_config(
    page_title="Placement de Bâtiments",
    page_icon="🏗️",
    layout="wide"
)

##############################
# CLASSE DE PLACEMENT
##############################
class PlacementBatiments:
    def __init__(self, terrain, batiments):
        self.terrain_original = terrain.copy()
        # Convertir le terrain en tableau d'objets (pour pouvoir mélanger nombres et chaînes)
        self.terrain = terrain.astype(object).copy()
        self.hauteur, self.largeur = terrain.shape
        self.batiments = batiments
        self.batiments_places = []
        self.carte_rayonnement = np.zeros((self.hauteur, self.largeur))
        self.production_priority = {'Guerison': 1, 'Nourriture': 2, 'Or': 3, '': 4}
    
    def est_case_libre(self, i, j):
        """
        Vérifie si une case est libre (valeur 1)
        """
        return self.terrain[i, j] == 1
    
    def est_case_occupee_par_batiment(self, i, j):
        """
        Vérifie si une case est occupée par un bâtiment (string)
        """
        return isinstance(self.terrain[i, j], str)
    
    def calculer_zone_rayonnement(self, x, y, longueur, largeur, rayonnement, valeur_culture):
        """
        Calcule la zone de rayonnement autour d'un bâtiment culturel
        """
        zone = []
        # Déterminer les limites de la zone de rayonnement
        x_min = max(0, x - rayonnement)
        x_max = min(self.hauteur, x + longueur + rayonnement)
        y_min = max(0, y - rayonnement)
        y_max = min(self.largeur, y + largeur + rayonnement)
        
        for i in range(x_min, x_max):
            for j in range(y_min, y_max):
                # Vérifier si la case est dans le rayonnement (pas à l'intérieur du bâtiment)
                if (i < x or i >= x + longueur or j < y or j >= y + largeur):
                    # Vérifier si la case est libre (1) ou contient déjà un bâtiment (string)
                    if self.est_case_libre(i, j) or self.est_case_occupee_par_batiment(i, j):
                        zone.append((i, j))
                        self.carte_rayonnement[i, j] += valeur_culture
        
        return zone
    
    def peut_placer_batiment(self, x, y, longueur, largeur):
        """
        Vérifie si un bâtiment peut être placé à la position (x,y)
        """
        if x + longueur > self.hauteur or y + largeur > self.largeur:
            return False
        
        # Vérifier que toutes les cases sont libres (valeur 1)
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                if not self.est_case_libre(i, j):  # Pas libre
                    return False
        return True
    
    def verifier_bande_1(self, x, y, longueur, largeur):
        """
        Vérifie si le placement crée des bandes de largeur 1
        """
        # Vérifier les bordures
        if x == 0 or x + longueur == self.hauteur or y == 0 or y + largeur == self.largeur:
            # Vérifier les bandes contre les bords
            if x > 0:
                # Vérifier bande horizontale au-dessus
                bande_valide = False
                for j in range(y, y + largeur):
                    if x - 1 >= 0 and (not self.est_case_libre(x-1, j) or j == 0 or j == self.largeur-1):
                        bande_valide = True
                        break
                if not bande_valide:
                    return False
            
            if x + longueur < self.hauteur:
                # Vérifier bande horizontale en-dessous
                bande_valide = False
                for j in range(y, y + largeur):
                    if x + longueur < self.hauteur and (not self.est_case_libre(x+longueur, j) or j == 0 or j == self.largeur-1):
                        bande_valide = True
                        break
                if not bande_valide:
                    return False
            
            if y > 0:
                # Vérifier bande verticale à gauche
                bande_valide = False
                for i in range(x, x + longueur):
                    if y - 1 >= 0 and (not self.est_case_libre(i, y-1) or i == 0 or i == self.hauteur-1):
                        bande_valide = True
                        break
                if not bande_valide:
                    return False
            
            if y + largeur < self.largeur:
                # Vérifier bande verticale à droite
                bande_valide = False
                for i in range(x, x + longueur):
                    if y + largeur < self.largeur and (not self.est_case_libre(i, y+largeur) or i == 0 or i == self.hauteur-1):
                        bande_valide = True
                        break
                if not bande_valide:
                    return False
        
        return True
    
    def calculer_culture_pour_position(self, x, y, longueur, largeur):
        """
        Calcule la culture totale reçue par un bâtiment à une position donnée
        """
        culture_totale = 0
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                culture_totale += self.carte_rayonnement[i, j]
        return culture_totale
    
    def trouver_meilleur_placement_culturel(self, batiment):
        """
        Trouve le meilleur placement pour un bâtiment culturel
        """
        meilleur_score = -1
        meilleur_placement = None
        
        # Essayer les deux orientations
        orientations = [(batiment['longueur'], batiment['largeur'])]
        if batiment['longueur'] != batiment['largeur']:
            orientations.append((batiment['largeur'], batiment['longueur']))
        
        for longueur, largeur in orientations:
            for i in range(self.hauteur - longueur + 1):
                for j in range(self.largeur - largeur + 1):
                    if self.peut_placer_batiment(i, j, longueur, largeur):
                        # Vérifier la contrainte de bande de largeur 1
                        if not self.verifier_bande_1(i, j, longueur, largeur):
                            continue
                        
                        # Calculer la zone de rayonnement
                        zone_temp = []
                        for ii in range(max(0, i - batiment['rayonnement']), 
                                       min(self.hauteur, i + longueur + batiment['rayonnement'])):
                            for jj in range(max(0, j - batiment['rayonnement']), 
                                          min(self.largeur, j + largeur + batiment['rayonnement'])):
                                if (ii < i or ii >= i + longueur or jj < j or jj >= j + largeur):
                                    if self.est_case_libre(ii, jj) or self.est_case_occupee_par_batiment(ii, jj):
                                        zone_temp.append((ii, jj))
                        
                        # Compter combien de producteurs déjà placés sont dans la zone
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
        """
        Trouve le meilleur placement pour un bâtiment producteur
        """
        meilleur_score = -1
        meilleur_placement = None
        meilleure_culture = -1
        
        # Essayer les deux orientations
        orientations = [(batiment['longueur'], batiment['largeur'])]
        if batiment['longueur'] != batiment['largeur']:
            orientations.append((batiment['largeur'], batiment['longueur']))
        
        for longueur, largeur in orientations:
            for i in range(self.hauteur - longueur + 1):
                for j in range(self.largeur - largeur + 1):
                    if self.peut_placer_batiment(i, j, longueur, largeur):
                        if not self.verifier_bande_1(i, j, longueur, largeur):
                            continue
                        
                        culture = self.calculer_culture_pour_position(i, j, longueur, largeur)
                        
                        # Priorité à la culture reçue, puis à la position
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
        
        # Placer le bâtiment
        for i in range(x, x + longueur):
            for j in range(y, y + largeur):
                self.terrain[i, j] = f"{nom_batiment[:3]}_{i}_{j}"
        
        # Si c'est un bâtiment culturel, calculer sa zone de rayonnement
        if batiment['type'] == 'culturel':
            self.calculer_zone_rayonnement(x, y, longueur, largeur, 
                                          batiment['rayonnement'], 
                                          batiment['culture'])
        
        # Déterminer le boost atteint
        boost = '0%'
        if batiment['type'] == 'producteur' and culture_recue > 0:
            if culture_recue >= batiment['boost_100']:
                boost = '100%'
            elif culture_recue >= batiment['boost_50']:
                boost = '50%'
            elif culture_recue >= batiment['boost_25']:
                boost = '25%'
        
        # Enregistrer le bâtiment placé
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
            'production': batiment['production']
        }
        self.batiments_places.append(batiment_place)
        
        return batiment_place
    
    def executer_placement(self):
        """
        Exécute l'algorithme de placement
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
        
        # Alterner le placement
        index_producteur = 0
        iteration = 0
        max_iterations = 100  # Sécurité pour éviter les boucles infinies
        
        while (culturels or index_producteur < len(producteurs)) and iteration < max_iterations:
            iteration += 1
            
            # Étape 1: Placer un bâtiment culturel
            if culturels:
                batiment = culturels.pop(0)
                placement = self.trouver_meilleur_placement_culturel(batiment)
                if placement:
                    x, y, longueur, largeur = placement
                    self.placer_batiment(batiment, x, y, longueur, largeur)
            
            # Étape 2: Placer des bâtiments producteurs
            places_dans_cette_iteration = 0
            while index_producteur < len(producteurs) and places_dans_cette_iteration < 3:
                batiment = producteurs[index_producteur]
                placement = self.trouver_meilleur_placement_producteur(batiment)
                
                if placement:
                    x, y, longueur, largeur, culture = placement
                    self.placer_batiment(batiment, x, y, longueur, largeur, culture)
                    index_producteur += 1
                    places_dans_cette_iteration += 1
                else:
                    # Si on ne peut pas placer ce producteur, on passe au suivant
                    index_producteur += 1
        
        return self.terrain, self.batiments_places
    
    def calculer_statistiques(self):
        """
        Calcule les statistiques finales
        """
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
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Onglet 1: Terrain original
        df_terrain_original = pd.DataFrame(terrain_original)
        df_terrain_original.to_excel(writer, sheet_name='Terrain_Original', index=False, header=False)
        
        # Onglet 2: Terrain avec bâtiments placés
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
        df_terrain_place.to_excel(writer, sheet_name='Terrain_Place', index=False, header=False)
        
        # Onglet 3: Liste des bâtiments placés
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
        df_placement.to_excel(writer, sheet_name='Placements', index=False)
        
        # Onglet 4: Bâtiments non placés
        # Créer un dictionnaire des quantités placées par type de bâtiment
        quantites_placees = {}
        for bat in batiments_places:
            if bat['nom'] not in quantites_placees:
                quantites_placees[bat['nom']] = 0
            quantites_placees[bat['nom']] += 1
        
        # Lister les bâtiments non placés
        non_places = []
        for bat in tous_les_batiments:
            nom = bat['nom']
            quantite_demandee = int(bat['quantite'])
            quantite_placee = quantites_placees.get(nom, 0)
            
            if quantite_placee < quantite_demandee:
                non_places.append({
                    'Nom': nom,
                    'Type': bat['type'],
                    'Longueur': bat['longueur'],
                    'Largeur': bat['largeur'],
                    'Quantite_demandee': quantite_demandee,
                    'Quantite_placee': quantite_placee,
                    'Reste_a_placer': quantite_demandee - quantite_placee,
                    'Culture': bat['culture'] if bat['type'] == 'culturel' else 'N/A',
                    'Rayonnement': bat['rayonnement'] if bat['type'] == 'culturel' else 'N/A',
                    'Production': bat['production'] if bat['type'] == 'producteur' else 'N/A'
                })
        
        df_non_places = pd.DataFrame(non_places)
        if not df_non_places.empty:
            df_non_places.to_excel(writer, sheet_name='Non_Places', index=False)
        
        # Onglet 5: Statistiques
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
        
        df_stats = pd.DataFrame(stats_data)
        df_stats.to_excel(writer, sheet_name='Statistiques', index=False)
        
        # Application des couleurs après avoir écrit tous les onglets
        workbook = writer.book
        
        # Définir les styles de couleur
        style_vert = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
        style_orange = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
        style_gris = workbook.add_format({'bg_color': '#F2F2F2', 'font_color': '#333333'})
        
        # Colorer l'onglet Terrain_Place
        worksheet_terrain = writer.sheets['Terrain_Place']
        for i in range(df_terrain_place.shape[0]):
            for j in range(df_terrain_place.shape[1]):
                cellule = str(df_terrain_place.iat[i, j])
                if cellule not in ['LIBRE', 'OCCUPE'] and '_' in cellule:
                    # C'est un bâtiment, chercher son type
                    nom_bat = cellule.split('_')[0]
                    for bat in batiments_places:
                        if bat['nom'][:3] == nom_bat or (len(bat['nom']) >= 3 and bat['nom'][:3] == nom_bat):
                            if bat['type'] == 'producteur':
                                worksheet_terrain.write(i, j, cellule, style_vert)
                            elif bat['type'] == 'culturel':
                                worksheet_terrain.write(i, j, cellule, style_orange)
                            else:
                                worksheet_terrain.write(i, j, cellule, style_gris)
                            break
        
        # Colorer l'onglet Placements
        worksheet_placement = writer.sheets['Placements']
        
        # Trouver l'index de la colonne Type
        type_col_idx = None
        for idx, col in enumerate(df_placement.columns):
            if col == 'Type':
                type_col_idx = idx
                break
        
        if type_col_idx is not None:
            for i, row in df_placement.iterrows():
                if row['Type'] == 'producteur':
                    worksheet_placement.write(i+1, type_col_idx, row['Type'], style_vert)
                elif row['Type'] == 'culturel':
                    worksheet_placement.write(i+1, type_col_idx, row['Type'], style_orange)
                else:
                    worksheet_placement.write(i+1, type_col_idx, row['Type'], style_gris)
        
        # Colorer l'onglet Non_Places si il existe
        if not df_non_places.empty:
            worksheet_non_places = writer.sheets['Non_Places']
            
            # Trouver l'index de la colonne Type
            type_non_places_idx = None
            for idx, col in enumerate(df_non_places.columns):
                if col == 'Type':
                    type_non_places_idx = idx
                    break
            
            if type_non_places_idx is not None:
                for i, row in df_non_places.iterrows():
                    if row['Type'] == 'producteur':
                        worksheet_non_places.write(i+1, type_non_places_idx, row['Type'], style_vert)
                    elif row['Type'] == 'culturel':
                        worksheet_non_places.write(i+1, type_non_places_idx, row['Type'], style_orange)
                    else:
                        worksheet_non_places.write(i+1, type_non_places_idx, row['Type'], style_gris)
    
    output.seek(0)
    return output


##############################
# INTERFACE STREAMLIT
##############################

# Titre
st.title("🏗️ Optimiseur de Placement de Bâtiments")
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
            st.info(f"📦 Total de bâtiments à placer: {total_batiments}")
        
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
st.markdown("🚀 Application développée pour l'optimisation de placement de bâtiments")