import pandas as pd
import numpy as np
import streamlit as st
import io
import copy
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# --- CONFIGURATION ---
COLORS = {
    'Culturel': 'FFFFA500',  # Orange
    'Producteur': 'FF008000', # Vert
    'Neutre': 'FF808080'      # Gris
}

class CityPlanner:
    def __init__(self, terrain_data, actuel_data=None):
        """
        Initialise le planner avec le terrain de base et optionnellement
        une configuration actuelle (bâtiments déjà placés)
        """
        # Initialiser le journal et les paramètres AVANT tout
        self.journal = []
        self.max_entries = 10000 
        self.interrupted = False
        self.best_solution = None
        self.best_score = -1
        
        self.rows = len(terrain_data)
        self.cols = len(terrain_data[0])
        
        # Grille des cases libres (1 = libre, 0 = occupée par le terrain)
        self.base_grid = np.zeros((self.rows, self.cols))
        self.border_mask = np.zeros((self.rows, self.cols), dtype=bool)
        self.initial_free_cells = 0
        
        for r in range(self.rows):
            for c in range(self.cols):
                val = str(terrain_data[r][c]).strip().upper()
                if val == '1': 
                    self.base_grid[r, c] = 1
                    self.initial_free_cells += 1
                elif val == 'X':
                    self.base_grid[r, c] = 0 
                    self.border_mask[r, c] = True
        
        # Grille de travail (copie de la base)
        self.grid = self.base_grid.copy()
        
        # Bâtiments placés
        self.placed_buildings = []
        
        # Si une configuration actuelle est fournie, charger les bâtiments
        if actuel_data is not None:
            self.load_existing_buildings(actuel_data)

    def log(self, msg):
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)
        else:
            self.interrupted = True

    def load_existing_buildings(self, actuel_data):
        """
        Charge les bâtiments déjà placés depuis l'onglet "Actuel"
        Version avec débogage amélioré
        """
        # Créer une grille de visite pour marquer les cellules déjà traitées
        visited = np.zeros((self.rows, self.cols), dtype=bool)
        buildings_found = []
        
        # Convertir actuel_data en tableau de strings
        actuel_str = []
        for r in range(min(len(actuel_data), self.rows)):
            row = []
            for c in range(min(len(actuel_data[0]), self.cols)):
                val = str(actuel_data[r][c]).strip()
                row.append(val)
            actuel_str.append(row)
        
        # Afficher les 10 premières cellules pour déboguer
        st.write("Aperçu des 10 premières cellules de l'onglet Actuel:")
        sample = []
        for r in range(min(5, len(actuel_str))):
            for c in range(min(5, len(actuel_str[0]))):
                if actuel_str[r][c] not in ['', 'X', '1', '0', 'nan', 'None']:
                    sample.append(f"({r},{c}): '{actuel_str[r][c]}'")
        st.write(sample[:10])
        
        for r in range(len(actuel_str)):
            for c in range(len(actuel_str[0])):
                if visited[r, c]:
                    continue
                    
                cell_value = actuel_str[r][c]
                
                # Ignorer les cases vides, les X et les 1/0
                if cell_value in ['', 'X', '1', '0', 'nan', 'None']:
                    continue
                
                # Déterminer les dimensions du bâtiment
                # Méthode simple : explorer vers la droite et vers le bas
                # tant que le nom est identique
                
                # Largeur
                w = 1
                for dc in range(1, self.cols - c):
                    if c + dc < len(actuel_str[0]):
                        if actuel_str[r][c + dc] == cell_value and not visited[r, c + dc]:
                            w = dc + 1
                        else:
                            break
                
                # Hauteur
                h = 1
                for dr in range(1, self.rows - r):
                    if r + dr < len(actuel_str):
                        if actuel_str[r + dr][c] == cell_value and not visited[r + dr, c]:
                            h = dr + 1
                        else:
                            break
                
                # Vérifier que tout le rectangle est valide
                valid = True
                for dr in range(h):
                    for dc in range(w):
                        if r + dr >= len(actuel_str) or c + dc >= len(actuel_str[0]):
                            valid = False
                            break
                        if actuel_str[r + dr][c + dc] != cell_value or visited[r + dr, c + dc]:
                            valid = False
                            break
                    if not valid:
                        break
                
                if valid:
                    # Marquer toutes les cellules comme visitées
                    for dr in range(h):
                        for dc in range(w):
                            if r + dr < self.rows and c + dc < self.cols:
                                visited[r + dr, c + dc] = True
                    
                    # Marquer la grille comme occupée
                    self.grid[r:r+h, c:c+w] = 0
                    
                    # Enregistrer le bâtiment
                    building_info = {
                        'nom_temp': cell_value,
                        'r': r,
                        'c': c,
                        'w': w,
                        'h': h,
                        'info': None
                    }
                    self.placed_buildings.append(building_info)
                    buildings_found.append(f"{cell_value} à ({r},{c}) dimensions {w}x{h}")
                    
                    self.log(f"Bâtiment existant trouvé: {cell_value} à ({r},{c}) dimensions {w}x{h}")
        
        self.log(f"Chargement de {len(self.placed_buildings)} bâtiments depuis la configuration actuelle")
        
        # Afficher les bâtiments trouvés dans Streamlit
        st.write(f"🏗️ Bâtiments détectés: {len(buildings_found)}")
        with st.expander("Voir la liste des bâtiments détectés"):
            for b in buildings_found:
                st.text(b)
        
        return buildings_found

    def match_buildings_with_info(self, buildings_info):
        """
        Associe les bâtiments placés à leurs caractéristiques
        Correction : accepte les orientations différentes
        """
        # Créer un dictionnaire des bâtiments par nom avec leurs quantités
        building_pool = {}
        for b in buildings_info:
            nom = b['Nom']
            if nom not in building_pool:
                building_pool[nom] = []
            building_pool[nom].append(b)
        
        # Traiter chaque bâtiment placé
        temp_buildings = self.placed_buildings.copy()
        self.placed_buildings = []
        unmatched = []
        matched = []
        
        for temp in temp_buildings:
            nom = temp['nom_temp']
            r, c, w, h = temp['r'], temp['c'], temp['w'], temp['h']
            
            if nom in building_pool and building_pool[nom]:
                # Chercher un bâtiment qui correspond aux dimensions (en acceptant les orientations)
                found = False
                for i, info in enumerate(building_pool[nom]):
                    # Vérifier les deux orientations possibles
                    if (info['Largeur'] == w and info['Longueur'] == h) or \
                       (info['Largeur'] == h and info['Longueur'] == w):
                        # Correspondance trouvée
                        info = building_pool[nom].pop(i)
                        self.placed_buildings.append({
                            'info': info,
                            'r': r,
                            'c': c,
                            'w': w,
                            'h': h
                        })
                        self.log(f"Bâtiment associé: {nom} à ({r},{c}) dimensions {w}x{h} (orientation acceptée)")
                        matched.append(f"{nom} à ({r},{c}) dims {w}x{h}")
                        found = True
                        break
                
                if not found:
                    unmatched.append(f"{nom} à ({r},{c}) dimensions {w}x{h}")
                    self.log(f"⚠️ Attention: bâtiment '{nom}' dimensions {w}x{h} ne correspond à aucune orientation")
            else:
                unmatched.append(f"{nom} à ({r},{c})")
                self.log(f"⚠️ Attention: bâtiment '{nom}' non trouvé dans la liste ou quantité insuffisante")
        
        # Afficher les résultats dans Streamlit
        st.write(f"✅ Bâtiments associés: {len(matched)}")
        with st.expander("Voir les bâtiments associés"):
            for m in matched:
                st.text(m)
        
        if unmatched:
            st.warning(f"⚠️ {len(unmatched)} bâtiments n'ont pas pu être associés")
            with st.expander("Voir les bâtiments non associés"):
                for u in unmatched:
                    st.text(u)
        
        # Recompter les cases libres
        self.initial_free_cells = np.sum(self.grid == 1)
        
        return unmatched

    def is_adjacent_to_X(self, r, c, w, h):
        """Vérifie si le bâtiment touche un bord (case X)"""
        r_min, r_max = max(0, r-1), min(self.rows, r+h+1)
        c_min, c_max = max(0, c-1), min(self.cols, c+w+1)
        return np.any(self.border_mask[r_min:r_max, c_min:c_max])

    def can_place(self, r, c, w, h, remaining_queue):
        """Vérifie si on peut placer un bâtiment à l'emplacement (r,c) avec dimensions (w,h)"""
        if r + h > self.rows or c + w > self.cols:
            return False
        if not np.all(self.grid[r:r+h, c:c+w] == 1):
            return False
        
        if remaining_queue:
            self.grid[r:r+h, c:c+w] = 0
            biggest = remaining_queue[0]
            bw, bh = biggest['Largeur'], biggest['Longueur']
            
            can_fit_biggest = False
            for orient in [(bw, bh), (bh, bw)]:
                ow, oh = orient
                for rr in range(self.rows - oh + 1):
                    for cc in range(self.cols - ow + 1):
                        if np.all(self.grid[rr:rr+oh, cc:cc+ow] == 1):
                            can_fit_biggest = True
                            break
                    if can_fit_biggest: break
                if can_fit_biggest: break
            
            self.grid[r:r+h, c:c+w] = 1
            if not can_fit_biggest:
                return False

        return True

    def solve(self, buildings, phase="initial"):
        """
        Algorithme de placement récursif avec backtracking
        Ne place que les bâtiments non encore placés
        """
        if not buildings or self.interrupted:
            return True

        b = buildings[0]
        self.log(f"Évaluation de : {b['Nom']} (Type: {b['Type']}) - Phase: {phase}")
        
        dims = [(b['Largeur'], b['Longueur']), (b['Longueur'], b['Largeur'])]
        if b['Largeur'] == b['Longueur']: 
            dims = [dims[0]]

        # Phase 1 : Placement prioritaire (bords pour les Neutres)
        for w, h in dims:
            for r in range(self.rows - h + 1):
                for c in range(self.cols - w + 1):
                    if b['Type'] == 'Neutre' and not self.is_adjacent_to_X(r, c, w, h):
                        continue
                    if self.can_place(r, c, w, h, buildings[1:]):
                        if self.try_placement(b, r, c, w, h, buildings, phase):
                            return True

        # Phase 2 : Placement de secours (pour Neutres si bords saturés)
        if b['Type'] == 'Neutre':
            self.log(f"Bords saturés, recherche interne pour : {b['Nom']}")
            for w, h in dims:
                for r in range(self.rows - h + 1):
                    for c in range(self.cols - w + 1):
                        if self.can_place(r, c, w, h, buildings[1:]):
                            if self.try_placement(b, r, c, w, h, buildings, phase):
                                return True
        return False

    def try_placement(self, b, r, c, w, h, buildings, phase):
        """Tente de placer un bâtiment et continue la récursion"""
        self.grid[r:r+h, c:c+w] = 0
        self.placed_buildings.append({
            'info': b, 
            'r': r, 
            'c': c, 
            'w': w, 
            'h': h
        })
        self.log(f"✓ Placé : {b['Nom']} en ({r},{c}) orientation {w}x{h}")
        
        if self.solve(buildings[1:], phase): 
            return True
        
        if not self.interrupted:
            self.log(f"✗ Enlevé : {b['Nom']} de ({r},{c})")
            self.grid[r:r+h, c:c+w] = 1
            self.placed_buildings.pop()
        return False
    
    def calculate_culture_for_position(self, building, r, c, cultural_buildings):
        """
        Calcule la culture qu'un bâtiment recevrait à une position donnée.
        """
        prod_r_start = r
        prod_r_end = r + building['h']
        prod_c_start = c
        prod_c_end = c + building['w']
        
        cultures_recues = []
        
        for cb in cultural_buildings:
            rayon = int(cb['info'].get('Rayonnement', 0))
            culture_value = cb['info'].get('Culture', 0)
            
            if rayon > 0 and culture_value > 0:
                cult_r_start = cb['r']
                cult_r_end = cb['r'] + cb['h']
                cult_c_start = cb['c']
                cult_c_end = cb['c'] + cb['w']
                
                est_dans_zone = False
                
                for pr in range(prod_r_start, prod_r_end):
                    for pc in range(prod_c_start, prod_c_end):
                        for cr in range(cult_r_start, cult_r_end):
                            for cc in range(cult_c_start, cult_c_end):
                                distance = abs(pr - cr) + abs(pc - cc)
                                if distance <= rayon:
                                    est_dans_zone = True
                                    break
                            if est_dans_zone:
                                break
                        if est_dans_zone:
                            break
                    if est_dans_zone:
                        break
                
                if est_dans_zone:
                    cultures_recues.append(culture_value)
        
        return sum(cultures_recues)
    
    def calculate_boost_from_culture(self, building_info, culture):
        """
        Détermine le boost à partir de la culture reçue.
        """
        boost = 0
        boost_100 = building_info.get('Boost 100%')
        boost_50 = building_info.get('Boost 50%')
        boost_25 = building_info.get('Boost 25%')
        
        if pd.notna(boost_100) and culture >= boost_100:
            boost = 100
        elif pd.notna(boost_50) and culture >= boost_50:
            boost = 50
        elif pd.notna(boost_25) and culture >= boost_25:
            boost = 25
        
        return boost
    
    def calculate_score_from_boost(self, boost):
        """
        Convertit un boost en points de score.
        """
        if boost == 100:
            return 4
        elif boost == 50:
            return 2
        elif boost == 25:
            return 1
        else:
            return 0

    def calculate_culture_and_score(self):
        """
        Calcule la culture reçue et un score basé sur les boosts.
        """
        cultural_buildings = [pb for pb in self.placed_buildings if pb['info']['Type'] == 'Culturel']
        
        prod_stats = []
        prod_by_type = {"Guerison": 0, "Nourriture": 0, "Or": 0, "Bijoux": 0, "Onguents": 0, 
                        "Cristal": 0, "Epices": 0, "Boiseries": 0, "Scriberie": 0, "Pommades": 0}
        
        total_culture = 0
        score = 0
        boost_counts = {100: 0, 50: 0, 25: 0, 0: 0}
        
        for pb in self.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                culture_recue = self.calculate_culture_for_position(pb, pb['r'], pb['c'], cultural_buildings)
                
                boost = self.calculate_boost_from_culture(pb['info'], culture_recue)
                
                boost_counts[boost] += 1
                score += self.calculate_score_from_boost(boost)
                
                prod_stats.append({
                    'Nom': pb['info']['Nom'],
                    'Culture reçue': culture_recue,
                    'Boost': f"{boost}%",
                    'Production': pb['info'].get('Production', '')
                })
                
                total_culture += culture_recue
                
                prod_type = str(pb['info'].get('Production', ''))
                if prod_type in prod_by_type:
                    prod_by_type[prod_type] += culture_recue
        
        return {
            'prod_stats': prod_stats,
            'prod_by_type': prod_by_type,
            'total_culture': total_culture,
            'score': score,
            'boost_counts': boost_counts
        }

    def optimize_placement(self, iterations=10):
        """
        Phase d'optimisation simple : essaie de déplacer les producteurs
        vers de meilleures positions
        """
        self.log("\n=== DÉBUT DE LA PHASE D'OPTIMISATION ===\n")
        
        # Sauvegarder la solution initiale
        initial_solution = {
            'grid': self.grid.copy(),
            'placed_buildings': copy.deepcopy(self.placed_buildings)
        }
        
        # Calculer le score initial
        initial_results = self.calculate_culture_and_score()
        self.best_score = initial_results['score']
        self.best_solution = initial_solution
        
        self.log(f"Score initial: {self.best_score} (Boosts: 100%:{initial_results['boost_counts'][100]}, 50%:{initial_results['boost_counts'][50]}, 25%:{initial_results['boost_counts'][25]})")
        
        # Identifier les bâtiments
        cultural_buildings = [b for b in self.placed_buildings if b['info']['Type'] == 'Culturel']
        producer_buildings = [b for b in self.placed_buildings if b['info']['Type'] == 'Producteur']
        
        # Quelques itérations d'optimisation
        for iteration in range(iterations):
            if self.interrupted:
                break
            
            self.log(f"\n--- Itération d'optimisation {iteration + 1} ---")
            improved = False
            
            # Pour chaque producteur, essayer de trouver une meilleure position
            for prod in producer_buildings:
                if self.interrupted:
                    break
                
                current_culture = self.calculate_culture_for_position(prod, prod['r'], prod['c'], cultural_buildings)
                
                # Sauvegarder l'état
                old_r, old_c = prod['r'], prod['c']
                old_w, old_h = prod['w'], prod['h']
                
                # Enlever le producteur
                self.grid[old_r:old_r+old_h, old_c:old_c+old_w] = 1
                
                # Chercher une meilleure position
                best_pos = None
                best_culture = current_culture
                
                for r in range(self.rows - old_h + 1):
                    for c in range(self.cols - old_w + 1):
                        if r == old_r and c == old_c:
                            continue
                        if np.all(self.grid[r:r+old_h, c:c+old_w] == 1):
                            culture_here = self.calculate_culture_for_position(prod, r, c, cultural_buildings)
                            if culture_here > best_culture:
                                best_culture = culture_here
                                best_pos = (r, c)
                
                if best_pos:
                    r, c = best_pos
                    self.grid[r:r+old_h, c:c+old_w] = 0
                    prod['r'], prod['c'] = r, c
                    
                    # Recalculer le score
                    new_results = self.calculate_culture_and_score()
                    
                    if new_results['score'] > self.best_score:
                        self.best_score = new_results['score']
                        self.best_solution = {
                            'grid': self.grid.copy(),
                            'placed_buildings': copy.deepcopy(self.placed_buildings)
                        }
                        self.log(f"  ✓ Producteur {prod['info']['Nom']} déplacé de ({old_r},{old_c}) à ({r},{c})")
                        self.log(f"    Nouveau score: {self.best_score}")
                        improved = True
                    else:
                        # Annuler le déplacement
                        self.grid[r:r+old_h, c:c+old_w] = 1
                        self.grid[old_r:old_r+old_h, old_c:old_c+old_w] = 0
                        prod['r'], prod['c'] = old_r, old_c
                else:
                    # Remettre le producteur
                    self.grid[old_r:old_r+old_h, old_c:old_c+old_w] = 0
            
            if not improved:
                self.log("  Aucune amélioration trouvée")
        
        # Restaurer la meilleure solution
        if self.best_solution:
            self.grid = self.best_solution['grid'].copy()
            self.placed_buildings = copy.deepcopy(self.best_solution['placed_buildings'])
            final_results = self.calculate_culture_and_score()
            self.log(f"\n=== OPTIMISATION TERMINÉE ===")
            self.log(f"Score final: {final_results['score']}")

# --- LOGIQUE D'EXPORT EXCEL ---
def generate_excel(planner, all_buildings):
    """Génère le fichier Excel de résultat"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Journal
        pd.DataFrame(planner.journal, columns=["Journal"]).to_excel(writer, sheet_name="Journal", index=False)
        
        # 2. Calcul de la culture
        culture_results = planner.calculate_culture_and_score()
        
        # 3. Statistiques de production
        prod_df = pd.DataFrame(culture_results['prod_stats'])
        if not prod_df.empty:
            prod_df.to_excel(writer, sheet_name="Production", index=False)
        
        # 4. Synthèse par type de production
        summary_types = pd.DataFrame([
            {"Type de Production": k, "Culture Totale": v} 
            for k, v in culture_results['prod_by_type'].items() if v > 0
        ])
        if not summary_types.empty:
            summary_types.to_excel(writer, sheet_name="Synthese_Production", index=False)
        
        # 5. Statistiques des boosts
        boost_stats = pd.DataFrame([
            {"Boost": "100%", "Nombre": culture_results['boost_counts'][100]},
            {"Boost": "50%", "Nombre": culture_results['boost_counts'][50]},
            {"Boost": "25%", "Nombre": culture_results['boost_counts'][25]},
            {"Boost": "0%", "Nombre": culture_results['boost_counts'][0]}
        ])
        boost_stats.to_excel(writer, sheet_name="Boosts", index=False)
        
        # 6. Création du plan du terrain avec les noms complets
        ws = writer.book.create_sheet("Plan_Terrain")
        
        for r in range(planner.rows):
            for c in range(planner.cols):
                cell = ws.cell(row=r+1, column=c+1)
                
                if planner.border_mask[r, c]:
                    cell.value = "X"
                    cell.fill = PatternFill(start_color='FF000000', fill_type='solid')
                else:
                    cell_value = ""
                    cell_color = None
                    
                    for pb in planner.placed_buildings:
                        if pb['r'] <= r < pb['r'] + pb['h'] and pb['c'] <= c < pb['c'] + pb['w']:
                            cell_value = pb['info']['Nom']
                            cell_color = COLORS.get(pb['info']['Type'], 'FFFFFFFF')
                            break
                    
                    if cell_value:
                        cell.value = cell_value
                        if cell_color:
                            cell.fill = PatternFill(start_color=cell_color, fill_type='solid')
        
        for col in range(1, planner.cols + 1):
            column_letter = get_column_letter(col)
            ws.column_dimensions[column_letter].width = 20
        
        # 7. Résumé
        # Compter les bâtiments placés par nom
        placed_counts = {}
        for pb in planner.placed_buildings:
            nom = pb['info']['Nom']
            placed_counts[nom] = placed_counts.get(nom, 0) + 1
        
        # Calculer les bâtiments non placés
        not_placed = []
        for b in all_buildings:
            nom = b['Nom']
            quantite = b['Quantite']
            placed = placed_counts.get(nom, 0)
            if placed < quantite:
                # Il en manque
                for _ in range(quantite - placed):
                    not_placed.append(b)
        
        cases_occupees = sum(p['w'] * p['h'] for p in planner.placed_buildings)
        
        summary_data = [
            ["Cases libres initiales", planner.initial_free_cells],
            ["Cases utilisées", cases_occupees],
            ["Cases non utilisées", planner.initial_free_cells - cases_occupees],
            ["Bâtiments placés", len(planner.placed_buildings)],
            ["Bâtiments non placés", len(not_placed)],
            ["Surface non placée (cases)", sum(b['Longueur'] * b['Largeur'] for b in not_placed)],
            ["Culture totale reçue (producteurs)", culture_results['total_culture']],
            ["Score d'optimisation", culture_results['score']],
            ["Statut", "STOP: LIMITE JOURNAL (10000 entrées)" if planner.interrupted else "OK"]
        ]
        
        summary_df = pd.DataFrame(summary_data, columns=["Métrique", "Valeur"])
        summary_df.to_excel(writer, sheet_name="Resume", index=False)
        
        # 8. Bâtiments non placés
        if not_placed:
            not_placed_df = pd.DataFrame(not_placed)
            not_placed_df.to_excel(writer, sheet_name="Non_Places", index=False)
    
    return output.getvalue()

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Optimiseur de Cité", page_icon="🏗️", layout="wide")

st.title("🏗️ Optimiseur de Cité avec Chargement de Configuration Existante")
st.markdown("""
Cet outil place automatiquement les bâtiments sur un terrain en partant d'une configuration existante.

**Fonctionnalités :**
- Lit le terrain de base (onglet 1)
- Lit la liste des bâtiments à placer (onglet 2)
- Lit la configuration actuelle (onglet 3) - peut être vide ou partiellement remplie
- **Gère les orientations différentes** des bâtiments
- Ne place que les bâtiments manquants
- Optimise la disposition finale

**Couleurs :**
- **Orange** : Bâtiments Culturels
- **Vert** : Bâtiments Producteurs
- **Gris** : Bâtiments Neutres
- **Noir** : Bords du terrain (X)
""")

uploaded = st.file_uploader("📂 Charger le fichier Excel (Ville.xlsx)", type="xlsx")

if uploaded:
    with st.spinner("Analyse du fichier en cours..."):
        try:
            # Lecture des trois onglets
            t_df = pd.read_excel(uploaded, sheet_name=0, header=None)
            b_df = pd.read_excel(uploaded, sheet_name=1)
            
            # Essayer de lire l'onglet 3
            buildings_detected = None
            try:
                a_df = pd.read_excel(uploaded, sheet_name=2, header=None)
                st.info("✅ Onglet 'Actuel' chargé avec succès")
            except:
                a_df = None
                st.info("ℹ️ Pas d'onglet 'Actuel' - départ terrain vide")
            
            b_df.columns = b_df.columns.str.strip()
            
            st.success("✅ Fichier chargé avec succès!")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.subheader("Terrain de base")
                st.dataframe(t_df.head(5))
            
            with col2:
                st.subheader("Bâtiments disponibles")
                total = sum(b_df['Quantite'])
                st.metric("Total bâtiments", total)
                st.dataframe(b_df[['Nom', 'Quantite']].head(10))
            
            with col3:
                st.subheader("Configuration actuelle")
                if a_df is not None:
                    st.dataframe(a_df.head(5))
                else:
                    st.write("(vide)")
            
            # Création du planner
            planner = CityPlanner(t_df.values, a_df.values if a_df is not None else None)
            
            # Construire la liste de tous les bâtiments à placer (avec quantités)
            all_buildings = []
            for _, row in b_df.iterrows():
                for _ in range(int(row['Quantite'])):
                    all_buildings.append(row.to_dict())
            
            # Associer les bâtiments existants avec leurs caractéristiques
            unmatched = planner.match_buildings_with_info(all_buildings)
            
            # Construire la liste des bâtiments restant à placer
            placed_counts = {}
            for pb in planner.placed_buildings:
                nom = pb['info']['Nom']
                placed_counts[nom] = placed_counts.get(nom, 0) + 1
            
            remaining_queue = []
            for _, row in b_df.iterrows():
                nom = row['Nom']
                quantite = row['Quantite']
                placed = placed_counts.get(nom, 0)
                for _ in range(max(0, quantite - placed)):
                    remaining_queue.append(row.to_dict())
            
            st.info(f"📊 Bâtiments déjà placés: {len(planner.placed_buildings)}")
            st.info(f"📊 Bâtiments restant à placer: {len(remaining_queue)}")
            
            # Vérification détaillée par type de bâtiment
            st.subheader("🔍 Vérification par type de bâtiment")
            check_data = []
            for _, row in b_df.iterrows():
                nom = row['Nom']
                quantite = row['Quantite']
                placed = placed_counts.get(nom, 0)
                check_data.append({
                    "Bâtiment": nom,
                    "Quantité": quantite,
                    "Placés": placed,
                    "Restants": quantite - placed
                })
            check_df = pd.DataFrame(check_data)
            st.dataframe(check_df)
            
            # Vérifier si le total correspond
            total_buildings = sum(b_df['Quantite'])
            if len(planner.placed_buildings) + len(remaining_queue) == total_buildings:
                st.success(f"✅ Total cohérent: {total_buildings} bâtiments")
            else:
                st.warning(f"⚠️ Incohérence: {len(planner.placed_buildings)} + {len(remaining_queue)} = {len(planner.placed_buildings) + len(remaining_queue)} mais devraient être {total_buildings}")
            
            if remaining_queue:
                # Trier les bâtiments restants
                neutres = [b for b in remaining_queue if b['Type'] == 'Neutre']
                culturels = [b for b in remaining_queue if b['Type'] == 'Culturel']
                producteurs = [b for b in remaining_queue if b['Type'] == 'Producteur']
                
                neutres.sort(key=lambda x: (x['Longueur'], x['Largeur']), reverse=True)
                culturels.sort(key=lambda x: (x['Longueur'], x['Largeur']), reverse=True)
                producteurs.sort(key=lambda x: (x['Longueur'], x['Largeur']), reverse=True)
                
                # Réorganiser la file d'attente
                full_queue = []
                full_queue.extend(neutres)
                
                max_len = max(len(culturels), len(producteurs))
                for i in range(max_len):
                    if i < len(culturels):
                        full_queue.append(culturels[i])
                    if i < len(producteurs):
                        full_queue.append(producteurs[i])
            else:
                full_queue = []
            
            if st.button("🚀 Lancer le placement et l'optimisation", type="primary"):
                with st.spinner("Placement en cours..."):
                    # Phase 1: Placement des bâtiments manquants
                    if full_queue:
                        st.info(f"Phase 1: Placement de {len(full_queue)} bâtiments...")
                        planner.solve(full_queue, phase="placement")
                    else:
                        st.info("Tous les bâtiments sont déjà placés")
                    
                    # Phase 2: Optimisation
                    st.info("Phase 2: Optimisation...")
                    planner.optimize_placement(iterations=10)
                    
                    # Résultats finaux
                    culture_results = planner.calculate_culture_and_score()
                    
                    st.success("✅ Terminé!")
                    
                    # Statistiques
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Bâtiments placés", len(planner.placed_buildings))
                    with col2:
                        st.metric("Cases utilisées", 
                                 sum(p['w'] * p['h'] for p in planner.placed_buildings))
                    with col3:
                        st.metric("Culture totale", f"{culture_results['total_culture']:.0f}")
                    with col4:
                        st.metric("Score", culture_results['score'])
                    
                    # Boosts
                    st.subheader("📊 Répartition des boosts")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Boost 100%", culture_results['boost_counts'][100])
                    with col2:
                        st.metric("Boost 50%", culture_results['boost_counts'][50])
                    with col3:
                        st.metric("Boost 25%", culture_results['boost_counts'][25])
                    with col4:
                        st.metric("Boost 0%", culture_results['boost_counts'][0])
                    
                    # Téléchargement
                    excel_data = generate_excel(planner, all_buildings)
                    st.download_button(
                        label="📥 Télécharger le résultat (Excel)",
                        data=excel_data,
                        file_name="Resultat_Cite.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement: {str(e)}")
            st.exception(e)

else:
    st.info("👆 Veuillez charger un fichier Excel pour commencer")