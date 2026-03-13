import pandas as pd
import numpy as np
import streamlit as st
import io
import copy
import random
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
    def __init__(self, terrain_data):
        self.rows = len(terrain_data)
        self.cols = len(terrain_data[0])
        self.grid = np.zeros((self.rows, self.cols))
        self.border_mask = np.zeros((self.rows, self.cols), dtype=bool)
        self.initial_free_cells = 0
        
        for r in range(self.rows):
            for c in range(self.cols):
                val = str(terrain_data[r][c]).strip().upper()
                if val == '1': 
                    self.grid[r,c] = 1
                    self.initial_free_cells += 1
                elif val == 'X':
                    self.grid[r,c] = 0 
                    self.border_mask[r,c] = True
        
        self.journal = []
        self.placed_buildings = []
        self.max_entries = 10000 
        self.interrupted = False
        self.best_solution = None
        self.best_score = -1
        self.best_boost_counts = None

    def log(self, msg):
        if len(self.journal) < self.max_entries:
            self.journal.append(msg)
        else:
            self.interrupted = True

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
        """Algorithme de placement récursif avec backtracking"""
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
            'h': h,
            'orientation': f"{w}x{h}",
            'id': id(b)  # Ajouter un identifiant unique
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
    
    def calculate_potential_gain(self, building, current_culture, cultural_buildings):
        """
        Calcule le gain potentiel si on déplaçait ce bâtiment.
        Retourne (gain_potentiel, points_gagnables, distance_au_prochain_palier)
        """
        info = building['info']
        current_boost = self.calculate_boost_from_culture(info, current_culture)
        
        # Trouver le prochain palier
        boost_25 = info.get('Boost 25%')
        boost_50 = info.get('Boost 50%')
        boost_100 = info.get('Boost 100%')
        
        best_gain = 0
        best_points = 0
        best_distance = float('inf')
        
        if current_boost == 0 and pd.notna(boost_25):
            # Peut gagner 1 point en atteignant 25%
            seuil = boost_25
            if seuil > current_culture:
                distance = seuil - current_culture
                points = 1
                # Plus la distance est petite, plus le gain potentiel est grand
                gain = points * (1000 / max(1, distance))
                if gain > best_gain:
                    best_gain = gain
                    best_points = points
                    best_distance = distance
        
        if current_boost <= 25 and pd.notna(boost_50):
            # Peut gagner 1 ou 2 points en atteignant 50%
            seuil = boost_50
            if seuil > current_culture:
                distance = seuil - current_culture
                points = 2 if current_boost == 0 else 1
                gain = points * (1000 / max(1, distance))
                if gain > best_gain:
                    best_gain = gain
                    best_points = points
                    best_distance = distance
        
        if current_boost <= 50 and pd.notna(boost_100):
            # Peut gagner 1, 2 ou 4 points en atteignant 100%
            seuil = boost_100
            if seuil > current_culture:
                distance = seuil - current_culture
                if current_boost == 0:
                    points = 4
                elif current_boost == 25:
                    points = 3
                else:  # current_boost == 50
                    points = 2
                gain = points * (1000 / max(1, distance))
                if gain > best_gain:
                    best_gain = gain
                    best_points = points
                    best_distance = distance
        
        return best_gain, best_points, best_distance

    def calculate_culture_and_score(self):
        """
        Calcule la culture reçue et un score basé sur les boosts.
        """
        cultural_buildings = [pb for pb in self.placed_buildings if pb['info']['Type'] == 'Culturel']
        
        prod_stats = []
        prod_by_type = {"Guerison": 0, "Nourriture": 0, "Or": 0, "Bijoux": 0, "Onguents": 0, 
                        "Cristal": 0, "Epices": 0, "Boiseries": 0, "Scriberie": 0}
        
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

    def find_best_position_for_building(self, building, cultural_buildings, producer_buildings):
        """
        Trouve la meilleure position pour un bâtiment donné en considérant
        les échanges possibles avec d'autres bâtiments.
        """
        best_positions = []
        w, h = building['w'], building['h']
        
        # Explorer toutes les positions possibles
        for r in range(self.rows - h + 1):
            for c in range(self.cols - w + 1):
                # Ignorer la position actuelle
                if r == building['r'] and c == building['c']:
                    continue
                
                # Vérifier si la zone est libre
                zone = self.grid[r:r+h, c:c+w]
                
                if np.all(zone == 1):
                    # Zone complètement libre
                    culture_ici = self.calculate_culture_for_position(building, r, c, cultural_buildings)
                    best_positions.append({
                        'r': r, 'c': c,
                        'culture': culture_ici,
                        'occupants': [],
                        'type': 'libre'
                    })
                else:
                    # Zone occupée - vérifier si on peut échanger avec les occupants
                    occupants = []
                    zones_occupees = []
                    
                    # Identifier tous les bâtiments qui occupent cette zone
                    for other in producer_buildings:
                        if other['id'] == building['id']:
                            continue
                        
                        # Vérifier si ce bâtiment chevauche la zone cible
                        if (other['r'] < r + h and other['r'] + other['h'] > r and
                            other['c'] < c + w and other['c'] + other['w'] > c):
                            occupants.append(other)
                            zones_occupees.append({
                                'r': other['r'], 'c': other['c'],
                                'w': other['w'], 'h': other['h']
                            })
                    
                    # Vérifier si la zone cible peut accueillir le bâtiment après avoir enlevé les occupants
                    # et si les occupants peuvent être replacés ailleurs
                    if occupants:
                        # Sauvegarder l'état
                        grid_backup = self.grid.copy()
                        
                        # Enlever temporairement les occupants
                        for occ in occupants:
                            self.grid[occ['r']:occ['r']+occ['h'], occ['c']:occ['c']+occ['w']] = 1
                        
                        # Vérifier si la zone est maintenant libre
                        if np.all(self.grid[r:r+h, c:c+w] == 1):
                            culture_ici = self.calculate_culture_for_position(building, r, c, cultural_buildings)
                            best_positions.append({
                                'r': r, 'c': c,
                                'culture': culture_ici,
                                'occupants': occupants,
                                'type': 'echange'
                            })
                        
                        # Restaurer l'état
                        self.grid = grid_backup
        
        # Trier par culture décroissante
        best_positions.sort(key=lambda x: x['culture'], reverse=True)
        return best_positions[:5]  # Garder les 5 meilleures

    def try_swap(self, building, target_pos, cultural_buildings, producer_buildings):
        """
        Tente d'échanger un bâtiment vers une nouvelle position, en déplaçant
        les occupants si nécessaire.
        """
        r, c = target_pos['r'], target_pos['c']
        occupants = target_pos['occupants']
        
        if not occupants:
            # Déplacement simple vers case libre
            old_r, old_c = building['r'], building['c']
            
            # Libérer l'ancienne position
            self.grid[old_r:old_r+building['h'], old_c:old_c+building['w']] = 1
            # Occuper la nouvelle position
            self.grid[r:r+building['h'], c:c+building['w']] = 0
            
            building['r'], building['c'] = r, c
            
            self.log(f"  Déplacement: {building['info']['Nom']} de ({old_r},{old_c}) à ({r},{c})")
            return True
        
        else:
            # Échange avec des occupants
            # Sauvegarder les positions
            old_r, old_c = building['r'], building['c']
            old_h, old_w = building['h'], building['w']
            
            # Libérer l'ancienne position du bâtiment principal
            self.grid[old_r:old_r+old_h, old_c:old_c+old_w] = 1
            
            # Libérer les positions des occupants
            for occ in occupants:
                self.grid[occ['r']:occ['r']+occ['h'], occ['c']:occ['c']+occ['w']] = 1
            
            # Vérifier si on peut replacer les occupants
            occupants_replaces = True
            new_positions = []
            
            for occ in occupants:
                placed = False
                # Chercher une nouvelle place pour cet occupant
                for orow in range(self.rows - occ['h'] + 1):
                    for ocol in range(self.cols - occ['w'] + 1):
                        if np.all(self.grid[orow:orow+occ['h'], ocol:ocol+occ['w']] == 1):
                            self.grid[orow:orow+occ['h'], ocol:ocol+occ['w']] = 0
                            new_positions.append((occ, orow, ocol))
                            placed = True
                            break
                    if placed:
                        break
                if not placed:
                    occupants_replaces = False
                    break
            
            if occupants_replaces:
                # Placer le bâtiment principal à sa nouvelle position
                self.grid[r:r+building['h'], c:c+building['w']] = 0
                building['r'], building['c'] = r, c
                
                # Mettre à jour les positions des occupants
                for occ, orow, ocol in new_positions:
                    occ['r'], occ['c'] = orow, ocol
                
                occupants_str = ", ".join([occ['info']['Nom'] for occ in occupants])
                self.log(f"  Échange: {building['info']['Nom']} ({old_r},{old_c}) <-> [{occupants_str}] vers ({r},{c})")
                return True
            else:
                # Restaurer l'état initial
                self.grid = self.best_solution['grid'].copy()
                return False

    def optimize_placement(self, buildings, iterations=30):
        """
        Phase d'optimisation avancée :
        - Permet les échanges entre bâtiments de tailles différentes
        - Priorise les producteurs proches d'un palier
        """
        self.log("\n=== DÉBUT DE LA PHASE D'OPTIMISATION AVANCÉE ===\n")
        
        # Sauvegarder la solution initiale
        current_solution = {
            'grid': self.grid.copy(),
            'placed_buildings': copy.deepcopy(self.placed_buildings)
        }
        
        # Calculer le score initial
        initial_results = self.calculate_culture_and_score()
        self.best_score = initial_results['score']
        self.best_solution = current_solution
        self.best_boost_counts = initial_results['boost_counts']
        
        self.log(f"Score initial: {self.best_score} (Boosts: 100%:{initial_results['boost_counts'][100]}, 50%:{initial_results['boost_counts'][50]}, 25%:{initial_results['boost_counts'][25]})")
        
        # Identifier les bâtiments
        cultural_buildings = [b for b in self.placed_buildings if b['info']['Type'] == 'Culturel']
        
        # Plusieurs itérations d'optimisation
        for iteration in range(iterations):
            if self.interrupted:
                break
            
            self.log(f"\n--- Itération d'optimisation {iteration + 1} ---")
            
            # Mettre à jour la liste des producteurs
            producer_buildings = [b for b in self.placed_buildings if b['info']['Type'] == 'Producteur']
            
            # Calculer le potentiel pour chaque producteur
            producer_data = []
            for prod in producer_buildings:
                current_culture = self.calculate_culture_for_position(prod, prod['r'], prod['c'], cultural_buildings)
                gain, points, distance = self.calculate_potential_gain(prod, current_culture, cultural_buildings)
                
                producer_data.append({
                    'prod': prod,
                    'current_culture': current_culture,
                    'gain': gain,
                    'points': points,
                    'distance': distance,
                    'boost': self.calculate_boost_from_culture(prod['info'], current_culture)
                })
            
            # Trier par gain potentiel
            producer_data.sort(key=lambda x: x['gain'], reverse=True)
            
            # Prendre les top producteurs (ceux avec le plus grand gain potentiel)
            top_producers = [p for p in producer_data if p['gain'] > 0][:10]
            
            if not top_producers:
                self.log("  Aucun producteur avec potentiel d'amélioration")
                continue
            
            improved = False
            changes_made = 0
            
            # Pour chaque producteur à fort potentiel
            for target_data in top_producers:
                if self.interrupted or changes_made >= 3:
                    break
                
                target_prod = target_data['prod']
                
                # Trouver les meilleures positions d'échange
                best_positions = self.find_best_position_for_building(
                    target_prod, cultural_buildings, producer_buildings
                )
                
                for pos in best_positions:
                    if pos['culture'] > target_data['current_culture']:
                        # Amélioration directe
                        if self.try_swap(target_prod, pos, cultural_buildings, producer_buildings):
                            changes_made += 1
                            improved = True
                            
                            # Recalculer le score
                            new_results = self.calculate_culture_and_score()
                            
                            if new_results['score'] > self.best_score:
                                self.best_score = new_results['score']
                                self.best_solution = {
                                    'grid': self.grid.copy(),
                                    'placed_buildings': copy.deepcopy(self.placed_buildings)
                                }
                                self.best_boost_counts = new_results['boost_counts']
                                self.log(f"  ✓ Nouveau meilleur score: {self.best_score}")
                                self.log(f"    Boosts: 100%:{new_results['boost_counts'][100]}, 50%:{new_results['boost_counts'][50]}, 25%:{new_results['boost_counts'][25]}")
                            
                            break
            
            if not improved:
                self.log("  Aucune amélioration trouvée dans cette itération")
        
        # Restaurer la meilleure solution trouvée
        if self.best_solution:
            self.grid = self.best_solution['grid'].copy()
            self.placed_buildings = copy.deepcopy(self.best_solution['placed_buildings'])
            final_results = self.calculate_culture_and_score()
            self.log(f"\n=== OPTIMISATION TERMINÉE ===")
            self.log(f"Score final: {final_results['score']}")
            self.log(f"Boosts finaux: 100%:{final_results['boost_counts'][100]}, 50%:{final_results['boost_counts'][50]}, 25%:{final_results['boost_counts'][25]}")
        
        return self.best_score

# --- LOGIQUE D'EXPORT EXCEL ---
def generate_excel(planner, full_queue):
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
            ws.column_dimensions[column_letter].width = 25
        
        # 7. Résumé - CORRECTION: Utiliser les noms des bâtiments comme clés
        placed_names = [pb['info']['Nom'] + str(pb['r']) + str(pb['c']) for pb in planner.placed_buildings]
        not_placed = []
        
        for b in full_queue:
            # Chercher si ce bâtiment est placé (par son nom et sa position)
            found = False
            for pb in planner.placed_buildings:
                if pb['info']['Nom'] == b['Nom'] and pb['r'] == b.get('r', -1) and pb['c'] == b.get('c', -1):
                    found = True
                    break
            if not found:
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

st.title("🏗️ Optimiseur de Placement de Bâtiments avec Optimisation Avancée V3")
st.markdown("""
Cet outil place automatiquement les bâtiments sur un terrain puis optimise leur disposition pour maximiser les boosts de production.

**Nouveautés de cette version :**
- ✅ Correction du bug d'affichage des bâtiments non placés
- ✅ Échanges entre bâtiments de tailles différentes
- ✅ Priorisation des producteurs proches d'un palier (Ferme luxueuse, Caserne, etc.)
- ✅ Algorithmes de réarrangement plus puissants

- **Orange** : Bâtiments Culturels
- **Vert** : Bâtiments Producteurs
- **Gris** : Bâtiments Neutres
- **Noir** : Bords du terrain (X)
""")

uploaded = st.file_uploader("📂 Charger le fichier Excel (Ville.xlsx)", type="xlsx")

if uploaded:
    with st.spinner("Analyse du fichier en cours..."):
        try:
            t_df = pd.read_excel(uploaded, sheet_name=0, header=None)
            b_df = pd.read_excel(uploaded, sheet_name=1)
            b_df.columns = b_df.columns.str.strip()
            
            st.success("✅ Fichier chargé avec succès!")
            
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Aperçu du terrain")
                st.dataframe(t_df.head(10))
            
            with col2:
                st.subheader("Aperçu des bâtiments")
                st.dataframe(b_df.head(10))
            
            st.subheader("🔄 Ordre de placement")
            
            neutres = b_df[b_df['Type'] == 'Neutre'].copy()
            culturels = b_df[b_df['Type'] == 'Culturel'].copy()
            producteurs = b_df[b_df['Type'] == 'Producteur'].copy()
            
            neutres = neutres.sort_values(['Longueur', 'Largeur'], ascending=False)
            culturels = culturels.sort_values(['Longueur', 'Largeur'], ascending=False)
            producteurs = producteurs.sort_values(['Longueur', 'Largeur'], ascending=False)
            
            full_queue = []
            
            for _, row in neutres.iterrows():
                for _ in range(int(row['Quantite'])):
                    full_queue.append(row.to_dict())
            
            c_list = culturels.to_dict('records')
            p_list = producteurs.to_dict('records')
            
            max_len = max(len(c_list), len(p_list))
            for i in range(max_len):
                if i < len(c_list):
                    for _ in range(int(c_list[i]['Quantite'])):
                        full_queue.append(c_list[i].copy())
                if i < len(p_list):
                    for _ in range(int(p_list[i]['Quantite'])):
                        full_queue.append(p_list[i].copy())
            
            st.info(f"📊 Ordre de placement: {len(full_queue)} bâtiments à placer")
            
            # Paramètres d'optimisation
            st.subheader("⚙️ Paramètres d'optimisation")
            col1, col2 = st.columns(2)
            with col1:
                optimize = st.checkbox("Activer l'optimisation avancée", value=True)
            with col2:
                iterations = st.slider("Itérations d'optimisation", min_value=10, max_value=100, value=30)
            
            if st.button("🚀 Lancer l'optimisation avancée V3", type="primary"):
                with st.spinner("Placement et optimisation en cours... (cela peut prendre plusieurs minutes)"):
                    planner = CityPlanner(t_df.values)
                    
                    # Phase 1: Placement initial
                    st.info("Phase 1: Placement initial...")
                    planner.solve(full_queue, phase="initiale")
                    
                    # Phase 2: Optimisation avancée
                    if optimize:
                        st.info(f"Phase 2: Optimisation avancée ({iterations} itérations)...")
                        planner.optimize_placement(full_queue, iterations=iterations)
                    
                    # Calculer les résultats finaux
                    culture_results = planner.calculate_culture_and_score()
                    
                    st.success("✅ Optimisation terminée!")
                    
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
                        st.metric("Score optimisation", culture_results['score'])
                    
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
                    
                    st.subheader("📊 Culture par type de production")
                    prod_df = pd.DataFrame([
                        {"Type": k, "Culture": v} 
                        for k, v in culture_results['prod_by_type'].items() if v > 0
                    ])
                    st.dataframe(prod_df)
                    
                    excel_data = generate_excel(planner, full_queue)
                    st.download_button(
                        label="📥 Télécharger le résultat (Excel)",
                        data=excel_data,
                        file_name="Resultat_Placement_Optimise_V3.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement: {str(e)}")
            st.exception(e)

else:
    st.info("👆 Veuillez charger un fichier Excel pour commencer")