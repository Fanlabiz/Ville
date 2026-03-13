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
            'orientation': f"{w}x{h}"
        })
        self.log(f"✓ Placé : {b['Nom']} en ({r},{c}) orientation {w}x{h}")
        
        if self.solve(buildings[1:], phase): 
            return True
        
        if not self.interrupted:
            self.log(f"✗ Enlevé : {b['Nom']} de ({r},{c})")
            self.grid[r:r+h, c:c+w] = 1
            self.placed_buildings.pop()
        return False
    
    def calculate_culture_and_score(self):
        """
        Calcule la culture reçue et un score basé sur les boosts.
        Le score favorise les boosts 100% puis 50% puis 25%.
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
                prod_r_start = pb['r']
                prod_r_end = pb['r'] + pb['h']
                prod_c_start = pb['c']
                prod_c_end = pb['c'] + pb['w']
                
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
                
                culture_recue = sum(cultures_recues)
                
                boost = 0
                boost_100 = pb['info'].get('Boost 100%')
                boost_50 = pb['info'].get('Boost 50%')
                boost_25 = pb['info'].get('Boost 25%')
                
                if pd.notna(boost_100) and culture_recue >= boost_100:
                    boost = 100
                elif pd.notna(boost_50) and culture_recue >= boost_50:
                    boost = 50
                elif pd.notna(boost_25) and culture_recue >= boost_25:
                    boost = 25
                
                boost_counts[boost] += 1
                
                # Score pondéré : 100% = 4 points, 50% = 2 points, 25% = 1 point
                if boost == 100:
                    score += 4
                elif boost == 50:
                    score += 2
                elif boost == 25:
                    score += 1
                
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

    def optimize_placement(self, buildings, iterations=5):
        """
        Phase d'optimisation : réarrange les bâtiments pour maximiser les boosts.
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
        
        # Identifier les bâtiments culturels et producteurs
        cultural_buildings = [b for b in self.placed_buildings if b['info']['Type'] == 'Culturel']
        producer_buildings = [b for b in self.placed_buildings if b['info']['Type'] == 'Producteur']
        
        # Plusieurs itérations d'optimisation
        for iteration in range(iterations):
            self.log(f"\n--- Itération d'optimisation {iteration + 1} ---")
            improved = False
            
            # Essayer de déplacer chaque producteur pour améliorer son boost
            for i, prod in enumerate(producer_buildings):
                if self.interrupted:
                    break
                
                # Sauvegarder l'état actuel
                current_grid = self.grid.copy()
                current_prod = copy.deepcopy(prod)
                
                # Enlever le producteur
                self.grid[prod['r']:prod['r']+prod['h'], prod['c']:prod['c']+prod['w']] = 1
                self.placed_buildings.remove(prod)
                
                # Chercher une meilleure position pour ce producteur
                best_pos = None
                best_culture = -1
                
                # Tester toutes les positions possibles
                for r in range(self.rows - prod['h'] + 1):
                    for c in range(self.cols - prod['w'] + 1):
                        if np.all(self.grid[r:r+prod['h'], c:c+prod['w']] == 1):
                            # Tester cette position
                            self.grid[r:r+prod['h'], c:c+prod['w']] = 0
                            
                            # Calculer la culture à cette position
                            test_prod = copy.deepcopy(prod)
                            test_prod['r'] = r
                            test_prod['c'] = c
                            
                            # Créer une liste temporaire avec ce producteur à la nouvelle position
                            temp_buildings = self.placed_buildings + [test_prod] + cultural_buildings
                            
                            # Calculer la culture pour ce producteur à cette position
                            culture_at_pos = 0
                            for cb in cultural_buildings:
                                rayon = int(cb['info'].get('Rayonnement', 0))
                                culture_value = cb['info'].get('Culture', 0)
                                
                                if rayon > 0 and culture_value > 0:
                                    cult_r_start = cb['r']
                                    cult_r_end = cb['r'] + cb['h']
                                    cult_c_start = cb['c']
                                    cult_c_end = cb['c'] + cb['w']
                                    
                                    for pr in range(r, r + prod['h']):
                                        for pc in range(c, c + prod['w']):
                                            for cr in range(cult_r_start, cult_r_end):
                                                for cc in range(cult_c_start, cult_c_end):
                                                    distance = abs(pr - cr) + abs(pc - cc)
                                                    if distance <= rayon:
                                                        culture_at_pos += culture_value
                                                        break
                                                if culture_at_pos > best_culture:
                                                    break
                                            if culture_at_pos > best_culture:
                                                break
                                        if culture_at_pos > best_culture:
                                            break
                            
                            self.grid[r:r+prod['h'], c:c+prod['w']] = 1
                            
                            if culture_at_pos > best_culture:
                                best_culture = culture_at_pos
                                best_pos = (r, c)
                
                # Si on a trouvé une meilleure position
                if best_pos and best_culture > 0:
                    r, c = best_pos
                    
                    # Placer le producteur à la nouvelle position
                    self.grid[r:r+prod['h'], c:c+prod['w']] = 0
                    prod['r'] = r
                    prod['c'] = c
                    self.placed_buildings.append(prod)
                    
                    self.log(f"  Producteur {prod['info']['Nom']} déplacé de ({current_prod['r']},{current_prod['c']}) à ({r},{c}) - Culture: {best_culture}")
                    improved = True
                else:
                    # Remettre le producteur à sa place
                    self.grid = current_grid
                    self.placed_buildings.append(current_prod)
            
            # Calculer le nouveau score après cette itération
            if improved:
                results = self.calculate_culture_and_score()
                if results['score'] > self.best_score:
                    self.best_score = results['score']
                    self.best_solution = {
                        'grid': self.grid.copy(),
                        'placed_buildings': copy.deepcopy(self.placed_buildings)
                    }
                    self.log(f"Nouveau meilleur score: {self.best_score} (Boosts: 100%:{results['boost_counts'][100]}, 50%:{results['boost_counts'][50]}, 25%:{results['boost_counts'][25]})")
        
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
        
        # 7. Résumé
        placed_ids = [id(p['info']) for p in planner.placed_buildings]
        not_placed = [b for b in full_queue if id(b) not in placed_ids]
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

st.title("🏗️ Optimiseur de Placement de Bâtiments avec Optimisation")
st.markdown("""
Cet outil place automatiquement les bâtiments sur un terrain puis optimise leur disposition pour maximiser les boosts de production.

**Processus en 2 phases :**
1. **Placement initial** : Placement selon les règles (neutres aux bords, alternance culturels/producteurs)
2. **Optimisation** : Réarrangement des producteurs pour maximiser les boosts 100%

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
                optimize = st.checkbox("Activer l'optimisation", value=True)
            with col2:
                iterations = st.slider("Nombre d'itérations d'optimisation", min_value=1, max_value=20, value=5)
            
            if st.button("🚀 Lancer l'optimisation", type="primary"):
                with st.spinner("Placement et optimisation en cours... (cela peut prendre quelques minutes)"):
                    planner = CityPlanner(t_df.values)
                    
                    # Phase 1: Placement initial
                    st.info("Phase 1: Placement initial...")
                    planner.solve(full_queue, phase="initiale")
                    
                    # Phase 2: Optimisation
                    if optimize:
                        st.info(f"Phase 2: Optimisation ({iterations} itérations)...")
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
                        file_name="Resultat_Placement_Optimise.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement: {str(e)}")
            st.exception(e)

else:
    st.info("👆 Veuillez charger un fichier Excel pour commencer")