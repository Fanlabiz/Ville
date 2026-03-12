import pandas as pd
import numpy as np
import streamlit as st
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill

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
        # 1. Vérification limites et collision
        if r + h > self.rows or c + w > self.cols:
            return False
        if not np.all(self.grid[r:r+h, c:c+w] == 1):
            return False
        
        # 2. Règle du plus grand restant
        if remaining_queue:
            # On simule le placement pour voir l'espace restant
            self.grid[r:r+h, c:c+w] = 0
            biggest = remaining_queue[0]
            bw, bh = biggest['Largeur'], biggest['Longueur']
            
            can_fit_biggest = False
            # On cherche au moins UNE place pour le plus grand (en testant les 2 sens)
            for orient in [(bw, bh), (bh, bw)]:
                ow, oh = orient
                for rr in range(self.rows - oh + 1):
                    for cc in range(self.cols - ow + 1):
                        if np.all(self.grid[rr:rr+oh, cc:cc+ow] == 1):
                            can_fit_biggest = True
                            break
                    if can_fit_biggest: break
                if can_fit_biggest: break
            
            # On remet la grille en état
            self.grid[r:r+h, c:c+w] = 1
            if not can_fit_biggest:
                return False

        return True

    def solve(self, buildings):
        """Algorithme de placement récursif avec backtracking"""
        if not buildings or self.interrupted:
            return True

        b = buildings[0]
        self.log(f"Évaluation de : {b['Nom']} (Type: {b['Type']})")
        
        # Toutes les orientations possibles
        dims = [(b['Largeur'], b['Longueur']), (b['Longueur'], b['Largeur'])]
        if b['Largeur'] == b['Longueur']: 
            dims = [dims[0]]

        # PHASE 1 : Placement prioritaire (bords pour les Neutres)
        for w, h in dims:
            for r in range(self.rows - h + 1):
                for c in range(self.cols - w + 1):
                    if b['Type'] == 'Neutre' and not self.is_adjacent_to_X(r, c, w, h):
                        continue
                    if self.can_place(r, c, w, h, buildings[1:]):
                        if self.try_placement(b, r, c, w, h, buildings):
                            return True

        # PHASE 2 : Placement de secours (pour Neutres si bords saturés)
        if b['Type'] == 'Neutre':
            self.log(f"Bords saturés, recherche interne pour : {b['Nom']}")
            for w, h in dims:
                for r in range(self.rows - h + 1):
                    for c in range(self.cols - w + 1):
                        if self.can_place(r, c, w, h, buildings[1:]):
                            if self.try_placement(b, r, c, w, h, buildings):
                                return True
        return False

    def try_placement(self, b, r, c, w, h, buildings):
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
        
        if self.solve(buildings[1:]): 
            return True
        
        if not self.interrupted:
            self.log(f"✗ Enlevé : {b['Nom']} de ({r},{c})")
            self.grid[r:r+h, c:c+w] = 1
            self.placed_buildings.pop()
        return False
    
    def calculate_culture(self):
        """
        Calcule la culture reçue par chaque bâtiment producteur.
        Retourne un dictionnaire avec les résultats.
        """
        # Étape 1: Créer une carte de la culture émise par les bâtiments culturels
        culture_map = np.zeros((self.rows, self.cols))
        
        # Pour chaque bâtiment culturel placé
        for pb in self.placed_buildings:
            if pb['info']['Type'] == 'Culturel':
                rayon = int(pb['info'].get('Rayonnement', 0))
                culture_value = pb['info'].get('Culture', 0)
                
                if rayon > 0 and culture_value > 0:
                    # La zone de rayonnement est une bande AUTOUR du bâtiment
                    # Le bâtiment lui-même n'est pas inclus dans sa zone de rayonnement
                    
                    # Coordonnées du bâtiment
                    r_start_bat = pb['r']
                    r_end_bat = pb['r'] + pb['h']
                    c_start_bat = pb['c']
                    c_end_bat = pb['c'] + pb['w']
                    
                    # Zone de rayonnement étendue (incluant le bâtiment)
                    r_start = max(0, pb['r'] - rayon)
                    r_end = min(self.rows, pb['r'] + pb['h'] + rayon)
                    c_start = max(0, pb['c'] - rayon)
                    c_end = min(self.cols, pb['c'] + pb['w'] + rayon)
                    
                    # Parcourir toutes les cases dans la zone étendue
                    for r in range(r_start, r_end):
                        for c in range(c_start, c_end):
                            # Vérifier si la case est DANS le bâtiment
                            if r_start_bat <= r < r_end_bat and c_start_bat <= c < c_end_bat:
                                continue  # On ignore les cases du bâtiment lui-même
                            
                            # Vérifier si la case est à la bonne distance (rayonnement)
                            # Calculer la distance de Manhattan minimum entre la case et le bâtiment
                            dist_min = float('inf')
                            
                            # Pour chaque case du bâtiment, calculer la distance
                            for r_bat in range(r_start_bat, r_end_bat):
                                for c_bat in range(c_start_bat, c_end_bat):
                                    dist = abs(r - r_bat) + abs(c - c_bat)
                                    dist_min = min(dist_min, dist)
                            
                            # Si la distance minimale est <= rayon, la case est dans la zone
                            if dist_min <= rayon:
                                culture_map[r, c] += culture_value
                    
                    self.log(f"  Culture émise: {pb['info']['Nom']} ajoute {culture_value} dans zone {rayon}")
        
        # Étape 2: Calculer la culture reçue par chaque bâtiment producteur
        prod_stats = []
        prod_by_type = {"Guerison": 0, "Nourriture": 0, "Or": 0, "Bijoux": 0, "Onguents": 0, "Cristal": 0, "Epices": 0, "Boiseries": 0, "Scriberie": 0}
        
        total_culture_all_prod = 0
        
        for pb in self.placed_buildings:
            if pb['info']['Type'] == 'Producteur':
                # Récupérer la somme de culture dans l'empreinte du bâtiment
                footprint = culture_map[pb['r']:pb['r']+pb['h'], pb['c']:pb['c']+pb['w']]
                
                # La culture reçue est la somme (car chaque case peut recevoir de multiples sources)
                # Mais selon la règle, si UNE case est dans la zone, le bâtiment reçoit la culture
                # Donc on prend le MAX plutôt que la somme
                culture_recue = np.max(footprint) if footprint.size > 0 else 0
                
                # Déterminer le boost
                boost = 0
                if culture_recue >= pb['info'].get('Boost 100%', float('inf')):
                    boost = 100
                elif culture_recue >= pb['info'].get('Boost 50%', float('inf')):
                    boost = 50
                elif culture_recue >= pb['info'].get('Boost 25%', float('inf')):
                    boost = 25
                
                prod_stats.append({
                    'Nom': pb['info']['Nom'],
                    'Culture reçue': culture_recue,
                    'Boost': f"{boost}%",
                    'Production': pb['info'].get('Production', '')
                })
                
                total_culture_all_prod += culture_recue
                
                # Ajouter aux totaux par type de production
                prod_type = str(pb['info'].get('Production', ''))
                if prod_type in prod_by_type:
                    prod_by_type[prod_type] += culture_recue
        
        # Log pour déboguer - afficher la culture reçue par la caserne de siège
        for stat in prod_stats:
            if stat['Nom'] == 'Caserne de siège':
                self.log(f"  DEBUG - Caserne de siège reçoit {stat['Culture reçue']} de culture")
        
        return {
            'prod_stats': prod_stats,
            'prod_by_type': prod_by_type,
            'total_culture': total_culture_all_prod,
            'culture_map': culture_map
        }

# --- LOGIQUE D'EXPORT EXCEL ---
def generate_excel(planner, full_queue):
    """Génère le fichier Excel de résultat"""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. Journal
        pd.DataFrame(planner.journal, columns=["Journal"]).to_excel(writer, sheet_name="Journal", index=False)
        
        # 2. Calcul de la culture
        culture_results = planner.calculate_culture()
        
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
        
        # 5. Création du plan du terrain avec les noms complets
        # Créer une feuille pour le plan
        ws = writer.book.create_sheet("Plan_Terrain")
        
        # Remplir les cellules avec les noms complets
        for r in range(planner.rows):
            for c in range(planner.cols):
                cell = ws.cell(row=r+1, column=c+1)
                
                # Si c'est un bord (X)
                if planner.border_mask[r, c]:
                    cell.value = "X"
                    cell.fill = PatternFill(start_color='FF000000', fill_type='solid')  # Noir
                else:
                    # Vérifier si cette case est occupée par un bâtiment
                    cell_value = ""
                    cell_color = None
                    
                    for pb in planner.placed_buildings:
                        if pb['r'] <= r < pb['r'] + pb['h'] and pb['c'] <= c < pb['c'] + pb['w']:
                            cell_value = pb['info']['Nom']  # Nom complet
                            cell_color = COLORS.get(pb['info']['Type'], 'FFFFFFFF')
                            break
                    
                    if cell_value:
                        cell.value = cell_value
                        if cell_color:
                            cell.fill = PatternFill(start_color=cell_color, fill_type='solid')
        
        # Ajuster la largeur des colonnes pour voir les noms
        for col in range(1, planner.cols + 1):
            ws.column_dimensions[chr(64 + col)].width = 20
        
        # 6. Résumé
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
            ["Statut", "STOP: LIMITE JOURNAL (10000 entrées)" if planner.interrupted else "OK"]
        ]
        
        summary_df = pd.DataFrame(summary_data, columns=["Métrique", "Valeur"])
        summary_df.to_excel(writer, sheet_name="Resume", index=False)
        
        # 7. Bâtiments non placés
        if not_placed:
            not_placed_df = pd.DataFrame(not_placed)
            not_placed_df.to_excel(writer, sheet_name="Non_Places", index=False)
    
    return output.getvalue()

# --- INTERFACE STREAMLIT ---
st.set_page_config(page_title="Optimiseur de Cité", page_icon="🏗️", layout="wide")

st.title("🏗️ Optimiseur de Placement de Bâtiments")
st.markdown("""
Cet outil place automatiquement les bâtiments sur un terrain en maximisant les boosts de production.
- **Orange** : Bâtiments Culturels
- **Vert** : Bâtiments Producteurs
- **Gris** : Bâtiments Neutres
- **Noir** : Bords du terrain (X)
""")

uploaded = st.file_uploader("📂 Charger le fichier Excel (Ville.xlsx)", type="xlsx")

if uploaded:
    with st.spinner("Analyse du fichier en cours..."):
        try:
            # Lecture du terrain (premier onglet, sans en-tête)
            t_df = pd.read_excel(uploaded, sheet_name=0, header=None)
            
            # Lecture des bâtiments (deuxième onglet)
            b_df = pd.read_excel(uploaded, sheet_name=1)
            b_df.columns = b_df.columns.str.strip()
            
            st.success("✅ Fichier chargé avec succès!")
            
            # Aperçu des données
            col1, col2 = st.columns(2)
            with col1:
                st.subheader("Aperçu du terrain")
                st.dataframe(t_df.head(10))
            
            with col2:
                st.subheader("Aperçu des bâtiments")
                st.dataframe(b_df.head(10))
            
            # Création de la file d'attente triée
            st.subheader("🔄 Ordre de placement")
            
            # Séparer par type
            neutres = b_df[b_df['Type'] == 'Neutre'].copy()
            culturels = b_df[b_df['Type'] == 'Culturel'].copy()
            producteurs = b_df[b_df['Type'] == 'Producteur'].copy()
            
            # Trier par taille (du plus grand au plus petit)
            neutres = neutres.sort_values(['Longueur', 'Largeur'], ascending=False)
            culturels = culturels.sort_values(['Longueur', 'Largeur'], ascending=False)
            producteurs = producteurs.sort_values(['Longueur', 'Largeur'], ascending=False)
            
            # Construire la file d'attente
            full_queue = []
            
            # 1. D'abord tous les Neutres (en priorité)
            for _, row in neutres.iterrows():
                for _ in range(int(row['Quantite'])):
                    full_queue.append(row.to_dict())
            
            # 2. Alternance Culturels/Producteurs
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
            
            # Lancer le placement
            if st.button("🚀 Lancer l'optimisation", type="primary"):
                with st.spinner("Placement des bâtiments en cours... (cela peut prendre quelques secondes)"):
                    planner = CityPlanner(t_df.values)
                    planner.solve(full_queue)
                    
                    # Calculer les résultats
                    culture_results = planner.calculate_culture()
                    
                    # Afficher les résultats
                    st.success("✅ Placement terminé!")
                    
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
                        st.metric("Journal", f"{len(planner.journal)} entrées")
                    
                    # Afficher les boosts par type
                    st.subheader("📊 Culture par type de production")
                    prod_df = pd.DataFrame([
                        {"Type": k, "Culture": v} 
                        for k, v in culture_results['prod_by_type'].items() if v > 0
                    ])
                    st.dataframe(prod_df)
                    
                    # Bouton de téléchargement
                    excel_data = generate_excel(planner, full_queue)
                    st.download_button(
                        label="📥 Télécharger le résultat (Excel)",
                        data=excel_data,
                        file_name="Resultat_Placement.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
        except Exception as e:
            st.error(f"❌ Erreur lors du traitement: {str(e)}")
            st.exception(e)

else:
    st.info("👆 Veuillez charger un fichier Excel pour commencer")