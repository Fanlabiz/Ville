import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import List, Tuple, Optional, Dict
import streamlit as st
from io import BytesIO
import copy

@dataclass
class Batiment:
    nom: str
    longueur: int
    largeur: int
    quantite: int
    type: str  # 'Culturel', 'Producteur', 'Neutre'
    culture: int = 0
    rayonnement: int = 0
    boost_25: int = 0
    boost_50: int = 0
    boost_100: int = 0
    production: str = ""
    
    def __str__(self):
        return f"{self.nom} ({self.longueur}x{self.largeur})"

@dataclass
class BatimentPlace:
    batiment: Batiment
    x: int
    y: int
    orientation: str  # 'H' ou 'V'
    
    def occupe_cases(self) -> List[Tuple[int, int]]:
        cases = []
        if self.orientation == 'H':
            for i in range(self.batiment.longueur):
                for j in range(self.batiment.largeur):
                    cases.append((self.x + i, self.y + j))
        else:  # Vertical
            for i in range(self.batiment.largeur):
                for j in range(self.batiment.longueur):
                    cases.append((self.x + i, self.y + j))
        return cases

class Terrain:
    def __init__(self, matrice: List[List[int]]):
        self.matrice = np.array(matrice)
        self.hauteur, self.largeur = self.matrice.shape
        self.cases_libres = [(i, j) for i in range(self.hauteur) 
                            for j in range(self.largeur) if self.matrice[i][j] == 1]
        self.batiments_places: List[BatimentPlace] = []
        self.cases_occupees: List[Tuple[int, int]] = []
        
    def est_place_valide(self, batiment: Batiment, x: int, y: int, orientation: str) -> bool:
        if orientation == 'H':
            if x + batiment.longueur > self.hauteur or y + batiment.largeur > self.largeur:
                return False
            for i in range(batiment.longueur):
                for j in range(batiment.largeur):
                    if (x + i, y + j) in self.cases_occupees or self.matrice[x + i][y + j] == 0:
                        return False
        else:  # Vertical
            if x + batiment.largeur > self.hauteur or y + batiment.longueur > self.largeur:
                return False
            for i in range(batiment.largeur):
                for j in range(batiment.longueur):
                    if (x + i, y + j) in self.cases_occupees or self.matrice[x + i][y + j] == 0:
                        return False
        return True
    
    def placer_batiment(self, batiment: Batiment, x: int, y: int, orientation: str):
        bp = BatimentPlace(batiment, x, y, orientation)
        self.batiments_places.append(bp)
        self.cases_occupees.extend(bp.occupe_cases())
        
    def enlever_batiment(self, index: int):
        bp = self.batiments_places.pop(index)
        for case in bp.occupe_cases():
            self.cases_occupees.remove(case)
        return bp
    
    def reste_assez_place(self, plus_grand_batiment: Batiment) -> bool:
        # Vérifie s'il reste assez de place pour le plus grand bâtiment non placé
        cases_libres_restantes = len([c for c in self.cases_libres if c not in self.cases_occupees])
        taille_batiment = plus_grand_batiment.longueur * plus_grand_batiment.largeur
        return cases_libres_restantes >= taille_batiment
    
    def trouver_emplacements_possibles(self, batiment: Batiment) -> List[Tuple[int, int, str]]:
        emplacements = []
        for i in range(self.hauteur):
            for j in range(self.largeur):
                # Essayer orientation horizontale
                if self.est_place_valide(batiment, i, j, 'H'):
                    emplacements.append((i, j, 'H'))
                # Essayer orientation verticale si différente
                if batiment.longueur != batiment.largeur:
                    if self.est_place_valide(batiment, i, j, 'V'):
                        emplacements.append((i, j, 'V'))
        return emplacements

class PlaceurBatiments:
    def __init__(self, terrain: Terrain, batiments: List[Batiment]):
        self.terrain = terrain
        self.batiments = batiments
        self.journal: List[str] = []
        self.max_journal_entries = 1000
        
    def journaliser(self, message: str):
        if len(self.journal) < self.max_journal_entries:
            self.journal.append(message)
            
    def calculer_culture_recue(self) -> Dict[str, float]:
        culture_totale = 0
        culture_par_type = {"Guérison": 0, "Nourriture": 0, "Or": 0}
        boosts_atteints = {"25%": 0, "50%": 0, "100%": 0}
        
        # Calcul de la culture reçue par chaque producteur
        for bp_producteur in [p for p in self.terrain.batiments_places 
                              if p.batiment.type == 'Producteur']:
            culture_recue = 0
            cases_producteur = bp_producteur.occupe_cases()
            
            for bp_culturel in [p for p in self.terrain.batiments_places 
                                if p.batiment.type == 'Culturel']:
                # Vérifier si le producteur est dans la zone de rayonnement du culturel
                zone_rayonnement = []
                for i in range(bp_culturel.x - bp_culturel.batiment.rayonnement, 
                               bp_culturel.x + bp_culturel.batiment.longueur + bp_culturel.batiment.rayonnement):
                    for j in range(bp_culturel.y - bp_culturel.batiment.rayonnement,
                                   bp_culturel.y + bp_culturel.batiment.largeur + bp_culturel.batiment.rayonnement):
                        if 0 <= i < self.terrain.hauteur and 0 <= j < self.terrain.largeur:
                            zone_rayonnement.append((i, j))
                
                if any(case in zone_rayonnement for case in cases_producteur):
                    culture_recue += bp_culturel.batiment.culture
            
            culture_totale += culture_recue
            
            # Déterminer le boost atteint
            if culture_recue >= bp_producteur.batiment.boost_100:
                boosts_atteints["100%"] += 1
                if bp_producteur.batiment.production in culture_par_type:
                    culture_par_type[bp_producteur.batiment.production] += culture_recue
            elif culture_recue >= bp_producteur.batiment.boost_50:
                boosts_atteints["50%"] += 1
            elif culture_recue >= bp_producteur.batiment.boost_25:
                boosts_atteints["25%"] += 1
        
        return {
            "culture_totale": culture_totale,
            "culture_par_type": culture_par_type,
            "boosts_atteints": boosts_atteints
        }
    
    def placer_batiments(self):
        # Tri des bâtiments par ordre de priorité
        neutres = [b for b in self.batiments if b.type == 'Neutre']
        culturels = [b for b in self.batiments if b.type == 'Culturel']
        producteurs = [b for b in self.batiments if b.type == 'Producteur']
        
        # Créer une liste de tous les bâtiments à placer avec leur quantité
        tous_batiments = []
        for b in neutres:
            for _ in range(b.quantite):
                tous_batiments.append(b)
        for b in culturels:
            for _ in range(b.quantite):
                tous_batiments.append(b)
        for b in producteurs:
            for _ in range(b.quantite):
                tous_batiments.append(b)
        
        # Trier par taille décroissante pour la vérification d'espace
        tous_batiments_tries = sorted(tous_batiments, 
                                      key=lambda x: x.longueur * x.largeur, 
                                      reverse=True)
        
        index = 0
        historique_placements = []
        
        while index < len(tous_batiments) and len(self.journal) < self.max_journal_entries:
            batiment = tous_batiments_tries[index]
            self.journaliser(f"Évaluation du bâtiment: {batiment}")
            
            # Trouver le plus grand bâtiment non placé
            non_places = tous_batiments_tries[index:]
            if non_places:
                plus_grand = max(non_places, key=lambda x: x.longueur * x.largeur)
            else:
                plus_grand = batiment
            
            # Chercher un emplacement
            emplacements = self.terrain.trouver_emplacements_possibles(batiment)
            
            place_trouve = False
            for x, y, orientation in emplacements:
                if self.terrain.reste_assez_place(plus_grand):
                    self.terrain.placer_batiment(batiment, x, y, orientation)
                    historique_placements.append(index)
                    self.journaliser(f"Bâtiment placé: {batiment} à ({x},{y}) en {orientation}")
                    place_trouve = True
                    break
            
            if place_trouve:
                index += 1
            else:
                # Retour en arrière
                if historique_placements:
                    dernier_index = historique_placements.pop()
                    batiment_enleve = tous_batiments_tries[dernier_index]
                    self.journaliser(f"Impossible de placer {batiment}, retrait de {batiment_enleve}")
                    
                    # Enlever le dernier bâtiment placé
                    for i, bp in enumerate(self.terrain.batiments_places):
                        if bp.batiment == batiment_enleve:
                            self.terrain.enlever_batiment(i)
                            break
                    
                    index = dernier_index
                else:
                    self.journaliser(f"Échec du placement pour {batiment}")
                    index += 1
    
    def generer_resultats(self) -> Dict:
        culture_data = self.calculer_culture_recue()
        
        # Calcul des cases non utilisées
        cases_non_utilisees = [c for c in self.terrain.cases_libres 
                               if c not in self.terrain.cases_occupees]
        
        # Bâtiments non placés
        batiments_places_noms = [bp.batiment.nom for bp in self.terrain.batiments_places]
        batiments_non_places = []
        cases_non_placees = 0
        
        for b in self.batiments:
            places = batiments_places_noms.count(b.nom)
            non_places = b.quantite - places
            if non_places > 0:
                batiments_non_places.append({
                    "nom": b.nom,
                    "non_places": non_places,
                    "cases": non_places * b.longueur * b.largeur
                })
                cases_non_placees += non_places * b.longueur * b.largeur
        
        return {
            "journal": self.journal,
            "culture_totale": culture_data["culture_totale"],
            "culture_par_type": culture_data["culture_par_type"],
            "boosts_atteints": culture_data["boosts_atteints"],
            "batiments_non_places": batiments_non_places,
            "cases_non_utilisees": len(cases_non_utilisees),
            "cases_non_placees": cases_non_placees
        }

def creer_terrain_depuis_excel(onglet_terrain: pd.DataFrame) -> Terrain:
    matrice = onglet_terrain.values.tolist()
    return Terrain(matrice)

def creer_batiments_depuis_excel(onglet_batiments: pd.DataFrame) -> List[Batiment]:
    batiments = []
    for _, row in onglet_batiments.iterrows():
        # Déterminer le type (par défaut Neutre si non spécifié)
        type_bat = row.get('Type', 'Neutre')
        if pd.isna(type_bat) or type_bat not in ['Culturel', 'Producteur']:
            type_bat = 'Neutre'
        
        batiment = Batiment(
            nom=str(row['Nom']),
            longueur=int(row['Longueur']),
            largeur=int(row['Largeur']),
            quantite=int(row['Quantite']),
            type=type_bat,
            culture=int(row.get('Culture', 0)) if not pd.isna(row.get('Culture', 0)) else 0,
            rayonnement=int(row.get('Rayonnement', 0)) if not pd.isna(row.get('Rayonnement', 0)) else 0,
            boost_25=int(row.get('Boost 25%', 0)) if not pd.isna(row.get('Boost 25%', 0)) else 0,
            boost_50=int(row.get('Boost 50%', 0)) if not pd.isna(row.get('Boost 50%', 0)) else 0,
            boost_100=int(row.get('Boost 100%', 0)) if not pd.isna(row.get('Boost 100%', 0)) else 0,
            production=str(row.get('Production', '')) if not pd.isna(row.get('Production', '')) else ''
        )
        batiments.append(batiment)
    return batiments

def exporter_resultats(terrain: Terrain, resultats: Dict, batiments: List[Batiment]) -> BytesIO:
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Feuille Journal
        journal_df = pd.DataFrame({
            'Étape': range(1, len(resultats['journal']) + 1),
            'Description': resultats['journal']
        })
        journal_df.to_excel(writer, sheet_name='Journal', index=False)
        
        # Feuille Statistiques
        stats_data = {
            'Métrique': ['Culture totale reçue', 
                        'Boost 25% atteints', 
                        'Boost 50% atteints', 
                        'Boost 100% atteints',
                        'Culture Guérison',
                        'Culture Nourriture',
                        'Culture Or',
                        'Cases non utilisées',
                        'Cases des bâtiments non placés'],
            'Valeur': [
                resultats['culture_totale'],
                resultats['boosts_atteints']['25%'],
                resultats['boosts_atteints']['50%'],
                resultats['boosts_atteints']['100%'],
                resultats['culture_par_type']['Guérison'],
                resultats['culture_par_type']['Nourriture'],
                resultats['culture_par_type']['Or'],
                resultats['cases_non_utilisees'],
                resultats['cases_non_placees']
            ]
        }
        stats_df = pd.DataFrame(stats_data)
        stats_df.to_excel(writer, sheet_name='Statistiques', index=False)
        
        # Feuille Bâtiments non placés
        if resultats['batiments_non_places']:
            non_places_df = pd.DataFrame(resultats['batiments_non_places'])
        else:
            non_places_df = pd.DataFrame({'nom': ['Aucun'], 'non_places': [0], 'cases': [0]})
        non_places_df.to_excel(writer, sheet_name='Non_places', index=False)
        
        # Feuille Terrain avec bâtiments placés
        terrain_viz = terrain.matrice.copy().astype(str)
        terrain_viz[terrain_viz == '1'] = '.'
        terrain_viz[terrain_viz == '0'] = '#'
        
        # Colorier les bâtiments placés
        for bp in terrain.batiments_places:
            couleur = {'Culturel': 'O', 'Producteur': 'P', 'Neutre': 'N'}[bp.batiment.type]
            for x, y in bp.occupe_cases():
                terrain_viz[x, y] = couleur
        
        terrain_df = pd.DataFrame(terrain_viz)
        terrain_df.to_excel(writer, sheet_name='Terrain_place', index=False, header=False)
    
    output.seek(0)
    return output

# Interface Streamlit
st.set_page_config(page_title="Placeur de Bâtiments", layout="wide")

st.title("🏗️ Placeur Automatique de Bâtiments")

st.markdown("""
### Instructions
1. Chargez votre fichier Excel avec deux onglets :
   - **Onglet 1** : Le terrain (matrice de 0 et 1)
   - **Onglet 2** : La liste des bâtiments
2. Cliquez sur "Lancer le placement"
3. Téléchargez les résultats
""")

uploaded_file = st.file_uploader("Choisissez votre fichier Excel", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Lecture du fichier Excel
        excel_file = pd.ExcelFile(uploaded_file)
        
        # Vérification des onglets
        if len(excel_file.sheet_names) < 2:
            st.error("Le fichier doit contenir au moins 2 onglets")
        else:
            st.success(f"Fichier chargé avec succès !")
            
            # Affichage des aperçus
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Aperçu du terrain")
                terrain_data = pd.read_excel(uploaded_file, sheet_name=excel_file.sheet_names[0])
                st.dataframe(terrain_data.head())
                
            with col2:
                st.subheader("Aperçu des bâtiments")
                batiments_data = pd.read_excel(uploaded_file, sheet_name=excel_file.sheet_names[1])
                st.dataframe(batiments_data.head())
            
            if st.button("🚀 Lancer le placement", type="primary"):
                with st.spinner("Placement en cours..."):
                    # Création du terrain et des bâtiments
                    terrain = creer_terrain_depuis_excel(terrain_data)
                    batiments = creer_batiments_depuis_excel(batiments_data)
                    
                    # Placement
                    placeur = PlaceurBatiments(terrain, batiments)
                    placeur.placer_batiments()
                    resultats = placeur.generer_resultats()
                    
                    # Affichage des résultats
                    st.success("Placement terminé !")
                    
                    # Métriques clés
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Culture totale", f"{resultats['culture_totale']:.0f}")
                    with col2:
                        st.metric("Boost 100%", resultats['boosts_atteints']['100%'])
                    with col3:
                        st.metric("Cases libres", resultats['cases_non_utilisees'])
                    with col4:
                        st.metric("Bâtiments non placés", 
                                 sum(b['non_places'] for b in resultats['batiments_non_places']))
                    
                    # Export des résultats
                    output_excel = exporter_resultats(terrain, resultats, batiments)
                    
                    st.download_button(
                        label="📥 Télécharger les résultats",
                        data=output_excel,
                        file_name="resultats_placement.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Aperçu du journal
                    with st.expander("📋 Voir le journal des placements"):
                        journal_df = pd.DataFrame({
                            'Étape': range(1, min(21, len(resultats['journal']) + 1)),
                            'Description': resultats['journal'][:20]
                        })
                        st.dataframe(journal_df)
                        if len(resultats['journal']) > 20:
                            st.info(f"... et {len(resultats['journal']) - 20} entrées supplémentaires")
    
    except Exception as e:
        st.error(f"Erreur lors du traitement : {str(e)}")
        st.exception(e)

st.markdown("---")
st.markdown("Développé pour le placement optimal de bâtiments")