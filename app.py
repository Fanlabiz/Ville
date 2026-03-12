import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import List, Tuple, Optional, Dict
import streamlit as st
from io import BytesIO
import copy
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

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
    
    def __repr__(self):
        return self.__str__()

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
        cases_libres_restantes = len([c for c in self.cases_libres if c not in self.cases_occupees])
        taille_batiment = plus_grand_batiment.longueur * plus_grand_batiment.largeur
        return cases_libres_restantes >= taille_batiment
    
    def trouver_emplacements_possibles(self, batiment: Batiment) -> List[Tuple[int, int, str]]:
        emplacements = []
        for i in range(self.hauteur):
            for j in range(self.largeur):
                if self.est_place_valide(batiment, i, j, 'H'):
                    emplacements.append((i, j, 'H'))
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
            
    def reclassifier_batiments(self):
        """Reclassifie les bâtiments Producteur qui ne produisent rien en Neutres"""
        for b in self.batiments:
            if b.type == 'Producteur' and (b.production == 'Rien' or b.production == '' or pd.isna(b.production)):
                b.type = 'Neutre'
                self.journaliser(f"Reclassement: {b.nom} passé de Producteur à Neutre (ne produit rien)")
    
    def calculer_culture_recue(self) -> Dict[str, float]:
        culture_totale = 0
        culture_par_type = {"Guérison": 0, "Nourriture": 0, "Or": 0, "Bijoux": 0, "Onguents": 0, "Cristal": 0, "Epices": 0, "Boiseries": 0, "Scriberie": 0}
        boosts_atteints = {"25%": 0, "50%": 0, "100%": 0}
        
        for bp_producteur in [p for p in self.terrain.batiments_places 
                              if p.batiment.type == 'Producteur']:
            culture_recue = 0
            cases_producteur = bp_producteur.occupe_cases()
            
            for bp_culturel in [p for p in self.terrain.batiments_places 
                                if p.batiment.type == 'Culturel']:
                # Calcul de la zone de rayonnement
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
            if culture_recue >= bp_producteur.batiment.boost_100 and bp_producteur.batiment.boost_100 > 0:
                boosts_atteints["100%"] += 1
            elif culture_recue >= bp_producteur.batiment.boost_50 and bp_producteur.batiment.boost_50 > 0:
                boosts_atteints["50%"] += 1
            elif culture_recue >= bp_producteur.batiment.boost_25 and bp_producteur.batiment.boost_25 > 0:
                boosts_atteints["25%"] += 1
            
            if bp_producteur.batiment.production in culture_par_type:
                culture_par_type[bp_producteur.batiment.production] += culture_recue
        
        return {
            "culture_totale": culture_totale,
            "culture_par_type": culture_par_type,
            "boosts_atteints": boosts_atteints
        }
    
    def placer_batiments(self):
        # Reclassification
        self.reclassifier_batiments()
        
        # Séparation par type
        neutres = [b for b in self.batiments if b.type == 'Neutre']
        culturels = [b for b in self.batiments if b.type == 'Culturel']
        producteurs = [b for b in self.batiments if b.type == 'Producteur']
        
        self.journaliser(f"=== RÉPARTITION ===")
        self.journaliser(f"Neutres: {len(neutres)} types")
        self.journaliser(f"Culturels: {len(culturels)} types")
        self.journaliser(f"Producteurs: {len(producteurs)} types")
        
        # Construction de la liste ordonnée
        tous_batiments = []
        
        # ÉTAPE 1: Tous les neutres d'abord
        neutres_tries = sorted(neutres, key=lambda x: x.longueur * x.largeur, reverse=True)
        self.journaliser(f"\n=== PLACEMENT DES NEUTRES (priorité) ===")
        for b in neutres_tries:
            for _ in range(b.quantite):
                tous_batiments.append(b)
                self.journaliser(f"Ajout à la file: {b} (Neutre)")
        
        # ÉTAPE 2: Alternance culturels et producteurs
        self.journaliser(f"\n=== PRÉPARATION DE L'ALTERNANCE ===")
        
        liste_culturels = []
        for b in culturels:
            for _ in range(b.quantite):
                liste_culturels.append(b)
                self.journaliser(f"Ajout Culturel: {b}")
        
        liste_producteurs = []
        for b in producteurs:
            for _ in range(b.quantite):
                liste_producteurs.append(b)
                self.journaliser(f"Ajout Producteur: {b}")
        
        # Alternance
        max_len = max(len(liste_culturels), len(liste_producteurs))
        for i in range(max_len):
            if i < len(liste_culturels):
                tous_batiments.append(liste_culturels[i])
                self.journaliser(f"Alternance - Ajout Culturel: {liste_culturels[i]}")
            if i < len(liste_producteurs):
                tous_batiments.append(liste_producteurs[i])
                self.journaliser(f"Alternance - Ajout Producteur: {liste_producteurs[i]}")
        
        self.journaliser(f"\n=== DÉBUT DU PLACEMENT ===")
        self.journaliser(f"Total bâtiments à placer: {len(tous_batiments)}")
        
        # Algorithme de placement avec backtracking
        index = 0
        historique_placements = []
        tentatives = {}
        
        while index < len(tous_batiments) and len(self.journal) < self.max_journal_entries:
            batiment = tous_batiments[index]
            self.journaliser(f"Évaluation du bâtiment: {batiment}")
            
            non_places = tous_batiments[index:]
            if non_places:
                plus_grand = max(non_places, key=lambda x: x.longueur * x.largeur)
            else:
                plus_grand = batiment
            
            emplacements = self.terrain.trouver_emplacements_possibles(batiment)
            
            place_trouve = False
            for x, y, orientation in emplacements:
                if self.terrain.reste_assez_place(plus_grand):
                    self.terrain.placer_batiment(batiment, x, y, orientation)
                    historique_placements.append(index)
                    self.journaliser(f"✅ Bâtiment placé: {batiment} à ({x},{y}) en {orientation}")
                    place_trouve = True
                    break
            
            if place_trouve:
                index += 1
            else:
                if historique_placements:
                    dernier_index = historique_placements.pop()
                    batiment_enleve = tous_batiments[dernier_index]
                    self.journaliser(f"❌ Impossible de placer {batiment}, RETRAIT de {batiment_enleve}")
                    
                    for i, bp in enumerate(self.terrain.batiments_places):
                        if bp.batiment == batiment_enleve:
                            self.terrain.enlever_batiment(i)
                            break
                    
                    index = dernier_index
                    
                    key = f"{batiment.nom}_{index}"
                    tentatives[key] = tentatives.get(key, 0) + 1
                    if tentatives[key] > 10:
                        self.journaliser(f"⚠️ Trop de tentatives pour {batiment.nom}, passage au suivant")
                        index += 1
                else:
                    self.journaliser(f"❌ Échec du placement pour {batiment}")
                    index += 1
    
    def generer_resultats(self) -> Dict:
        culture_data = self.calculer_culture_recue()
        
        cases_non_utilisees = [c for c in self.terrain.cases_libres 
                               if c not in self.terrain.cases_occupees]
        
        batiments_places_noms = [bp.batiment.nom for bp in self.terrain.batiments_places]
        batiments_non_places = []
        cases_non_placees = 0
        
        for b in self.batiments:
            places = batiments_places_noms.count(b.nom)
            non_places = b.quantite - places
            if non_places > 0:
                batiments_non_places.append({
                    "nom": b.nom,
                    "type": b.type,
                    "non_places": non_places,
                    "cases": non_places * b.longueur * b.largeur
                })
                cases_non_placees += non_places * b.longueur * b.largeur
        
        stats_placement = {
            "total_batiments": sum(b.quantite for b in self.batiments),
            "batiments_places": len(self.terrain.batiments_places),
            "neutres_places": len([b for b in self.terrain.batiments_places if b.batiment.type == 'Neutre']),
            "culturels_places": len([b for b in self.terrain.batiments_places if b.batiment.type == 'Culturel']),
            "producteurs_places": len([b for b in self.terrain.batiments_places if b.batiment.type == 'Producteur']),
        }
        
        return {
            "journal": self.journal,
            "stats_placement": stats_placement,
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
        nom = str(row['Nom']).strip()
        type_bat = str(row.get('Type', 'Neutre')).strip() if not pd.isna(row.get('Type', 'Neutre')) else 'Neutre'
        production = str(row.get('Production', '')).strip() if not pd.isna(row.get('Production', '')) else ''
        
        if type_bat == 'Producteur' and (production == 'Rien' or production == '' or production == 'nan'):
            type_bat = 'Neutre'
        
        batiment = Batiment(
            nom=nom,
            longueur=int(row['Longueur']),
            largeur=int(row['Largeur']),
            quantite=int(row['Quantite']),
            type=type_bat,
            culture=int(row.get('Culture', 0)) if not pd.isna(row.get('Culture', 0)) else 0,
            rayonnement=int(row.get('Rayonnement', 0)) if not pd.isna(row.get('Rayonnement', 0)) else 0,
            boost_25=int(row.get('Boost 25%', 0)) if not pd.isna(row.get('Boost 25%', 0)) else 0,
            boost_50=int(row.get('Boost 50%', 0)) if not pd.isna(row.get('Boost 50%', 0)) else 0,
            boost_100=int(row.get('Boost 100%', 0)) if not pd.isna(row.get('Boost 100%', 0)) else 0,
            production=production
        )
        batiments.append(batiment)
    return batiments

def afficher_terrain_couleurs_streamlit(terrain: Terrain):
    """Affiche le terrain avec des couleurs dans Streamlit"""
    # Définition des couleurs
    couleurs = {
        'Culturel': '#FFA500',  # Orange
        'Producteur': '#90EE90',  # Vert clair
        'Neutre': '#D3D3D3',  # Gris clair
        'Occupee': '#404040',  # Gris foncé
        'Libre': '#FFFFFF'  # Blanc
    }
    
    # Légende
    st.markdown("### 🎨 Légende des couleurs")
    cols = st.columns(5)
    with cols[0]:
        st.markdown(f"<div style='background-color: {couleurs['Culturel']}; padding: 10px; text-align: center; border-radius: 5px; color: black; font-weight: bold;'>Culturel</div>", unsafe_allow_html=True)
    with cols[1]:
        st.markdown(f"<div style='background-color: {couleurs['Producteur']}; padding: 10px; text-align: center; border-radius: 5px; color: black; font-weight: bold;'>Producteur</div>", unsafe_allow_html=True)
    with cols[2]:
        st.markdown(f"<div style='background-color: {couleurs['Neutre']}; padding: 10px; text-align: center; border-radius: 5px; color: black; font-weight: bold;'>Neutre</div>", unsafe_allow_html=True)
    with cols[3]:
        st.markdown(f"<div style='background-color: {couleurs['Occupee']}; padding: 10px; text-align: center; border-radius: 5px; color: white; font-weight: bold;'>Occupé</div>", unsafe_allow_html=True)
    with cols[4]:
        st.markdown(f"<div style='background-color: {couleurs['Libre']}; padding: 10px; text-align: center; border-radius: 5px; color: black; border: 1px solid black; font-weight: bold;'>Libre</div>", unsafe_allow_html=True)
    
    # Création du tableau HTML
    html = "<div style='overflow-x: auto;'><table style='border-collapse: collapse; margin-top: 20px;'>"
    
    # Ajouter les en-têtes de colonnes
    html += "<tr><th style='background-color: #f0f0f0; border: 1px solid #ccc; padding: 5px;'>Y/X</th>"
    for j in range(terrain.largeur):
        html += f"<th style='background-color: #f0f0f0; border: 1px solid #ccc; padding: 5px;'>{j}</th>"
    html += "</tr>"
    
    for i in range(terrain.hauteur):
        html += f"<tr><th style='background-color: #f0f0f0; border: 1px solid #ccc; padding: 5px;'>{i}</th>"
        for j in range(terrain.largeur):
            case = (i, j)
            batiment_trouve = None
            for bp in terrain.batiments_places:
                if case in bp.occupe_cases():
                    batiment_trouve = bp
                    break
            
            if batiment_trouve:
                couleur = couleurs[batiment_trouve.batiment.type]
                texte = batiment_trouve.batiment.nom[0]  # Première lettre
                titre = batiment_trouve.batiment.nom
                couleur_texte = 'black' if batiment_trouve.batiment.type in ['Producteur', 'Neutre'] else 'white'
            else:
                if terrain.matrice[i][j] == 1:
                    couleur = couleurs['Libre']
                    texte = '.'
                    titre = 'Case libre'
                    couleur_texte = 'black'
                else:
                    couleur = couleurs['Occupee']
                    texte = '#'
                    titre = 'Case occupée (0)'
                    couleur_texte = 'white'
            
            html += f"<td style='background-color: {couleur}; width: 35px; height: 35px; text-align: center; border: 1px solid #666; color: {couleur_texte}; font-weight: bold; cursor: help;' title='{titre}'>{texte}</td>"
        html += "</tr>"
    
    html += "</table></div>"
    
    st.markdown(html, unsafe_allow_html=True)

def exporter_resultats_excel(terrain: Terrain, resultats: Dict, batiments: List[Batiment]) -> BytesIO:
    """Exporte les résultats avec couleurs dans Excel"""
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
            'Métrique': [
                'Total bâtiments', 'Bâtiments placés', 'Neutres placés', 'Culturels placés', 'Producteurs placés',
                'Culture totale reçue', 
                'Boost 25% atteints', 'Boost 50% atteints', 'Boost 100% atteints',
                'Culture Guérison', 'Culture Nourriture', 'Culture Or',
                'Cases non utilisées', 'Cases des bâtiments non placés'
            ],
            'Valeur': [
                resultats['stats_placement']['total_batiments'],
                resultats['stats_placement']['batiments_places'],
                resultats['stats_placement']['neutres_places'],
                resultats['stats_placement']['culturels_places'],
                resultats['stats_placement']['producteurs_places'],
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
            non_places_df = pd.DataFrame({'nom': ['Aucun'], 'type': [''], 'non_places': [0], 'cases': [0]})
        non_places_df.to_excel(writer, sheet_name='Non_places', index=False)
        
        # Feuille Placements détaillés
        placements_data = []
        for i, bp in enumerate(terrain.batiments_places):
            placements_data.append({
                'Ordre': i+1,
                'Nom': bp.batiment.nom,
                'Type': bp.batiment.type,
                'Position X': bp.x,
                'Position Y': bp.y,
                'Orientation': bp.orientation,
                'Taille': f"{bp.batiment.longueur}x{bp.batiment.largeur}"
            })
        placements_df = pd.DataFrame(placements_data)
        placements_df.to_excel(writer, sheet_name='Placements', index=False)
        
        # Feuille Terrain avec couleurs
        terrain_viz = terrain.matrice.copy()
        terrain_df = pd.DataFrame(terrain_viz)
        terrain_df.to_excel(writer, sheet_name='Terrain_couleurs', index=False, header=False)
        
        # Récupérer le workbook pour appliquer les couleurs
        workbook = writer.book
        sheet = workbook['Terrain_couleurs']
        
        # Définir les couleurs (format Excel)
        couleur_culturel = PatternFill(start_color='FFA500', end_color='FFA500', fill_type='solid')  # Orange
        couleur_producteur = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')  # Vert clair
        couleur_neutre = PatternFill(start_color='D3D3D3', end_color='D3D3D3', fill_type='solid')  # Gris clair
        couleur_occupee = PatternFill(start_color='404040', end_color='404040', fill_type='solid')  # Gris foncé
        couleur_libre = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')  # Blanc
        
        # Appliquer les couleurs
        for i in range(terrain.hauteur):
            for j in range(terrain.largeur):
                cell = sheet.cell(row=i+1, column=j+1)
                
                case = (i, j)
                batiment_trouve = None
                for bp in terrain.batiments_places:
                    if case in bp.occupe_cases():
                        batiment_trouve = bp
                        break
                
                if batiment_trouve:
                    if batiment_trouve.batiment.type == 'Culturel':
                        cell.fill = couleur_culturel
                        cell.font = Font(color='000000')  # Noir
                    elif batiment_trouve.batiment.type == 'Producteur':
                        cell.fill = couleur_producteur
                        cell.font = Font(color='000000')
                    else:  # Neutre
                        cell.fill = couleur_neutre
                        cell.font = Font(color='000000')
                    cell.value = batiment_trouve.batiment.nom[:3]  # 3 premières lettres
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    if terrain.matrice[i][j] == 1:
                        cell.fill = couleur_libre
                        cell.value = '.'
                        cell.font = Font(color='000000')
                    else:
                        cell.fill = couleur_occupee
                        cell.value = '#'
                        cell.font = Font(color='FFFFFF')  # Blanc
                
                # Bordures
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                
                # Ajuster la largeur
                sheet.column_dimensions[get_column_letter(j+1)].width = 6
        
        # Ajouter une feuille de légende
        sheet_legende = workbook.create_sheet('Légende')
        legendes = [
            ['Couleur', 'Type', 'Description'],
            ['Orange', 'Culturel', 'Bâtiments culturels'],
            ['Vert clair', 'Producteur', 'Bâtiments producteurs'],
            ['Gris clair', 'Neutre', 'Bâtiments neutres'],
            ['Gris foncé', 'Occupé', 'Cases occupées (0)'],
            ['Blanc', 'Libre', 'Cases libres (1)']
        ]
        
        for i, ligne in enumerate(legendes, 1):
            for j, valeur in enumerate(ligne, 1):
                cell = sheet_legende.cell(row=i, column=j, value=valeur)
                cell.border = Border(
                    left=Side(style='thin'),
                    right=Side(style='thin'),
                    top=Side(style='thin'),
                    bottom=Side(style='thin')
                )
                if i == 1:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color='F0F0F0', end_color='F0F0F0', fill_type='solid')
                elif valeur == 'Orange':
                    cell.fill = couleur_culturel
                elif valeur == 'Vert clair':
                    cell.fill = couleur_producteur
                elif valeur == 'Gris clair':
                    cell.fill = couleur_neutre
                elif valeur == 'Gris foncé':
                    cell.fill = couleur_occupee
                    cell.font = Font(color='FFFFFF')
                elif valeur == 'Blanc':
                    cell.fill = couleur_libre
        
        # Ajuster les colonnes de la légende
        sheet_legende.column_dimensions['A'].width = 15
        sheet_legende.column_dimensions['B'].width = 15
        sheet_legende.column_dimensions['C'].width = 25
    
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
3. Visualisez le résultat en couleurs
4. Téléchargez les résultats
""")

uploaded_file = st.file_uploader("Choisissez votre fichier Excel", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Lecture du fichier Excel
        excel_file = pd.ExcelFile(uploaded_file)
        sheet_names = excel_file.sheet_names
        
        if len(sheet_names) < 2:
            st.error("Le fichier doit contenir au moins 2 onglets")
        else:
            # Identification des onglets
            terrain_sheet = next((s for s in sheet_names if 'terrain' in s.lower()), sheet_names[0])
            batiments_sheet = next((s for s in sheet_names if 'batiment' in s.lower()), sheet_names[1])
            
            st.success(f"✅ Fichier chargé avec succès !")
            st.info(f"📄 Onglet Terrain: **{terrain_sheet}** | 📄 Onglet Bâtiments: **{batiments_sheet}**")
            
            # Chargement des données
            terrain_data = pd.read_excel(uploaded_file, sheet_name=terrain_sheet, header=None)
            batiments_data = pd.read_excel(uploaded_file, sheet_name=batiments_sheet)
            
            # Affichage des aperçus
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("📊 Aperçu du terrain")
                st.dataframe(terrain_data.head(10), use_container_width=True)
                st.text(f"📐 Dimensions: {terrain_data.shape[0]} lignes x {terrain_data.shape[1]} colonnes")
                
            with col2:
                st.subheader("🏭 Aperçu des bâtiments")
                st.dataframe(batiments_data.head(10), use_container_width=True)
                st.text(f"📦 Nombre de types: {len(batiments_data)}")
            
            if st.button("🚀 Lancer le placement", type="primary", use_container_width=True):
                with st.spinner("Placement en cours... (cela peut prendre quelques secondes)"):
                    # Création du terrain et des bâtiments
                    terrain = creer_terrain_depuis_excel(terrain_data)
                    batiments = creer_batiments_depuis_excel(batiments_data)
                    
                    # Placement
                    placeur = PlaceurBatiments(terrain, batiments)
                    placeur.placer_batiments()
                    resultats = placeur.generer_resultats()
                    
                    # Affichage des résultats
                    st.success("✅ Placement terminé avec succès !")
                    
                    # Visualisation du terrain en couleurs
                    st.subheader("🗺️ Visualisation du terrain placé")
                    afficher_terrain_couleurs_streamlit(terrain)
                    
                    # Métriques clés
                    st.subheader("📊 Métriques clés")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("🎨 Culture totale", f"{resultats['culture_totale']:.0f}")
                    with col2:
                        st.metric("⚡ Boost 100%", resultats['boosts_atteints']['100%'])
                    with col3:
                        st.metric("⬜ Cases libres", resultats['cases_non_utilisees'])
                    with col4:
                        non_places = sum(b['non_places'] for b in resultats['batiments_non_places'])
                        st.metric("📦 Bâtiments non placés", non_places)
                    
                    # Statistiques détaillées
                    with st.expander("📈 Voir les statistiques détaillées"):
                        col1, col2 = st.columns(2)
                        with col1:
                            st.json({
                                "Total bâtiments": resultats['stats_placement']['total_batiments'],
                                "Bâtiments placés": resultats['stats_placement']['batiments_places'],
                                "Neutres placés": resultats['stats_placement']['neutres_places'],
                                "Culturels placés": resultats['stats_placement']['culturels_places'],
                                "Producteurs placés": resultats['stats_placement']['producteurs_places']
                            })
                        with col2:
                            st.json({
                                "Culture par type": resultats['culture_par_type'],
                                "Boosts atteints": resultats['boosts_atteints']
                            })
                    
                    # Bâtiments non placés
                    if resultats['batiments_non_places']:
                        with st.expander("⚠️ Bâtiments non placés"):
                            non_places_df = pd.DataFrame(resultats['batiments_non_places'])
                            st.dataframe(non_places_df, use_container_width=True)
                    
                    # Journal
                    with st.expander("📋 Voir le journal des placements"):
                        journal_df = pd.DataFrame({
                            'Étape': range(1, min(51, len(resultats['journal']) + 1)),
                            'Description': resultats['journal'][:50]
                        })
                        st.dataframe(journal_df, use_container_width=True)
                        if len(resultats['journal']) > 50:
                            st.info(f"... et {len(resultats['journal']) - 50} entrées supplémentaires")
                    
                    # Export des résultats
                    st.subheader("💾 Téléchargement")
                    output_excel = exporter_resultats_excel(terrain, resultats, batiments)
                    
                    st.download_button(
                        label="📥 Télécharger les résultats (Excel avec couleurs)",
                        data=output_excel,
                        file_name="resultats_placement.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
    
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement : {str(e)}")
        st.exception(e)

st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    Développé pour le placement optimal de bâtiments<br>
    Version avec visualisation couleur 🌈
</div>
""", unsafe_allow_html=True)