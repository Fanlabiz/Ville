import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import List, Tuple, Optional, Dict
import copy

@dataclass
class Batiment:
    nom: str
    longueur: int
    largeur: int
    quantite: int
    type: str  # "culturel", "producteur", "neutre"
    culture: int = 0
    rayonnement: int = 0
    boost_25: int = 0
    boost_50: int = 0
    boost_100: int = 0
    production: str = ""
    
@dataclass
class Position:
    x: int
    y: int
    orientation: str  # "H" ou "V"
    
class PlacementBatiments:
    def __init__(self, fichier_excel: str):
        self.fichier_excel = fichier_excel
        self.terrain = None
        self.hauteur = 0
        self.largeur = 0
        self.batiments = []
        self.batiments_places = []  # Liste des (batiment, position)
        self.culture_par_case = None  # Carte de la culture émise
        self.boosts_obtenus = {}  # Par batiment producteur
        
    def charger_donnees(self):
        """Charge les données depuis le fichier Excel"""
        # Lecture du terrain
        df_terrain = pd.read_excel(self.fichier_excel, sheet_name=0, header=None)
        self.terrain = df_terrain.values
        self.hauteur, self.largeur = self.terrain.shape
        
        # Initialisation de la carte de culture
        self.culture_par_case = np.zeros((self.hauteur, self.largeur))
        
        # Lecture des bâtiments
        df_batiments = pd.read_excel(self.fichier_excel, sheet_name=1)
        for _, row in df_batiments.iterrows():
            # Déterminer le type
            if pd.isna(row['Production']) or row['Production'] == "":
                if row['Culture'] > 0:
                    type_bat = "culturel"
                else:
                    type_bat = "neutre"
            else:
                type_bat = "producteur"
                
            batiment = Batiment(
                nom=row['Nom'],
                longueur=int(row['Longueur']),
                largeur=int(row['Largeur']),
                quantite=int(row['Quantite']),
                type=type_bat,
                culture=row['Culture'] if not pd.isna(row['Culture']) else 0,
                rayonnement=row['Rayonnement'] if not pd.isna(row['Rayonnement']) else 0,
                boost_25=row['Boost 25%'] if not pd.isna(row['Boost 25%']) else 0,
                boost_50=row['Boost 50%'] if not pd.isna(row['Boost 50%']) else 0,
                boost_100=row['Boost 100%'] if not pd.isna(row['Boost 100%']) else 0,
                production=row['Production'] if not pd.isna(row['Production']) else ""
            )
            
            # Ajouter plusieurs fois selon la quantité
            for _ in range(batiment.quantite):
                self.batiments.append(copy.deepcopy(batiment))
    
    def est_case_libre(self, x: int, y: int) -> bool:
        """Vérifie si une case est libre (dans le terrain et non occupée)"""
        if x < 0 or x >= self.hauteur or y < 0 or y >= self.largeur:
            return False
        return self.terrain[x, y] == 1
    
    def peut_placer_batiment(self, batiment: Batiment, pos: Position) -> bool:
        """Vérifie si on peut placer un bâtiment à une position donnée"""
        if pos.orientation == "H":
            longueur, largeur = batiment.longueur, batiment.largeur
        else:
            longueur, largeur = batiment.largeur, batiment.longueur
            
        # Vérifier que toutes les cases sont libres
        for i in range(longueur):
            for j in range(largeur):
                x = pos.x + i
                y = pos.y + j
                if not self.est_case_libre(x, y):
                    return False
        return True
    
    def placer_batiment(self, batiment: Batiment, pos: Position):
        """Place un bâtiment sur le terrain"""
        if pos.orientation == "H":
            longueur, largeur = batiment.longueur, batiment.largeur
        else:
            longueur, largeur = batiment.largeur, batiment.longueur
            
        # Marquer les cases comme occupées
        for i in range(longueur):
            for j in range(largeur):
                x = pos.x + i
                y = pos.y + j
                self.terrain[x, y] = 0  # 0 = occupé
        
        # Ajouter aux bâtiments placés
        self.batiments_places.append((batiment, pos))
        
        # Si c'est un bâtiment culturel, mettre à jour la culture
        if batiment.type == "culturel":
            self.ajouter_culture(batiment, pos)
    
    def enlever_dernier_batiment(self):
        """Enlève le dernier bâtiment placé (backtracking)"""
        if not self.batiments_places:
            return
        
        batiment, pos = self.batiments_places.pop()
        
        # Si c'était un bâtiment culturel, enlever sa culture
        if batiment.type == "culturel":
            self.enlever_culture(batiment, pos)
        
        # Remettre les cases comme libres
        if pos.orientation == "H":
            longueur, largeur = batiment.longueur, batiment.largeur
        else:
            longueur, largeur = batiment.largeur, batiment.longueur
            
        for i in range(longueur):
            for j in range(largeur):
                x = pos.x + i
                y = pos.y + j
                self.terrain[x, y] = 1  # 1 = libre
        
        return batiment, pos
    
    def ajouter_culture(self, batiment: Batiment, pos: Position):
        """Ajoute la culture émise par un bâtiment culturel"""
        # Déterminer les limites du rayonnement
        if pos.orientation == "H":
            x_min = max(0, pos.x - batiment.rayonnement)
            x_max = min(self.hauteur, pos.x + batiment.longueur + batiment.rayonnement)
            y_min = max(0, pos.y - batiment.rayonnement)
            y_max = min(self.largeur, pos.y + batiment.largeur + batiment.rayonnement)
        else:
            x_min = max(0, pos.x - batiment.rayonnement)
            x_max = min(self.hauteur, pos.x + batiment.largeur + batiment.rayonnement)
            y_min = max(0, pos.y - batiment.rayonnement)
            y_max = min(self.largeur, pos.y + batiment.longueur + batiment.rayonnement)
        
        # Ajouter la culture à toutes les cases dans le rayonnement
        for x in range(x_min, x_max):
            for y in range(y_min, y_max):
                self.culture_par_case[x, y] += batiment.culture
    
    def enlever_culture(self, batiment: Batiment, pos: Position):
        """Enlève la culture émise par un bâtiment culturel"""
        # Mêmes calculs que pour ajouter_culture
        if pos.orientation == "H":
            x_min = max(0, pos.x - batiment.rayonnement)
            x_max = min(self.hauteur, pos.x + batiment.longueur + batiment.rayonnement)
            y_min = max(0, pos.y - batiment.rayonnement)
            y_max = min(self.largeur, pos.y + batiment.largeur + batiment.rayonnement)
        else:
            x_min = max(0, pos.x - batiment.rayonnement)
            x_max = min(self.hauteur, pos.x + batiment.largeur + batiment.rayonnement)
            y_min = max(0, pos.y - batiment.rayonnement)
            y_max = min(self.largeur, pos.y + batiment.longueur + batiment.rayonnement)
        
        for x in range(x_min, x_max):
            for y in range(y_min, y_max):
                self.culture_par_case[x, y] -= batiment.culture
    
    def calculer_culture_recue(self, batiment: Batiment, pos: Position) -> int:
        """Calcule la culture totale reçue par un bâtiment producteur à une position"""
        if batiment.type != "producteur":
            return 0
            
        culture_totale = 0
        if pos.orientation == "H":
            longueur, largeur = batiment.longueur, batiment.largeur
        else:
            longueur, largeur = batiment.largeur, batiment.longueur
            
        for i in range(longueur):
            for j in range(largeur):
                x = pos.x + i
                y = pos.y + j
                culture_totale += self.culture_par_case[x, y]
        
        return culture_totale
    
    def trouver_positions_possibles(self, batiment: Batiment) -> List[Position]:
        """Trouve toutes les positions possibles pour un bâtiment"""
        positions = []
        
        # Essayer les deux orientations
        for orientation in ["H", "V"]:
            if orientation == "H":
                max_x = self.hauteur - batiment.longueur + 1
                max_y = self.largeur - batiment.largeur + 1
            else:
                max_x = self.hauteur - batiment.largeur + 1
                max_y = self.largeur - batiment.longueur + 1
            
            for x in range(max_x):
                for y in range(max_y):
                    pos = Position(x, y, orientation)
                    if self.peut_placer_batiment(batiment, pos):
                        positions.append(pos)
        
        return positions
    
    def calculer_taille_plus_grand_restant(self, index_batiment: int) -> int:
        """Calcule la taille (surface) du plus grand bâtiment restant à placer"""
        batiments_restants = self.batiments[index_batiment:]
        if not batiments_restants:
            return 0
        
        # Surface du plus grand (longueur * largeur)
        return max(b.longueur * b.largeur for b in batiments_restants)
    
    def assez_de_place_pour_plus_grand(self, index_batiment: int) -> bool:
        """Vérifie s'il reste assez de place pour le plus grand bâtiment restant"""
        taille_plus_grand = self.calculer_taille_plus_grand_restant(index_batiment)
        if taille_plus_grand == 0:
            return True
        
        # Compter les cases libres
        cases_libres = np.sum(self.terrain == 1)
        return cases_libres >= taille_plus_grand
    
    def evaluer_position_culturel(self, batiment: Batiment, pos: Position) -> float:
        """Évalue la qualité d'une position pour un bâtiment culturel"""
        score = 0
        
        # Simuler le placement pour voir la culture ajoutée
        self.placer_batiment(batiment, pos)
        
        # Parcourir toutes les cases pour voir si elles pourraient booster des producteurs
        for x in range(self.hauteur):
            for y in range(self.largeur):
                if self.terrain[x, y] == 0:  # Case occupée par un producteur
                    # Vérifier si ce producteur a besoin de boost
                    # (on ne peut pas savoir directement, on estime)
                    if self.culture_par_case[x, y] < 100:  # Seuil arbitraire
                        score += self.culture_par_case[x, y]
        
        # Enlever le bâtiment
        self.enlever_dernier_batiment()
        
        return score
    
    def evaluer_position_producteur(self, batiment: Batiment, pos: Position) -> float:
        """Évalue la qualité d'une position pour un bâtiment producteur"""
        culture_recue = self.calculer_culture_recue(batiment, pos)
        
        # Déterminer le boost obtenu
        if culture_recue >= batiment.boost_100:
            boost = 100
        elif culture_recue >= batiment.boost_50:
            boost = 50
        elif culture_recue >= batiment.boost_25:
            boost = 25
        else:
            boost = 0
        
        # Le score est la culture reçue (on maximise les boosts)
        return culture_recue
    
    def evaluer_position_neutre(self, batiment: Batiment, pos: Position) -> float:
        """Évalue la qualité d'une position pour un bâtiment neutre"""
        # Favoriser les bords du terrain
        if pos.orientation == "H":
            if pos.x == 0 or pos.x + batiment.longueur == self.hauteur:
                return 100  # Bord haut ou bas
            if pos.y == 0 or pos.y + batiment.largeur == self.largeur:
                return 50   # Bord gauche ou droit
        else:
            if pos.x == 0 or pos.x + batiment.largeur == self.hauteur:
                return 100
            if pos.y == 0 or pos.y + batiment.longueur == self.largeur:
                return 50
        
        return 10  # Intérieur
    
    def placer_batiment_avec_regles(self, batiment: Batiment, index_batiment: int) -> bool:
        """Place un bâtiment selon les règles spécifiques à son type"""
        positions = self.trouver_positions_possibles(batiment)
        
        if not positions:
            return False
        
        # Évaluer chaque position selon le type
        positions_evaluees = []
        for pos in positions:
            if batiment.type == "culturel":
                score = self.evaluer_position_culturel(batiment, pos)
            elif batiment.type == "producteur":
                score = self.evaluer_position_producteur(batiment, pos)
            else:  # neutre
                score = self.evaluer_position_neutre(batiment, pos)
            
            positions_evaluees.append((score, pos))
        
        # Trier par score décroissant
        positions_evaluees.sort(reverse=True)
        
        # Essayer chaque position dans l'ordre
        for score, pos in positions_evaluees:
            # Placer le bâtiment
            self.placer_batiment(batiment, pos)
            
            # Vérifier s'il reste assez de place pour le plus grand bâtiment restant
            if self.assez_de_place_pour_plus_grand(index_batiment + 1):
                return True
            else:
                # Pas assez de place, backtracking
                print(f"Backtracking - retour au bâtiment {batiment.nom}")
                self.enlever_dernier_batiment()
        
        return False
    
    def placer_tous_batiments(self) -> bool:
        """Place tous les bâtiments en suivant l'ordre spécifié"""
        # Séparer les bâtiments par type
        neutres = [b for b in self.batiments if b.type == "neutre"]
        culturels = [b for b in self.batiments if b.type == "culturel"]
        producteurs = [b for b in self.batiments if b.type == "producteur"]
        
        # Ordre de placement : d'abord tous les neutres, puis alternance culturel/producteur
        ordre_placement = neutres.copy()
        
        # Ajouter en alternance
        max_len = max(len(culturels), len(producteurs))
        for i in range(max_len):
            if i < len(culturels):
                ordre_placement.append(culturels[i])
            if i < len(producteurs):
                ordre_placement.append(producteurs[i])
        
        # Placer chaque bâtiment dans l'ordre
        for i, batiment in enumerate(ordre_placement):
            print(f"Placement du bâtiment {batiment.nom} ({i+1}/{len(ordre_placement)})")
            
            if not self.placer_batiment_avec_regles(batiment, i):
                print(f"Impossible de placer {batiment.nom}, backtracking nécessaire")
                # Ici, il faudrait implémenter un backtracking plus poussé
                # qui remonte plusieurs bâtiments en arrière
                return False
        
        return True
    
    def calculer_boosts_finaux(self):
        """Calcule les boosts obtenus par chaque bâtiment producteur"""
        boosts = {}
        for batiment, pos in self.batiments_places:
            if batiment.type == "producteur":
                culture_recue = self.calculer_culture_recue(batiment, pos)
                
                if culture_recue >= batiment.boost_100:
                    boost = 100
                elif culture_recue >= batiment.boost_50:
                    boost = 50
                elif culture_recue >= batiment.boost_25:
                    boost = 25
                else:
                    boost = 0
                
                boosts[batiment.nom] = {
                    "culture": culture_recue,
                    "boost": boost,
                    "production": batiment.production
                }
        
        return boosts
    
    def afficher_resultats(self):
        """Affiche les résultats du placement"""
        print("\n" + "="*50)
        print("RÉSULTATS DU PLACEMENT")
        print("="*50)
        
        # Afficher le terrain final
        print("\nTerrain final (0=occupé, 1=libre):")
        for i in range(self.hauteur):
            ligne = ""
            for j in range(self.largeur):
                if self.terrain[i, j] == 1:
                    ligne += "□ "  # Case libre
                else:
                    # Chercher quel bâtiment occupe cette case
                    trouve = False
                    for batiment, pos in self.batiments_places:
                        if pos.orientation == "H":
                            if pos.x <= i < pos.x + batiment.longueur and pos.y <= j < pos.y + batiment.largeur:
                                if batiment.type == "culturel":
                                    ligne += "C "  # Culturel
                                elif batiment.type == "producteur":
                                    ligne += "P "  # Producteur
                                else:
                                    ligne += "N "  # Neutre
                                trouve = True
                                break
                        else:
                            if pos.x <= i < pos.x + batiment.largeur and pos.y <= j < pos.y + batiment.longueur:
                                if batiment.type == "culturel":
                                    ligne += "C "
                                elif batiment.type == "producteur":
                                    ligne += "P "
                                else:
                                    ligne += "N "
                                trouve = True
                                break
                    
                    if not trouve:
                        ligne += "■ "  # Case occupée par obstacle initial
            print(ligne)
        
        # Calculer et afficher les boosts
        boosts = self.calculer_boosts_finaux()
        print("\nBoosts obtenus:")
        for nom, infos in boosts.items():
            print(f"{nom}: {infos['culture']} culture → boost {infos['boost']}%")
        
        # Taux d'occupation
        cases_libres = np.sum(self.terrain == 1)
        cases_occupees = self.hauteur * self.largeur - cases_libres
        print(f"\nTaux d'occupation: {cases_occupees}/{self.hauteur * self.largeur} cases")

def main():
    # Utilisation
    placement = PlacementBatiments("terrain_batiments.xlsx")
    placement.charger_donnees()
    
    print("Début du placement des bâtiments...")
    success = placement.placer_tous_batiments()
    
    if success:
        print("\n✅ Tous les bâtiments ont été placés avec succès!")
        placement.afficher_resultats()
    else:
        print("\n❌ Échec: impossible de placer tous les bâtiments")

if __name__ == "__main__":
    main()