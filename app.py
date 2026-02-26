import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import List, Tuple, Optional
import copy

@dataclass
class Building:
    """Classe représentant un bâtiment"""
    name: str
    length: int
    width: int
    quantity: int
    culture: float
    radius: int
    boost_25: float
    boost_50: float
    boost_100: float
    
    def get_dimensions(self, orientation: str) -> Tuple[int, int]:
        """Retourne les dimensions selon l'orientation"""
        if orientation == 'H':
            return self.length, self.width
        else:  # 'V'
            return self.width, self.length

class BuildingPlacer:
    def __init__(self, terrain_file: str):
        """
        Initialise le placeur de bâtiments
        
        Args:
            terrain_file: Chemin vers le fichier Excel
        """
        self.terrain_file = terrain_file
        self.terrain_grid = None
        self.buildings = []
        self.placement_grid = None
        self.boost_grid = None
        self.placed_buildings = []
        
        self.load_data()
        
    def load_data(self):
        """Charge les données depuis le fichier Excel"""
        try:
            # Chargement du terrain
            df_terrain = pd.read_excel(self.terrain_file, sheet_name=0, header=None)
            self.terrain_grid = df_terrain.values.astype(int)
            
            # Chargement des bâtiments
            df_buildings = pd.read_excel(self.terrain_file, sheet_name=1)
            
            for _, row in df_buildings.iterrows():
                building = Building(
                    name=row['Nom'],
                    length=int(row['longueur']),
                    width=int(row['largeur']),
                    quantity=int(row['quantité']),
                    culture=float(row['culture']),
                    radius=int(row['rayonnement']),
                    boost_25=float(row['Boost 25%']),
                    boost_50=float(row['Boost 50%']),
                    boost_100=float(row['Boost 100%'])
                )
                self.buildings.append(building)
            
            # Initialisation des grilles
            self.placement_grid = np.zeros_like(self.terrain_grid)
            self.boost_grid = np.zeros_like(self.terrain_grid, dtype=float)
            
            # Marquer les cases occupées
            self.placement_grid[self.terrain_grid == 1] = -1
            
            print(f"Terrain chargé: {self.terrain_grid.shape}")
            print(f"Nombre de bâtiments à placer: {sum(b.quantity for b in self.buildings)}")
            
        except Exception as e:
            print(f"Erreur lors du chargement des données: {e}")
            raise
    
    def can_place_building(self, x: int, y: int, length: int, width: int) -> bool:
        """
        Vérifie si un bâtiment peut être placé à la position donnée
        
        Args:
            x, y: Position du coin supérieur gauche
            length, width: Dimensions du bâtiment
            
        Returns:
            True si le placement est possible
        """
        # Vérifier les limites du terrain
        if x + length > self.terrain_grid.shape[0] or y + width > self.terrain_grid.shape[1]:
            return False
        
        # Vérifier que toutes les cases sont libres
        for i in range(length):
            for j in range(width):
                if self.placement_grid[x + i, y + j] != 0:
                    return False
        
        return True
    
    def calculate_boost(self, building: Building, x: int, y: int, 
                       length: int, width: int) -> float:
        """
        Calcule le boost pour un bâtiment à la position donnée
        
        Returns:
            Multiplicateur de boost (1.0 = 100%, 1.25 = 125%, etc.)
        """
        total_culture = 0
        center_x = x + length // 2
        center_y = y + width // 2
        
        # Parcourir toutes les cases dans le rayon
        for i in range(max(0, center_x - building.radius), 
                      min(self.terrain_grid.shape[0], center_x + building.radius + 1)):
            for j in range(max(0, center_y - building.radius),
                          min(self.terrain_grid.shape[1], center_y + building.radius + 1)):
                # Vérifier si c'est un bâtiment qui produit de la culture
                if self.placement_grid[i, j] > 0:
                    placed_building = self.placed_buildings[self.placement_grid[i, j] - 1]
                    total_culture += placed_building.culture
        
        # Déterminer le niveau de boost
        if total_culture >= building.boost_100:
            return 2.0  # +100%
        elif total_culture >= building.boost_50:
            return 1.5  # +50%
        elif total_culture >= building.boost_25:
            return 1.25  # +25%
        else:
            return 1.0  # Pas de boost
    
    def place_building(self, building: Building, x: int, y: int, 
                      orientation: str) -> bool:
        """
        Place un bâtiment sur le terrain
        
        Returns:
            True si le placement a réussi
        """
        length, width = building.get_dimensions(orientation)
        
        if not self.can_place_building(x, y, length, width):
            return False
        
        # Placer le bâtiment
        building_id = len(self.placed_buildings) + 1
        for i in range(length):
            for j in range(width):
                self.placement_grid[x + i, y + j] = building_id
        
        # Enregistrer le bâtiment placé
        placed_info = {
            'building': building,
            'x': x,
            'y': y,
            'orientation': orientation,
            'length': length,
            'width': width,
            'building_id': building_id
        }
        self.placed_buildings.append(placed_info)
        
        return True
    
    def calculate_total_production(self) -> float:
        """Calcule la production totale avec les boosts"""
        total = 0
        
        # Réinitialiser la grille de boost
        self.boost_grid.fill(0)
        
        # Calculer les boosts pour chaque bâtiment
        for placed in self.placed_buildings:
            building = placed['building']
            boost = self.calculate_boost(
                building, 
                placed['x'], 
                placed['y'],
                placed['length'],
                placed['width']
            )
            placed['current_boost'] = boost
            total += building.culture * boost
        
        return total
    
    def find_best_placement(self, building: Building) -> Optional[dict]:
        """
        Trouve la meilleure position pour un bâtiment
        
        Returns:
            Dictionnaire avec les infos de placement ou None
        """
        best_score = -1
        best_placement = None
        
        # Essayer les deux orientations
        for orientation in ['H', 'V']:
            length, width = building.get_dimensions(orientation)
            
            # Parcourir toutes les positions possibles
            for x in range(self.terrain_grid.shape[0] - length + 1):
                for y in range(self.terrain_grid.shape[1] - width + 1):
                    if self.can_place_building(x, y, length, width):
                        # Calculer le boost potentiel
                        boost = self.calculate_boost(building, x, y, length, width)
                        
                        # Score basé sur le boost et la position
                        score = boost * building.culture
                        
                        if score > best_score:
                            best_score = score
                            best_placement = {
                                'x': x,
                                'y': y,
                                'orientation': orientation,
                                'boost': boost,
                                'score': score
                            }
        
        return best_placement
    
    def place_all_buildings(self) -> dict:
        """
        Place tous les bâtiments en optimisant les boosts
        
        Returns:
            Statistiques du placement
        """
        # Trier les bâtiments par importance (ceux qui boostent en premier)
        sorted_buildings = sorted(self.buildings, 
                                 key=lambda b: b.boost_100, 
                                 reverse=True)
        
        for building in sorted_buildings:
            for _ in range(building.quantity):
                placement = self.find_best_placement(building)
                
                if placement:
                    self.place_building(
                        building,
                        placement['x'],
                        placement['y'],
                        placement['orientation']
                    )
                    print(f"Placé {building.name} à ({placement['x']}, {placement['y']}) "
                          f"avec boost {placement['boost']:.2f}")
                else:
                    print(f"Impossible de placer {building.name} - plus d'espace disponible")
                    break
        
        # Calculer la production finale
        total_production = self.calculate_total_production()
        
        return {
            'total_production': total_production,
            'buildings_placed': len(self.placed_buildings),
            'grid': self.placement_grid.copy(),
            'placed_buildings': copy.deepcopy(self.placed_buildings)
        }
    
    def save_results(self, output_file: str):
        """
        Sauvegarde les résultats dans un fichier
        
        Args:
            output_file: Chemin du fichier de sortie
        """
        # Créer un DataFrame pour les résultats
        results = []
        for placed in self.placed_buildings:
            building = placed['building']
            results.append({
                'Nom': building.name,
                'Position_X': placed['x'],
                'Position_Y': placed['y'],
                'Orientation': placed['orientation'],
                'Boost_actuel': placed.get('current_boost', 1.0),
                'Production': building.culture * placed.get('current_boost', 1.0)
            })
        
        df_results = pd.DataFrame(results)
        
        # Sauvegarder dans un fichier Excel
        with pd.ExcelWriter(output_file) as writer:
            df_results.to_excel(writer, sheet_name='Placements', index=False)
            
            # Sauvegarder la grille de placement
            df_grid = pd.DataFrame(self.placement_grid)
            df_grid.to_excel(writer, sheet_name='Grille_placement', index=False)
            
            # Statistiques
            stats = pd.DataFrame([{
                'Total_production': self.calculate_total_production(),
                'Nombre_batiments_places': len(self.placed_buildings),
                'Cases_occupees': np.sum(self.placement_grid > 0)
            }])
            stats.to_excel(writer, sheet_name='Statistiques', index=False)
        
        print(f"Résultats sauvegardés dans {output_file}")
    
    def visualize_grid(self):
        """Affiche une visualisation textuelle de la grille"""
        print("\nGrille de placement:")
        print("-" * (self.placement_grid.shape[1] * 3 + 1))
        
        for i in range(self.placement_grid.shape[0]):
            row = "|"
            for j in range(self.placement_grid.shape[1]):
                if self.placement_grid[i, j] == -1:
                    row += " X "
                elif self.placement_grid[i, j] == 0:
                    row += " . "
                else:
                    row += f" {self.placement_grid[i, j]:1d} "
            row += "|"
            print(row)
        
        print("-" * (self.placement_grid.shape[1] * 3 + 1))

# Fonction principale
def main():
    """Fonction principale d'exécution"""
    
    # Configuration
    input_file = "terrain_batiments.xlsx"  # À modifier avec votre fichier
    output_file = "resultats_placement.xlsx"
    
    print("Démarrage du placement de bâtiments...")
    
    try:
        # Créer le placeur
        placer = BuildingPlacer(input_file)
        
        # Placer tous les bâtiments
        stats = placer.place_all_buildings()
        
        # Afficher les résultats
        print(f"\nRésultats du placement:")
        print(f"Bâtiments placés: {stats['buildings_placed']}")
        print(f"Production totale: {stats['total_production']:.2f}")
        
        # Visualiser la grille
        placer.visualize_grid()
        
        # Sauvegarder les résultats
        placer.save_results(output_file)
        
    except FileNotFoundError:
        print(f"Erreur: Le fichier {input_file} n'a pas été trouvé.")
        print("Veuillez vérifier le chemin du fichier.")
    except Exception as e:
        print(f"Erreur inattendue: {e}")

if __name__ == "__main__":
    main()