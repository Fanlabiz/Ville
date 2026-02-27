import pandas as pd
import numpy as np
from dataclasses import dataclass
from typing import List, Tuple, Optional
import copy
import sys
import os

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
            
            print(f"✅ Terrain chargé: {self.terrain_grid.shape}")
            print(f"✅ Nombre de bâtiments à placer: {sum(b.quantity for b in self.buildings)}")
            
        except Exception as e:
            print(f"❌ Erreur lors du chargement des données: {e}")
            raise
    
    def can_place_building(self, x: int, y: int, length: int, width: int) -> bool:
        """
        Vérifie si un bâtiment peut être placé à la position donnée
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
                    total_culture += placed_building['building'].culture
        
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
        """
        print("\n🏗️  Début du placement des bâtiments...")
        
        # Réinitialiser
        self.placement_grid = np.zeros_like(self.terrain_grid)
        self.placement_grid[self.terrain_grid == 1] = -1
        self.placed_buildings = []
        
        # Trier les bâtiments par importance (ceux qui boostent en premier)
        sorted_buildings = sorted(self.buildings, 
                                 key=lambda b: b.boost_100, 
                                 reverse=True)
        
        placement_log = []
        
        for building in sorted_buildings:
            print(f"   Placement de {building.quantity} {building.name}...")
            placed_count = 0
            
            for _ in range(building.quantity):
                placement = self.find_best_placement(building)
                
                if placement:
                    self.place_building(
                        building,
                        placement['x'],
                        placement['y'],
                        placement['orientation']
                    )
                    placed_count += 1
                    placement_log.append({
                        'building': building.name,
                        'x': placement['x'],
                        'y': placement['y'],
                        'orientation': placement['orientation'],
                        'boost': placement['boost']
                    })
                else:
                    print(f"   ⚠️  Plus d'espace pour {building.name}")
                    break
            
            print(f"   ✅ {placed_count}/{building.quantity} placés")
        
        # Calculer la production finale
        total_production = self.calculate_total_production()
        
        print(f"\n📊 Résultats finaux:")
        print(f"   Bâtiments placés: {len(self.placed_buildings)}")
        print(f"   Production totale: {total_production:.2f}")
        
        return {
            'total_production': total_production,
            'buildings_placed': len(self.placed_buildings),
            'grid': self.placement_grid.copy(),
            'placed_buildings': copy.deepcopy(self.placed_buildings),
            'placement_log': placement_log
        }
    
    def save_results(self, output_file: str, results: dict):
        """
        Sauvegarde les résultats dans un fichier Excel
        
        Args:
            output_file: Chemin du fichier de sortie
            results: Résultats de l'optimisation
        """
        print(f"\n💾 Sauvegarde des résultats dans {output_file}...")
        
        # Créer un DataFrame pour les résultats détaillés
        placement_data = []
        for placed in results['placed_buildings']:
            building = placed['building']
            placement_data.append({
                'Nom': building.name,
                'Position_X': placed['x'],
                'Position_Y': placed['y'],
                'Orientation': placed['orientation'],
                'Longueur': placed['length'],
                'Largeur': placed['width'],
                'Culture_base': building.culture,
                'Boost': placed.get('current_boost', 1.0),
                'Production_finale': building.culture * placed.get('current_boost', 1.0)
            })
        
        df_placements = pd.DataFrame(placement_data)
        
        # Créer un DataFrame pour la grille
        df_grid = pd.DataFrame(results['grid'])
        
        # Créer un DataFrame pour les statistiques
        stats_data = [{
            'Production_totale': results['total_production'],
            'Batiments_places': results['buildings_placed'],
            'Batiments_total': sum(b.quantity for b in self.buildings),
            'Cases_occupees': np.sum(results['grid'] > 0),
            'Cases_libres': np.sum(results['grid'] == 0),
            'Cases_obstruees': np.sum(results['grid'] == -1),
            'Taux_occupation': f"{np.sum(results['grid'] > 0) / results['grid'].size * 100:.1f}%"
        }]
        df_stats = pd.DataFrame(stats_data)
        
        # Sauvegarder dans un fichier Excel
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df_placements.to_excel(writer, sheet_name='Placements', index=False)
            df_grid.to_excel(writer, sheet_name='Grille_placement', index=False, header=False)
            df_stats.to_excel(writer, sheet_name='Statistiques', index=False)
        
        print(f"✅ Résultats sauvegardés avec succès!")
        print(f"   - {len(placement_data)} bâtiments placés")
        print(f"   - Production totale: {results['total_production']:.2f}")
    
    def print_grid_summary(self):
        """Affiche un résumé de la grille"""
        print("\n🗺️  Résumé de la grille:")
        print(f"   Dimensions: {self.placement_grid.shape[0]}x{self.placement_grid.shape[1]}")
        print(f"   Cases libres: {np.sum(self.placement_grid == 0)}")
        print(f"   Cases occupées (bâtiments): {np.sum(self.placement_grid > 0)}")
        print(f"   Cases obstruées: {np.sum(self.placement_grid == -1)}")

def main():
    """Fonction principale d'exécution"""
    
    # Vérifier les arguments en ligne de commande
    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        # Demander le fichier d'entrée
        input_file = input("📂 Entrez le chemin du fichier Excel d'entrée: ").strip()
    
    # Définir le fichier de sortie
    if len(sys.argv) > 2:
        output_file = sys.argv[2]
    else:
        # Générer un nom de fichier de sortie par défaut
        base_name = os.path.splitext(input_file)[0]
        output_file = f"{base_name}_resultats.xlsx"
    
    print("=" * 50)
    print("🏗️  OPTIMISEUR DE PLACEMENT DE BÂTIMENTS")
    print("=" * 50)
    
    try:
        # Vérifier que le fichier d'entrée existe
        if not os.path.exists(input_file):
            print(f"❌ Erreur: Le fichier {input_file} n'existe pas.")
            return
        
        print(f"\n📂 Fichier d'entrée: {input_file}")
        print(f"📂 Fichier de sortie: {output_file}")
        
        # Créer le placeur
        placer = BuildingPlacer(input_file)
        
        # Afficher le résumé initial
        placer.print_grid_summary()
        
        # Placer tous les bâtiments
        results = placer.place_all_buildings()
        
        # Sauvegarder les résultats
        placer.save_results(output_file, results)
        
        print("\n" + "=" * 50)
        print("✅ OPTIMISATION TERMINÉE AVEC SUCCÈS")
        print("=" * 50)
        
    except FileNotFoundError:
        print(f"❌ Erreur: Le fichier {input_file} n'a pas été trouvé.")
        print("Veuillez vérifier le chemin du fichier.")
    except Exception as e:
        print(f"❌ Erreur inattendue: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()