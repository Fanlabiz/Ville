def match_buildings_with_info(self, buildings_info):
    """
    Associe les bâtiments placés à leurs caractéristiques
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
                    found = True
                    break
            
            if not found:
                unmatched.append(f"{nom} à ({r},{c}) dimensions {w}x{h}")
                self.log(f"⚠️ Attention: bâtiment '{nom}' dimensions {w}x{h} ne correspond à aucune orientation")
        else:
            unmatched.append(f"{nom} à ({r},{c})")
            self.log(f"⚠️ Attention: bâtiment '{nom}' non trouvé dans la liste ou quantité insuffisante")
    
    if unmatched:
        st.warning(f"Bâtiments non associés: {len(unmatched)}")
        for u in unmatched[:10]:  # Afficher les 10 premiers
            st.write(f"  - {u}")
    
    # Recompter les cases libres
    self.initial_free_cells = np.sum(self.grid == 1)