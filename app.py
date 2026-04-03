import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import copy
from itertools import product, combinations

st.set_page_config(page_title="Optimiseur de Ville", layout="wide")
st.title("🏙️ Optimiseur de placement de bâtiments")

# ─────────────────────────────────────────────
# LECTURE DU FICHIER INPUT
# ─────────────────────────────────────────────

def lire_fichier(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    
    # Terrain (sans header)
    terrain_df = pd.read_excel(uploaded_file, sheet_name=xl.sheet_names[0], header=None)
    
    # Bâtiments (première ligne = header)
    bat_df = pd.read_excel(uploaded_file, sheet_name=xl.sheet_names[1], header=0)
    bat_df.columns = [str(c).strip() for c in bat_df.columns]
    
    return terrain_df, bat_df

def construire_terrain(terrain_df):
    """
    Retourne :
      - grid : matrice numpy de strings ('' = vide, 'X' = bord, nom = bâtiment)
      - inside_mask : booléens, True si la case est à l'intérieur du périmètre X
    """
    rows, cols = terrain_df.shape
    grid = np.full((rows, cols), '', dtype=object)
    
    for r in range(rows):
        for c in range(cols):
            val = terrain_df.iloc[r, c]
            if pd.notna(val) and str(val).strip() != '':
                grid[r, c] = str(val).strip()
    
    # Calcul du masque intérieur : flood-fill depuis l'extérieur
    border_mask = (grid == 'X')
    inside_mask = np.zeros((rows, cols), dtype=bool)
    
    # BFS depuis les bords non-X
    from collections import deque
    outside = np.zeros((rows, cols), dtype=bool)
    queue = deque()
    for r in range(rows):
        for c in range(cols):
            if not border_mask[r, c]:
                if r == 0 or r == rows-1 or c == 0 or c == cols-1:
                    outside[r, c] = True
                    queue.append((r, c))
    
    while queue:
        r, c = queue.popleft()
        for dr, dc in [(-1,0),(1,0),(0,-1),(0,1)]:
            nr, nc = r+dr, c+dc
            if 0 <= nr < rows and 0 <= nc < cols:
                if not outside[nr, nc] and not border_mask[nr, nc]:
                    outside[nr, nc] = True
                    queue.append((nr, nc))
    
    inside_mask = ~outside & ~border_mask
    return grid, inside_mask, border_mask

def extraire_batiments(grid, inside_mask, bat_df):
    """
    Parcourt la grille et identifie les bâtiments placés (position top-left).
    Retourne une liste de dicts.
    """
    rows, cols = grid.shape
    
    # Index des bâtiments connus
    bat_info = {}
    for _, row in bat_df.iterrows():
        nom = str(row.get('Nom', '')).strip()
        if nom:
            bat_info[nom] = {
                'longueur': int(row.get('Longueur', 1)),
                'largeur': int(row.get('Largeur', 1)),
                'type': str(row.get('Type', 'Neutre')).strip(),
                'culture': float(row.get('Culture', 0) or 0),
                'rayonnement': int(row.get('Rayonnement', 0) or 0),
                'boost25': float(row.get('Boost 25%', 0) or 0),
                'boost50': float(row.get('Boost 50%', 0) or 0),
                'boost100': float(row.get('Boost 100%', 0) or 0),
                'production': str(row.get('Production', 'Rien')).strip(),
                'quantite': float(row.get('Quantite', 0) or 0),
                'priorite': int(row.get('Priorite', 0) or 0),
            }
    
    visited = np.zeros((rows, cols), dtype=bool)
    placed = []
    
    for r in range(rows):
        for c in range(cols):
            val = grid[r, c]
            if val and val != 'X' and not visited[r, c]:
                nom = val
                if nom not in bat_info:
                    continue
                info = bat_info[nom]
                # Déterminer orientation : on teste les deux
                # La case (r,c) est le coin top-left
                L, l = info['longueur'], info['largeur']
                # Chercher toutes les cases de ce bâtiment autour de (r,c)
                placed_bat = None
                for orientation in ['H', 'V']:
                    if orientation == 'H':
                        h_cells, w_cells = l, L   # hauteur, largeur en cases
                    else:
                        h_cells, w_cells = L, l
                    
                    ok = True
                    cells = []
                    for dr in range(h_cells):
                        for dc in range(w_cells):
                            nr, nc = r+dr, c+dc
                            if nr >= rows or nc >= cols:
                                ok = False; break
                            if not (grid[nr, nc] == nom or grid[nr, nc] == ''):
                                # accepter si c'est le même nom
                                if grid[nr, nc] != nom:
                                    ok = False; break
                            cells.append((nr, nc))
                        if not ok:
                            break
                    
                    if ok:
                        # Vérifier que toutes les cellules portent ce nom ou sont vides mais adjacentes
                        # Approche simplifiée : vérifier que la cellule top-left est bien 'nom'
                        placed_bat = {
                            'nom': nom,
                            'row': r, 'col': c,
                            'h': h_cells, 'w': w_cells,
                            'orientation': orientation,
                            **{k: v for k, v in info.items()}
                        }
                        for nr, nc in cells:
                            visited[nr, nc] = True
                        break
                
                if placed_bat:
                    placed.append(placed_bat)
    
    return placed, bat_info

# ─────────────────────────────────────────────
# CALCUL DE LA CULTURE ET DES BOOSTS
# ─────────────────────────────────────────────

def calculer_culture(placed_list, grid_shape):
    """Pour chaque bâtiment producteur, calcule la culture reçue."""
    rows, cols = grid_shape
    
    # Carte de culture : somme de culture en chaque case
    culture_map = np.zeros((rows, cols), dtype=float)
    
    for bat in placed_list:
        if bat['type'] == 'Culturel' and bat['rayonnement'] > 0:
            ray = bat['rayonnement']
            r0, c0 = bat['row'], bat['col']
            r1, c1 = r0 + bat['h'] - 1, c0 + bat['w'] - 1
            
            # Zone de rayonnement : bande autour du bâtiment
            for r in range(r0 - ray, r1 + ray + 1):
                for c in range(c0 - ray, c1 + ray + 1):
                    if 0 <= r < rows and 0 <= c < cols:
                        # Vérifier que c'est dans la bande (pas à l'intérieur du bâtiment)
                        inside_bat = (r0 <= r <= r1) and (c0 <= c <= c1)
                        if not inside_bat:
                            culture_map[r, c] += bat['culture']
    
    # Maintenant calculer la culture reçue par chaque producteur
    for bat in placed_list:
        if bat['type'] == 'Producteur':
            r0, c0 = bat['row'], bat['col']
            r1, c1 = r0 + bat['h'] - 1, c0 + bat['w'] - 1
            
            # Culture reçue = max dans les cases du bâtiment? Non : somme des culturels qui le touchent
            # Règle : si le producteur est (partiellement) dans le rayonnement d'un culturel → reçoit sa culture
            # On a déjà tout dans culture_map, mais il faut éviter les doublons
            # → Recalcul direct par bâtiment culturel
            culture_recue = 0.0
            for cult in placed_list:
                if cult['type'] == 'Culturel' and cult['rayonnement'] > 0:
                    ray = cult['rayonnement']
                    cr0, cc0 = cult['row'], cult['col']
                    cr1, cc1 = cr0 + cult['h'] - 1, cc0 + cult['w'] - 1
                    
                    zone_r0 = cr0 - ray
                    zone_c0 = cc0 - ray
                    zone_r1 = cr1 + ray
                    zone_c1 = cc1 + ray
                    
                    # Le producteur est-il dans cette zone ?
                    overlap_r = (r0 <= zone_r1) and (r1 >= zone_r0)
                    overlap_c = (c0 <= zone_c1) and (c1 >= zone_c0)
                    
                    if overlap_r and overlap_c:
                        culture_recue += cult['culture']
            
            bat['culture_recue'] = culture_recue
            
            # Calcul du boost
            b25 = bat.get('boost25', 0)
            b50 = bat.get('boost50', 0)
            b100 = bat.get('boost100', 0)
            
            if b100 and culture_recue >= b100:
                bat['boost'] = 100
                bat['boost_factor'] = 2.0
            elif b50 and culture_recue >= b50:
                bat['boost'] = 50
                bat['boost_factor'] = 1.5
            elif b25 and culture_recue >= b25:
                bat['boost'] = 25
                bat['boost_factor'] = 1.25
            else:
                bat['boost'] = 0
                bat['boost_factor'] = 1.0
            
            bat['prod_boosted'] = bat['quantite'] * bat['boost_factor']
        else:
            bat['culture_recue'] = 0.0
            bat['boost'] = 0
            bat['boost_factor'] = 1.0
            bat['prod_boosted'] = 0.0
    
    return culture_map

def score_total(placed_list, prio_order=None):
    """
    Score à deux composantes pour respecter strictement la hiérarchie de priorité :

    Composante 1 (dominante) : ratio de boost normalisé, pondéré par priorité individuelle
        → un bâtiment priorité 1 à 0% boost pèse PLUS qu'un bâtiment priorité 4 à 100%
        → indépendant de la quantité produite (évite que les gros producteurs écrasent les petits)

    Composante 2 (secondaire, facteur 1e-6) : production absolue boostée × poids type
        → départage à égalité de boost, favorise les gros producteurs et les types prioritaires
    """
    if prio_order is None:
        prio_order = ['Guérison', 'Nourriture', 'Or']

    base_weights = {}
    for i, p in enumerate(prio_order):
        base_weights[p] = 10 ** (len(prio_order) - i)

    MAX_PRIO = 4

    score = 0.0
    for bat in placed_list:
        if bat['type'] != 'Producteur' or bat['production'] == 'Rien':
            continue

        w_prod = base_weights.get(bat['production'], 1)
        prio   = bat['priorite'] or 0
        w_prio = (MAX_PRIO - prio + 1) if 1 <= prio <= MAX_PRIO else 1

        boost_f       = bat['boost_factor']
        prod_actuelle = bat['prod_boosted']
        qte           = bat['quantite']

        ratio_atteint = boost_f - 1.0
        c1 = w_prio * (ratio_atteint + 0.15 * (1.0 - ratio_atteint))

        # Composante 2 : départage par production absolue
        b100 = bat['boost100'] or 0
        b50  = bat['boost50']  or 0
        b25  = bat['boost25']  or 0
        cr   = bat['culture_recue']
        if b100 and cr < b100:
            gain_pot = qte * 2.0 - prod_actuelle
        elif b50 and cr < b50:
            gain_pot = qte * 1.5 - prod_actuelle
        elif b25 and cr < b25:
            gain_pot = qte * 1.25 - prod_actuelle
        else:
            gain_pot = 0
        c2 = w_prod * (prod_actuelle + 0.1 * gain_pot)

        score += c1 * 1e6 + c2

    return score


def optimiser(placed_list, inside_mask, grid_shape, prio_order, n_passes=3, progress_cb=None):
    """
    Optimisation rapide en 4 phases par passe.
    Techniques de performance :
    - Calcul incrémental de culture_recue (évite de tout recalculer)
    - Modification in-place + restauration (évite deepcopy)
    - Occupation sous forme de set de cases (test O(1))
    - Positions candidates pré-filtrées (cases libres d'abord)
    """
    # ── Représentation interne rapide ──────────────────────────────────────
    # On travaille sur placed_list directement (in-place), avec sauvegarde/restore.
    # Chaque bâtiment est un dict mutable.

    rows, cols = grid_shape
    MAX_PRIO = 4  # niveaux de priorité individuelle (1=plus important)
    best = copy.deepcopy(placed_list)
    calculer_culture(best, grid_shape)
    best_score = score_total(best, prio_order)

    # ── Helpers ────────────────────────────────────────────────────────────

    def cells_of(b):
        """Cases occupées par un bâtiment."""
        return [(b['row']+dr, b['col']+dc)
                for dr in range(b['h']) for dc in range(b['w'])]

    def build_occ(state, exclude_ids=()):
        """Set des cases occupées (excluant certains indices)."""
        s = set()
        for k, b in enumerate(state):
            if k not in exclude_ids:
                for dr in range(b['h']):
                    for dc in range(b['w']):
                        s.add((b['row']+dr, b['col']+dc))
        return s

    def free_positions(h, w, occ):
        """Positions (r,c) valides : dans le terrain et non occupées."""
        libres, occupees = [], []
        for r in range(rows - h + 1):
            for c in range(cols - w + 1):
                cases = [(r+dr, c+dc) for dr in range(h) for dc in range(w)]
                if not all(inside_mask[r2, c2] for r2, c2 in cases):
                    continue
                if any(p in occ for p in cases):
                    occupees.append((r, c))
                else:
                    libres.append((r, c))
        return libres + occupees   # cases libres testées en premier

    # ── Recalcul de la culture ──────────────────────────────────────────────
    # recalc_all : recalcule TOUS les producteurs en Python pur (rapide sur petits N)
    # recalc_partial : recalcule seulement les producteurs affectés par un déplacement

    def _culture_pour_prod(b, state_cults):
        """Culture reçue par un producteur depuis une liste de culturels."""
        r0, c0 = b['row'], b['col']
        r1, c1 = r0+b['h']-1, c0+b['w']-1
        total = 0.0
        for cult in state_cults:
            ray = cult['rayonnement']
            cr0, cc0 = cult['row'], cult['col']
            if r0 <= cr0+cult['h']-1+ray and r1 >= cr0-ray and                c0 <= cc0+cult['w']-1+ray and c1 >= cc0-ray:
                total += cult['culture']
        return total

    def _apply_boost_bat(b, cr):
        b['culture_recue'] = cr
        b25 = b.get('boost25', 0) or 0
        b50 = b.get('boost50', 0) or 0
        b100 = b.get('boost100', 0) or 0
        if b100 and cr >= b100:
            b['boost'], b['boost_factor'] = 100, 2.0
        elif b50 and cr >= b50:
            b['boost'], b['boost_factor'] = 50, 1.5
        elif b25 and cr >= b25:
            b['boost'], b['boost_factor'] = 25, 1.25
        else:
            b['boost'], b['boost_factor'] = 0, 1.0
        b['prod_boosted'] = b['quantite'] * b['boost_factor']

    def apply_boost(bat):
        _apply_boost_bat(bat, bat.get('culture_recue', 0))

    def culture_recue_bat(prod, state):
        cults = [b for b in state if b['type'] == 'Culturel' and b['rayonnement'] > 0]
        return _culture_pour_prod(prod, cults)

    def recalc_all_numpy(state):
        """Recalcule tous les producteurs."""
        cults = [b for b in state if b['type'] == 'Culturel' and b['rayonnement'] > 0]
        for b in state:
            if b['type'] == 'Producteur':
                _apply_boost_bat(b, _culture_pour_prod(b, cults))
            else:
                b['culture_recue'] = 0.0
                b['boost'] = 0; b['boost_factor'] = 1.0; b['prod_boosted'] = 0.0

    def recalc_producteurs(state, moved_cult=None, old_pos=None):
        """
        Recalcule culture et boost.
        moved_cult : index du culturel déplacé (recalc partiel)
        old_pos    : (r,c,h,w) ancienne position du culturel
        """
        if moved_cult is not None and old_pos is not None:
            cult = state[moved_cult]
            ray = cult['rayonnement']
            or0, oc0, oh, ow = old_pos
            or1, oc1 = or0+oh-1, oc0+ow-1
            nr0, nc0 = cult['row'], cult['col']
            nr1, nc1 = nr0+cult['h']-1, nc0+cult['w']-1
            cults = [b for b in state if b['type'] == 'Culturel' and b['rayonnement'] > 0]
            for b in state:
                if b['type'] != 'Producteur':
                    continue
                r0b, c0b = b['row'], b['col']
                r1b, c1b = r0b+b['h']-1, c0b+b['w']-1
                in_old = r0b<=or1+ray and r1b>=or0-ray and c0b<=oc1+ray and c1b>=oc0-ray
                in_new = r0b<=nr1+ray and r1b>=nr0-ray and c0b<=nc1+ray and c1b>=nc0-ray
                if in_old or in_new:
                    _apply_boost_bat(b, _culture_pour_prod(b, cults))
        else:
            recalc_all_numpy(state)

    def recalc_two_cults(state, ci1, old1, ci2, old2):
        """
        Recalc partiel après déplacement simultané de deux culturels.
        Recalcule uniquement les producteurs touchés par l'ancienne OU la nouvelle
        zone de rayonnement de l'un ou l'autre des deux culturels.
        old1, old2 : (r, c, h, w) — anciennes positions des deux culturels.
        """
        c1b = state[ci1]
        c2b = state[ci2]
        ray1 = c1b['rayonnement']
        ray2 = c2b['rayonnement']

        # Zones touchées (ancienne + nouvelle) pour chaque culturel
        or1, oc1, oh1, ow1 = old1
        nr1, nc1 = c1b['row'], c1b['col']
        nh1, nw1 = c1b['h'], c1b['w']

        or2, oc2, oh2, ow2 = old2
        nr2, nc2 = c2b['row'], c2b['col']
        nh2, nw2 = c2b['h'], c2b['w']

        cults = [b for b in state if b['type'] == 'Culturel' and b['rayonnement'] > 0]

        for b in state:
            if b['type'] != 'Producteur':
                continue
            r0b, c0b = b['row'], b['col']
            r1b, c1b_ = r0b + b['h'] - 1, c0b + b['w'] - 1

            # Zone old c1
            in_old1 = (r0b <= or1+oh1-1+ray1 and r1b >= or1-ray1 and
                       c0b <= oc1+ow1-1+ray1 and c1b_ >= oc1-ray1)
            # Zone new c1
            in_new1 = (r0b <= nr1+nh1-1+ray1 and r1b >= nr1-ray1 and
                       c0b <= nc1+nw1-1+ray1 and c1b_ >= nc1-ray1)
            # Zone old c2
            in_old2 = (r0b <= or2+oh2-1+ray2 and r1b >= or2-ray2 and
                       c0b <= oc2+ow2-1+ray2 and c1b_ >= oc2-ray2)
            # Zone new c2
            in_new2 = (r0b <= nr2+nh2-1+ray2 and r1b >= nr2-ray2 and
                       c0b <= nc2+nw2-1+ray2 and c1b_ >= nc2-ray2)

            if in_old1 or in_new1 or in_old2 or in_new2:
                _apply_boost_bat(b, _culture_pour_prod(b, cults))


    # Précalculer les poids pour score_fast
    _base_weights_fast = {}
    for _i, _p in enumerate(prio_order):
        _base_weights_fast[_p] = 10 ** (len(prio_order) - _i)
    _MAX_PRIO_FAST = MAX_PRIO

    def score_fast(state):
        """Score rapide : lit directement les champs déjà calculés, sans dict.get()."""
        s = 0.0
        for b in state:
            if b['type'] != 'Producteur' or b['production'] == 'Rien':
                continue
            prio = b['priorite'] or 0
            w_prio = (_MAX_PRIO_FAST - prio + 1) if 1 <= prio <= _MAX_PRIO_FAST else 1
            ratio = b['boost_factor'] - 1.0
            c1 = w_prio * (ratio + 0.15 * (1.0 - ratio))
            c2 = _base_weights_fast.get(b['production'], 1) * b['prod_boosted']
            s += c1 * 1e6 + c2
        return s

    def score_incremental(state):
        return score_fast(state)

    def move_bat(b, r, c, h, w, ori):
        """Applique un déplacement in-place, retourne l'ancien état."""
        old = (b['row'], b['col'], b['h'], b['w'], b['orientation'])
        b['row'], b['col'], b['h'], b['w'], b['orientation'] = r, c, h, w, ori
        return old

    def restore_bat(b, old):
        b['row'], b['col'], b['h'], b['w'], b['orientation'] = old

    def prio_sort_key(bat):
        if bat['type'] == 'Producteur' and bat.get('priorite', 0) > 0:
            return (0, bat['priorite'])
        elif bat['type'] == 'Producteur':
            return (1, 99)
        else:
            return (2, 0)

    # ── Boucle principale ──────────────────────────────────────────────────
    for pass_num in range(1, n_passes + 1):
        if progress_cb:
            progress_cb(pass_num, n_passes, best_score)

        n = len(best)
        order = sorted(range(n), key=lambda i: prio_sort_key(best[i]))
        # Indices des producteurs prioritaires (pour guider les culturels en Phase A)
        prod_prio_a = [i for i in range(n)
                       if best[i]['type']=='Producteur' and best[i].get('priorite',0)>0]

        # ── Phase A : déplacement individuel exhaustif (in-place) ──────────
        # Barycentre des producteurs prioritaires pour guider les culturels
        prio_prod_centers = [(best[i]['row']+best[i]['h']/2.0, best[i]['col']+best[i]['w']/2.0)
                             for i in prod_prio_a]

        for i in order:
            bat = best[i]
            occ = build_occ(best, exclude_ids={i})

            orientations = ['H','V'] if bat['orientation'] == 'H' else ['V','H']
            best_s_a = best_score
            best_move = None

            for orientation in orientations:
                h = bat['largeur'] if orientation == 'H' else bat['longueur']
                w = bat['longueur'] if orientation == 'H' else bat['largeur']

                # Trier et limiter les positions candidates pour la performance
                bat_cr = bat['row'] + h / 2.0
                bat_cc = bat['col'] + w / 2.0
                candidates_a = []
                for r, c in free_positions(h, w, occ):
                    if r == bat['row'] and c == bat['col'] and h == bat['h'] and w == bat['w']:
                        continue
                    if any((r+dr, c+dc) in occ for dr in range(h) for dc in range(w)):
                        continue
                    cr2, cc2 = r + h/2.0, c + w/2.0
                    if bat['type'] == 'Culturel' and prio_prod_centers:
                        # Culturels : trier par distance min aux producteurs prioritaires
                        d = min(abs(cr2-pr)+abs(cc2-pc) for pr,pc in prio_prod_centers)
                    else:
                        # Producteurs et autres : trier par distance à leur position actuelle
                        # (explorer d'abord les positions proches → convergence rapide)
                        d = abs(cr2 - bat_cr) + abs(cc2 - bat_cc)
                    candidates_a.append((d, r, c))

                candidates_a.sort()
                # Limites : 150 pour culturels, 200 pour producteurs
                limit = 60 if bat['type'] == 'Culturel' else 40
                candidates_a = candidates_a[:limit]

                for _, r, c in candidates_a:
                    old = move_bat(bat, r, c, h, w, orientation)
                    if bat['type'] == 'Culturel':
                        recalc_producteurs(best, moved_cult=i, old_pos=old[:4])
                    else:
                        recalc_producteurs(best)
                    s = score_incremental(best)
                    if s > best_s_a + 1e-6:
                        best_s_a = s
                        best_move = (r, c, h, w, orientation)
                    restore_bat(bat, old)

            if best_move:
                r, c, h, w, ori = best_move
                move_bat(bat, r, c, h, w, ori)
                recalc_producteurs(best)
                best_score = best_s_a

        # ── Phase B : swaps de paires (limité aux paires proches) ────────────
        # Trier les paires par distance Manhattan entre les deux bâtiments
        # Ne tester que les paires à moins de MAX_SWAP_DIST cases l'une de l'autre
        MAX_SWAP_DIST = 12
        pairs_b = []
        for i in range(n):
            for j in range(i+1, n):
                bi, bj = best[i], best[j]
                dist = abs(bi['row']-bj['row']) + abs(bi['col']-bj['col'])
                if dist <= MAX_SWAP_DIST:
                    pairs_b.append((dist, i, j))
        pairs_b.sort()

        for _, i, j in pairs_b:
            bi, bj = best[i], best[j]
            ri, ci_col = bi['row'], bi['col']
            rj, cj_col = bj['row'], bj['col']
            occ_sans = build_occ(best, exclude_ids={i, j})

            for ori_i in ['H', 'V']:
                hi = bi['largeur'] if ori_i == 'H' else bi['longueur']
                wi = bi['longueur'] if ori_i == 'H' else bi['largeur']
                for ori_j in ['H', 'V']:
                    hj = bj['largeur'] if ori_j == 'H' else bj['longueur']
                    wj = bj['longueur'] if ori_j == 'H' else bj['largeur']

                    def fits(h_, w_, r_, c_):
                        if r_+h_ > rows or c_+w_ > cols: return False
                        return all(inside_mask[r_+dr, c_+dc] and (r_+dr, c_+dc) not in occ_sans
                                   for dr in range(h_) for dc in range(w_))

                    if not fits(hi, wi, rj, cj_col) or not fits(hj, wj, ri, ci_col):
                        continue
                    cells_i = {(rj+dr, cj_col+dc) for dr in range(hi) for dc in range(wi)}
                    cells_j = {(ri+dr, ci_col+dc) for dr in range(hj) for dc in range(wj)}
                    if cells_i & cells_j:
                        continue

                    old_i = move_bat(bi, rj, cj_col, hi, wi, ori_i)
                    old_j = move_bat(bj, ri, ci_col, hj, wj, ori_j)
                    recalc_producteurs(best)
                    s = score_incremental(best)
                    if s > best_score + 1e-6:
                        best_score = s
                    else:
                        restore_bat(bi, old_i)
                        restore_bat(bj, old_j)
                        recalc_producteurs(best)

        # ── Phase C : déplacements groupés de 2 culturels (amélioration 1) ──
        prod_prio = sorted(
            [i for i in range(n) if best[i]['type']=='Producteur' and best[i].get('priorite',0)>0],
            key=lambda i: best[i].get('priorite', 99)
        )
        cult_indices = [i for i in range(n) if best[i]['type']=='Culturel']

        for pi in prod_prio:
            prod = best[pi]
            cr = prod.get('culture_recue', 0)
            next_seuil = next(
                (s for s in [prod.get('boost25',0), prod.get('boost50',0), prod.get('boost100',0)]
                 if s and cr < s), None)
            if next_seuil is None:
                continue

            r0p, c0p = prod['row'], prod['col']
            r1p, c1p = r0p + prod['h'] - 1, c0p + prod['w'] - 1

            def touche(r, c, h, w, ray):
                return r0p <= r+h-1+ray and r1p >= r-ray and c0p <= c+w-1+ray and c1p >= c-ray

            # Pré-calculer positions touchantes par culturel (limitées à 8, triées par distance)
            prod_cr_c = r0p + prod['h'] / 2.0
            prod_cc_c = c0p + prod['w'] / 2.0
            cult_pos = {}
            for ci in cult_indices:
                cult = best[ci]
                ray = cult['rayonnement']
                occ_sans = build_occ(best, exclude_ids={ci})
                pos = []
                for orientation in ['H','V']:
                    h = cult['largeur'] if orientation=='H' else cult['longueur']
                    w = cult['longueur'] if orientation=='H' else cult['largeur']
                    for r, c in free_positions(h, w, occ_sans):
                        if any((r+dr2, c+dc2) in occ_sans for dr2 in range(h) for dc2 in range(w)):
                            continue
                        if touche(r, c, h, w, ray):
                            d = abs(r+h/2.0-prod_cr_c) + abs(c+w/2.0-prod_cc_c)
                            pos.append((d, r, c, h, w, orientation))
                pos.sort()
                cult_pos[ci] = [(r,c,h,w,ori) for _,r,c,h,w,ori in pos[:8]]

            # Tester paires de culturels (in-place)
            for ci1, ci2 in combinations(cult_indices, 2):
                c1_bat, c2_bat = best[ci1], best[ci2]
                # occ de base : sans les deux culturels
                occ_base = build_occ(best, exclude_ids={ci1, ci2})

                for r1, c1_, h1, w1, ori1 in cult_pos.get(ci1, []):
                    cells1 = {(r1+dr, c1_+dc) for dr in range(h1) for dc in range(w1)}
                    # Vérifier que la position de c1 est libre (dans occ_base)
                    if any(p in occ_base for p in cells1):
                        continue

                    # Déplacer c1 temporairement
                    old1 = move_bat(c1_bat, r1, c1_, h1, w1, ori1)
                    # occ après déplacement de c1 (pour valider c2) : occ_base + cells1
                    occ_avec_c1 = occ_base | cells1

                    accepted_c2 = False
                    for r2, c2_, h2, w2, ori2 in cult_pos.get(ci2, []):
                        cells2 = {(r2+dr, c2_+dc) for dr in range(h2) for dc in range(w2)}
                        # c2 ne doit pas chevaucher occ_base ni cells1
                        if any(p in occ_avec_c1 for p in cells2):
                            continue

                        old2 = move_bat(c2_bat, r2, c2_, h2, w2, ori2)
                        # Recalc partiel : seulement les producteurs touchés par c1 ou c2
                        recalc_two_cults(best, ci1, old1[:4], ci2, old2[:4])
                        s = score_incremental(best)
                        new_boost = best[pi].get('boost', 0)
                        old_boost = prod.get('boost', 0)
                        if s > best_score + 1e-6 or (new_boost > old_boost and s >= best_score * 0.90):
                            # Accepter les deux déplacements : ne pas restaurer
                            best_score = max(best_score, s)
                            accepted_c2 = True
                            break   # on garde c1 ET c2 déplacés, on passe à la paire suivante
                        else:
                            restore_bat(c2_bat, old2)
                            # c2 restauré → recalc partiel sur c2 seulement
                            recalc_producteurs(best, moved_cult=ci2, old_pos=old2[:4])

                    if accepted_c2:
                        # c1 ET c2 sont déjà en place, on sort de la boucle c1 aussi
                        break
                    else:
                        # Aucun c2 valide trouvé : restaurer c1
                        restore_bat(c1_bat, old1)
                        recalc_producteurs(best, moved_cult=ci1, old_pos=old1[:4])

        # ── Phase D : regroupement individuel de culturels ──────────────────
        for pi in prod_prio:
            prod = best[pi]
            cr = prod.get('culture_recue', 0)
            next_seuil = next(
                (s for s in [prod.get('boost25',0), prod.get('boost50',0), prod.get('boost100',0)]
                 if s and cr < s), None)
            if next_seuil is None:
                continue
            r0p, c0p = prod['row'], prod['col']
            r1p, c1p = r0p + prod['h'] - 1, c0p + prod['w'] - 1

            for ci in cult_indices:
                cult = best[ci]
                ray = cult['rayonnement']
                occ_sans = build_occ(best, exclude_ids={ci})
                best_new_cr = prod.get('culture_recue', 0)
                best_move_d = None

                for orientation in ['H','V']:
                    h = cult['largeur'] if orientation=='H' else cult['longueur']
                    w = cult['longueur'] if orientation=='H' else cult['largeur']
                    for r, c in free_positions(h, w, occ_sans):
                        if r == cult['row'] and c == cult['col'] and h == cult['h'] and w == cult['w']:
                            continue
                        # S'assurer que la position est vraiment libre (free_positions inclut aussi les occupées)
                        if any((r+dr, c+dc) in occ_sans for dr in range(h) for dc in range(w)):
                            continue
                        if not (r0p <= r+h-1+ray and r1p >= r-ray and c0p <= c+w-1+ray and c1p >= c-ray):
                            continue
                        old = move_bat(cult, r, c, h, w, orientation)
                        recalc_producteurs(best, moved_cult=ci, old_pos=old[:4])
                        new_cr = best[pi].get('culture_recue', 0)
                        s = score_incremental(best)
                        if new_cr > best_new_cr and s >= best_score * 0.90:
                            best_new_cr = new_cr
                            best_move_d = (r, c, h, w, orientation, s)
                        restore_bat(cult, old)
                        recalc_producteurs(best, moved_cult=ci, old_pos=(r,c,h,w))

                if best_move_d:
                    r, c, h, w, ori, s = best_move_d
                    move_bat(cult, r, c, h, w, ori)
                    recalc_producteurs(best)
                    if s > best_score:
                        best_score = s
                    prod = best[pi]

        # ── Phase E : relocalisation des culturels éloignés ────────────────
        # Pour chaque culturel dont la distance au barycentre des producteurs
        # prioritaires non saturés est > seuil, chercher une position plus proche.
        # Cela évite que des culturels restent "coincés" loin des producteurs utiles.

        # Calculer le barycentre des producteurs prioritaires qui n'ont pas atteint 100%
        targets = [best[i] for i in prod_prio
                   if best[i].get('boost', 0) < 100 and best[i].get('boost100', 0)]
        if targets:
            # Barycentre pondéré par priorité inverse (priorité 1 = poids fort)
            total_w = 0.0
            bar_r, bar_c = 0.0, 0.0
            for t in targets:
                w = MAX_PRIO - t.get('priorite', MAX_PRIO) + 1
                cr_t = t['row'] + t['h'] / 2.0
                cc_t = t['col'] + t['w'] / 2.0
                bar_r += w * cr_t
                bar_c += w * cc_t
                total_w += w
            bar_r /= total_w
            bar_c /= total_w

            # Pour chaque culturel, calculer sa distance au barycentre
            for ci in cult_indices:
                cult = best[ci]
                cr_cult = cult['row'] + cult['h'] / 2.0
                cc_cult = cult['col'] + cult['w'] / 2.0
                dist = abs(cr_cult - bar_r) + abs(cc_cult - bar_c)  # distance Manhattan

                # Seuil : si le culturel est à plus de 10 cases du barycentre
                if dist <= 10:
                    continue

                # Chercher la position libre la plus proche du barycentre
                # qui touche au moins un producteur prioritaire non saturé
                ray = cult['rayonnement']
                occ_sans = build_occ(best, exclude_ids={ci})
                best_dist = dist
                best_move_e = None

                # Collecter candidats triés par distance au barycentre, limités à 20
                candidates_e = []
                for orientation in ['H', 'V']:
                    h = cult['largeur'] if orientation == 'H' else cult['longueur']
                    w = cult['longueur'] if orientation == 'H' else cult['largeur']
                    for r, c in free_positions(h, w, occ_sans):
                        if any((r+dr, c+dc) in occ_sans for dr in range(h) for dc in range(w)):
                            continue
                        touche_cible = any(
                            t['row'] <= r+h-1+ray and t['row']+t['h']-1 >= r-ray and
                            t['col'] <= c+w-1+ray and t['col']+t['w']-1 >= c-ray
                            for t in targets
                        )
                        if not touche_cible:
                            continue
                        new_dist = abs(r + h/2.0 - bar_r) + abs(c + w/2.0 - bar_c)
                        if new_dist < best_dist:
                            candidates_e.append((new_dist, r, c, h, w, orientation))
                candidates_e.sort()
                for new_dist, r, c, h, w, orientation in candidates_e[:8]:
                    old = move_bat(cult, r, c, h, w, orientation)
                    recalc_producteurs(best, moved_cult=ci, old_pos=old[:4])
                    s = score_incremental(best)
                    if s >= best_score * 0.95:
                        best_dist = new_dist
                        best_move_e = (r, c, h, w, orientation, s)
                    restore_bat(cult, old)
                    recalc_producteurs(best, moved_cult=ci, old_pos=(r,c,h,w))
                if best_move_e:
                    r, c, h, w, ori, s = best_move_e
                    move_bat(cult, r, c, h, w, ori)
                    recalc_producteurs(best)
                    if s > best_score:
                        best_score = s

    recalc_producteurs(best)
    return best, score_total(best, prio_order)


# ─────────────────────────────────────────────
# GÉNÉRATION DU FICHIER OUTPUT
# ─────────────────────────────────────────────

ORANGE = "FFFFA500"
VERT   = "FF90EE90"
GRIS   = "FFD3D3D3"
BLEU_CLAIR = "FFCFE2F3"
JAUNE  = "FFFFFF99"
BLANC  = "FFFFFFFF"
ROUGE_CLAIR = "FFFFCCCC"

def fill(hex6):
    return PatternFill("solid", fgColor=hex6)

def border_thin():
    s = Side(style='thin')
    return Border(left=s, right=s, top=s, bottom=s)

def ecrire_output(placed_init, placed_opt, inside_mask, grid_shape, bat_info, prio_order):
    wb = openpyxl.Workbook()
    
    # ── Onglet 1 : Liste des bâtiments placés ──
    ws1 = wb.active
    ws1.title = "Batiments places"

    headers = [
        ("Nom",            22),
        ("Type",           13),
        ("Production",     14),
        ("Priorité",        9),
        ("Ligne",           7),
        ("Colonne",         8),
        ("Hauteur",         8),
        ("Largeur",         8),
        ("Rayonnement",    13),
        ("Culture donnée", 15),
        ("Boost 25%",      11),
        ("Boost 50%",      11),
        ("Boost 100%",     11),
        ("Culture reçue",  14),
        ("Boost atteint %",15),
        ("Qté/h base",     12),
        ("Qté/h boostée",  14),
    ]
    for c, (h, w) in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = fill(BLEU_CLAIR)
        cell.border = border_thin()
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws1.column_dimensions[get_column_letter(c)].width = w
    ws1.row_dimensions[1].height = 30

    for i, bat in enumerate(placed_opt, 2):
        # Colonnes dépendantes du type
        if bat["type"] == "Culturel":
            culture_donnee = bat.get("culture", 0)
            rayonnement    = bat.get("rayonnement", 0)
            boost25 = boost50 = boost100 = ""
            culture_recue  = ""
            boost_atteint  = ""
            qte_base       = ""
            qte_boostee    = ""
        else:
            culture_donnee = ""
            rayonnement    = ""
            boost25        = bat.get("boost25", 0)
            boost50        = bat.get("boost50", 0)
            boost100       = bat.get("boost100", 0)
            culture_recue  = bat.get("culture_recue", 0)
            boost_atteint  = bat.get("boost", 0)
            qte_base       = bat["quantite"]
            qte_boostee    = bat.get("prod_boosted", bat["quantite"])

        vals = [
            bat["nom"], bat["type"], bat["production"],
            bat.get("priorite", 0) or "",
            bat["row"]+1, bat["col"]+1, bat["h"], bat["w"],
            rayonnement, culture_donnee,
            boost25, boost50, boost100,
            culture_recue, boost_atteint,
            qte_base, qte_boostee,
        ]
        bg = ORANGE if bat["type"] == "Culturel" else (VERT if bat["type"] == "Producteur" else GRIS)
        for c, v in enumerate(vals, 1):
            cell = ws1.cell(row=i, column=c, value=v)
            cell.fill = fill(bg)
            cell.border = border_thin()
            cell.alignment = Alignment(horizontal="center", vertical="center")
        ws1.cell(row=i, column=1).alignment = Alignment(horizontal="left", vertical="center")
    
    # ── Onglet 2 : Synthèse par type de production ──
    ws2 = wb.create_sheet("Synthese")
    ws2.cell(1,1,"Type de production").font = Font(bold=True)
    ws2.cell(1,2,"Culture totale recue").font = Font(bold=True)
    ws2.cell(1,3,"Boost atteint (%)").font = Font(bold=True)
    ws2.cell(1,4,"Qte/h total").font = Font(bold=True)
    ws2.cell(1,5,"Qte/h avant optim").font = Font(bold=True)
    ws2.cell(1,6,"Gain/Perte Qte/h").font = Font(bold=True)
    
    for c in range(1,7):
        ws2.cell(1,c).fill = fill(BLEU_CLAIR)
        ws2.cell(1,c).border = border_thin()
    
    # Grouper par production
    from collections import defaultdict
    prod_groups_opt = defaultdict(list)
    prod_groups_init = defaultdict(list)
    
    for bat in placed_opt:
        if bat['type'] == 'Producteur' and bat['production'] != 'Rien':
            prod_groups_opt[bat['production']].append(bat)
    
    calculer_culture(placed_init, grid_shape)
    for bat in placed_init:
        if bat['type'] == 'Producteur' and bat['production'] != 'Rien':
            prod_groups_init[bat['production']].append(bat)
    
    row = 2
    all_prods = list(set(list(prod_groups_opt.keys()) + list(prod_groups_init.keys())))
    all_prods.sort(key=lambda x: prio_order.index(x) if x in prio_order else 99)
    
    for prod in all_prods:
        bats_opt = prod_groups_opt.get(prod, [])
        bats_init = prod_groups_init.get(prod, [])
        
        cult_tot = sum(b.get('culture_recue',0) for b in bats_opt)
        boosts = [b.get('boost',0) for b in bats_opt]
        boost_moyen = max(boosts) if boosts else 0
        qte_opt = sum(b.get('prod_boosted', b['quantite']) for b in bats_opt)
        qte_init = sum(b.get('prod_boosted', b['quantite']) for b in bats_init)
        gain = qte_opt - qte_init
        
        vals = [prod, cult_tot, boost_moyen, qte_opt, qte_init, gain]
        for c, v in enumerate(vals, 1):
            cell = ws2.cell(row=row, column=c, value=round(v,1) if isinstance(v, float) else v)
            cell.border = border_thin()
            if c == 6:
                cell.fill = fill(VERT if gain >= 0 else ROUGE_CLAIR)
        row += 1
    
    synthese_widths = [20, 22, 18, 14, 18, 16]
    for c, w in enumerate(synthese_widths, 1):
        ws2.column_dimensions[get_column_letter(c)].width = w
    ws2.row_dimensions[1].height = 30
    for c in range(1, 7):
        ws2.cell(1, c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # ── Onglet 3 : Bâtiments déplacés ──
    ws3 = wb.create_sheet("Deplacements")
    hdrs = ["Nom", "Ligne avant", "Col avant", "Ligne apres", "Col apres", "Note"]
    for c, h in enumerate(hdrs, 1):
        cell = ws3.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = fill(BLEU_CLAIR)
        cell.border = border_thin()
    
    # Correspondance init → opt par nom
    init_by_nom = {}
    for bat in placed_init:
        n = bat['nom']
        init_by_nom.setdefault(n, []).append((bat['row'], bat['col']))
    opt_by_nom = {}
    for bat in placed_opt:
        n = bat['nom']
        opt_by_nom.setdefault(n, []).append((bat['row'], bat['col']))
    
    row = 2
    moved = []
    for nom in set(list(init_by_nom.keys()) + list(opt_by_nom.keys())):
        inits = sorted(init_by_nom.get(nom, []))
        opts = sorted(opt_by_nom.get(nom, []))
        for idx, (ri, ci) in enumerate(inits):
            ro, co = opts[idx] if idx < len(opts) else (ri, ci)
            if ri != ro or ci != co:
                moved.append((nom, ri+1, ci+1, ro+1, co+1))
    
    for nom, ri, ci, ro, co in moved:
        ws3.cell(row=row, column=1, value=nom).border = border_thin()
        ws3.cell(row=row, column=2, value=ri).border = border_thin()
        ws3.cell(row=row, column=3, value=ci).border = border_thin()
        ws3.cell(row=row, column=4, value=ro).border = border_thin()
        ws3.cell(row=row, column=5, value=co).border = border_thin()
        ws3.cell(row=row, column=6, value="Déplacé").border = border_thin()
        row += 1
    
    if row == 2:
        ws3.cell(row=2, column=1, value="Aucun déplacement effectué")
    
    for c in range(1,7):
        ws3.column_dimensions[get_column_letter(c)].width = 18
    
    # ── Onglet 4 : Séquence d'opérations ──
    ws4 = wb.create_sheet("Sequence operations")
    ws4.cell(1,1,"Etape").font = Font(bold=True)
    ws4.cell(1,2,"Operation").font = Font(bold=True)
    ws4.cell(1,3,"Batiment").font = Font(bold=True)
    ws4.cell(1,4,"Detail").font = Font(bold=True)
    for c in range(1,5):
        ws4.cell(1,c).fill = fill(BLEU_CLAIR)
        ws4.cell(1,c).border = border_thin()
    
    etape = 2
    for nom, ri, ci, ro, co in moved:
        ws4.cell(etape, 1, etape-1).border = border_thin()
        ws4.cell(etape, 2, "Déplacer").border = border_thin()
        ws4.cell(etape, 3, nom).border = border_thin()
        ws4.cell(etape, 4, f"De (ligne {ri}, col {ci}) vers (ligne {ro}, col {co})").border = border_thin()
        etape += 1
    
    if etape == 2:
        ws4.cell(2,1,"-").border = border_thin()
        ws4.cell(2,2,"Aucun déplacement nécessaire").border = border_thin()
    
    for c in range(1,5):
        ws4.column_dimensions[get_column_letter(c)].width = 40
    
    # ── Onglet 5 : Carte du terrain ──
    ws5 = wb.create_sheet("Terrain optimise")
    rows, cols = grid_shape

    FILL_X     = PatternFill("solid", fgColor="FF404040")
    FILL_OUT   = PatternFill("solid", fgColor="FFFFFFFF")
    FILL_CULT  = PatternFill("solid", fgColor="FFFFA500")
    FILL_PROD  = PatternFill("solid", fgColor="FF90EE90")
    FILL_NEUT  = PatternFill("solid", fgColor="FFD3D3D3")
    FILL_EMPTY = PatternFill("solid", fgColor="FFF5F5F5")

    CELL_W = 7
    CELL_H = 18

    # Dimensions fixes pour toutes les colonnes/lignes
    for r in range(rows):
        ws5.row_dimensions[r + 1].height = CELL_H
    for c in range(cols):
        ws5.column_dimensions[get_column_letter(c + 1)].width = CELL_W

    # Passe 1 : fond de carte (intérieur vide + extérieur blanc)
    for r in range(rows):
        for c in range(cols):
            cell = ws5.cell(row=r + 1, column=c + 1)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border_thin()
            cell.font = Font(size=7)
            if inside_mask[r, c]:
                cell.fill = FILL_EMPTY
            else:
                cell.fill = FILL_OUT

    # Passe 2 : cases X — reproduire fidèlement le terrain input
    for r in range(rows):
        for c in range(cols):
            if not inside_mask[r, c]:
                # Détecter si c'est un X : case non-intérieure adjacente à l'intérieur
                # ou sur la bordure de la grille non-intérieure avec voisin intérieur
                is_x = False
                for dr, dc in [(-1,0),(1,0),(0,-1),(0,1)]:
                    nr, nc = r+dr, c+dc
                    if 0 <= nr < rows and 0 <= nc < cols and inside_mask[nr, nc]:
                        is_x = True
                        break
                # Aussi : si la case fait partie de la première/dernière ligne
                # ou première/dernière colonne non-nulle du terrain
                if is_x:
                    cell = ws5.cell(row=r + 1, column=c + 1)
                    cell.fill = FILL_X
                    cell.font = Font(size=7, color="FFFFFFFF", bold=True)
                    cell.value = "X"
                    cell.alignment = Alignment(horizontal="center", vertical="center")

    # Passe 3 : bâtiments avec cellules fusionnées
    thick = Side(style="medium")
    thin  = Side(style="thin")

    # Éliminer les doublons de position (deux bâtiments sur la même case après opt)
    seen_positions = set()
    bats_uniques = []
    for bat in placed_opt:
        key = (bat["row"], bat["col"], bat["h"], bat["w"])
        if key not in seen_positions:
            seen_positions.add(key)
            bats_uniques.append(bat)

    for bat in bats_uniques:
        r0, c0 = bat["row"], bat["col"]
        r1, c1 = r0 + bat["h"] - 1, c0 + bat["w"] - 1

        # Vérifier que le bâtiment est dans le terrain
        if r1 >= rows or c1 >= cols:
            continue

        typ = bat["type"]
        if typ == "Culturel":
            f = FILL_CULT
        elif typ == "Producteur":
            f = FILL_PROD
        else:
            f = FILL_NEUT

        boost = bat.get("boost", 0)
        label = bat["nom"]
        if boost > 0:
            label += f"\n+{boost}%"

        # Défusionner la zone si elle était déjà fusionnée (sécurité)
        try:
            ws5.unmerge_cells(
                start_row=r0+1, start_column=c0+1,
                end_row=r1+1,   end_column=c1+1
            )
        except Exception:
            pass

        # Colorier toutes les cases et appliquer les bordures
        for dr in range(bat["h"]):
            for dc in range(bat["w"]):
                r, c = r0 + dr, c0 + dc
                cell = ws5.cell(row=r + 1, column=c + 1)
                cell.fill = f
                top    = thick if r == r0 else thin
                bottom = thick if r == r1 else thin
                left   = thick if c == c0 else thin
                right  = thick if c == c1 else thin
                cell.border = Border(top=top, bottom=bottom, left=left, right=right)

        # Fusionner
        if bat["h"] > 1 or bat["w"] > 1:
            try:
                ws5.merge_cells(
                    start_row=r0 + 1, start_column=c0 + 1,
                    end_row=r1 + 1,   end_column=c1 + 1
                )
            except Exception:
                pass

        # Libellé sur la cellule top-left (après fusion, openpyxl la renvoie correctement)
        tl = ws5.cell(row=r0 + 1, column=c0 + 1)
        tl.value = label
        tl.font = Font(size=7, bold=True)
        tl.fill = f
        tl.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ─────────────────────────────────────────────
# INTERFACE STREAMLIT
# ─────────────────────────────────────────────

# Session state : conserver les résultats entre re-renders
for key in ['placed_opt', 'score_opt', 'score_init', 'output_bytes',
            'rows_data', 'moved_count', 'placed_init_snap',
            'inside_mask_snap', 'grid_shape_snap', 'bat_info_snap', 'prio_snap']:
    if key not in st.session_state:
        st.session_state[key] = None

uploaded_file = st.file_uploader(
    "📂 Choisir le fichier Excel de la ville",
    type=["xlsx"],
    help="Fichier avec onglets Terrain et Batiments"
)

if uploaded_file:
    try:
        terrain_df, bat_df = lire_fichier(uploaded_file)
        grid, inside_mask, border_mask = construire_terrain(terrain_df)
        placed_init, bat_info = extraire_batiments(grid, inside_mask, bat_df)

        n_inside = inside_mask.sum()
        n_placed = len(placed_init)

        col1, col2, col3 = st.columns(3)
        col1.metric("Cases intérieures", int(n_inside))
        col2.metric("Bâtiments détectés", n_placed)

        calculer_culture(placed_init, grid.shape)
        score_init = score_total(placed_init)
        col3.metric("Score initial", f"{score_init:,.0f}")

        st.markdown("---")
        st.subheader("⚙️ Options d'optimisation")

        prio_options = ['Guérison', 'Nourriture', 'Or', 'Autre']
        prio_order = st.multiselect(
            "Ordre de priorité des productions (du plus au moins important)",
            options=prio_options,
            default=['Guérison', 'Nourriture', 'Or']
        )
        n_passes = st.slider("Nombre de passes d'optimisation", 1, 5, 2)

        if st.button("🚀 Lancer l'optimisation", type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()

            def progress_cb(p, total, score):
                progress_bar.progress(p / total)
                status_text.text(f"Passe {p}/{total} — Score : {score:,.0f}")

            prio = prio_order if prio_order else ['Guérison', 'Nourriture', 'Or']

            with st.spinner("Optimisation en cours..."):
                placed_opt, score_opt = optimiser(
                    copy.deepcopy(placed_init), inside_mask, grid.shape,
                    prio, n_passes=n_passes, progress_cb=progress_cb
                )

            progress_bar.progress(1.0)
            calculer_culture(placed_opt, grid.shape)
            gain = score_opt - score_init
            status_text.text(f"✅ Terminée — Score : {score_opt:,.0f} (gain : +{gain:,.0f})")

            # Résumé par production
            from collections import defaultdict
            prod_summary = defaultdict(lambda: {'qte_init': 0, 'qte_opt': 0, 'boost_max': 0})
            for bat in placed_opt:
                if bat['type'] == 'Producteur' and bat['production'] != 'Rien':
                    p = bat['production']
                    prod_summary[p]['qte_opt'] += bat.get('prod_boosted', bat['quantite'])
                    prod_summary[p]['boost_max'] = max(prod_summary[p]['boost_max'], bat.get('boost', 0))
            calculer_culture(placed_init, grid.shape)
            for bat in placed_init:
                if bat['type'] == 'Producteur' and bat['production'] != 'Rien':
                    prod_summary[bat['production']]['qte_init'] += bat.get('prod_boosted', bat['quantite'])

            rows_data = []
            for p, d in sorted(prod_summary.items()):
                gain_p = d['qte_opt'] - d['qte_init']
                rows_data.append({
                    "Production": p,
                    "Qté/h avant": f"{d['qte_init']:,.0f}",
                    "Qté/h après": f"{d['qte_opt']:,.0f}",
                    "Boost max": f"{d['boost_max']}%",
                    "Gain": f"+{gain_p:,.0f}" if gain_p >= 0 else f"{gain_p:,.0f}"
                })

            # Compter les déplacements : apparier init→opt par nom (ordre d'apparition)
            from collections import defaultdict
            init_positions = defaultdict(list)
            for bi in placed_init:
                init_positions[bi['nom']].append((bi['row'], bi['col']))
            opt_positions = defaultdict(list)
            for bo in placed_opt:
                opt_positions[bo['nom']].append((bo['row'], bo['col']))
            moved_count = 0
            for nom in init_positions:
                inits = sorted(init_positions[nom])
                opts  = sorted(opt_positions.get(nom, []))
                for idx, pos_i in enumerate(inits):
                    pos_o = opts[idx] if idx < len(opts) else pos_i
                    if pos_i != pos_o:
                        moved_count += 1

            # Générer le fichier Excel
            calculer_culture(placed_init, grid.shape)
            output_bytes = ecrire_output(
                copy.deepcopy(placed_init), placed_opt,
                inside_mask, grid.shape, bat_info, prio
            )

            # Stocker tout dans session_state → persiste après re-render
            st.session_state.placed_opt        = placed_opt
            st.session_state.score_opt         = score_opt
            st.session_state.score_init        = score_init
            st.session_state.output_bytes      = output_bytes.getvalue()
            st.session_state.rows_data         = rows_data
            st.session_state.moved_count       = moved_count
            st.session_state.prio_snap         = prio

        # ── Affichage des résultats (persistant) ──────────────────────────
        if st.session_state.placed_opt is not None:
            st.markdown("---")
            st.subheader("📊 Résultats")

            gain_total = st.session_state.score_opt - st.session_state.score_init
            r1, r2, r3 = st.columns(3)
            r1.metric("Score initial",   f"{st.session_state.score_init:,.0f}")
            r2.metric("Score optimisé",  f"{st.session_state.score_opt:,.0f}")
            r3.metric("Gain",            f"+{gain_total:,.0f}")

            if st.session_state.rows_data:
                st.dataframe(
                    pd.DataFrame(st.session_state.rows_data),
                    hide_index=True, use_container_width=True
                )
            st.info(f"🔄 Bâtiments déplacés : {st.session_state.moved_count}")

            st.markdown("---")
            st.subheader("💾 Télécharger les résultats")
            st.download_button(
                label="📥 Télécharger le fichier résultats Excel",
                data=st.session_state.output_bytes,
                file_name="resultats_optimisation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_btn"
            )

    except Exception as e:
        st.error(f"Erreur lors du traitement : {e}")
        import traceback
        st.code(traceback.format_exc())
else:
    st.info("👆 Chargez votre fichier Excel pour commencer.")
    st.markdown("""
    ### Structure attendue du fichier Excel
    **Onglet 1 – Terrain** : grille avec `X` en bordure, noms de bâtiments dans les cases  
    **Onglet 2 – Batiments** : colonnes `Nom, Longueur, Largeur, Nombre, Type, Culture, Rayonnement, Boost 25%, Boost 50%, Boost 100%, Production, Quantite, Priorite`
    """)
