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
    """Score pondéré par gain potentiel réel entre paliers."""
    if prio_order is None:
        prio_order = ['Guérison', 'Nourriture', 'Or']

    # Poids de base par ordre de priorité (exponentiel pour garantir la hiérarchie)
    base_weights = {}
    for i, p in enumerate(prio_order):
        base_weights[p] = 10 ** (len(prio_order) - i)

    score = 0.0
    for bat in placed_list:
        if bat['type'] != 'Producteur' or bat['production'] == 'Rien':
            continue
        w = base_weights.get(bat['production'], 1)

        # Amélioration 4 : pondérer par le gain potentiel jusqu'au prochain palier
        # → favorise les bâtiments proches d'un franchissement de seuil
        b25  = bat.get('boost25',  0) or 0
        b50  = bat.get('boost50',  0) or 0
        b100 = bat.get('boost100', 0) or 0
        cr   = bat.get('culture_recue', 0)
        qte  = bat['quantite']

        # Production actuelle boostée
        prod_actuelle = bat.get('prod_boosted', qte)

        # Gain potentiel max encore atteignable (vers 100% si pas encore atteint)
        if b100 and cr < b100:
            gain_potentiel = qte * 2.0 - prod_actuelle   # ce qu'on gagnerait si 100%
        elif b50 and cr < b50:
            gain_potentiel = qte * 1.5 - prod_actuelle
        elif b25 and cr < b25:
            gain_potentiel = qte * 1.25 - prod_actuelle
        else:
            gain_potentiel = 0

        # Score = production boostée + bonus proportionnel au gain potentiel pondéré par priorité
        score += w * (prod_actuelle + 0.1 * gain_potentiel)

    return score


def optimiser(placed_list, inside_mask, grid_shape, prio_order, n_passes=3, progress_cb=None):
    """
    Optimisation en 4 phases par passe :
    A) Déplacements individuels exhaustifs — ordre : producteurs prioritaires puis culturels
       Amélioration 2 : orientation pré-triée par gain attendu
       Amélioration 3 : cases libres explorées en priorité
    B) Swaps de paires (toutes tailles, toutes orientations)
    C) Déplacements groupés de culturels (2 à la fois) pour franchir un palier de boost
       Amélioration 1 : combinaisons de 2 culturels testées simultanément
    D) Regroupement individuel de culturels autour des producteurs prioritaires
    """
    best = copy.deepcopy(placed_list)
    calculer_culture(best, grid_shape)
    best_score = score_total(best, prio_order)

    rows, cols = grid_shape

    def prio_sort_key(bat):
        if bat['type'] == 'Producteur' and bat.get('priorite', 0) > 0:
            return (0, bat['priorite'])
        elif bat['type'] == 'Producteur':
            return (1, 99)
        else:
            return (2, 0)

    def occupation_mask(state, exclude_ids=()):
        occ = np.zeros((rows, cols), dtype=bool)
        for k, b in enumerate(state):
            if k in exclude_ids:
                continue
            for dr in range(b['h']):
                for dc in range(b['w']):
                    r2, c2 = b['row']+dr, b['col']+dc
                    if 0 <= r2 < rows and 0 <= c2 < cols:
                        occ[r2, c2] = True
        return occ

    def positions_candidates(bat, occ, orientation):
        """Amélioration 2+3 : liste des positions valides, cases libres en tête."""
        h = bat['largeur'] if orientation == 'H' else bat['longueur']
        w = bat['longueur'] if orientation == 'H' else bat['largeur']
        libres, occupees = [], []
        for r in range(rows - h + 1):
            for c in range(cols - w + 1):
                ok = True
                any_occupied = False
                for dr in range(h):
                    for dc in range(w):
                        r2, c2 = r+dr, c+dc
                        if not inside_mask[r2, c2]:
                            ok = False; break
                        if occ[r2, c2]:
                            any_occupied = True
                    if not ok: break
                if ok:
                    if any_occupied:
                        occupees.append((r, c))
                    else:
                        libres.append((r, c))
        # Cases libres d'abord (amélioration 3), puis occupées
        return libres + occupees

    def meilleure_position(idx, state):
        """Retourne la meilleure position pour state[idx] parmi toutes les positions valides."""
        bat = state[idx]
        occ = occupation_mask(state, exclude_ids={idx})
        cur_s = score_total(state, prio_order)
        best_s = cur_s
        best_r, best_c, best_h, best_w, best_ori = bat['row'], bat['col'], bat['h'], bat['w'], bat['orientation']
        found = False

        # Amélioration 2 : tester en priorité l'orientation actuelle, puis l'autre
        orientations = ['H', 'V']
        if bat['orientation'] == 'V':
            orientations = ['V', 'H']

        for orientation in orientations:
            h = bat['largeur'] if orientation == 'H' else bat['longueur']
            w = bat['longueur'] if orientation == 'H' else bat['largeur']
            for r, c in positions_candidates(bat, occ, orientation):
                if r == best_r and c == best_c and h == best_h and w == best_w:
                    continue
                # Vérification rapide sans copy
                ok = True
                for dr in range(h):
                    for dc in range(w):
                        r2, c2 = r+dr, c+dc
                        if occ[r2, c2]:
                            ok = False; break
                    if not ok: break
                if not ok:
                    continue
                cand = copy.deepcopy(state)
                cand[idx].update({'row': r, 'col': c, 'h': h, 'w': w, 'orientation': orientation})
                calculer_culture(cand, grid_shape)
                s = score_total(cand, prio_order)
                if s > best_s + 1e-6:
                    best_s = s
                    best_r, best_c, best_h, best_w, best_ori = r, c, h, w, orientation
                    found = True
        return found, best_r, best_c, best_h, best_w, best_ori, best_s

    for pass_num in range(1, n_passes + 1):
        if progress_cb:
            progress_cb(pass_num, n_passes, best_score)

        n = len(best)
        order = sorted(range(n), key=lambda i: prio_sort_key(best[i]))

        # ── Phase A : déplacement individuel exhaustif ──
        for i in order:
            found, r, c, h, w, ori, s = meilleure_position(i, best)
            if found:
                best[i].update({'row': r, 'col': c, 'h': h, 'w': w, 'orientation': ori})
                calculer_culture(best, grid_shape)
                best_score = s

        # ── Phase B : swaps de paires (toutes tailles) ──
        for i in range(n):
            for j in range(i + 1, n):
                bi, bj = best[i], best[j]
                ri, ci_col = bi['row'], bi['col']
                rj, cj_col = bj['row'], bj['col']
                occ_sans_ij = occupation_mask(best, exclude_ids={i, j})

                for ori_i in ['H', 'V']:
                    hi = bi['largeur'] if ori_i == 'H' else bi['longueur']
                    wi = bi['longueur'] if ori_i == 'H' else bi['largeur']
                    for ori_j in ['H', 'V']:
                        hj = bj['largeur'] if ori_j == 'H' else bj['longueur']
                        wj = bj['longueur'] if ori_j == 'H' else bj['largeur']

                        # Vérifier que chaque bâtiment tient dans la position de l'autre
                        def fits(h_, w_, r_, c_):
                            if r_ + h_ > rows or c_ + w_ > cols:
                                return False
                            for dr in range(h_):
                                for dc in range(w_):
                                    r2, c2 = r_+dr, c_+dc
                                    if not inside_mask[r2, c2] or occ_sans_ij[r2, c2]:
                                        return False
                            return True

                        if fits(hi, wi, rj, cj_col) and fits(hj, wj, ri, ci_col):
                            # Vérifier absence de chevauchement entre les deux nouvelles positions
                            overlap = False
                            for dr in range(hi):
                                for dc in range(wi):
                                    r2, c2 = rj+dr, cj_col+dc
                                    for dr2 in range(hj):
                                        for dc2 in range(wj):
                                            if (ri+dr2 == r2 and ci_col+dc2 == c2):
                                                overlap = True
                            if not overlap:
                                cand = copy.deepcopy(best)
                                cand[i].update({'row': rj, 'col': cj_col, 'h': hi, 'w': wi, 'orientation': ori_i})
                                cand[j].update({'row': ri, 'col': ci_col, 'h': hj, 'w': wj, 'orientation': ori_j})
                                calculer_culture(cand, grid_shape)
                                s = score_total(cand, prio_order)
                                if s > best_score + 1e-6:
                                    best = cand
                                    best_score = s

        # ── Phase C : déplacements groupés de 2 culturels simultanément ──
        # (Amélioration 1) Pour chaque producteur prioritaire qui n'a pas atteint son palier max,
        # tester toutes les combinaisons de 2 culturels vers des positions qui touchent ce producteur.
        prod_prio = sorted(
            [i for i in range(n) if best[i]['type'] == 'Producteur' and best[i].get('priorite', 0) > 0],
            key=lambda i: best[i].get('priorite', 99)
        )
        cult_indices = [i for i in range(n) if best[i]['type'] == 'Culturel']

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

            def touche_producteur(r, c, h, w, ray):
                zr0, zr1 = r - ray, r + h - 1 + ray
                zc0, zc1 = c - ray, c + w - 1 + ray
                return r0p <= zr1 and r1p >= zr0 and c0p <= zc1 and c1p >= zc0

            # Pré-calculer les meilleures positions "touchantes" pour chaque culturel
            cult_positions = {}  # ci_idx → liste de (r, c, h, w, ori)
            for ci_idx in cult_indices:
                cult = best[ci_idx]
                ray = cult['rayonnement']
                occ_sans = occupation_mask(best, exclude_ids={ci_idx})
                positions = []
                for orientation in ['H', 'V']:
                    h = cult['largeur'] if orientation == 'H' else cult['longueur']
                    w = cult['longueur'] if orientation == 'H' else cult['largeur']
                    for r, c in positions_candidates(cult, occ_sans, orientation):
                        ok = all(not occ_sans[r+dr, c+dc]
                                 for dr in range(h) for dc in range(w))
                        if ok and touche_producteur(r, c, h, w, ray):
                            positions.append((r, c, h, w, orientation))
                cult_positions[ci_idx] = positions

            # Tester toutes les paires de culturels
            from itertools import combinations
            for ci1, ci2 in combinations(cult_indices, 2):
                for r1, c1_, h1, w1, ori1 in cult_positions.get(ci1, []):
                    for r2, c2_, h2, w2, ori2 in cult_positions.get(ci2, []):
                        # Vérifier que les deux nouvelles positions ne se chevauchent pas
                        occ_base = occupation_mask(best, exclude_ids={ci1, ci2})
                        ok = True
                        tmp = np.zeros((rows, cols), dtype=bool)
                        for dr in range(h1):
                            for dc in range(w1):
                                r_, c_ = r1+dr, c1_+dc
                                if occ_base[r_, c_] or tmp[r_, c_]:
                                    ok = False; break
                                tmp[r_, c_] = True
                            if not ok: break
                        if not ok: continue
                        for dr in range(h2):
                            for dc in range(w2):
                                r_, c_ = r2+dr, c2_+dc
                                if occ_base[r_, c_] or tmp[r_, c_]:
                                    ok = False; break
                                tmp[r_, c_] = True
                            if not ok: break
                        if not ok: continue

                        cand = copy.deepcopy(best)
                        cand[ci1].update({'row': r1, 'col': c1_, 'h': h1, 'w': w1, 'orientation': ori1})
                        cand[ci2].update({'row': r2, 'col': c2_, 'h': h2, 'w': w2, 'orientation': ori2})
                        calculer_culture(cand, grid_shape)
                        new_cr = cand[pi].get('culture_recue', 0)
                        gs = score_total(cand, prio_order)
                        # Accepter si ça améliore le score global OU si ça franchit un nouveau palier
                        # sans trop dégrader le score (≥ 90%)
                        new_boost = cand[pi].get('boost', 0)
                        old_boost = prod.get('boost', 0)
                        if gs > best_score + 1e-6 or (new_boost > old_boost and gs >= best_score * 0.90):
                            best = cand
                            best_score = max(best_score, gs)

        # ── Phase D : regroupement individuel de culturels ──
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

            for ci_idx in cult_indices:
                cult = best[ci_idx]
                ray = cult['rayonnement']
                best_new_cr = prod.get('culture_recue', 0)
                best_pos = None
                occ_sans = occupation_mask(best, exclude_ids={ci_idx})

                for orientation in ['H', 'V']:
                    h = cult['largeur'] if orientation == 'H' else cult['longueur']
                    w = cult['longueur'] if orientation == 'H' else cult['largeur']
                    for r, c in positions_candidates(cult, occ_sans, orientation):
                        if r == cult['row'] and c == cult['col'] and h == cult['h'] and w == cult['w']:
                            continue
                        ok = all(not occ_sans[r+dr, c+dc] for dr in range(h) for dc in range(w))
                        if not ok:
                            continue
                        zr0, zr1 = r-ray, r+h-1+ray
                        zc0, zc1 = c-ray, c+w-1+ray
                        if not (r0p <= zr1 and r1p >= zr0 and c0p <= zc1 and c1p >= zc0):
                            continue
                        cand = copy.deepcopy(best)
                        cand[ci_idx].update({'row': r, 'col': c, 'h': h, 'w': w, 'orientation': orientation})
                        calculer_culture(cand, grid_shape)
                        new_cr = cand[pi].get('culture_recue', 0)
                        gs = score_total(cand, prio_order)
                        if new_cr > best_new_cr and gs >= best_score * 0.90:
                            best_new_cr = new_cr
                            best_pos = (r, c, h, w, orientation, gs, cand)

                if best_pos is not None:
                    r, c, h, w, ori, gs, cand = best_pos
                    best = cand
                    if gs > best_score:
                        best_score = gs
                    prod = best[pi]

    calculer_culture(best, grid_shape)
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
    
    for c in range(1,7):
        ws2.column_dimensions[get_column_letter(c)].width = 22
    
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

    for bat in placed_opt:
        r0, c0 = bat["row"], bat["col"]
        r1, c1 = r0 + bat["h"] - 1, c0 + bat["w"] - 1

        typ = bat["type"]
        if typ == "Culturel":
            f = FILL_CULT
        elif typ == "Producteur":
            f = FILL_PROD
        else:
            f = FILL_NEUT

        # Colorier toutes les cases avant fusion (nécessaire pour openpyxl)
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
            ws5.merge_cells(
                start_row=r0 + 1, start_column=c0 + 1,
                end_row=r1 + 1,   end_column=c1 + 1
            )

        # Libellé sur la cellule top-left (seule active après fusion)
        boost = bat.get("boost", 0)
        label = bat["nom"]
        if boost > 0:
            label += f"\n+{boost}%"

        cell = ws5.cell(row=r0 + 1, column=c0 + 1)
        cell.value = label
        cell.font = Font(size=7, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ─────────────────────────────────────────────
# INTERFACE STREAMLIT
# ─────────────────────────────────────────────

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
        
        # Options d'optimisation
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
            
            with st.spinner("Optimisation en cours..."):
                placed_opt, score_opt = optimiser(
                    copy.deepcopy(placed_init),
                    inside_mask,
                    grid.shape,
                    prio_order if prio_order else ['Guérison', 'Nourriture', 'Or'],
                    n_passes=n_passes,
                    progress_cb=progress_cb
                )
            
            progress_bar.progress(1.0)
            gain = score_opt - score_init
            status_text.text(f"✅ Optimisation terminée — Score : {score_opt:,.0f} (gain : +{gain:,.0f})")
            
            # Résumé
            st.markdown("---")
            st.subheader("📊 Résultats")
            
            calculer_culture(placed_opt, grid.shape)
            
            from collections import defaultdict
            prod_summary = defaultdict(lambda: {'qte_init':0, 'qte_opt':0, 'boost_max':0})
            
            init_by_nom = defaultdict(list)
            for b in placed_init:
                init_by_nom[b['nom']].append(b)
            
            for bat in placed_opt:
                if bat['type'] == 'Producteur' and bat['production'] != 'Rien':
                    prod = bat['production']
                    prod_summary[prod]['qte_opt'] += bat.get('prod_boosted', bat['quantite'])
                    prod_summary[prod]['boost_max'] = max(
                        prod_summary[prod]['boost_max'], bat.get('boost', 0))
            
            calculer_culture(placed_init, grid.shape)
            for bat in placed_init:
                if bat['type'] == 'Producteur' and bat['production'] != 'Rien':
                    prod = bat['production']
                    prod_summary[prod]['qte_init'] += bat.get('prod_boosted', bat['quantite'])
            
            rows_data = []
            for prod, d in sorted(prod_summary.items()):
                rows_data.append({
                    "Production": prod,
                    "Qté/h avant": f"{d['qte_init']:,.0f}",
                    "Qté/h après": f"{d['qte_opt']:,.0f}",
                    "Boost max": f"{d['boost_max']}%",
                    "Gain": f"+{d['qte_opt']-d['qte_init']:,.0f}" if d['qte_opt'] >= d['qte_init']
                             else f"{d['qte_opt']-d['qte_init']:,.0f}"
                })
            
            if rows_data:
                st.dataframe(pd.DataFrame(rows_data), hide_index=True, use_container_width=True)
            else:
                st.info("Aucun bâtiment producteur trouvé.")
            
            # Déplacements
            moved_count = 0
            for b_opt in placed_opt:
                for b_ini in placed_init:
                    if b_opt['nom'] == b_ini['nom']:
                        if b_opt['row'] != b_ini['row'] or b_opt['col'] != b_ini['col']:
                            moved_count += 1
                        break
            
            st.info(f"🔄 Bâtiments déplacés : {moved_count}")
            
            # Export
            st.markdown("---")
            st.subheader("💾 Télécharger les résultats")
            
            calculer_culture(placed_init, grid.shape)
            output_bytes = ecrire_output(
                copy.deepcopy(placed_init),
                placed_opt,
                inside_mask,
                grid.shape,
                bat_info,
                prio_order if prio_order else ['Guérison', 'Nourriture', 'Or']
            )
            
            st.download_button(
                label="📥 Télécharger le fichier résultats Excel",
                data=output_bytes,
                file_name="resultats_optimisation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
