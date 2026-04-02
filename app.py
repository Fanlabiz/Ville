import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io
import copy
from itertools import product

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
    """Score pondéré selon les priorités de production."""
    if prio_order is None:
        prio_order = ['Guérison', 'Nourriture', 'Or']
    
    score = 0.0
    weights = {}
    for i, p in enumerate(prio_order):
        weights[p] = 10 ** (len(prio_order) - i)
    
    for bat in placed_list:
        if bat['type'] == 'Producteur' and bat['production'] != 'Rien':
            w = weights.get(bat['production'], 1)
            score += w * bat['prod_boosted']
    return score

# ─────────────────────────────────────────────
# OPTIMISATION
# ─────────────────────────────────────────────

def cases_libres(placed_list, inside_mask, grid_shape):
    """Retourne le masque des cases libres (inside et non occupées)."""
    rows, cols = grid_shape
    occupied = np.zeros((rows, cols), dtype=bool)
    for bat in placed_list:
        for dr in range(bat['h']):
            for dc in range(bat['w']):
                r, c = bat['row'] + dr, bat['col'] + dc
                if 0 <= r < rows and 0 <= c < cols:
                    occupied[r, c] = True
    free = inside_mask & ~occupied
    return free

def peut_placer(bat, row, col, orientation, placed_list, inside_mask, grid_shape, exclude_id=None):
    """Vérifie si un bâtiment peut être placé en (row, col) avec cette orientation."""
    rows, cols = grid_shape
    info = bat
    if orientation == 'H':
        h, w = info['largeur'], info['longueur']
    else:
        h, w = info['longueur'], info['largeur']
    
    occupied = np.zeros((rows, cols), dtype=bool)
    for i, b in enumerate(placed_list):
        if exclude_id is not None and i == exclude_id:
            continue
        for dr in range(b['h']):
            for dc in range(b['w']):
                r2, c2 = b['row'] + dr, b['col'] + dc
                if 0 <= r2 < rows and 0 <= c2 < cols:
                    occupied[r2, c2] = True
    
    for dr in range(h):
        for dc in range(w):
            r2, c2 = row + dr, col + dc
            if r2 >= rows or c2 >= cols:
                return False
            if not inside_mask[r2, c2]:
                return False
            if occupied[r2, c2]:
                return False
    return True

def optimiser(placed_list, inside_mask, grid_shape, prio_order, n_passes=3, progress_cb=None):
    """
    Optimisation par échanges de positions entre bâtiments.
    Chaque passe tente tous les swaps de paires et tous les déplacements simples.
    """
    best = copy.deepcopy(placed_list)
    calculer_culture(best, grid_shape)
    best_score = score_total(best, prio_order)
    
    rows, cols = grid_shape
    improved = True
    pass_num = 0
    
    while improved and pass_num < n_passes:
        improved = False
        pass_num += 1
        if progress_cb:
            progress_cb(pass_num, n_passes, best_score)
        
        n = len(best)
        
        # Passe 1 : déplacer chaque bâtiment vers une meilleure position
        for i in range(n):
            bat = best[i]
            orig_row, orig_col, orig_h, orig_w = bat['row'], bat['col'], bat['h'], bat['w']
            
            for orientation in ['H', 'V']:
                if orientation == 'H':
                    h, w = bat['largeur'], bat['longueur']
                else:
                    h, w = bat['longueur'], bat['largeur']
                
                for r in range(rows):
                    for c in range(cols):
                        if r == orig_row and c == orig_col and h == orig_h and w == orig_w:
                            continue
                        if peut_placer(bat, r, c, orientation, best, inside_mask, grid_shape, exclude_id=i):
                            candidate = copy.deepcopy(best)
                            candidate[i]['row'] = r
                            candidate[i]['col'] = c
                            candidate[i]['h'] = h
                            candidate[i]['w'] = w
                            candidate[i]['orientation'] = orientation
                            calculer_culture(candidate, grid_shape)
                            s = score_total(candidate, prio_order)
                            if s > best_score + 1e-6:
                                best = candidate
                                best_score = s
                                improved = True
                                break
                    else:
                        continue
                    break
        
        # Passe 2 : swaps de paires
        for i in range(n):
            for j in range(i+1, n):
                bi, bj = best[i], best[j]
                
                for ori_i in ['H', 'V']:
                    hi = bi['largeur'] if ori_i == 'H' else bi['longueur']
                    wi = bi['longueur'] if ori_i == 'H' else bi['largeur']
                    
                    for ori_j in ['H', 'V']:
                        hj = bj['largeur'] if ori_j == 'H' else bj['longueur']
                        wj = bj['longueur'] if ori_j == 'H' else bj['largeur']
                        
                        # Essayer de placer bi à la position de bj et vice-versa
                        candidate = copy.deepcopy(best)
                        
                        rj, cj = bj['row'], bj['col']
                        ri, ci = bi['row'], bi['col']
                        
                        # Retirer les deux bâtiments puis vérifier placements
                        candidate[i]['row'] = rj
                        candidate[i]['col'] = cj
                        candidate[i]['h'] = hi
                        candidate[i]['w'] = wi
                        candidate[i]['orientation'] = ori_i
                        
                        candidate[j]['row'] = ri
                        candidate[j]['col'] = ci
                        candidate[j]['h'] = hj
                        candidate[j]['w'] = wj
                        candidate[j]['orientation'] = ori_j
                        
                        # Vérification de validité simplifiée (taille différente → on ignore)
                        if hi == hj and wi == wj:
                            calculer_culture(candidate, grid_shape)
                            s = score_total(candidate, prio_order)
                            if s > best_score + 1e-6:
                                best = candidate
                                best_score = s
                                improved = True
    
    return best, best_score

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
    
    headers = ["Nom", "Type", "Production", "Ligne", "Colonne",
               "Hauteur", "Largeur", "Culture recue", "Boost %",
               "Qte/h base", "Qte/h boostee"]
    for c, h in enumerate(headers, 1):
        cell = ws1.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.fill = fill(BLEU_CLAIR)
        cell.border = border_thin()
    
    for i, bat in enumerate(placed_opt, 2):
        vals = [
            bat['nom'], bat['type'], bat['production'],
            bat['row']+1, bat['col']+1, bat['h'], bat['w'],
            bat.get('culture_recue', 0), bat.get('boost', 0),
            bat['quantite'], bat.get('prod_boosted', bat['quantite'])
        ]
        bg = ORANGE if bat['type'] == 'Culturel' else (VERT if bat['type'] == 'Producteur' else GRIS)
        for c, v in enumerate(vals, 1):
            cell = ws1.cell(row=i, column=c, value=v)
            cell.fill = fill(bg)
            cell.border = border_thin()
    
    for c in range(1, len(headers)+1):
        ws1.column_dimensions[get_column_letter(c)].width = 18
    
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
    
    # Rebuild grid opt
    grid_opt = np.full((rows, cols), '', dtype=object)
    for bat in placed_opt:
        for dr in range(bat['h']):
            for dc in range(bat['w']):
                r2, c2 = bat['row']+dr, bat['col']+dc
                if 0 <= r2 < rows and 0 <= c2 < cols:
                    grid_opt[r2, c2] = bat['nom']
    
    # Créer un dict bâtiment par position top-left pour le boost
    boost_at = {}
    for bat in placed_opt:
        boost_at[(bat['row'], bat['col'])] = bat.get('boost', 0)
    
    bat_at = {}
    for bat in placed_opt:
        bat_at[(bat['row'], bat['col'])] = bat
    
    FILL_X    = PatternFill("solid", fgColor="FF404040")
    FILL_OUT  = PatternFill("solid", fgColor="FFFFFFFF")
    FILL_CULT = PatternFill("solid", fgColor="FFFFA500")
    FILL_PROD = PatternFill("solid", fgColor="FF90EE90")
    FILL_NEUT = PatternFill("solid", fgColor="FFD3D3D3")
    FILL_EMPTY= PatternFill("solid", fgColor="FFF5F5F5")
    
    for r in range(rows):
        ws5.row_dimensions[r+1].height = 18
        for c in range(cols):
            cell = ws5.cell(row=r+1, column=c+1)
            ws5.column_dimensions[get_column_letter(c+1)].width = 6
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.font = Font(size=7)
            cell.border = border_thin()
            
            val = grid_opt[r, c]
            
            # Bord X
            if val == 'X':
                cell.fill = FILL_X
                cell.font = Font(size=7, color="FFFFFFFF", bold=True)
                cell.value = 'X'
            elif not inside_mask[r, c]:
                cell.fill = FILL_OUT
            elif val == '':
                cell.fill = FILL_EMPTY
            else:
                # Trouver le bâtiment
                # Chercher dans placed_opt lequel occupe cette case
                found_bat = None
                for bat in placed_opt:
                    if bat['row'] <= r <= bat['row']+bat['h']-1 and \
                       bat['col'] <= c <= bat['col']+bat['w']-1:
                        found_bat = bat
                        break
                
                if found_bat:
                    typ = found_bat['type']
                    if typ == 'Culturel':
                        cell.fill = FILL_CULT
                    elif typ == 'Producteur':
                        cell.fill = FILL_PROD
                    else:
                        cell.fill = FILL_NEUT
                    
                    # Afficher le nom et boost uniquement sur la case top-left
                    if r == found_bat['row'] and c == found_bat['col']:
                        boost = found_bat.get('boost', 0)
                        label = found_bat['nom']
                        if boost > 0:
                            label += f"\n+{boost}%"
                        cell.value = label
                        cell.font = Font(size=7, bold=True)
    
    # Largeur uniforme
    for c in range(cols):
        ws5.column_dimensions[get_column_letter(c+1)].width = 7
    
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
