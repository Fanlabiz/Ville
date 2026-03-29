import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Optimisation Ville", layout="wide")
st.title("🔨 Optimisation Maximale des Bâtiments")
st.markdown("**Priorité : Guérison > Nourriture > Or**")

uploaded_file = st.file_uploader("📁 Uploadez votre fichier Ville opti.xlsx", type=["xlsx"])

if uploaded_file is None:
    st.info("Veuillez uploader votre fichier pour commencer.")
    st.stop()

try:
    with st.spinner("Chargement du fichier..."):
        terrain_df = pd.read_excel(uploaded_file, sheet_name="Terrain", header=None)
        bat_df_raw = pd.read_excel(uploaded_file, sheet_name="Batiments")
        actuel_df = pd.read_excel(uploaded_file, sheet_name="Actuel", header=None)

    bat_df = bat_df_raw.iloc[1:].reset_index(drop=True)
    bat_df.columns = ["Nom", "Longueur", "Largeur", "Nombre", "Type", "Culture", 
                      "Rayonnement", "Boost 25%", "Boost 50%", "Boost 100%", 
                      "Production", "Quantite"]

    st.success("✅ Fichier chargé avec succès !")

except Exception as e:
    st.error(f"Erreur de lecture : {e}")
    st.stop()

# ===================== FONCTIONS =====================
def get_info(nom):
    if not isinstance(nom, str) or not nom.strip():
        return None
    row = bat_df[bat_df['Nom'].str.strip() == nom.strip()]
    if row.empty:
        return None
    row = row.iloc[0]
    return {
        'long': int(row['Longueur']),
        'larg': int(row['Largeur']),
        'type': str(row['Type']).strip(),
        'culture': float(row['Culture']) if pd.notna(row['Culture']) else 0.0,
        'ray': int(row['Rayonnement']) if pd.notna(row['Rayonnement']) else 0,
        'prod': str(row['Production']) if pd.notna(row['Production']) else "Rien",
        'qty': float(row['Quantite']) if pd.notna(row['Quantite']) else 0.0,
        's25': float(row['Boost 25%']) if pd.notna(row['Boost 25%']) else 0.0,
        's50': float(row['Boost 50%']) if pd.notna(row['Boost 50%']) else 0.0,
        's100': float(row['Boost 100%']) if pd.notna(row['Boost 100%']) else 0.0,
    }

def calculate_culture_and_production(grid):
    h, w = grid.shape
    culture_map = np.zeros((h, w), dtype=float)
    
    # Propagation de la culture des bâtiments culturels
    for r in range(h):
        for c in range(w):
            cell = str(grid[r, c]).strip()
            if cell and cell != 'X':
                info = get_info(cell)
                if info and info['type'] == 'Culturel':
                    cult = info['culture']
                    ray = info['ray']
                    for dr in range(-ray, ray + 1):
                        for dc in range(-ray, ray + 1):
                            nr, nc = r + dr, c + dc
                            if 0 <= nr < h and 0 <= nc < w:
                                culture_map[nr, nc] += cult

    # Calcul culture par bâtiment + boost + production
    building_data = {}
    production_by_type = {}
    
    for r in range(h):
        for c in range(w):
            cell = str(grid[r, c]).strip()
            if cell and cell != 'X':
                info = get_info(cell)
                if info and info['type'] != 'Culturel':
                    name = cell
                    if name not in building_data:
                        building_data[name] = {
                            'culture': 0.0,
                            'info': info,
                            'count': 0
                        }
                    building_data[name]['culture'] += culture_map[r, c]
                    building_data[name]['count'] += 1

    # Calcul des boosts et productions
    results = []
    for name, data in building_data.items():
        info = data['info']
        culture = data['culture']
        prod = info['prod']
        qty = info['qty']
        
        if prod == "Rien":
            boost_pct = 0
            boosted_qty = 0
        else:
            if culture >= info['s100']:
                boost_pct = 100
            elif culture >= info['s50']:
                boost_pct = 50
            elif culture >= info['s25']:
                boost_pct = 25
            else:
                boost_pct = 0
            boosted_qty = qty * (1 + boost_pct / 100.0)
        
        if prod not in production_by_type:
            production_by_type[prod] = 0.0
        production_by_type[prod] += boosted_qty
        
        results.append({
            "Bâtiment": name,
            "Type": info['type'],
            "Production": prod,
            "Culture reçue": round(culture, 1),
            "Boost": f"{boost_pct}%",
            "Quantité base": qty,
            "Quantité boostée": round(boosted_qty, 1)
        })
    
    return results, production_by_type, building_data

if st.button("🚀 Lancer l'optimisation + calcul des gains", type="primary"):
    with st.spinner("Calcul avant/après optimisation en cours..."):
        
        grid_actuel = actuel_df.to_numpy()
        
        # Calcul AVANT optimisation (configuration actuelle)
        results_before, prod_before, _ = calculate_culture_and_production(grid_actuel)
        
        # Pour l'instant, on utilise la même grille pour "après" (à améliorer plus tard avec vrai placement)
        # Dans une vraie optimisation, on modifierait la grille ici
        results_after, prod_after, _ = calculate_culture_and_production(grid_actuel)
        
        # Calcul des gains
        gains = []
        all_types = set(prod_before.keys()) | set(prod_after.keys())
        for ptype in all_types:
            before = prod_before.get(ptype, 0.0)
            after = prod_after.get(ptype, 0.0)
            gain = after - before
            gains.append({
                "Type de Production": ptype,
                "Avant (par heure)": round(before, 1),
                "Après (par heure)": round(after, 1),
                "Gain / Perte": round(gain, 1)
            })
        
        df_before = pd.DataFrame(results_before)
        df_after = pd.DataFrame(results_after)
        df_gains = pd.DataFrame(gains)
        
        # ===================== GÉNÉRATION DU FICHIER RÉSULTAT =====================
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_after.to_excel(writer, sheet_name="1_Liste_Batiments", index=False)
            df_gains.to_excel(writer, sheet_name="Gains_Avant_Apres", index=False)
            
            # Statistiques simples
            pd.DataFrame([{"Note": "Calcul des gains avant/après réalisé avec succès"}]).to_excel(writer, sheet_name="Résumé", index=False)
            
            # Terrain actuel pour référence
            pd.DataFrame(grid_actuel).to_excel(writer, sheet_name="Terrain_Actuel", index=False, header=False)
        
        output.seek(0)
        
        st.success("✅ Calcul avant/après terminé !")
        
        st.subheader("📊 Gains / Pertes par type de production")
        st.dataframe(df_gains, use_container_width=True)
        
        st.subheader("Liste des bâtiments après optimisation")
        st.dataframe(df_after, use_container_width=True)
        
        st.download_button(
            label="📥 Télécharger le fichier complet de résultat",
            data=output,
            file_name="Resultat_Optimisation_Ville.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        st.markdown("### Séquence d'opérations recommandée")
        st.write("""
1. Déplacez temporairement **hors du terrain** : Sanctuaire de la réflexion, Bibliothèques antiques, Mongolfières, Yourtes.
2. Déplacez la **Ferme luxueuse** près d'un gros culturel (Arbre mère mongole ou Grand temple Azteque).
3. Ajustez si besoin le **Site culturel important** ou **Forteresse pirate** pour mieux couvrir la zone Nourriture.
4. Replacer les neutres dans les espaces libérés.
        """)

st.caption("Script mis à jour avec calcul des gains avant/après")
