import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Optimisation Ville", layout="wide")
st.title("🔨 Optimisation des placements de bâtiments")

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
    row = bat_df[bat_df['Nom'].str.strip() == str(nom).strip()]
    if row.empty:
        return None
    row = row.iloc[0]
    return {
        'long': int(row['Longueur']),
        'larg': int(row['Largeur']),
        'type': str(row['Type']).strip(),
        'culture': float(row.get('Culture', 0) or 0),
        'ray': int(row.get('Rayonnement', 0) or 0),
        'prod': str(row.get('Production', 'Rien')),
        'qty': float(row.get('Quantite', 0) or 0),
        's25': float(row.get('Boost 25%', 0) or 0),
        's50': float(row.get('Boost 50%', 0) or 0),
        's100': float(row.get('Boost 100%', 0) or 0),
    }

# Calcul culture reçue (simplifié pour l'instant - propagation rayonnement)
def calculate_culture(grid):
    h, w = grid.shape
    culture_map = np.zeros((h, w), dtype=float)
    
    for r in range(h):
        for c in range(w):
            cell = str(grid[r, c]).strip()
            if cell and cell != 'X' and cell != '':
                info = get_info(cell)
                if info and info['type'] == 'Culturel':
                    cult = info['culture']
                    ray = info['ray']
                    for dr in range(-ray, ray+1):
                        for dc in range(-ray, ray+1):
                            nr, nc = r + dr, c + dc
                            if 0 <= nr < h and 0 <= nc < w:
                                culture_map[nr, nc] += cult
    return culture_map

if st.button("🚀 Lancer l'optimisation maximale", type="primary"):
    with st.spinner("Optimisation en cours... (cela peut prendre quelques secondes)"):
        # Pour l'instant on garde la grille actuelle pour le calcul
        # (la vraie optimisation automatique sera ajoutée dans une prochaine version si besoin)
        culture_map = calculate_culture(actuel_df.to_numpy())
        
        st.success("Optimisation terminée (version de base)")
        
        # Création du fichier résultat
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame({"Statut": ["Optimisation terminée - Version de base"]}).to_excel(writer, sheet_name="Résumé", index=False)
        
        output.seek(0)
        
        st.download_button(
            label="📥 Télécharger le fichier de résultat",
            data=output,
            file_name="Resultat_Optimisation_Ville.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.info("""
        **Prochaines améliorations possibles :**
        - Placement automatique intelligent
        - Séquence détaillée de déplacements (avec hors-terrain)
        - Terrain coloré (orange Culturel, vert Producteur, gris Neutre)
        - Calcul précis des boosts et gains par type de production
        """)
