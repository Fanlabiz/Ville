import pandas as pd
import streamlit as st
import os
from io import BytesIO

st.set_page_config(page_title="Optimisation Ville", layout="wide")
st.title("🔨 Optimisation des placements de bâtiments")

st.markdown("""
**Instructions :**  
1. Uploadez votre fichier **Ville opti.xlsx**  
2. Cliquez sur "Analyser le fichier"  
3. Lancez ensuite l'optimisation
""")

# ===================== UPLOAD DU FICHIER =====================
uploaded_file = st.file_uploader("📁 Uploadez votre fichier Excel (Ville opti.xlsx)", 
                                 type=["xlsx"], 
                                 help="Fichier contenant les onglets Terrain, Batiments et Actuel")

if uploaded_file is None:
    st.info("👆 Veuillez uploader votre fichier Excel pour commencer.")
    st.stop()

# ===================== LECTURE DU FICHIER UPLOADÉ =====================
try:
    with st.spinner("Chargement du fichier..."):
        terrain_df = pd.read_excel(uploaded_file, sheet_name="Terrain", header=None)
        bat_df_raw = pd.read_excel(uploaded_file, sheet_name="Batiments")
        actuel_df = pd.read_excel(uploaded_file, sheet_name="Actuel", header=None)

    # Nettoyage de l'onglet Batiments
    bat_df = bat_df_raw.iloc[1:].reset_index(drop=True)
    bat_df.columns = ["Nom", "Longueur", "Largeur", "Nombre", "Type", "Culture", 
                      "Rayonnement", "Boost 25%", "Boost 50%", "Boost 100%", 
                      "Production", "Quantite"]

    st.success("✅ Fichier chargé avec succès !")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Lignes Terrain", terrain_df.shape[0])
    with col2:
        st.metric("Types de bâtiments", len(bat_df))
    with col3:
        st.metric("Taille grille actuelle", f"{actuel_df.shape[0]} × {actuel_df.shape[1]}")

    # Aperçu rapide
    if st.checkbox("Afficher aperçu du terrain (premières lignes)"):
        st.dataframe(terrain_df.iloc[:10, :15], use_container_width=True)

except Exception as e:
    st.error(f"❌ Erreur lors de la lecture du fichier : {e}")
    st.stop()

# ===================== FONCTION INFO BÂTIMENT =====================
def get_building_info(nom):
    nom = str(nom).strip()
    row = bat_df[bat_df['Nom'].str.strip() == nom]
    if row.empty:
        return None
    row = row.iloc[0]
    return {
        'nom': row['Nom'],
        'longueur': int(row['Longueur']),
        'largeur': int(row['Largeur']),
        'type': str(row['Type']).strip(),
        'culture': float(row['Culture']) if pd.notna(row['Culture']) else 0.0,
        'rayonnement': int(row['Rayonnement']) if pd.notna(row['Rayonnement']) else 0,
        'production': str(row['Production']) if pd.notna(row['Production']) else "Rien",
        'quantite': float(row['Quantite']) if pd.notna(row['Quantite']) else 0.0,
        'seuil_25': float(row['Boost 25%']) if pd.notna(row['Boost 25%']) else 0.0,
        'seuil_50': float(row['Boost 50%']) if pd.notna(row['Boost 50%']) else 0.0,
        'seuil_100': float(row['Boost 100%']) if pd.notna(row['Boost 100%']) else 0.0,
    }

st.divider()
st.write("**Fichier prêt pour optimisation.**")

# Bouton pour lancer l'optimisation (on ajoutera la logique complète juste après)
if st.button("🚀 Lancer l'optimisation maximale (priorité Guérison > Nourriture > Or)", type="primary"):
    st.info("Optimisation en cours... Cette partie sera développée dans la prochaine version.")
    # Ici on mettra tout le code d'optimisation + génération du fichier résultat

# Option pour télécharger le fichier exemple (si besoin)
with open("/home/workdir/Ville opti.xlsx", "rb") as f:
    st.download_button(
        label="📥 Télécharger le fichier exemple (Ville opti.xlsx)",
        data=f,
        file_name="Ville opti.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
