import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Optimisation Ville", layout="wide")
st.title("🔨 Optimisation des placements de bâtiments")

st.markdown("""
**Comment utiliser :**  
1. Uploadez votre fichier **Ville opti.xlsx** (celui avec les 3 onglets : Terrain, Batiments, Actuel)  
2. Cliquez sur **Analyser le fichier**  
3. Lancez l'optimisation
""")

# ===================== UPLOAD DU FICHIER =====================
uploaded_file = st.file_uploader(
    "📁 Uploadez votre fichier Excel (Ville opti.xlsx)", 
    type=["xlsx"],
    help="Fichier contenant les onglets : Terrain, Batiments et Actuel"
)

if uploaded_file is None:
    st.info("👆 Veuillez uploader votre fichier pour commencer.")
    st.stop()

# ===================== LECTURE DU FICHIER =====================
try:
    with st.spinner("Chargement et analyse du fichier..."):
        # Lecture des 3 onglets
        terrain_df = pd.read_excel(uploaded_file, sheet_name="Terrain", header=None)
        bat_df_raw = pd.read_excel(uploaded_file, sheet_name="Batiments")
        actuel_df = pd.read_excel(uploaded_file, sheet_name="Actuel", header=None)

        # Nettoyage Batiments
        bat_df = bat_df_raw.iloc[1:].reset_index(drop=True)
        bat_df.columns = ["Nom", "Longueur", "Largeur", "Nombre", "Type", "Culture", 
                          "Rayonnement", "Boost 25%", "Boost 50%", "Boost 100%", 
                          "Production", "Quantite"]

    st.success("✅ Fichier chargé avec succès !")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Lignes dans Terrain", terrain_df.shape[0])
    with col2:
        st.metric("Types de bâtiments", len(bat_df))
    with col3:
        st.metric("Taille de la grille actuelle", f"{actuel_df.shape[0]} × {actuel_df.shape[1]}")

    # Aperçu optionnel
    if st.checkbox("Voir aperçu des 10 premières lignes du terrain"):
        st.dataframe(terrain_df.iloc[:10, :20])

except Exception as e:
    st.error(f"❌ Erreur pendant le chargement du fichier :\n{e}")
    st.stop()

# ===================== FONCTION UTILE =====================
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
st.success("Le fichier est prêt pour l'optimisation.")

# Bouton principal
if st.button("🚀 Lancer l'optimisation maximale", type="primary", use_container_width=True):
    st.info("🔄 Optimisation en cours... (la logique complète arrive dans la prochaine mise à jour)")
    # Ici on ajoutera tout le code d'optimisation + génération du fichier résultat

st.caption("Développé pour maximiser Guérison → Nourriture → Or")
