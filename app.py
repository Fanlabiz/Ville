import pandas as pd
import numpy as np
import os
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

# ===================== CONFIGURATION =====================
# Chemin absolu vers le fichier (fonctionne dans Streamlit Cloud / votre setup)
EXCEL_FILE = "/home/workdir/Ville opti.xlsx"
OUTPUT_FILE = "Resultat_Optimisation_Ville.xlsx"

st.title("Optimisation Ville - Placement Bâtiments")

st.write("Répertoire actuel :", os.getcwd())
st.write("Fichier Excel :", EXCEL_FILE)

if not os.path.exists(EXCEL_FILE):
    st.error(f"❌ Fichier non trouvé : {EXCEL_FILE}")
    st.stop()

# ===================== LECTURE DES DONNÉES =====================
try:
    terrain_df = pd.read_excel(EXCEL_FILE, sheet_name="Terrain", header=None)
    bat_df_raw = pd.read_excel(EXCEL_FILE, sheet_name="Batiments")
    actuel_df = pd.read_excel(EXCEL_FILE, sheet_name="Actuel", header=None)

    # Nettoyage Batiments
    bat_df = bat_df_raw.iloc[1:].reset_index(drop=True)
    bat_df.columns = ["Nom", "Longueur", "Largeur", "Nombre", "Type", "Culture", 
                      "Rayonnement", "Boost 25%", "Boost 50%", "Boost 100%", 
                      "Production", "Quantite"]

    terrain = terrain_df.to_numpy()
    grid_actuel = actuel_df.to_numpy()

    st.success("✅ Fichier Excel chargé avec succès !")
    st.write(f"Terrain : {terrain.shape[0]} lignes × {terrain.shape[1]} colonnes")
    st.write(f"{len(bat_df)} types de bâtiments détectés")

except Exception as e:
    st.error(f"Erreur lors du chargement : {e}")
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

st.write("Script prêt pour l'optimisation.")

# Bouton pour lancer l'optimisation (on va l'implémenter ensuite)
if st.button("🚀 Lancer l'optimisation maximale"):
    st.info("Optimisation en cours... (version complète à venir)")
    # Ici on mettra la logique complète
