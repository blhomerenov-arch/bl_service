import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

# Authentification (déjà faite)
if "authenticated" not in st.session_state or not st.session_state.authenticated:
    st.stop()

st.title("📊 Rapports et Statistiques")

try:
    etat_df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
    motif_df = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF")

    st.success("✅ Fichiers chargés avec succès")

    st.subheader("📋 Colonnes détectées dans ETAT FTTH RTC RTCL")
    st.write(etat_df.columns.tolist())

    st.subheader("📋 Colonnes détectées dans MOTIF TOTAL")
    st.write(motif_df.columns.tolist())

    st.divider()

    # Tentative de détection
    motif_col = None
    for col in motif_df.columns:
        if 'motif' in str(col).lower() or 'detail' in str(col).lower():
            motif_col = col
            break

    if motif_col:
        st.success(f"✅ Colonne Motif trouvée : **{motif_col}**")
        motif_count = motif_df[motif_col].value_counts().head(10)
        st.bar_chart(motif_count)
    else:
        st.error("❌ Aucune colonne Motif trouvée")

except Exception as e:
    st.error(f"Erreur : {str(e)}")

st.caption("Si tu vois les colonnes, dis-moi lesquelles apparaissent pour que je puisse améliorer la détection.")
