import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

st.markdown("""
    <style>
    .header {background-color: #0E7CFF; color: white; padding: 15px; border-radius: 8px; text-align: center; margin-bottom: 15px;}
    .success {background-color: #d4edda; padding: 12px; border-radius: 8px;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>Gestion Chantier Fibre & RTC - MHAMID</h2></div>', unsafe_allow_html=True)

page = st.radio(
    "Navigation",
    ["📝 INSTANCES", "📊 RAPPORTS", "⚠️ DÉRANGEMENTS", "🔧 FIABILISATION", "⚖️ LITIGES"],
    horizontal=True,
    label_visibility="collapsed"
)

# ====================== FONCTION DÉTECTION AMÉLIORÉE ======================
def find_column(df, keywords):
    if df is None or df.empty:
        return None
    for col in df.columns:
        col_str = str(col).lower()
        if any(k in col_str for k in keywords):
            return col
    # Backup : prendre la première colonne qui contient beaucoup de texte
    for col in df.columns:
        if df[col].dtype == "object" and df[col].notna().sum() > 20:
            return col
    return None

# ====================== PAGE INSTANCES ======================
if page == "📝 INSTANCES":
    st.subheader("📝 Saisie du Motif Journalier")

    with st.form("saisie_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            demande = st.text_input("Demande*", placeholder="000D740B")
            nom = st.text_input("Nom")
            contact = st.text_input("Contact")
            adresse = st.text_area("Adresse", height=80)

        with col2:
            telecopie = st.text_input("N° de Téléscopie*", placeholder="525311326")
            date_reception = st.date_input("Date de réception", datetime.now().date())
            secteur = st.selectbox("Secteur", ["MHAMID", "BOUAAKAZ", "Province M'HAMID"])
            agent = st.selectbox("Agent", ["hamid", "SHAKHMAN"])

        motif_options = [
            "Adresse erronée", "Client refuse installation", "Transport saturé",
            "PC saturé", "INJOINABLE", "Local fermé + injoignable",
            "Création PC", "ETUDE CREATION PC", "MSAN saturé", "Autre"
        ]
        motif = st.selectbox("Motif", motif_options)

        if motif == "Autre":
            motif = st.text_input("Précisez le motif")

        if st.form_submit_button("✅ Valider et Enregistrer", type="primary", use_container_width=True):
            if demande and telecopie and motif:
                st.success(f"✅ Motif enregistré pour la demande **{demande}**")
                st.balloons()
            else:
                st.error("❌ Demande, Téléscopie et Motif sont obligatoires")

    st.subheader("Liste des Instances")
    try:
        df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        st.dataframe(df, use_container_width=True, height=500)
    except:
        st.warning("Impossible de charger le fichier ETAT")

# ====================== PAGE RAPPORTS ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")

    try:
        etat_df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        motif_df = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF")

        # Détection des colonnes
        motif_col = find_column(motif_df, ['motif', 'detail', 'pc mauvais', 'adresse', 'refuse'])
        secteur_col = find_column(etat_df, ['secteur', 'sector'])
        etat_col = find_column(etat_df, ['etat', 'état', 'state'])
        delai_col = find_column(etat_df, ['délai', 'delai'])

        st.success("✅ Fichiers chargés avec succès")

        # KPIs
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Commandes", len(etat_df))
        col2.metric("Total Motifs", len(motif_df))
        col3.metric("Délai Moyen", round(etat_df[delai_col].mean(), 1) if delai_col else "N/A")
        col4.metric("Commandes VA", len(etat_df[etat_df[etat_col].astype(str).str.upper() == 'VA']) if etat_col else 0)

        st.divider()

        # Graphique Motifs
        if motif_col:
            motif_series = motif_df[motif_col].astype(str).str.strip()
            motif_series = motif_series[(motif_series != 'nan') & (motif_series != '')]
            motif_count = motif_series.value_counts().head(
