import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

# Style
st.markdown("""
    <style>
    .header {background-color: #0E7CFF; color: white; padding: 15px; border-radius: 8px; text-align: center; margin-bottom: 15px;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>Gestion Chantier Fibre & RTC - MHAMID</h2></div>', unsafe_allow_html=True)

# Navigation horizontale
page = st.radio(
    "Navigation",
    ["📝 INSTANCES", "📊 RAPPORTS", "⚠️ DÉRANGEMENTS", "🔧 FIABILISATION", "⚖️ LITIGES"],
    horizontal=True,
    label_visibility="collapsed"
)

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

        submitted = st.form_submit_button("✅ Valider et Enregistrer", type="primary", use_container_width=True)

    if submitted:
        if demande and telecopie and motif:
            st.success(f"✅ Motif enregistré pour la demande **{demande}**")
            st.balloons()
        else:
            st.error("Demande, Téléscopie et Motif sont obligatoires")

    # Tableau des instances
    st.subheader("Liste des Instances")
    try:
        df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        st.dataframe(df, use_container_width=True, height=500)
    except:
        st.warning("Impossible de charger le fichier ETAT FTTH RTC RTCL.xlsx")

# ====================== PAGE RAPPORTS ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")

    try:
        df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        motif_df = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF") if "MOTIF TOTAL (1).xlsx" else pd.DataFrame()

        col1, col2, col3 = st.columns(3)
        col1.metric("Total Commandes", len(df))
        col2.metric("Motifs Cumulés", len(motif_df))
        col3.metric("Délai Moyen (jours)", round(df['Délai(j)'].mean(), 1) if 'Délai(j)' in df.columns else "N/A")

        st.subheader("Top 10 des Motifs les plus fréquents")
        if not motif_df.empty and 'Motif' in motif_df.columns:
            top_motifs = motif_df['Motif'].value_counts().head(10)
            st.bar_chart(top_motifs)
            st.dataframe(top_motifs.rename("Nombre"), use_container_width=True)
        else:
            st.info("Aucun motif chargé pour le moment")

    except Exception as e:
        st.error("Erreur lors du chargement des rapports")

# ====================== AUTRES PAGES ======================
elif page == "⚠️ DÉRANGEMENTS":
    st.subheader("⚠️ Dérangements")
    st.info("Liste des dérangements en cours - À compléter avec tes données")

elif page == "🔧 FIABILISATION":
    st.subheader("🔧 Fiabilisation")
    st.info("Suivi des actions de fiabilisation du réseau")

elif page == "⚖️ LITIGES":
    st.subheader("⚖️ Litiges")
    st.info("Gestion des litiges clients et techniques")

st.caption("Application de gestion de chantier MHAMID - Fibre & RTC")
