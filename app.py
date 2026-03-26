import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

# Style pour ressembler à l'application interne
st.markdown("""
    <style>
    .main-header {
        background-color: #0E7CFF;
        color: white;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-header"><h2>Gestion Chantier Fibre & RTC - MHAMID</h2></div>', unsafe_allow_html=True)

# ====================== NAVIGATION RÉELLE ======================
page = st.sidebar.radio(
    "Navigation",
    ["📊 RAPPORTS", "📝 INSTANCES", "⚠️ DÉRANGEMENTS", "🔧 FIABILISATION", "⚖️ LITIGES"],
    label_visibility="collapsed"
)

# ====================== PAGE INSTANCES (la plus importante) ======================
if page == "📝 INSTANCES":
    st.subheader("📝 Saisie du Motif Journalier")

    with st.form("saisie_form", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            demande = st.text_input("Demande*", placeholder="000D740B")
            nom = st.text_input("Nom")
            contact = st.text_input("Contact", placeholder="0666549488")
            adresse = st.text_area("Adresse", height=85)

        with col2:
            telecopie = st.text_input("N° de Téléscopie*", placeholder="525311326")
            date_reception = st.date_input("Date de réception", datetime.now().date())
            secteur = st.selectbox("Secteur", ["MHAMID", "BOUAAKAZ", "Province M'HAMID"])
            agent = st.selectbox("Agent", ["hamid", "SHAKHMAN"])

        categorie = st.selectbox("Catégorie", ["RTC DTL", "GPON", "GPON DFO"])
        type_install = st.selectbox("Type d'installation", ["Installation Fixe"])

        motif_list = [
            "Adresse erronée", "Client refuse installation", "Transport saturé",
            "PC saturé", "INJOINABLE", "Local fermé + injoignable",
            "Création PC", "ETUDE CREATION PC", "MSAN saturé", "Autre"
        ]
        motif = st.selectbox("Motif", motif_list)

        if motif == "Autre":
            motif = st.text_input("Précisez le motif")

        submitted = st.form_submit_button("✅ Valider et Enregistrer", type="primary", use_container_width=True)

    if submitted:
        if demande and telecopie and motif:
            st.success(f"✅ Motif enregistré pour la demande **{demande}**")
            st.balloons()
        else:
            st.error("Demande, Téléscopie et Motif sont obligatoires")

    # Tableau
    st.subheader("Liste des Instances")
    try:
        df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        st.dataframe(df, use_container_width=True, height=500)
    except:
        st.warning("Fichier ETAT FTTH RTC RTCL.xlsx non trouvé")

# ====================== AUTRES PAGES (à compléter plus tard) ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")
    st.info("Page Rapports - En cours de développement")

elif page == "⚠️ DÉRANGEMENTS":
    st.subheader("⚠️ Dérangements")
    st.info("Page Dérangements - En cours de développement")

elif page == "🔧 FIABILISATION":
    st.subheader("🔧 Fiabilisation")
    st.info("Page Fiabilisation - En cours de développement")

elif page == "⚖️ LITIGES":
    st.subheader("⚖️ Litiges")
    st.info("Page Litiges - En cours de développement")

st.caption("Application Chantier MHAMID - Fibre & RTC")
