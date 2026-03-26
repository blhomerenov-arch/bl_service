import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Saisie Instance", layout="wide")

# Style pour ressembler à l'application interne
st.markdown("""
    <style>
    .header {
        background-color: #0E7CFF;
        color: white;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
        margin-bottom: 20px;
    }
    .form-box {
        background-color: #f8f9fa;
        padding: 25px;
        border-radius: 10px;
        border: 1px solid #ddd;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>INSTANCES</h2></div>', unsafe_allow_html=True)

# Navigation
page = st.radio("Navigation", 
                ["📝 INSTANCES", "📊 RAPPORTS", "⚠️ DÉRANGEMENTS", "🔧 FIABILISATION", "⚖️ LITIGES"],
                horizontal=True)

if page == "📝 INSTANCES":
    st.subheader("Nouvelle Saisie")

    with st.form("saisie", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            demande = st.text_input("Demande*", placeholder="000D740B")
            nom = st.text_input("Nom")
            contact = st.text_input("Contact", placeholder="0666549488")
            adresse = st.text_area("Adresse", height=85)

        with col2:
            telecopie = st.text_input("N° de Téléscopie*", placeholder="525311326")
            date_reception = st.date_input("Date de réception", value=datetime.now().date())
            
            col_a, col_b = st.columns(2)
            with col_a:
                secteur = st.selectbox("Secteur", ["MHAMID", "BOUAAKAZ"])
            with col_b:
                agent = st.selectbox("Agent", ["hamid", "SHAKHMAN"])

            type_install = st.selectbox("Type d'installation", ["Installation Fixe"])

        # Motif
        st.write("**Motif**")
        motif_options = [
            "Adresse erronée", "Client refuse installation", "Transport saturé",
            "PC saturé", "INJOINABLE", "Local fermé + injoignable",
            "Création PC", "ETUDE CREATION PC", "MSAN saturé", "Autre"
        ]
        motif = st.selectbox("", motif_options, label_visibility="collapsed")

        if motif == "Autre":
            motif = st.text_input("Précisez le motif")

        submitted = st.form_submit_button("Valider", type="primary", use_container_width=True)

    if submitted:
        if demande and telecopie:
            st.success(f"✅ Enregistré - Demande : {demande}")
            st.balloons()
        else:
            st.error("Demande et N° Téléscopie sont obligatoires")

    # Tableau
    st.subheader("Liste des Instances")
    try:
        df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        st.dataframe(df, use_container_width=True, height=450)
    except:
        st.info("Fichier ETAT non chargé")

else:
    st.info(f"Page **{page}** en cours de développement")

st.caption("Application de gestion de chantier - MHAMID")
