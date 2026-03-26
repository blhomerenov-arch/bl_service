import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Saisie Instance", layout="wide")

# Style CSS pour ressembler à l'application interne
st.markdown("""
    <style>
    .main-header {
        background-color: #0E7CFF;
        color: white;
        padding: 12px 20px;
        border-radius: 8px;
        text-align: center;
        margin-bottom: 20px;
    }
    .form-container {
        background-color: #f8f9fa;
        padding: 25px;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
    }
    .stTextInput > div > div > input, .stSelectbox > div > div {
        background-color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

# En-tête bleu comme sur ta photo
st.markdown('<div class="main-header"><h2>INSTANCES</h2></div>', unsafe_allow_html=True)

# Navigation
st.tabs(["RAPPORTS", "INSTANCES (9)", "DÉRANGEMENTS (0)", "FIABILISATION", "LITIGES"])

st.subheader("Nouvelle Saisie")

# ====================== FORMULAIRE PLUS PROCHE DE TA PHOTO ======================
with st.form("saisie_form", clear_on_submit=True):
    col1, col2 = st.columns([1, 1])

    with col1:
        demande = st.text_input("Demande*", placeholder="000D740B", key="demande")
        nom = st.text_input("Nom", placeholder="")
        contact = st.text_input("Contact", placeholder="0666549488")
        adresse = st.text_area("Adresse", height=90, placeholder="MHAMID Lotissement Mhamid 9 ...")

    with col2:
        telecopie = st.text_input("N° de Téléscopie*", placeholder="525311326")
        date_reception = st.date_input("Date de réception", value=datetime.now().date())
        
        col_a, col_b = st.columns(2)
        with col_a:
            secteur = st.selectbox("Secteur", ["MHAMID", "BOUAAKAZ", "Province M'HAMID"])
        with col_b:
            agent = st.selectbox("Agent", ["hamid", "SHAKHMAN"])

        categorie = st.selectbox("Catégorie", ["RTC DTL", "GPON", "GPON DFO", "RTC"])
        type_install = st.selectbox("Type d'installation", ["Installation Fixe"])

    # Motif
    st.markdown("**Motif**")
    motifs = [
        "Adresse erronée", "Client refuse installation", "Transport saturé Mha", 
        "PC saturé", "INJOINABLE", "Local fermé + injoignable", 
        "Création PC", "ETUDE CREATION PC", "MSAN saturé", "Autre"
    ]
    motif_selection = st.selectbox("", motifs, label_visibility="collapsed")

    if motif_selection == "Autre":
        motif_final = st.text_input("Précisez le motif détaillé")
    else:
        motif_final = motif_selection

    # Bouton centré et large
    submitted = st.form_submit_button("Valider", type="primary", use_container_width=True)

# ====================== TRAITEMENT ======================
if submitted:
    if not demande or not telecopie or not motif_final:
        st.error("⚠️ Les champs Demande, N° de Téléscopie et Motif sont obligatoires")
    else:
        st.success(f"✅ Enregistré avec succès - Demande : **{demande}**")
        st.balloons()

# ====================== TABLEAU ======================
st.subheader("Liste des Instances")

try:
    df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
    search = st.text_input("🔍 Rechercher (Demande, Nom, Adresse...)")
    
    if search:
        mask = df.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        display_df = df[mask]
    else:
        display_df = df
    
    st.dataframe(display_df, use_container_width=True, height=450)
    
except Exception as e:
    st.warning("Impossible de charger le fichier ETAT FTTH RTC RTCL.xlsx. Vérifiez que le fichier est bien dans le repository.")

st.caption("Application de gestion de chantier MHAMID - Fibre & RTC")

 
