import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

# ====================== AUTHENTIFICATION SIMPLIFIÉE ======================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔐 Connexion - Gestion Chantier MHAMID")
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("### Identifiez-vous")
        username = st.text_input("Nom d'utilisateur", placeholder="admin")
        password = st.text_input("Mot de passe", type="password", placeholder="1234")
        
        if st.button("Se connecter", type="primary", use_container_width=True):
            if username == "admin" and password == "1234":
                st.session_state.authenticated = True
                st.success("✅ Connexion réussie !")
                st.rerun()
            else:
                st.error("❌ Identifiants incorrects. Essayez admin / 1234")
    
    st.caption("Utilisateur par défaut : **admin** | Mot de passe : **1234**")
    st.stop()

# ====================== APPLICATION PRINCIPALE ======================
st.markdown("""
    <style>
    .header {background-color: #0E7CFF; color: white; padding: 15px; border-radius: 8px; text-align: center; margin-bottom: 15px;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>Gestion Chantier Fibre & RTC - MHAMID</h2></div>', unsafe_allow_html=True)

# Navigation
page = st.radio(
    "Navigation",
    ["📝 INSTANCES", "📊 RAPPORTS", "⚠️ DÉRANGEMENTS", "🔧 FIABILISATION", "⚖️ LITIGES"],
    horizontal=True,
    label_visibility="collapsed"
)

# ====================== FONCTION DÉTECTION ======================
def find_column(df, keywords):
    if df is None or df.empty:
        return None
    for col in df.columns:
        if any(k in str(col).lower() for k in keywords):
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

        motif = st.selectbox("Motif", [
            "Adresse erronée", "Client refuse installation", "Transport saturé",
            "PC saturé", "INJOINABLE", "Local fermé + injoignable",
            "Création PC", "ETUDE CREATION PC", "MSAN saturé", "Autre"
        ])
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
        st.warning("Impossible de charger le fichier ETAT FTTH RTC RTCL.xlsx")

# ====================== PAGE RAPPORTS ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")
    st.info("Les graphiques et statistiques seront affichés ici une fois les colonnes correctement détectées.")

else:
    st.info(f"Page **{page}** en cours de développement.")

st.caption("Connecté en tant que **admin** | Application MHAMID")
