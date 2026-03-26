import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

st.title("🚀 Gestion Chantier Fibre & RTC - MHAMID")
st.markdown("**Application de saisie des motifs journaliers**")

# Chargement des fichiers Excel
@st.cache_data
def load_data():
    try:
        etat = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
    except:
        etat = pd.DataFrame()
    
    try:
        motif = pd.read_excel("MOTIF.xlsx", sheet_name="MOTIF")
    except:
        motif = pd.DataFrame()
    
    try:
        motif_total = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF")
    except:
        motif_total = pd.DataFrame()
    
    return etat, motif, motif_total

etat_df, motif_df, motif_total_df = load_data()

# Formulaire de saisie
st.subheader("📝 Saisie du Motif Journalier")

with st.form("saisie_motif", clear_on_submit=True):
    col1, col2 = st.columns(2)
    
    with col1:
        demande = st.text_input("Demande*", placeholder="Ex: 000D740B")
        nom = st.text_input("Nom")
        contact = st.text_input("Contact", placeholder="0666549488")
        telecopie = st.text_input("N° de Téléscopie*", placeholder="525311326")
    
    with col2:
        adresse = st.text_area("Adresse", height=100)
        date_reception = st.date_input("Date de réception", value=datetime.now().date())
        secteur = st.selectbox("Secteur", ["MHAMID", "BOUAAKAZ", "Province M'HAMID"])
        agent = st.selectbox("Agent", ["hamid", "SHAKHMAN"])
    
    type_install = st.selectbox("Type d'installation", ["Installation Fixe"])
    
    motifs_communs = [
        "Adresse erronée", "Client refuse installation", "Transport saturé", 
        "PC saturé", "INJOINABLE", "Local fermé", "Création PC", 
        "ETUDE CREATION PC", "MSAN saturé", "Autre"
    ]
    motif_selection = st.selectbox("Motif", motifs_communs)
    
    if motif_selection == "Autre":
        motif_libre = st.text_input("Précisez le motif")
        motif_final = motif_libre
    else:
        motif_final = motif_selection
    
    submitted = st.form_submit_button("✅ Valider et Enregistrer", type="primary")

if submitted:
    if not demande or not telecopie or not motif_final:
        st.error("❌ Demande, Téléscopie et Motif sont obligatoires !")
    else:
        nouvelle_ligne = {
            "Date": datetime.now().strftime("%d/%m/%Y"),
            "Demande": demande,
            "Nom": nom,
            "Contact": contact,
            "Adresse": adresse,
            "Téléscopie": telecopie,
            "Motif": motif_final,
            "Secteur": secteur,
            "Agent": agent
        }
        
        new_df = pd.DataFrame([nouvelle_ligne])
        motif_df = pd.concat([motif_df, new_df], ignore_index=True)
        motif_total_df = pd.concat([motif_total_df, new_df], ignore_index=True)
        
        st.success(f"✅ Motif enregistré pour la demande **{demande}**")
        st.balloons()

# Tableau des instances
st.subheader("📋 Liste des Instances")

if not etat_df.empty:
    search = st.text_input("🔍 Rechercher")
    if search:
        mask = etat_df.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
        display_df = etat_df[mask]
    else:
        display_df = etat_df
    st.dataframe(display_df, use_container_width=True)
else:
    st.info("Le fichier ETAT FTTH RTC RTCL.xlsx n'est pas encore chargé.")

# Statistiques
st.subheader("📊 Statistiques")
col1, col2, col3 = st.columns(3)
col1.metric("Instances", len(etat_df))
col2.metric("Motifs aujourd'hui", len(motif_df))
col3.metric("Total Motifs", len(motif_total_df))

st.caption("Note : Pour le moment les fichiers Excel ne sont pas dans le repo. L'application fonctionne mais sans les données réelles.")
