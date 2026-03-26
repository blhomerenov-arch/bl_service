import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

# En-tête bleu comme sur ta photo
st.markdown("""
    <style>
    .header {background-color: #0E7CFF; color: white; padding: 15px; border-radius: 5px; text-align: center;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>🚀 Gestion Chantier Fibre & RTC - MHAMID</h2></div>', unsafe_allow_html=True)

# Navigation comme sur la photo
tabs = st.tabs(["RAPPORTS", "INSTANCES (9)", "DÉRANGEMENTS (0)", "FIABILISATION", "LITIGES"])

with tabs[1]:  # Onglet INSTANCES actif
    st.subheader("📝 Saisie du Motif Journalier")

    # Formulaire amélioré
    with st.form("saisie_motif", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            demande = st.text_input("Demande*", placeholder="000D740B", help="Numéro de la commande")
            nom = st.text_input("Nom")
            contact = st.text_input("Contact", placeholder="0666549488")
            telecopie = st.text_input("N° de Téléscopie*", placeholder="525311326")
        
        with col2:
            adresse = st.text_area("Adresse", height=80, placeholder="MHAMID Lotissement Mhamid 9 ...")
            date_reception = st.date_input("Date de réception", value=datetime.now().date())
            secteur = st.selectbox("Secteur", ["MHAMID", "BOUAAKAZ", "Province M'HAMID"])
            agent = st.selectbox("Agent", ["hamid", "SHAKHMAN", "Autre"])
        
        col3, col4 = st.columns(2)
        with col3:
            type_install = st.selectbox("Type d'installation", ["Installation Fixe"])
        with col4:
            categorie = st.selectbox("Catégorie", ["RTC DTL", "GPON", "GPON DFO", "RTC"])
        
        # Motif
        motifs_communs = [
            "Adresse erronée", "Client refuse installation", "Transport saturé", 
            "PC saturé", "INJOINABLE", "Local fermé + injoignable", 
            "Création PC", "ETUDE CREATION PC", "MSAN saturé", "Autre"
        ]
        motif_selection = st.selectbox("Motif", motifs_communs)
        
        if motif_selection == "Autre":
            motif_final = st.text_input("Précisez le motif")
        else:
            motif_final = motif_selection
        
        submitted = st.form_submit_button("✅ Valider et Enregistrer", type="primary", use_container_width=True)

    if submitted:
        if not demande or not telecopie or not motif_final:
            st.error("❌ Les champs Demande, N° Téléscopie et Motif sont obligatoires !")
        else:
            nouvelle_ligne = {
                "Date": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "Demande": demande,
                "Nom": nom,
                "Contact": contact,
                "Adresse": adresse,
                "Téléscopie": telecopie,
                "Motif": motif_final,
                "Secteur": secteur,
                "Agent": agent,
                "Catégorie": categorie,
                "Type": type_install
            }
            
            # Simulation de sauvegarde (à améliorer plus tard)
            st.success(f"✅ Motif enregistré avec succès pour la demande **{demande}**")
            st.balloons()

    # Tableau des instances
    st.subheader("📋 Liste des Instances")

    try:
        etat_df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        search = st.text_input("🔍 Rechercher par Demande, Nom ou Adresse")
        
        if search:
            mask = etat_df.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)
            display_df = etat_df[mask]
        else:
            display_df = etat_df
            
        st.dataframe(display_df, use_container_width=True, height=500)
        
    except Exception as e:
        st.error("Impossible de charger le fichier ETAT FTTH RTC RTCL.xlsx")

# Statistiques rapides
st.subheader("📊 Statistiques Rapides")
col1, col2, col3 = st.columns(3)
col1.metric("Nombre d'Instances", len(etat_df) if 'etat_df' in locals() else 0)
col2.metric("Motifs saisis aujourd'hui", "0")  # À compléter plus tard
col3.metric("Total Motifs", "0")

st.caption("Application développée pour le chantier MHAMID - Fibre & RTC")
