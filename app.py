import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
import io  # ← Ajouté ici

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

# ====================== AUTHENTIFICATION ======================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔐 Connexion à l'application")
    username = st.text_input("Nom d'utilisateur")
    password = st.text_input("Mot de passe", type="password")
   
    if st.button("Se connecter"):
        if username == "admin" and password == "1234":  # Change le mot de passe !
            st.session_state.authenticated = True
            st.success("Connexion réussie !")
            st.rerun()
        else:
            st.error("Identifiants incorrects")
    st.stop()

# ====================== STYLE ======================
st.markdown("""
    <style>
    .header {background-color: #0E7CFF; color: white; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px;}
    .success {background-color: #d4edda; padding: 12px; border-radius: 8px;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>🚧 Gestion Chantier Fibre & RTC - MHAMID</h2></div>', unsafe_allow_html=True)

# ====================== FONCTION DÉTECTION COLONNES ======================
def find_column(df, keywords):
    if df is None or df.empty:
        return None
    for col in df.columns:
        col_str = str(col).lower()
        if any(k in col_str for k in [k.lower() for k in keywords]):
            return col
    return None

# ====================== INITIALISATION SESSION STATE ======================
if "instances" not in st.session_state:
    st.session_state.instances = pd.DataFrame(columns=[
        "Demande", "Nom", "Contact", "Adresse", "Téléscopie", 
        "Date Réception", "Secteur", "Agent", "Motif", "Date Saisie"
    ])

# ====================== NAVIGATION ======================
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
            demande = st.text_input("Demande*", placeholder="000D740B", help="Numéro de la demande")
            nom = st.text_input("Nom du client")
            contact = st.text_input("Contact")
            adresse = st.text_area("Adresse", height=100)
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
            motif = st.text_input("Précisez le motif *")

        submitted = st.form_submit_button("✅ Valider et Enregistrer", type="primary", use_container_width=True)

        if submitted:
            if demande and telecopie and motif:
                new_row = pd.DataFrame([{
                    "Demande": demande,
                    "Nom": nom,
                    "Contact": contact,
                    "Adresse": adresse,
                    "Téléscopie": telecopie,
                    "Date Réception": date_reception,
                    "Secteur": secteur,
                    "Agent": agent,
                    "Motif": motif,
                    "Date Saisie": datetime.now()
                }])
                
                st.session_state.instances = pd.concat([st.session_state.instances, new_row], ignore_index=True)
                
                st.success(f"✅ Motif enregistré pour la demande **{demande}**")
                st.balloons()
            else:
                st.error("❌ Les champs Demande, Téléscopie et Motif sont obligatoires")

    # Affichage des instances saisies
    st.subheader("📋 Instances saisies aujourd’hui")
    if not st.session_state.instances.empty:
        st.dataframe(st.session_state.instances, use_container_width=True, height=400)
        
        # Option d'export des instances saisies
        csv = st.session_state.instances.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="⬇️ Télécharger les instances (CSV)",
            data=csv,
            file_name=f"instances_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.info("Aucune instance saisie pour le moment.")

    st.divider()

    # Chargement du fichier Excel existant
    st.subheader("📂 Fichier de référence")
    try:
        df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        st.dataframe(df, use_container_width=True, height=500)
    except Exception as e:
        st.warning("Impossible de charger le fichier **ETAT FTTH RTC RTCL.xlsx**. Vérifie qu’il est dans le même dossier.")

# ====================== PAGE RAPPORTS ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")

    try:
        etat_df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        motif_df = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF")

        # Détection des colonnes
        motif_col = find_column(motif_df, ['motif', 'detail motif'])
        secteur_col = find_column(etat_df, ['secteur', 'sector'])
        etat_col = find_column(etat_df, ['etat', 'état', 'state'])
        delai_col = find_column(etat_df, ['délai', 'delai', 'délai(j)'])

        # KPIs
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Commandes", len(etat_df))
        col2.metric("Total Motifs", len(motif_df))
        col3.metric("Délai Moyen (jours)", 
                    round(etat_df[delai_col].mean(), 1) if delai_col and delai_col in etat_df.columns else "N/A")
        col4.metric("Commandes VA", 
                    len(etat_df[etat_df[etat_col].astype(str).str.upper().str.strip() == 'VA']) 
                    if etat_col else 0)

        st.divider()

        # Graphique Top Motifs
        if motif_col and motif_col in motif_df.columns:
            motif_count = motif_df[motif_col].astype(str).str.strip().value_counts().head(15)
            fig = px.bar(
                x=motif_count.index, 
                y=motif_count.values, 
                title=f"Top 15 Motifs les plus fréquents",
                labels={"x": "Motif", "y": "Nombre"}
            )
            fig.update_layout(xaxis_tickangle=-45, height=500)
            st.plotly_chart(fig, use_container_width=True)

        # Export complet
        if st.button("📄 Générer Rapport Complet"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                etat_df.to_excel(writer, sheet_name="Etat", index=False)
                motif_df.to_excel(writer, sheet_name="Motifs", index=False)
                if not st.session_state.instances.empty:
                    st.session_state.instances.to_excel(writer, sheet_name="Instances Saisies", index=False)
            
            output.seek(0)
            st.download_button(
                label="⬇️ Télécharger le Rapport Excel",
                data=output,
                file_name=f"Rapport_Chantier_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des données : {str(e)}")
        st.info("Vérifie que les fichiers **ETAT FTTH RTC RTCL.xlsx** et **MOTIF TOTAL (1).xlsx** sont présents.")

# ====================== AUTRES PAGES ======================
else:
    st.info(f"Page **{page}** est en cours de développement. Elle sera disponible bientôt.")

st.caption("Application de gestion de chantier MHAMID - Fibre & RTC | Connecté en tant que **admin**")
