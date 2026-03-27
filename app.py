import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
import io

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

# ====================== AUTHENTIFICATION ======================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    st.title("🔐 Connexion à l'application")
    username = st.text_input("Nom d'utilisateur")
    password = st.text_input("Mot de passe", type="password")
   
    if st.button("Se connecter"):
        if username == "admin" and password == "1234":
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

# ====================== INITIALISATION ======================
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
            demande = st.text_input("Demande*", placeholder="000D740B")
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

        if st.form_submit_button("✅ Valider et Enregistrer", type="primary", use_container_width=True):
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

    st.subheader("📋 Instances saisies")
    if not st.session_state.instances.empty:
        st.dataframe(st.session_state.instances, use_container_width=True)
        csv = st.session_state.instances.to_csv(index=False).encode('utf-8')
        st.download_button("⬇️ Télécharger Instances (CSV)", csv, f"instances_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")
    else:
        st.info("Aucune instance saisie pour le moment.")

    try:
        df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        st.subheader("📂 Fichier de référence")
        st.dataframe(df, use_container_width=True, height=400)
    except:
        st.warning("Impossible de charger ETAT FTTH RTC RTCL.xlsx")

# ====================== PAGE RAPPORTS (CORRIGÉE) ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")

    try:
        etat_df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        motif_df = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Lignes", len(motif_df))
        col2.metric("Colonnes", len(motif_df.columns))

        st.divider()
        st.subheader("📈 Statistiques des Motifs")

        # Détection intelligente de la colonne motif
        motif_col = find_column(motif_df, ['detail motif', 'détail motif', 'motif', 'pc mauvais', 'code motif'])

        # Si pas trouvé, on prend la 6ème colonne (index 5)
        if not motif_col and len(motif_df.columns) > 5:
            motif_col = motif_df.columns[5]
            st.info(f"✅ Colonne utilisée automatiquement : **{motif_col}**")

        if motif_col and motif_col in motif_df.columns:
            motif_series = motif_df[motif_col].astype(str).str.strip()
            motif_series = motif_series[(motif_series != "") & (motif_series != "nan") & (motif_series != "None")]

            motif_count = motif_series.value_counts().head(15)

            if not motif_count.empty:
                # Graphique en barres
                st.subheader("📊 Top 15 des Motifs")
                fig_bar = px.bar(
                    x=motif_count.index,
                    y=motif_count.values,
                    title=f"Top 15 Motifs - {motif_col}",
                    labels={"x": "Motif", "y": "Nombre"},
                    text=motif_count.values
                )
                fig_bar.update_layout(xaxis_tickangle=-45, height=550, margin=dict(b=200))
                st.plotly_chart(fig_bar, use_container_width=True)

                # Graphique en cercle
                st.subheader("🥧 Répartition des Motifs")
                top10 = motif_series.value_counts().head(10)
                fig_pie = px.pie(
                    names=top10.index,
                    values=top10.values,
                    title="Répartition en pourcentage"
                )
                fig_pie.update_traces(textinfo='percent+label')
                st.plotly_chart(fig_pie, use_container_width=True)

                st.dataframe(
                    motif_count.reset_index().rename(columns={"index": "Motif", "count": "Nombre"}),
                    use_container_width=True
                )
            else:
                st.warning("La colonne des motifs existe mais est encore vide. Remplis-la dans ton fichier Excel.")
        else:
            st.error("Impossible de détecter la colonne des motifs.")
            st.write("Colonnes disponibles :", list(motif_df.columns))

    except Exception as e:
        st.error(f"Erreur de chargement des fichiers : {str(e)}")

# ====================== AUTRES PAGES ======================
else:
    st.info(f"Page **{page}** est en cours de développement.")

st.caption("Application Gestion Chantier MHAMID - Fibre & RTC | Connecté en tant que admin")
