import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

st.markdown("""
    <style>
    .header {background-color: #0E7CFF; color: white; padding: 15px; border-radius: 8px; text-align: center; margin-bottom: 15px;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>Gestion Chantier Fibre & RTC - MHAMID</h2></div>', unsafe_allow_html=True)

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

    st.subheader("Liste des Instances")
    try:
        df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        st.dataframe(df, use_container_width=True, height=500)
    except:
        st.warning("Impossible de charger le fichier ETAT FTTH RTC RTCL.xlsx")

# ====================== PAGE RAPPORTS (version robuste) ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")

    try:
        etat_df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        
        # Chargement sécurisé du fichier Motif
        try:
            motif_df = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF")
        except:
            motif_df = pd.DataFrame()

        # KPIs
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Commandes", len(etat_df))
        col2.metric("Motifs Cumulés", len(motif_df))
        col3.metric("Délai Moyen", round(etat_df.get('Délai(j)', pd.Series([0])).mean(), 1))
        col4.metric("Commandes VA", len(etat_df[etat_df.get('Etat', pd.Series([])) == 'VA']))

        st.divider()

        st.subheader("Répartition des Motifs")

        # Recherche intelligente de la colonne "Motif"
        motif_column = None
        for col in motif_df.columns:
            if 'motif' in str(col).lower():
                motif_column = col
                break

        if motif_column and not motif_df.empty:
            motif_count = motif_df[motif_column].value_counts().head(15)

            # Graphique Barres
            fig1 = px.bar(
                x=motif_count.index, 
                y=motif_count.values,
                title="Top 15 des Motifs les plus fréquents",
                labels={"x": "Motif", "y": "Nombre"},
                color=motif_count.values,
                color_continuous_scale="blues"
            )
            fig1.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig1, use_container_width=True)

            # Graphique Camembert
            fig2 = px.pie(
                values=motif_count.values,
                names=motif_count.index,
                title="Répartition en pourcentage des motifs"
            )
            st.plotly_chart(fig2, use_container_width=True)

            # Tableau
            st.subheader("Détail des Motifs")
            summary = motif_df[motif_column].value_counts().reset_index()
            summary.columns = ['Motif', 'Nombre']
            st.dataframe(summary, use_container_width=True)

        else:
            st.info("Aucune colonne 'Motif' trouvée dans le fichier MOTIF TOTAL. Vérifiez le nom des colonnes.")

        # Par Secteur
        if not etat_df.empty:
            st.subheader("Commandes par Secteur")
            secteur_col = next((col for col in etat_df.columns if 'secteur' in str(col).lower()), None)
            if secteur_col:
                secteur_count = etat_df[secteur_col].value_counts()
                fig3 = px.bar(secteur_count, title="Nombre de commandes par Secteur")
                st.plotly_chart(fig3, use_container_width=True)

    except Exception as e:
        st.error(f"Erreur lors du chargement des données : {str(e)}")

else:
    st.subheader(page)
    st.info(f"Page **{page}** en cours de développement.")

st.caption("Application de gestion de chantier MHAMID - Fibre & RTC")
