import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

st.markdown("""
    <style>
    .header {background-color: #0E7CFF; color: white; padding: 15px; border-radius: 8px; text-align: center; margin-bottom: 15px;}
    .info-box {background-color: #f0f8ff; padding: 15px; border-radius: 8px; border-left: 5px solid #0E7CFF;}
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

# ====================== PAGE RAPPORTS - DÉTECTION ULTRA ROBUSTE ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")

    try:
        etat_df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        motif_df = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF")

        # ==================== KPIs ====================
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Commandes", len(etat_df))
        col2.metric("Total Motifs", len(motif_df))
        col3.metric("Délai Moyen", round(etat_df.get('Délai(j)', pd.Series([0])).mean(), 1))
        col4.metric("Commandes VA", len(etat_df[etat_df.get('Etat', pd.Series([])) == 'VA']))

        st.divider()

        # ==================== DÉTECTION MOTIF ====================
        st.subheader("Répartition des Motifs")

        motif_col = None
        motif_keywords = ['motif', 'detail motif', 'pc mauvais', 'adresse erron', 'refuse', 'saturé', 'injoinable', 'création', 'etude']

        for col in motif_df.columns:
            col_str = str(col).lower()
            if any(kw in col_str for kw in motif_keywords):
                motif_col = col
                break

        # Backup : prendre la première colonne texte non vide
        if not motif_col and not motif_df.empty:
            for col in motif_df.columns:
                if motif_df[col].dtype == "object" and motif_df[col].notna().sum() > 10:
                    motif_col = col
                    break

        if motif_col:
            motif_series = motif_df[motif_col].astype(str).str.strip()
            motif_series = motif_series[(motif_series != 'nan') & (motif_series != '')]
            motif_count = motif_series.value_counts().head(15)

            fig1 = px.bar(
                x=motif_count.index, 
                y=motif_count.values,
                title=f"Top 15 Motifs (Colonne : {motif_col})",
                labels={"x": "Motif", "y": "Nombre"},
                color=motif_count.values,
                color_continuous_scale="blues"
            )
            fig1.update_layout(xaxis_tickangle=-45, height=520)
            st.plotly_chart(fig1, use_container_width=True)

            fig2 = px.pie(values=motif_count.values, names=motif_count.index, title="Répartition % des Motifs")
            st.plotly_chart(fig2, use_container_width=True)

            st.subheader("Détail des Motifs")
            st.dataframe(motif_count.reset_index().rename(columns={motif_col: 'Motif', 0: 'Nombre'}), 
                        use_container_width=True)

            st.success(f"✅ Colonne Motif détectée : **{motif_col}**")
        else:
            st.warning("Impossible de détecter la colonne Motif")
            st.write("Colonnes disponibles :", motif_df.columns.tolist())

        st.divider()

        # ==================== DÉTECTION SECTEUR ====================
        st.subheader("Commandes par Secteur")

        secteur_col = None
        secteur_keywords = ['secteur', 'sector', 'mhami', 'bouaakaz', 'province']

        for col in etat_df.columns:
            col_str = str(col).lower()
            if any(kw in col_str for kw in secteur_keywords):
                secteur_col = col
                break

        if not secteur_col and not etat_df.empty:
            for col in etat_df.columns:
                if etat_df[col].dtype == "object" and etat_df[col].notna().sum() > 5:
                    secteur_col = col
                    break

        if secteur_col:
            secteur_count = etat_df[secteur_col].value_counts()
            fig3 = px.bar(
                x=secteur_count.index, 
                y=secteur_count.values,
                title=f"Commandes par Secteur (Colonne : {secteur_col})",
                labels={"x": "Secteur", "y": "Nombre"}
            )
            st.plotly_chart(fig3, use_container_width=True)

            st.success(f"✅ Colonne Secteur détectée : **{secteur_col}**")
        else:
            st.warning("Impossible de détecter la colonne Secteur")
            st.write("Colonnes disponibles dans ETAT :", etat_df.columns.tolist())

    except Exception as e:
        st.error(f"Erreur : {str(e)}")

else:
    st.subheader(page)
    st.info(f"Page **{page}** en cours de développement.")

st.caption("Application de gestion de chantier MHAMID - Fibre & RTC")
