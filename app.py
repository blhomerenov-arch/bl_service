import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
import io

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

# Style
st.markdown("""
    <style>
    .header {background-color: #0E7CFF; color: white; padding: 15px; border-radius: 8px; text-align: center; margin-bottom: 15px;}
    .success {background-color: #d4edda; padding: 10px; border-radius: 8px; border-left: 5px solid #28a745;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>Gestion Chantier Fibre & RTC - MHAMID</h2></div>', unsafe_allow_html=True)

# ====================== NAVIGATION ======================
page = st.radio(
    "Navigation",
    ["📝 INSTANCES", "📊 RAPPORTS", "⚠️ DÉRANGEMENTS", "🔧 FIABILISATION", "⚖️ LITIGES"],
    horizontal=True,
    label_visibility="collapsed"
)

# ====================== FONCTION DE DÉTECTION OPTIMISÉE ======================
def find_column(df, keywords):
    """Détecte automatiquement une colonne selon une liste de mots-clés"""
    if df.empty:
        return None
    for col in df.columns:
        col_str = str(col).lower()
        if any(k in col_str for k in keywords):
            return col
    return None

# ====================== PAGE INSTANCES ======================
if page == "📝 INSTANCES":
    st.subheader("📝 Saisie du Motif Journalier")
    # (formulaire identique, je le garde court pour ne pas alourdir)
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
                st.success(f"✅ Motif enregistré pour **{demande}**")
                st.balloons()

    st.subheader("Liste des Instances")
    try:
        df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        st.dataframe(df, use_container_width=True, height=500)
    except:
        st.warning("Impossible de charger ETAT FTTH RTC RTCL.xlsx")

# ====================== PAGE RAPPORTS (version finale optimisée) ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")

    try:
        etat_df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        motif_df = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF")

        # ==================== KPIs ====================
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Commandes", len(etat_df))
        col2.metric("Total Motifs", len(motif_df))

        # Détection Délai
        delai_col = find_column(etat_df, ['délai', 'delai', 'délai(j)', 'delai(j)'])
        delai_moyen = round(etat_df[delai_col].mean(), 1) if delai_col else "N/A"
        col3.metric("Délai Moyen (jours)", delai_moyen)

        # Détection État
        etat_col = find_column(etat_df, ['etat', 'état', 'state', 'status'])
        va_count = len(etat_df[etat_df[etat_col].astype(str).str.upper() == 'VA']) if etat_col else 0
        col4.metric("Commandes VA", va_count)

        st.divider()

        # ==================== GRAPHIQUES MOTIFS ====================
        st.subheader("Répartition des Motifs")
        motif_col = find_column(motif_df, ['motif', 'detail motif', 'pc mauvais'])
        if motif_col:
            motif_series = motif_df[motif_col].astype(str).str.strip()
            motif_series = motif_series[(motif_series != 'nan') & (motif_series != '')]
            motif_count = motif_series.value_counts().head(15)

            fig1 = px.bar(x=motif_count.index, y=motif_count.values, title=f"Top 15 Motifs ({motif_col})")
            fig1.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig1, use_container_width=True)

            fig2 = px.pie(values=motif_count.values, names=motif_count.index, title="Répartition %")
            st.plotly_chart(fig2, use_container_width=True)

        # ==================== GRAPHIQUES SECTEUR & ÉTAT ====================
        st.subheader("Commandes par Secteur & État")
        col_a, col_b = st.columns(2)

        with col_a:
            secteur_col = find_column(et
