import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import os

# ====================== CONFIGURATION ======================
st.set_page_config(
    page_title="Gestion Chantier MHAMID", 
    layout="wide",
    page_icon="🏗️",
    initial_sidebar_state="expanded"
)

# ====================== STYLES CSS ======================
st.markdown("""
    <style>
    .header {
        background: linear-gradient(135deg, #0E7CFF 0%, #0056b3 100%);
        color: white;
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .success {
        background-color: #d4edda;
        padding: 12px;
        border-radius: 8px;
        border-left: 4px solid #28a745;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #dee2e6;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>🏗️ Gestion Chantier Fibre & RTC - MHAMID</h2></div>', unsafe_allow_html=True)

# ====================== CHEMINS FICHIERS ======================
FICHIER_INSTANCES = "instances_log.xlsx"
FICHIER_ETAT = "ETAT FTTH RTC RTCL.xlsx"
FICHIER_MOTIF = "MOTIF TOTAL (1).xlsx"
FICHIER_DERANGEMENTS = "derangements.xlsx"
FICHIER_FIABILISATION = "fiabilisation.xlsx"
FICHIER_LITIGES = "litiges.xlsx"

# ====================== FONCTIONS UTILITAIRES ======================
@st.cache_data(ttl=300)  # Cache de 5 minutes
def charger_excel(fichier, sheet_name=0):
    """Charge un fichier Excel avec gestion d'erreur"""
    try:
        if os.path.exists(fichier):
            return pd.read_excel(fichier, sheet_name=sheet_name)
        else:
            return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Erreur de chargement {fichier}: {str(e)}")
        return pd.DataFrame()

def sauvegarder_excel(df, fichier, sheet_name="Sheet1"):
    """Sauvegarde un DataFrame dans Excel"""
    try:
        if os.path.exists(fichier):
            with pd.ExcelWriter(fichier, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            df.to_excel(fichier, sheet_name=sheet_name, index=False)
        return True
    except Exception as e:
        st.error(f"❌ Erreur de sauvegarde: {str(e)}")
        return False

def find_column(df, keywords):
    """Détecte intelligemment une colonne par mots-clés"""
    if df is None or df.empty:
        return None
    for col in df.columns:
        col_str = str(col).lower()
        if any(k in col_str for k in keywords):
            return col
    # Backup : première colonne texte
    for col in df.columns:
        if df[col].dtype == "object" and df[col].notna().sum() > 5:
            return col
    return None

def export_dataframe_excel(df, nom_fichier):
    """Exporte un DataFrame en Excel téléchargeable"""
    from io import BytesIO
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Export')
    return buffer.getvalue()

# ====================== SIDEBAR ======================
with st.sidebar:
    st.image("https://via.placeholder.com/150x50/0E7CFF/FFFFFF?text=MHAMID", use_container_width=True)
    st.markdown("---")
    
    page = st.radio(
        "🧭 Navigation",
        ["📝 INSTANCES", "📊 RAPPORTS", "⚠️ DÉRANGEMENTS", "🔧 FIABILISATION", "⚖️ LITIGES"],
        label_visibility="visible"
    )
    
    st.markdown("---")
    st.info(f"**Date:** {datetime.now().strftime('%d/%m/%Y')}\n\n**Heure:** {datetime.now().strftime('%H:%M')}")
    
    # Stats rapides
    try:
        df_instances = charger_excel(FICHIER_INSTANCES)
        if not df_instances.empty:
            st.metric("📋 Total Instances", len(df_instances))
            st.metric("📅 Aujourd'hui", len(df_instances[df_instances['Date'].dt.date == datetime.now().date()]) if 'Date' in df_instances.columns else 0)
    except:
        pass

# ====================== PAGE INSTANCES ======================
if page == "📝 INSTANCES":
    st.subheader("📝 Saisie du Motif Journalier")
    
    # Charger les instances existantes
    df_instances = charger_excel(FICHIER_INSTANCES)
    if df_instances.empty:
        df_instances = pd.DataFrame(columns=[
            "Date", "Demande", "Nom", "Contact", "Adresse", "Telecopie",
            "Date_Reception", "Secteur", "Agent", "Motif", "Statut"
        ])

    with st.form("saisie_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            demande = st.text_input("N° Demande*", placeholder="000D740B", help="Numéro de demande obligatoire")
            nom = st.text_input("Nom du client")
            contact = st.text_input("Contact téléphonique", placeholder="0612345678")
            adresse = st.text_area("Adresse complète", height=80)

        with col2:
            telecopie = st.text_input("N° de Télécopie*", placeholder="525311326", help="Numéro obligatoire")
            date_reception = st.date_input("Date de réception", datetime.now().date())
            secteur = st.selectbox("Secteur", ["MHAMID", "BOUAAKAZ", "Province M'HAMID", "Autre"])
            agent = st.selectbox("Agent", ["hamid", "SHAKHMAN", "Autre"])

        motif_options = [
            "Adresse erronée",
            "Client refuse installation",
            "Transport saturé",
            "PC saturé",
            "INJOINABLE",
            "Local fermé + injoignable",
            "Création PC",
            "ETUDE CREATION PC",
            "MSAN saturé",
            "Câble défectueux",
            "Problème technique",
            "En attente matériel",
            "Autre"
        ]
        motif = st.selectbox("Motif de l'instance", motif_options)

        if motif == "Autre":
            motif = st.text_input("Précisez le motif")

        statut = st.selectbox("Statut", ["En cours", "Résolu", "En attente", "Annulé"])
        
        col_submit1, col_submit2 = st.columns([3, 1])
        with col_submit1:
            submitted = st.form_submit_button("✅ Valider et Enregistrer", type="primary", use_container_width=True)
        with col_submit2:
            clear = st.form_submit_button("🗑️ Effacer", use_container_width=True)

        if submitted:
            if demande and telecopie and motif:
                # Ajouter nouvelle ligne
                nouvelle_ligne = pd.DataFrame([{
                    "Date": datetime.now(),
                    "Demande": demande,
                    "Nom": nom,
                    "Contact": contact,
                    "Adresse": adresse,
                    "Telecopie": telecopie,
                    "Date_Reception": date_reception,
                    "Secteur": secteur,
                    "Agent": agent,
                    "Motif": motif,
                    "Statut": statut
                }])
                
                df_instances = pd.concat([df_instances, nouvelle_ligne], ignore_index=True)
                
                # Sauvegarder
                if sauvegarder_excel(df_instances, FICHIER_INSTANCES):
                    st.success(f"✅ Motif enregistré pour la demande **{demande}**")
                    st.balloons()
                    st.cache_data.clear()  # Rafraîchir le cache
                else:
                    st.error("❌ Erreur lors de la sauvegarde")
            else:
                st.error("❌ Demande, Télécopie et Motif sont obligatoires")

    st.markdown("---")
    
    # Affichage des instances avec filtres
    st.subheader("📋 Liste des Instances Enregistrées")
    
    if not df_instances.empty:
        col_filter1, col_filter2, col_filter3, col_filter4 = st.columns(4)
        
        with col_filter1:
            filtre_secteur = st.multiselect("Filtrer par Secteur", df_instances['Secteur'].unique() if 'Secteur' in df_instances.columns else [])
        with col_filter2:
            filtre_agent = st.multiselect("Filtrer par Agent", df_instances['Agent'].unique() if 'Agent' in df_instances.columns else [])
        with col_filter3:
            filtre_statut = st.multiselect("Filtrer par Statut", df_instances['Statut'].unique() if 'Statut' in df_instances.columns else [])
        with col_filter4:
            filtre_motif = st.multiselect("Filtrer par Motif", df_instances['Motif'].unique() if 'Motif' in df_instances.columns else [])

        # Appliquer les filtres
        df_filtre = df_instances.copy()
        if filtre_secteur:
            df_filtre = df_filtre[df_filtre['Secteur'].isin(filtre_secteur)]
        if filtre_agent:
            df_filtre = df_filtre[df_filtre['Agent'].isin(filtre_agent)]
        if filtre_statut:
            df_filtre = df_filtre[df_filtre['Statut'].isin(filtre_statut)]
        if filtre_motif:
            df_filtre = df_filtre[df_filtre['Motif'].isin(filtre_motif)]

        st.dataframe(
            df_filtre.sort_values('Date', ascending=False) if 'Date' in df_filtre.columns else df_filtre,
            use_container_width=True,
            height=400
        )
        
        # Export
        col_exp1, col_exp2, col_exp3 = st.columns([2, 1, 1])
        with col_exp2:
            excel_data = export_dataframe_excel(df_filtre)
            st.download_button(
                "📥 Télécharger Excel",
                excel_data,
                f"instances_{datetime.now().strftime('%Y%m%d')}.xlsx",
                "application/vnd.ms-excel",
                use_container_width=True
            )
        with col_exp3:
            csv = df_filtre.to_csv(index=False).encode('utf-8')
            st.download_button(
                "📥 Télécharger CSV",
                csv,
                f"instances_{datetime.now().strftime('%Y%m%d')}.csv",
                "text/csv",
                use_container_width=True
            )
    else:
        st.info("ℹ️ Aucune instance enregistrée pour le moment")

    # Afficher aussi le fichier ETAT FTTH
    st.markdown("---")
    st.subheader("📊 État FTTH RTC RTCL")
    try:
        df_etat = charger_excel(FICHIER_ETAT, sheet_name="SITUATION14.15")
        if not df_etat.empty:
            st.dataframe(df_etat, use_container_width=True, height=400)
        else:
            st.warning("⚠️ Fichier ETAT FTTH RTC RTCL.xlsx non trouvé")
    except Exception as e:
        st.warning(f"⚠️ Impossible de charger le fichier: {str(e)}")

# ====================== PAGE RAPPORTS ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques Détaillés")

    tab1, tab2, tab3 = st.tabs(["📈 Vue Générale", "📊 Analyse Motifs", "🗺️ Analyse Secteurs"])
    
    with tab1:
        try:
            etat_df = charger_excel(FICHIER_ETAT, sheet_name="SITUATION14.15")
            motif_df = charger_excel(FICHIER_MOTIF, sheet_name="MOTIF")

            if not etat_df.empty and not motif_df.empty:
                # Détection des colonnes
                motif_col = find_column(motif_df, ['motif', 'detail', 'pc mauvais'])
                secteur_col = find_column(etat_df, ['secteur', 'sector'])
                etat_col = find_column(etat_df, ['etat', 'état', 'state'])
                delai_col = find_column(etat_df, ['délai', 'delai'])

                st.success("✅ Fichiers chargés avec succès")

                # KPIs principaux
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric(
                        "📦 Total Commandes",
                        f"{len(etat_df):,}",
                        help="Nombre total de commandes"
                    )
                
                with col2:
                    st.metric(
                        "📝 Total Motifs",
                        f"{len(motif_df):,}",
                        help="Nombre total de motifs enregistrés"
                    )
                
                with col3:
                    if delai_col:
                        delai_moyen = round(etat_df[delai_col].mean(), 1)
                        st.metric(
                            "⏱️ Délai Moyen",
                            f"{delai_moyen} j",
                            help="Délai moyen en jours"
                        )
                    else:
                        st.metric("⏱️ Délai Moyen", "N/A")
                
                with col4:
                    if etat_col:
                        nb_va = len(etat_df[etat_df[etat_col].astype(str).str.upper() == 'VA'])
                        st.metric(
                            "✅ Commandes VA",
                            f"{nb_va:,}",
                            help="Commandes validées"
                        )
                    else:
                        st.metric("✅ Commandes VA", "0")

                st.markdown("---")

                # Graphique évolution temporelle (si date disponible)
                date_col = find_column(etat_df, ['date', 'Date'])
                if date_col:
                    st.subheader("📅 Évolution Temporelle")
                    etat_df[date_col] = pd.to_datetime(etat_df[date_col], errors='coerce')
                    evolution = etat_df.groupby(etat_df[date_col].dt.to_period('M')).size()
                    
                    fig_evolution = px.line(
                        x=evolution.index.astype(str),
                        y=evolution.values,
                        title="Évolution mensuelle des commandes",
                        labels={'x': 'Mois', 'y': 'Nombre de commandes'}
                    )
                    st.plotly_chart(fig_evolution, use_container_width=True)

            else:
                st.warning("⚠️ Fichiers non trouvés ou vides")

        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")

    with tab2:
        st.subheader("🔍 Analyse Détaillée des Motifs")
        
        try:
            motif_df = charger_excel(FICHIER_MOTIF, sheet_name="MOTIF")
            
            if not motif_df.empty:
                motif_col = find_column(motif_df, ['motif', 'detail', 'pc mauvais'])
                
                if motif_col:
                    st.info(f"📌 Colonne analysée : **{motif_col}**")
                    
                    # Nettoyage des données
                    motif_series = motif_df[motif_col].astype(str).str.strip()
                    motif_series = motif_series[(motif_series != 'nan') & (motif_series != '') & (motif_series != 'None')]
                    motif_count = motif_series.value_counts().head(15)

                    # Graphique en barres
                    col_chart1, col_chart2 = st.columns(2)
                    
                    with col_chart1:
                        fig1 = px.bar(
                            x=motif_count.index,
                            y=motif_count.values,
                            title="🏆 Top 15 des Motifs",
                            labels={'x': 'Motif', 'y': 'Nombre'},
                            color=motif_count.values,
                            color_continuous_scale='Blues'
                        )
                        fig1.update_layout(xaxis_tickangle=-45, showlegend=False)
                        st.plotly_chart(fig1, use_container_width=True)

                    with col_chart2:
                        fig2 = px.pie(
                            values=motif_count.values,
                            names=motif_count.index,
                            title="📊 Répartition des Motifs",
                            hole=0.4
                        )
                        st.plotly_chart(fig2, use_container_width=True)

                    # Tableau détaillé
                    st.subheader("📋 Détails des Motifs")
                    df_motifs_detail = pd.DataFrame({
                        'Motif': motif_count.index,
                        'Nombre': motif_count.values,
                        'Pourcentage': (motif_count.values / motif_count.sum() * 100).round(2)
                    })
                    st.dataframe(df_motifs_detail, use_container_width=True)

                    # Export
                    excel_motifs = export_dataframe_excel(df_motifs_detail)
                    st.download_button(
                        "📥 Télécharger l'analyse",
                        excel_motifs,
                        f"analyse_motifs_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        use_container_width=False
                    )
                else:
                    st.error("❌ Colonne 'Motif' introuvable")
            else:
                st.warning("⚠️ Fichier motifs vide")

        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")

    with tab3:
        st.subheader("🗺️ Analyse par Secteur")
        
        try:
            etat_df = charger_excel(FICHIER_ETAT, sheet_name="SITUATION14.15")
            
            if not etat_df.empty:
                secteur_col = find_column(etat_df, ['secteur', 'sector'])
                
                if secteur_col:
                    secteur_count = etat_df[secteur_col].value_counts()
                    
                    col_sect1, col_sect2 = st.columns(2)
                    
                    with col_sect1:
                        fig3 = px.bar(
                            x=secteur_count.index,
                            y=secteur_count.values,
                            title="📍 Commandes par Secteur",
                            labels={'x': 'Secteur', 'y': 'Nombre de commandes'},
                            color=secteur_count.values,
                            color_continuous_scale='Viridis'
                        )
                        st.plotly_chart(fig3, use_container_width=True)
                    
                    with col_sect2:
                        fig4 = px.pie(
                            values=secteur_count.values,
                            names=secteur_count.index,
                            title="🥧 Répartition Géographique"
                        )
                        st.plotly_chart(fig4, use_container_width=True)

                    # Statistiques détaillées par secteur
                    st.subheader("📊 Statistiques Détaillées")
                    for secteur in secteur_count.index[:5]:  # Top 5
                        with st.expander(f"📌 {secteur} ({secteur_count[secteur]} commandes)"):
                            df_secteur = etat_df[etat_df[secteur_col] == secteur]
                            
                            col_a, col_b, col_c = st.columns(3)
                            col_a.metric("Total", len(df_secteur))
                            
                            # Analyse des délais si disponible
                            delai_col = find_column(df_secteur, ['délai', 'delai'])
                            if delai_col:
                                col_b.metric("Délai Moyen", f"{df_secteur[delai_col].mean():.1f}j")
                                col_c.metric("Délai Max", f"{df_secteur[delai_col].max():.0f}j")
                else:
                    st.warning("⚠️ Colonne secteur introuvable")
            else:
                st.warning("⚠️ Fichier état vide")

        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")

# ====================== PAGE DÉRANGEMENTS ======================
elif page == "⚠️ DÉRANGEMENTS":
    st.subheader("⚠️ Gestion des Dérangements")
    
    # Charger les dérangements
    df_derangements = charger_excel(FICHIER_DERANGEMENTS)
    if df_derangements.empty:
        df_derangements = pd.DataFrame(columns=[
            "Date", "N_Ticket", "Client", "Adresse", "Type_Derangement",
            "Priorite", "Agent_Assigne", "Statut", "Date_Resolution", "Commentaire"
        ])

    tab1, tab2 = st.tabs(["📝 Nouveau Dérangement", "📋 Liste des Dérangements"])
    
    with tab1:
        with st.form("derangement_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                n_ticket = st.text_input("N° Ticket*", placeholder="DR-2024-001")
                client = st.text_input("Nom Client*")
                adresse = st.text_area("Adresse")
                type_derang = st.selectbox("Type de dérangement", [
                    "Pas de connexion",
                    "Connexion intermittente",
                    "Débit faible",
                    "Problème matériel",
                    "Câble endommagé",
                    "Autre"
                ])
            
            with col2:
                priorite = st.select_slider("Priorité", ["Basse", "Normale", "Haute", "Urgente"])
                agent = st.selectbox("Agent assigné", ["hamid", "SHAKHMAN", "Non assigné"])
                statut = st.selectbox("Statut", ["Nouveau", "En cours", "Résolu", "Fermé"])
                date_resolution = st.date_input("Date de résolution prévue")
            
            commentaire = st.text_area("Commentaire / Description détaillée")
            
            if st.form_submit_button("✅ Enregistrer le dérangement", type="primary", use_container_width=True):
                if n_ticket and client:
                    nouvelle_ligne = pd.DataFrame([{
                        "Date": datetime.now(),
                        "N_Ticket": n_ticket,
                        "Client": client,
                        "Adresse": adresse,
                        "Type_Derangement": type_derang,
                        "Priorite": priorite,
                        "Agent_Assigne": agent,
                        "Statut": statut,
                        "Date_Resolution": date_resolution,
                        "Commentaire": commentaire
                    }])
                    
                    df_derangements = pd.concat([df_derangements, nouvelle_ligne], ignore_index=True)
                    
                    if sauvegarder_excel(df_derangements, FICHIER_DERANGEMENTS):
                        st.success(f"✅ Dérangement **{n_ticket}** enregistré")
                        st.cache_data.clear()
                    else:
                        st.error("❌ Erreur de sauvegarde")
                else:
                    st.error("❌ N° Ticket et Client obligatoires")
    
    with tab2:
        if not df_derangements.empty:
            # Filtres
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                filtre_priorite = st.multiselect("Priorité", df_derangements['Priorite'].unique())
            with col_f2:
                filtre_statut = st.multiselect("Statut", df_derangements['Statut'].unique())
            with col_f3:
                filtre_agent = st.multiselect("Agent", df_derangements['Agent_Assigne'].unique())
            
            df_filtre = df_derangements.copy()
            if filtre_priorite:
                df_filtre = df_filtre[df_filtre['Priorite'].isin(filtre_priorite)]
            if filtre_statut:
                df_filtre = df_filtre[df_filtre['Statut'].isin(filtre_statut)]
            if filtre_agent:
                df_filtre = df_filtre[df_filtre['Agent_Assigne'].isin(filtre_agent)]
            
            # KPIs
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("Total", len(df_filtre))
            col_m2.metric("En cours", len(df_filtre[df_filtre['Statut'] == 'En cours']))
            col_m3.metric("Résolus", len(df_filtre[df_filtre['Statut'] == 'Résolu']))
            col_m4.metric("Urgents", len(df_filtre[df_filtre['Priorite'] == 'Urgente']))
            
            st.dataframe(df_filtre.sort_values('Date', ascending=False), use_container_width=True, height=400)
            
            # Export
            excel_data = export_dataframe_excel(df_filtre)
            st.download_button(
                "📥 Télécharger",
                excel_data,
                f"derangements_{datetime.now().strftime('%Y%m%d')}.xlsx",
                use_container_width=False
            )
        else:
            st.info("ℹ️ Aucun dérangement enregistré")

# ====================== PAGE FIABILISATION ======================
elif page == "🔧 FIABILISATION":
    st.subheader("🔧 Programme de Fiabilisation")
    
    df_fiab = charger_excel(FICHIER_FIABILISATION)
    if df_fiab.empty:
        df_fiab = pd.DataFrame(columns=[
            "Date", "Zone", "Type_Intervention", "PC_Concerne",
            "Agent", "Probleme_Detecte", "Action_Corrective", "Statut", "Date_Planifiee"
        ])

    tab1, tab2 = st.tabs(["📝 Nouvelle Intervention", "📋 Suivi des Interventions"])
    
    with tab1:
        with st.form("fiab_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                zone = st.selectbox("Zone", ["MHAMID", "BOUAAKAZ", "Province M'HAMID"])
                type_inter = st.selectbox("Type d'intervention", [
                    "Maintenance préventive",
                    "Renforcement réseau",
                    "Remplacement matériel",
                    "Nettoyage PC",
                    "Vérification câblage",
                    "Autre"
                ])
                pc_concerne = st.text_input("PC concerné")
                agent = st.selectbox("Agent responsable", ["hamid", "SHAKHMAN"])
            
            with col2:
                probleme = st.text_area("Problème détecté", height=100)
                action = st.text_area("Action corrective", height=100)
                statut = st.selectbox("Statut", ["Planifié", "En cours", "Terminé", "Reporté"])
                date_plan = st.date_input("Date planifiée")
            
            if st.form_submit_button("✅ Enregistrer l'intervention", type="primary", use_container_width=True):
                nouvelle_ligne = pd.DataFrame([{
                    "Date": datetime.now(),
                    "Zone": zone,
                    "Type_Intervention": type_inter,
                    "PC_Concerne": pc_concerne,
                    "Agent": agent,
                    "Probleme_Detecte": probleme,
                    "Action_Corrective": action,
                    "Statut": statut,
                    "Date_Planifiee": date_plan
                }])
                
                df_fiab = pd.concat([df_fiab, nouvelle_ligne], ignore_index=True)
                
                if sauvegarder_excel(df_fiab, FICHIER_FIABILISATION):
                    st.success("✅ Intervention enregistrée")
                    st.cache_data.clear()
    
    with tab2:
        if not df_fiab.empty:
            # Filtres
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                filtre_zone = st.multiselect("Zone", df_fiab['Zone'].unique())
            with col_f2:
                filtre_statut_fiab = st.multiselect("Statut", df_fiab['Statut'].unique())
            
            df_filtre = df_fiab.copy()
            if filtre_zone:
                df_filtre = df_filtre[df_filtre['Zone'].isin(filtre_zone)]
            if filtre_statut_fiab:
                df_filtre = df_filtre[df_filtre['Statut'].isin(filtre_statut_fiab)]
            
            # KPIs
            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("Total Interventions", len(df_filtre))
            col_m2.metric("En cours", len(df_filtre[df_filtre['Statut'] == 'En cours']))
            col_m3.metric("Terminées", len(df_filtre[df_filtre['Statut'] == 'Terminé']))
            
            st.dataframe(df_filtre.sort_values('Date', ascending=False), use_container_width=True, height=400)
            
            # Graphique par type
            if 'Type_Intervention' in df_filtre.columns:
                fig = px.bar(df_filtre['Type_Intervention'].value_counts(), title="Interventions par type")
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("ℹ️ Aucune intervention enregistrée")

# ====================== PAGE LITIGES ======================
elif page == "⚖️ LITIGES":
    st.subheader("⚖️ Gestion des Litiges")
    
    df_litiges = charger_excel(FICHIER_LITIGES)
    if df_litiges.empty:
        df_litiges = pd.DataFrame(columns=[
            "Date", "N_Litige", "Client", "Type_Litige", "Description",
            "Montant", "Statut", "Agent_Responsable", "Date_Resolution", "Commentaire"
        ])

    tab1, tab2 = st.tabs(["📝 Nouveau Litige", "📋 Suivi des Litiges"])
    
    with tab1:
        with st.form("litige_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            
            with col1:
                n_litige = st.text_input("N° Litige*", placeholder="LIT-2024-001")
                client = st.text_input("Client*")
                type_litige = st.selectbox("Type de litige", [
                    "Facturation incorrecte",
                    "Service non conforme",
                    "Retard d'installation",
                    "Dommages matériels",
                    "Désaccord contractuel",
                    "Autre"
                ])
                montant = st.number_input("Montant (DH)", min_value=0.0, step=100.0)
            
            with col2:
                description = st.text_area("Description détaillée", height=100)
                statut = st.selectbox("Statut", ["Ouvert", "En traitement", "Résolu", "Fermé", "Escaladé"])
                agent = st.selectbox("Agent responsable", ["hamid", "SHAKHMAN", "Direction"])
                date_resol = st.date_input("Date résolution prévue")
            
            commentaire = st.text_area("Commentaire / Action entreprise")
            
            if st.form_submit_button("✅ Enregistrer le litige", type="primary", use_container_width=True):
                if n_litige and client:
                    nouvelle_ligne = pd.DataFrame([{
                        "Date": datetime.now(),
                        "N_Litige": n_litige,
                        "Client": client,
                        "Type_Litige": type_litige,
                        "Description": description,
                        "Montant": montant,
                        "Statut": statut,
                        "Agent_Responsable": agent,
                        "Date_Resolution": date_resol,
                        "Commentaire": commentaire
                    }])
                    
                    df_litiges = pd.concat([df_litiges, nouvelle_ligne], ignore_index=True)
                    
                    if sauvegarder_excel(df_litiges, FICHIER_LITIGES):
                        st.success(f"✅ Litige **{n_litige}** enregistré")
                        st.cache_data.clear()
                else:
                    st.error("❌ N° Litige et Client obligatoires")
    
    with tab2:
        if not df_litiges.empty:
            # Filtres
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                filtre_type = st.multiselect("Type", df_litiges['Type_Litige'].unique())
            with col_f2:
                filtre_statut_lit = st.multiselect("Statut", df_litiges['Statut'].unique())
            
            df_filtre = df_litiges.copy()
            if filtre_type:
                df_filtre = df_filtre[df_filtre['Type_Litige'].isin(filtre_type)]
            if filtre_statut_lit:
                df_filtre = df_filtre[df_filtre['Statut'].isin(filtre_statut_lit)]
            
            # KPIs
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("Total Litiges", len(df_filtre))
            col_m2.metric("En traitement", len(df_filtre[df_filtre['Statut'] == 'En traitement']))
            col_m3.metric("Résolus", len(df_filtre[df_filtre['Statut'] == 'Résolu']))
            col_m4.metric("Montant Total", f"{df_filtre['Montant'].sum():,.0f} DH")
            
            st.dataframe(df_filtre.sort_values('Date', ascending=False), use_container_width=True, height=400)
            
            # Graphiques
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                fig1 = px.bar(df_filtre['Type_Litige'].value_counts(), title="Litiges par type")
                st.plotly_chart(fig1, use_container_width=True)
            with col_g2:
                fig2 = px.pie(df_filtre, names='Statut', title="Répartition par statut")
                st.plotly_chart(fig2, use_container_width=True)
            
            # Export
            excel_data = export_dataframe_excel(df_filtre)
            st.download_button(
                "📥 Télécharger",
                excel_data,
                f"litiges_{datetime.now().strftime('%Y%m%d')}.xlsx",
                use_container_width=False
            )
        else:
            st.info("ℹ️ Aucun litige enregistré")

# ====================== FOOTER ======================
st.markdown("---")
col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
with col_f1:
    st.caption("🏗️ Application MHAMID - Gestion Fibre & RTC")
with col_f2:
    st.caption(f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M')}")
with col_f3:
    st.caption("v2.0 - Améliorée")
