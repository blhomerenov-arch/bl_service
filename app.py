import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
import os
from io import BytesIO
import json

# ====================== CONFIGURATION ======================
st.set_page_config(
    page_title="Gestion Chantier MHAMID v3.0",
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
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>🏗️ Gestion Chantier Fibre & RTC - MHAMID v3.0</h2></div>', unsafe_allow_html=True)

# ====================== CHEMINS FICHIERS ======================
FICHIER_INSTANCES = "instances_log.xlsx"
FICHIER_ETAT = "ETAT FTTH RTC RTCL.xlsx"
FICHIER_MOTIF = "MOTIF TOTAL (1).xlsx"
FICHIER_DERANGEMENTS = "derangements.xlsx"
FICHIER_FIABILISATION = "fiabilisation.xlsx"
FICHIER_LITIGES = "litiges.xlsx"

# ====================== FONCTIONS UTILITAIRES ======================
@st.cache_data(ttl=300)
def charger_excel(fichier, sheet_name=0):
    """Charge un fichier Excel avec gestion d'erreur"""
    try:
        if os.path.exists(fichier):
            return pd.read_excel(fichier, sheet_name=sheet_name, engine='openpyxl')
        else:
            return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Erreur de chargement {fichier}: {str(e)}")
        return pd.DataFrame()

def sauvegarder_excel(df, fichier, sheet_name="Sheet1"):
    """Sauvegarde un DataFrame dans Excel"""
    try:
        if os.path.exists(fichier):
            with pd.ExcelWriter(fichier, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        else:
            df.to_excel(fichier, sheet_name=sheet_name, index=False, engine='openpyxl')
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
    for col in df.columns:
        if df[col].dtype == "object" and df[col].notna().sum() > 5:
            return col
    return None

def export_excel(df, nom_fichier):
    """Exporte un DataFrame en Excel téléchargeable"""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Export')
    return buffer.getvalue()

# ====================== SIDEBAR ======================
with st.sidebar:
    st.markdown("### 🏗️ MHAMID")
    st.markdown("---")
    
    page = st.radio(
        "🧭 Navigation",
        ["🏠 DASHBOARD", "📝 INSTANCES", "📊 RAPPORTS", "⚠️ DÉRANGEMENTS", 
         "🔧 FIABILISATION", "⚖️ LITIGES"],
        label_visibility="visible"
    )
    
    st.markdown("---")
    
    # Stats rapides sidebar
    st.subheader("📊 Stats Rapides")
    try:
        df_instances = charger_excel(FICHIER_INSTANCES)
        df_derangements = charger_excel(FICHIER_DERANGEMENTS)
        
        if not df_instances.empty:
            st.metric("📋 Instances", len(df_instances))
        if not df_derangements.empty:
            st.metric("⚠️ Dérangements", len(df_derangements))
    except:
        pass
    
    st.markdown("---")
    st.info(f"**Date:** {datetime.now().strftime('%d/%m/%Y')}\n\n**Heure:** {datetime.now().strftime('%H:%M')}")

# ====================== PAGE DASHBOARD ======================
if page == "🏠 DASHBOARD":
    st.title("🏠 Dashboard Principal")
    
    # Charger toutes les données
    df_instances = charger_excel(FICHIER_INSTANCES)
    df_derangements = charger_excel(FICHIER_DERANGEMENTS)
    df_litiges = charger_excel(FICHIER_LITIGES)
    
    # KPIs Principaux
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("📋 Instances", len(df_instances) if not df_instances.empty else 0)
    
    with col2:
        st.metric("⚠️ Dérangements", len(df_derangements) if not df_derangements.empty else 0)
    
    with col3:
        st.metric("⚖️ Litiges", len(df_litiges) if not df_litiges.empty else 0)
    
    with col4:
        if not df_instances.empty and 'Statut' in df_instances.columns:
            taux = round(len(df_instances[df_instances['Statut'] == 'Résolu']) / len(df_instances) * 100, 1) if len(df_instances) > 0 else 0
            st.metric("✅ Taux Résolution", f"{taux}%")
        else:
            st.metric("✅ Taux Résolution", "N/A")
    
    st.markdown("---")
    
    # Graphiques
    if not df_instances.empty:
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            if 'Secteur' in df_instances.columns:
                secteur_count = df_instances['Secteur'].value_counts()
                fig1 = px.pie(values=secteur_count.values, names=secteur_count.index, 
                             title="📍 Instances par Secteur", hole=0.4)
                st.plotly_chart(fig1, use_container_width=True)
        
        with col_g2:
            if 'Motif' in df_instances.columns:
                motif_count = df_instances['Motif'].value_counts().head(5)
                fig2 = px.bar(x=motif_count.index, y=motif_count.values, 
                             title="🏆 Top 5 Motifs")
                st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("ℹ️ Aucune donnée à afficher. Commencez par saisir des instances.")

# ====================== PAGE INSTANCES ======================
elif page == "📝 INSTANCES":
    st.subheader("📝 Saisie du Motif Journalier")
    
    df_instances = charger_excel(FICHIER_INSTANCES)
    if df_instances.empty:
        df_instances = pd.DataFrame(columns=[
            "Date", "Demande", "Nom", "Contact", "Adresse", "Telecopie",
            "Date_Reception", "Secteur", "Agent", "Motif", "Statut"
        ])

    with st.form("saisie_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            demande = st.text_input("N° Demande*", placeholder="000D740B")
            nom = st.text_input("Nom du client")
            contact = st.text_input("Contact téléphonique", placeholder="0612345678")
            adresse = st.text_area("Adresse complète", height=80)

        with col2:
            telecopie = st.text_input("N° de Télécopie*", placeholder="525311326")
            date_reception = st.date_input("Date de réception", datetime.now().date())
            secteur = st.selectbox("Secteur", ["MHAMID", "BOUAAKAZ", "Province M'HAMID", "Autre"])
            agent = st.selectbox("Agent", ["hamid", "SHAKHMAN", "Autre"])

        motif_options = [
            "Adresse erronée", "Client refuse installation", "Transport saturé",
            "PC saturé", "INJOINABLE", "Local fermé + injoignable",
            "Création PC", "ETUDE CREATION PC", "MSAN saturé",
            "Câble défectueux", "Problème technique", "En attente matériel", "Autre"
        ]
        motif = st.selectbox("Motif de l'instance", motif_options)

        if motif == "Autre":
            motif = st.text_input("Précisez le motif")

        statut = st.selectbox("Statut", ["En cours", "Résolu", "En attente", "Annulé"])
        
        submitted = st.form_submit_button("✅ Valider et Enregistrer", type="primary", use_container_width=True)

        if submitted:
            if demande and telecopie and motif:
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
                
                if sauvegarder_excel(df_instances, FICHIER_INSTANCES):
                    st.success(f"✅ Motif enregistré pour la demande **{demande}**")
                    st.balloons()
                    st.cache_data.clear()
                else:
                    st.error("❌ Erreur lors de la sauvegarde")
            else:
                st.error("❌ Demande, Télécopie et Motif sont obligatoires")

    st.markdown("---")
    
    # Affichage avec filtres
    st.subheader("📋 Liste des Instances Enregistrées")
    
    if not df_instances.empty:
        col_filter1, col_filter2, col_filter3 = st.columns(3)
        
        with col_filter1:
            filtre_secteur = st.multiselect("Secteur", df_instances['Secteur'].unique() if 'Secteur' in df_instances.columns else [])
        with col_filter2:
            filtre_agent = st.multiselect("Agent", df_instances['Agent'].unique() if 'Agent' in df_instances.columns else [])
        with col_filter3:
            filtre_statut = st.multiselect("Statut", df_instances['Statut'].unique() if 'Statut' in df_instances.columns else [])

        df_filtre = df_instances.copy()
        if filtre_secteur:
            df_filtre = df_filtre[df_filtre['Secteur'].isin(filtre_secteur)]
        if filtre_agent:
            df_filtre = df_filtre[df_filtre['Agent'].isin(filtre_agent)]
        if filtre_statut:
            df_filtre = df_filtre[df_filtre['Statut'].isin(filtre_statut)]

        st.dataframe(
            df_filtre.sort_values('Date', ascending=False) if 'Date' in df_filtre.columns else df_filtre,
            use_container_width=True,
            height=400
        )
        
        # Export
        col_exp1, col_exp2 = st.columns([3, 1])
        with col_exp2:
            excel_data = export_excel(df_filtre)
            st.download_button(
                "📥 Télécharger Excel",
                excel_data,
                f"instances_{datetime.now().strftime('%Y%m%d')}.xlsx",
                use_container_width=True
            )
    else:
        st.info("ℹ️ Aucune instance enregistrée pour le moment")

# ====================== PAGE RAPPORTS ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")

    tab1, tab2 = st.tabs(["📈 Vue Générale", "📊 Analyse Motifs"])
    
    with tab1:
        try:
            etat_df = charger_excel(FICHIER_ETAT, sheet_name="SITUATION14.15")
            motif_df = charger_excel(FICHIER_MOTIF, sheet_name="MOTIF")

            if not etat_df.empty and not motif_df.empty:
                st.success("✅ Fichiers chargés avec succès")

                # KPIs
                col1, col2, col3, col4 = st.columns(4)
                
                col1.metric("📦 Total Commandes", f"{len(etat_df):,}")
                col2.metric("📝 Total Motifs", f"{len(motif_df):,}")
                
                delai_col = find_column(etat_df, ['délai', 'delai'])
                if delai_col:
                    col3.metric("⏱️ Délai Moyen", f"{round(etat_df[delai_col].mean(), 1)} j")
                else:
                    col3.metric("⏱️ Délai Moyen", "N/A")
                
                etat_col = find_column(etat_df, ['etat', 'état', 'state'])
                if etat_col:
                    nb_va = len(etat_df[etat_df[etat_col].astype(str).str.upper() == 'VA'])
                    col4.metric("✅ Commandes VA", f"{nb_va:,}")
                else:
                    col4.metric("✅ Commandes VA", "0")

                st.markdown("---")

                # Graphique Secteur
                secteur_col = find_column(etat_df, ['secteur', 'sector'])
                if secteur_col:
                    secteur_count = etat_df[secteur_col].value_counts()
                    fig = px.bar(x=secteur_count.index, y=secteur_count.values, 
                                title="📍 Commandes par Secteur")
                    st.plotly_chart(fig, use_container_width=True)

            else:
                st.warning("⚠️ Fichiers ETAT ou MOTIF non trouvés")

        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")

    with tab2:
        st.subheader("🔍 Analyse des Motifs")
        
        try:
            motif_df = charger_excel(FICHIER_MOTIF, sheet_name="MOTIF")
            
            if not motif_df.empty:
                motif_col = find_column(motif_df, ['motif', 'detail', 'pc mauvais'])
                
                if motif_col:
                    st.info(f"📌 Colonne analysée : **{motif_col}**")
                    
                    motif_series = motif_df[motif_col].astype(str).str.strip()
                    motif_series = motif_series[(motif_series != 'nan') & (motif_series != '') & (motif_series != 'None')]
                    motif_count = motif_series.value_counts().head(15)

                    col_chart1, col_chart2 = st.columns(2)
                    
                    with col_chart1:
                        fig1 = px.bar(x=motif_count.index, y=motif_count.values,
                                     title="🏆 Top 15 des Motifs",
                                     color=motif_count.values,
                                     color_continuous_scale='Blues')
                        fig1.update_layout(xaxis_tickangle=-45, showlegend=False)
                        st.plotly_chart(fig1, use_container_width=True)

                    with col_chart2:
                        fig2 = px.pie(values=motif_count.values, names=motif_count.index,
                                     title="📊 Répartition des Motifs", hole=0.4)
                        st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.error("❌ Colonne motif introuvable")
            else:
                st.warning("⚠️ Fichier motifs vide")

        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")

# ====================== PAGE DÉRANGEMENTS ======================
elif page == "⚠️ DÉRANGEMENTS":
    st.subheader("⚠️ Gestion des Dérangements")
    
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
                    "Pas de connexion", "Connexion intermittente", "Débit faible",
                    "Problème matériel", "Câble endommagé", "Autre"
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
            # KPIs
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("Total", len(df_derangements))
            col_m2.metric("En cours", len(df_derangements[df_derangements['Statut'] == 'En cours']) if 'Statut' in df_derangements.columns else 0)
            col_m3.metric("Résolus", len(df_derangements[df_derangements['Statut'] == 'Résolu']) if 'Statut' in df_derangements.columns else 0)
            col_m4.metric("Urgents", len(df_derangements[df_derangements['Priorite'] == 'Urgente']) if 'Priorite' in df_derangements.columns else 0)
            
            st.dataframe(df_derangements.sort_values('Date', ascending=False) if 'Date' in df_derangements.columns else df_derangements, 
                        use_container_width=True, height=400)
            
            # Export
            excel_data = export_excel(df_derangements)
            st.download_button(
                "📥 Télécharger Excel",
                excel_data,
                f"derangements_{datetime.now().strftime('%Y%m%d')}.xlsx"
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
                    "Maintenance préventive", "Renforcement réseau", "Remplacement matériel",
                    "Nettoyage PC", "Vérification câblage", "Autre"
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
            col_m1, col_m2, col_m3 = st.columns(3)
            col_m1.metric("Total Interventions", len(df_fiab))
            col_m2.metric("En cours", len(df_fiab[df_fiab['Statut'] == 'En cours']) if 'Statut' in df_fiab.columns else 0)
            col_m3.metric("Terminées", len(df_fiab[df_fiab['Statut'] == 'Terminé']) if 'Statut' in df_fiab.columns else 0)
            
            st.dataframe(df_fiab.sort_values('Date', ascending=False) if 'Date' in df_fiab.columns else df_fiab, 
                        use_container_width=True, height=400)
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
                    "Facturation incorrecte", "Service non conforme", "Retard d'installation",
                    "Dommages matériels", "Désaccord contractuel", "Autre"
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
            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("Total Litiges", len(df_litiges))
            col_m2.metric("En traitement", len(df_litiges[df_litiges['Statut'] == 'En traitement']) if 'Statut' in df_litiges.columns else 0)
            col_m3.metric("Résolus", len(df_litiges[df_litiges['Statut'] == 'Résolu']) if 'Statut' in df_litiges.columns else 0)
            col_m4.metric("Montant Total", f"{df_litiges['Montant'].sum():,.0f} DH" if 'Montant' in df_litiges.columns else "0 DH")
            
            st.dataframe(df_litiges.sort_values('Date', ascending=False) if 'Date' in df_litiges.columns else df_litiges, 
                        use_container_width=True, height=400)
            
            # Graphiques
            if 'Type_Litige' in df_litiges.columns:
                fig = px.bar(df_litiges['Type_Litige'].value_counts(), title="Litiges par type")
                st.plotly_chart(fig, use_container_width=True)
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
    st.caption("v3.0 - Edition Corrigée")
