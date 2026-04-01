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
    page_title="SGIORN - Système de Gestion MHAMID",
    layout="wide",
    page_icon="🏗️",
    initial_sidebar_state="expanded"
)

# ====================== FICHIERS DE CONFIGURATION ======================
FICHIER_CONFIG = "config_app.json"

def charger_config():
    """Charge la configuration (agents, secteurs, etc.)"""
    try:
        if os.path.exists(FICHIER_CONFIG):
            with open(FICHIER_CONFIG, 'r', encoding='utf-8') as f:
                return json.load(f)
        else:
            return {
                "agents": ["hamid", "SHAKHMAN"],
                "secteurs": ["MHAMID", "BOUAAKAZ", "Province M'HAMID"],
                "nom_application": "Système de Gestion Intégré des Opérations Réseau Fibre & RTC",
                "sous_titre": "Région MHAMID - Maroc Telecom",
                "version": "3.1"
            }
    except:
        return {
            "agents": ["hamid", "SHAKHMAN"],
            "secteurs": ["MHAMID", "BOUAAKAZ", "Province M'HAMID"],
            "nom_application": "Système de Gestion Intégré des Opérations Réseau Fibre & RTC",
            "sous_titre": "Région MHAMID - Maroc Telecom",
            "version": "3.1"
        }

def sauvegarder_config(config):
    """Sauvegarde la configuration"""
    try:
        with open(FICHIER_CONFIG, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        st.error(f"Erreur sauvegarde: {str(e)}")
        return False

# Charger la config au démarrage
if 'config' not in st.session_state:
    st.session_state.config = charger_config()

config = st.session_state.config

# ====================== STYLES CSS ======================
st.markdown("""
    <style>
    .header {
        background: linear-gradient(135deg, #0E7CFF 0%, #0056b3 100%);
        color: white;
        padding: 25px;
        border-radius: 10px;
        text-align: center;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .header h1 {
        margin: 0;
        font-size: 28px;
        font-weight: bold;
    }
    .header p {
        margin: 5px 0 0 0;
        font-size: 14px;
        opacity: 0.9;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .metric-card {
        background: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        border-left: 4px solid #0E7CFF;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown(f'''
    <div class="header">
        <h1>🏗️ {config["nom_application"]}</h1>
        <p>{config["sous_titre"]} • Version {config["version"]}</p>
    </div>
''', unsafe_allow_html=True)

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
         "🔧 FIABILISATION", "⚖️ LITIGES", "⚙️ CONFIGURATION"],
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
    
    # Infos système
    st.subheader("ℹ️ Système")
    st.info(f"**Agents:** {len(config['agents'])}\n\n**Secteurs:** {len(config['secteurs'])}")
    st.caption(f"📅 {datetime.now().strftime('%d/%m/%Y')}")
    st.caption(f"🕐 {datetime.now().strftime('%H:%M')}")

# ====================== PAGE CONFIGURATION ======================
if page == "⚙️ CONFIGURATION":
    st.title("⚙️ Configuration du Système")
    
    tab1, tab2, tab3 = st.tabs(["👥 Gestion des Agents", "📍 Gestion des Secteurs", "🎨 Personnalisation"])
    
    # ==================== GESTION DES AGENTS ====================
    with tab1:
        st.subheader("👥 Gestion des Agents / Utilisateurs")
        
        col_agent1, col_agent2 = st.columns([2, 1])
        
        with col_agent1:
            st.markdown("### 📋 Liste des Agents")
            
            if config["agents"]:
                # Afficher les agents existants avec option de suppression
                for i, agent in enumerate(config["agents"]):
                    col_a, col_b = st.columns([3, 1])
                    with col_a:
                        st.info(f"👤 **{agent}**")
                    with col_b:
                        if st.button("🗑️ Suppr.", key=f"del_agent_{i}"):
                            config["agents"].remove(agent)
                            if sauvegarder_config(config):
                                st.success(f"✅ Agent {agent} supprimé")
                                st.rerun()
            else:
                st.warning("⚠️ Aucun agent configuré")
        
        with col_agent2:
            st.markdown("### ➕ Ajouter un Agent")
            
            with st.form("form_ajout_agent"):
                nouveau_agent = st.text_input("Nom de l'agent*", placeholder="Ex: Mohammed")
                
                if st.form_submit_button("✅ Ajouter l'Agent", type="primary", use_container_width=True):
                    if nouveau_agent:
                        if nouveau_agent not in config["agents"]:
                            config["agents"].append(nouveau_agent)
                            if sauvegarder_config(config):
                                st.success(f"✅ Agent **{nouveau_agent}** ajouté avec succès !")
                                st.balloons()
                                st.rerun()
                        else:
                            st.error("❌ Cet agent existe déjà !")
                    else:
                        st.error("❌ Le nom de l'agent est obligatoire")
        
        st.markdown("---")
        
        # Statistiques
        st.subheader("📊 Statistiques Agents")
        col_stat1, col_stat2, col_stat3 = st.columns(3)
        
        with col_stat1:
            st.metric("👥 Total Agents", len(config["agents"]))
        
        with col_stat2:
            # Compter instances par agent
            df_instances = charger_excel(FICHIER_INSTANCES)
            if not df_instances.empty and 'Agent' in df_instances.columns:
                agent_actif = df_instances['Agent'].value_counts().index[0] if len(df_instances) > 0 else "N/A"
                st.metric("🏆 Plus Actif", agent_actif)
            else:
                st.metric("🏆 Plus Actif", "N/A")
        
        with col_stat3:
            if not df_instances.empty:
                st.metric("📝 Total Instances", len(df_instances))
            else:
                st.metric("📝 Total Instances", "0")
    
    # ==================== GESTION DES SECTEURS ====================
    with tab2:
        st.subheader("📍 Gestion des Secteurs Géographiques")
        
        col_secteur1, col_secteur2 = st.columns([2, 1])
        
        with col_secteur1:
            st.markdown("### 📋 Liste des Secteurs")
            
            if config["secteurs"]:
                for i, secteur in enumerate(config["secteurs"]):
                    col_s, col_d = st.columns([3, 1])
                    with col_s:
                        st.success(f"📍 **{secteur}**")
                    with col_d:
                        if st.button("🗑️ Suppr.", key=f"del_secteur_{i}"):
                            config["secteurs"].remove(secteur)
                            if sauvegarder_config(config):
                                st.success(f"✅ Secteur {secteur} supprimé")
                                st.rerun()
            else:
                st.warning("⚠️ Aucun secteur configuré")
        
        with col_secteur2:
            st.markdown("### ➕ Ajouter un Secteur")
            
            with st.form("form_ajout_secteur"):
                nouveau_secteur = st.text_input("Nom du secteur*", placeholder="Ex: Zagora")
                
                if st.form_submit_button("✅ Ajouter le Secteur", type="primary", use_container_width=True):
                    if nouveau_secteur:
                        if nouveau_secteur not in config["secteurs"]:
                            config["secteurs"].append(nouveau_secteur)
                            if sauvegarder_config(config):
                                st.success(f"✅ Secteur **{nouveau_secteur}** ajouté avec succès !")
                                st.balloons()
                                st.rerun()
                        else:
                            st.error("❌ Ce secteur existe déjà !")
                    else:
                        st.error("❌ Le nom du secteur est obligatoire")
        
        st.markdown("---")
        
        # Statistiques secteurs
        st.subheader("📊 Statistiques par Secteur")
        
        df_instances = charger_excel(FICHIER_INSTANCES)
        if not df_instances.empty and 'Secteur' in df_instances.columns:
            secteur_count = df_instances['Secteur'].value_counts()
            
            col_graph1, col_graph2 = st.columns(2)
            
            with col_graph1:
                fig1 = px.bar(x=secteur_count.index, y=secteur_count.values, 
                             title="📊 Instances par Secteur",
                             labels={'x': 'Secteur', 'y': 'Nombre'},
                             color=secteur_count.values,
                             color_continuous_scale='Blues')
                st.plotly_chart(fig1, use_container_width=True)
            
            with col_graph2:
                fig2 = px.pie(values=secteur_count.values, names=secteur_count.index,
                             title="🥧 Répartition Géographique", hole=0.4)
                st.plotly_chart(fig2, use_container_width=True)
        else:
            st.info("ℹ️ Aucune donnée disponible pour les statistiques")
    
    # ==================== PERSONNALISATION ====================
    with tab3:
        st.subheader("🎨 Personnalisation de l'Application")
        
        with st.form("form_personnalisation"):
            st.markdown("### 📝 Informations de l'Application")
            
            nom_app = st.text_input(
                "Nom de l'application*",
                value=config["nom_application"],
                help="Nom principal affiché en haut de l'application"
            )
            
            sous_titre = st.text_input(
                "Sous-titre",
                value=config["sous_titre"],
                help="Description ou localisation"
            )
            
            version = st.text_input(
                "Version",
                value=config["version"],
                help="Numéro de version (ex: 3.1)"
            )
            
            st.markdown("### 🎨 Suggestions de Noms")
            
            suggestions = [
                "Système de Gestion Intégré des Opérations Réseau Fibre & RTC",
                "Plateforme de Supervision des Infrastructures Télécoms MHAMID",
                "Centre de Gestion et Maintenance Réseau Fibre Optique",
                "Système Intelligent de Pilotage des Opérations Télécom",
                "Plateforme Unifiée de Gestion des Services Réseau FTTH/RTC"
            ]
            
            for sugg in suggestions:
                if st.checkbox(sugg, key=sugg):
                    nom_app = sugg
            
            if st.form_submit_button("💾 Sauvegarder les Modifications", type="primary", use_container_width=True):
                if nom_app:
                    config["nom_application"] = nom_app
                    config["sous_titre"] = sous_titre
                    config["version"] = version
                    
                    if sauvegarder_config(config):
                        st.success("✅ Modifications sauvegardées avec succès !")
                        st.info("🔄 Actualisez la page pour voir les changements")
                        st.balloons()
                else:
                    st.error("❌ Le nom de l'application est obligatoire")
        
        st.markdown("---")
        
        # Export/Import de configuration
        st.subheader("💾 Sauvegarde & Restauration")
        
        col_exp1, col_exp2 = st.columns(2)
        
        with col_exp1:
            st.markdown("#### 📤 Exporter la Configuration")
            config_json = json.dumps(config, indent=4, ensure_ascii=False)
            st.download_button(
                "📥 Télécharger config.json",
                config_json,
                file_name=f"config_mhamid_{datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json",
                use_container_width=True
            )
        
        with col_exp2:
            st.markdown("#### 📥 Importer une Configuration")
            uploaded_config = st.file_uploader("Choisir un fichier config.json", type=['json'])
            
            if uploaded_config:
                try:
                    new_config = json.load(uploaded_config)
                    if st.button("✅ Appliquer cette Configuration", type="primary"):
                        st.session_state.config = new_config
                        if sauvegarder_config(new_config):
                            st.success("✅ Configuration importée avec succès !")
                            st.rerun()
                except Exception as e:
                    st.error(f"❌ Erreur: {str(e)}")

# ====================== PAGE DASHBOARD ======================
elif page == "🏠 DASHBOARD":
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
            
            # Utiliser les secteurs configurés
            secteur = st.selectbox("Secteur", config["secteurs"] + ["Autre"])
            if secteur == "Autre":
                secteur = st.text_input("Précisez le secteur")
            
            # Utiliser les agents configurés
            agent = st.selectbox("Agent", config["agents"] + ["Autre"])
            if agent == "Autre":
                agent = st.text_input("Précisez l'agent")

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

# ====================== AUTRES PAGES (identiques au code précédent) ======================
# ... (Gardez le reste du code pour RAPPORTS, DÉRANGEMENTS, FIABILISATION, LITIGES)

elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")
    # ... (code identique à la version précédente)
    st.info("Page Rapports - Code identique à la version précédente")

elif page == "⚠️ DÉRANGEMENTS":
    st.subheader("⚠️ Gestion des Dérangements")
    # ... (code identique, mais utilisez config["agents"] et config["secteurs"])
    st.info("Page Dérangements - À compléter avec le code précédent")

elif page == "🔧 FIABILISATION":
    st.subheader("🔧 Programme de Fiabilisation")
    # ... (code identique, mais utilisez config["agents"] et config["secteurs"])
    st.info("Page Fiabilisation - À compléter avec le code précédent")

elif page == "⚖️ LITIGES":
    st.subheader("⚖️ Gestion des Litiges")
    # ... (code identique, mais utilisez config["agents"])
    st.info("Page Litiges - À compléter avec le code précédent")

# ====================== FOOTER ======================
st.markdown("---")
col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
with col_f1:
    st.caption(f"🏗️ {config['nom_application']}")
with col_f2:
    st.caption(f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M')}")
with col_f3:
    st.caption(f"v{config['version']} - Édition Professionnelle")
