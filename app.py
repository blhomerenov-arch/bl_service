import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path
import os
from io import BytesIO
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
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
    .success {
        background-color: #d4edda;
        padding: 12px;
        border-radius: 8px;
        border-left: 4px solid #28a745;
    }
    .warning {
        background-color: #fff3cd;
        padding: 12px;
        border-radius: 8px;
        border-left: 4px solid #ffc107;
    }
    .danger {
        background-color: #f8d7da;
        padding: 12px;
        border-radius: 8px;
        border-left: 4px solid #dc3545;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #dee2e6;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>🏗️ Gestion Chantier Fibre & RTC - MHAMID v3.0</h2><p>Avec Alertes Email & Statistiques Avancées</p></div>', unsafe_allow_html=True)

# ====================== CHEMINS FICHIERS ======================
FICHIER_INSTANCES = "instances_log.xlsx"
FICHIER_ETAT = "ETAT FTTH RTC RTCL.xlsx"
FICHIER_MOTIF = "MOTIF TOTAL (1).xlsx"
FICHIER_DERANGEMENTS = "derangements.xlsx"
FICHIER_FIABILISATION = "fiabilisation.xlsx"
FICHIER_LITIGES = "litiges.xlsx"
FICHIER_CONFIG_EMAIL = "config_email.json"
FICHIER_ALERTES = "alertes_config.json"

# ====================== CONFIGURATION EMAIL ======================
def charger_config_email():
    """Charge la configuration email depuis un fichier JSON"""
    try:
        if os.path.exists(FICHIER_CONFIG_EMAIL):
            with open(FICHIER_CONFIG_EMAIL, 'r') as f:
                return json.load(f)
        else:
            return {
                "smtp_server": "smtp.gmail.com",
                "smtp_port": 587,
                "email_expediteur": "",
                "password": "",
                "emails_destinataires": []
            }
    except:
        return {}

def sauvegarder_config_email(config):
    """Sauvegarde la configuration email"""
    try:
        with open(FICHIER_CONFIG_EMAIL, 'w') as f:
            json.dump(config, f, indent=4)
        return True
    except:
        return False

def charger_config_alertes():
    """Charge la configuration des alertes"""
    try:
        if os.path.exists(FICHIER_ALERTES):
            with open(FICHIER_ALERTES, 'r') as f:
                return json.load(f)
        else:
            return {
                "alerte_delai": True,
                "seuil_delai_jours": 7,
                "alerte_derangement_urgent": True,
                "alerte_litige_escalade": True,
                "alerte_quotidienne": False,
                "heure_alerte_quotidienne": "08:00"
            }
    except:
        return {}

def sauvegarder_config_alertes(config):
    """Sauvegarde la configuration des alertes"""
    try:
        with open(FICHIER_ALERTES, 'w') as f:
            json.dump(config, f, indent=4)
        return True
    except:
        return False

# ====================== FONCTION EMAIL ======================
def envoyer_email(sujet, corps, destinataires, piece_jointe=None, nom_fichier=None):
    """Envoie un email avec pièce jointe optionnelle"""
    config = charger_config_email()
    
    if not config.get("email_expediteur") or not config.get("password"):
        return False, "Configuration email manquante"
    
    try:
        msg = MIMEMultipart()
        msg['From'] = config["email_expediteur"]
        msg['To'] = ", ".join(destinataires)
        msg['Subject'] = sujet
        
        # Corps du message
        msg.attach(MIMEText(corps, 'html'))
        
        # Pièce jointe
        if piece_jointe and nom_fichier:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(piece_jointe)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={nom_fichier}')
            msg.attach(part)
        
        # Envoi
        with smtplib.SMTP(config["smtp_server"], config["smtp_port"]) as server:
            server.starttls()
            server.login(config["email_expediteur"], config["password"])
            server.send_message(msg)
        
        return True, "Email envoyé avec succès"
    
    except Exception as e:
        return False, f"Erreur: {str(e)}"

def generer_rapport_html(titre, dataframe, stats=None):
    """Génère un rapport HTML pour email"""
    html = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; }}
            h2 {{ color: #0E7CFF; }}
            table {{ border-collapse: collapse; width: 100%; margin-top: 20px; }}
            th {{ background-color: #0E7CFF; color: white; padding: 10px; text-align: left; }}
            td {{ border: 1px solid #ddd; padding: 8px; }}
            tr:nth-child(even) {{ background-color: #f2f2f2; }}
            .stats {{ background-color: #f8f9fa; padding: 15px; border-radius: 5px; margin: 20px 0; }}
            .metric {{ display: inline-block; margin: 10px 20px 10px 0; }}
            .metric-value {{ font-size: 24px; font-weight: bold; color: #0E7CFF; }}
            .metric-label {{ font-size: 12px; color: #666; }}
        </style>
    </head>
    <body>
        <h2>🏗️ {titre}</h2>
        <p><strong>Date:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
    """
    
    if stats:
        html += '<div class="stats">'
        for key, value in stats.items():
            html += f'''
            <div class="metric">
                <div class="metric-value">{value}</div>
                <div class="metric-label">{key}</div>
            </div>
            '''
        html += '</div>'
    
    if not dataframe.empty:
        html += dataframe.head(50).to_html(index=False, escape=False)
        if len(dataframe) > 50:
            html += f"<p><em>... et {len(dataframe) - 50} lignes supplémentaires</em></p>"
    else:
        html += "<p>Aucune donnée disponible</p>"
    
    html += """
    <hr>
    <p style="color: #666; font-size: 12px;">
        Ce rapport a été généré automatiquement par le système de gestion MHAMID.
    </p>
    </body>
    </html>
    """
    return html

# ====================== FONCTIONS UTILITAIRES ======================
@st.cache_data(ttl=300)
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

def export_excel_multi_sheets(dict_dataframes, nom_fichier):
    """Exporte plusieurs DataFrames dans un fichier Excel avec plusieurs feuilles"""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        for sheet_name, df in dict_dataframes.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Auto-ajuster la largeur des colonnes
            worksheet = writer.sheets[sheet_name]
            for idx, col in enumerate(df.columns):
                max_length = max(
                    df[col].astype(str).apply(len).max(),
                    len(str(col))
                )
                worksheet.column_dimensions[chr(65 + idx)].width = min(max_length + 2, 50)
    
    return buffer.getvalue()

def import_excel_vers_df(fichier_upload):
    """Importe un fichier Excel uploadé vers DataFrame"""
    try:
        xls = pd.ExcelFile(fichier_upload)
        sheets = {}
        for sheet_name in xls.sheet_names:
            sheets[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
        return sheets
    except Exception as e:
        st.error(f"❌ Erreur d'import: {str(e)}")
        return None

# ====================== STATISTIQUES AVANCÉES ======================
def calculer_statistiques_avancees(df_instances, df_derangements, df_litiges):
    """Calcule des statistiques avancées pour le dashboard"""
    stats = {}
    
    # Instances
    if not df_instances.empty:
        stats['total_instances'] = len(df_instances)
        if 'Date' in df_instances.columns:
            df_instances['Date'] = pd.to_datetime(df_instances['Date'], errors='coerce')
            stats['instances_7j'] = len(df_instances[df_instances['Date'] >= datetime.now() - timedelta(days=7)])
            stats['instances_30j'] = len(df_instances[df_instances['Date'] >= datetime.now() - timedelta(days=30)])
        
        if 'Motif' in df_instances.columns:
            stats['motif_principal'] = df_instances['Motif'].value_counts().index[0] if len(df_instances) > 0 else "N/A"
        
        if 'Statut' in df_instances.columns:
            stats['taux_resolution'] = round(
                len(df_instances[df_instances['Statut'] == 'Résolu']) / len(df_instances) * 100, 1
            ) if len(df_instances) > 0 else 0
    
    # Dérangements
    if not df_derangements.empty:
        stats['total_derangements'] = len(df_derangements)
        if 'Priorite' in df_derangements.columns:
            stats['derangements_urgents'] = len(df_derangements[df_derangements['Priorite'] == 'Urgente'])
        if 'Statut' in df_derangements.columns:
            stats['derangements_ouverts'] = len(df_derangements[df_derangements['Statut'].isin(['Nouveau', 'En cours'])])
    
    # Litiges
    if not df_litiges.empty:
        stats['total_litiges'] = len(df_litiges)
        if 'Montant' in df_litiges.columns:
            stats['montant_total_litiges'] = df_litiges['Montant'].sum()
        if 'Statut' in df_litiges.columns:
            stats['litiges_ouverts'] = len(df_litiges[df_litiges['Statut'].isin(['Ouvert', 'En traitement'])])
    
    return stats

def generer_graphiques_avances(df):
    """Génère des graphiques avancés"""
    figs = []
    
    if not df.empty and 'Date' in df.columns:
        df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
        df = df.dropna(subset=['Date'])
        
        # Évolution temporelle
        df['Mois'] = df['Date'].dt.to_period('M')
        evolution = df.groupby('Mois').size()
        
        fig1 = go.Figure()
        fig1.add_trace(go.Scatter(
            x=evolution.index.astype(str),
            y=evolution.values,
            mode='lines+markers',
            name='Évolution',
            line=dict(color='#0E7CFF', width=3),
            marker=dict(size=8)
        ))
        fig1.update_layout(
            title="📈 Évolution Mensuelle",
            xaxis_title="Mois",
            yaxis_title="Nombre",
            hovermode='x unified'
        )
        figs.append(fig1)
        
        # Heatmap hebdomadaire
        df['Jour_Semaine'] = df['Date'].dt.day_name()
        df['Heure'] = df['Date'].dt.hour
        heatmap_data = df.groupby(['Jour_Semaine', 'Heure']).size().reset_index(name='Count')
        
        if not heatmap_data.empty:
            pivot = heatmap_data.pivot(index='Jour_Semaine', columns='Heure', values='Count').fillna(0)
            
            fig2 = go.Figure(data=go.Heatmap(
                z=pivot.values,
                x=pivot.columns,
                y=pivot.index,
                colorscale='Blues'
            ))
            fig2.update_layout(title="🔥 Heatmap Activité (Jour/Heure)")
            figs.append(fig2)
    
    return figs

# ====================== SIDEBAR ======================
with st.sidebar:
   st.image("https://via.placeholder.com/150x50/0E7CFF/FFFFFF?text=MHAMID", use_column_width=True)
    st.markdown("---")
    
    page = st.radio(
        "🧭 Navigation",
        ["🏠 DASHBOARD", "📝 INSTANCES", "📊 RAPPORTS", "⚠️ DÉRANGEMENTS", 
         "🔧 FIABILISATION", "⚖️ LITIGES", "📧 ALERTES & CONFIG", "📥 IMPORT/EXPORT"],
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
    df_fiab = charger_excel(FICHIER_FIABILISATION)
    
    # Calculer les statistiques
    stats = calculer_statistiques_avancees(df_instances, df_derangements, df_litiges)
    
    # KPIs Principaux
    st.subheader("📊 Vue d'Ensemble")
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric(
            "📋 Instances",
            stats.get('total_instances', 0),
            f"+{stats.get('instances_7j', 0)} cette semaine"
        )
    
    with col2:
        st.metric(
            "⚠️ Dérangements",
            stats.get('total_derangements', 0),
            f"{stats.get('derangements_urgents', 0)} urgents"
        )
    
    with col3:
        st.metric(
            "⚖️ Litiges",
            stats.get('total_litiges', 0),
            f"{stats.get('litiges_ouverts', 0)} ouverts"
        )
    
    with col4:
        st.metric(
            "✅ Taux Résolution",
            f"{stats.get('taux_resolution', 0)}%"
        )
    
    with col5:
        st.metric(
            "💰 Montant Litiges",
            f"{stats.get('montant_total_litiges', 0):,.0f} DH"
        )
    
    st.markdown("---")
    
    # Graphiques du Dashboard
    tab1, tab2, tab3 = st.tabs(["📈 Tendances", "📊 Répartitions", "🔥 Analyses"])
    
    with tab1:
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            if not df_instances.empty and 'Date' in df_instances.columns:
                df_instances['Date'] = pd.to_datetime(df_instances['Date'], errors='coerce')
                df_temp = df_instances.dropna(subset=['Date'])
                df_temp['Mois'] = df_temp['Date'].dt.to_period('M')
                evolution = df_temp.groupby('Mois').size()
                
                fig = px.line(
                    x=evolution.index.astype(str),
                    y=evolution.values,
                    title="📈 Évolution Instances",
                    markers=True
                )
                st.plotly_chart(fig, use_column_width=True)
        
        with col_g2:
            if not df_derangements.empty and 'Date' in df_derangements.columns:
                df_derangements['Date'] = pd.to_datetime(df_derangements['Date'], errors='coerce')
                df_temp = df_derangements.dropna(subset=['Date'])
                df_temp['Mois'] = df_temp['Date'].dt.to_period('M')
                evolution = df_temp.groupby('Mois').size()
                
                fig = px.line(
                    x=evolution.index.astype(str),
                    y=evolution.values,
                    title="📈 Évolution Dérangements",
                    markers=True
                )
                st.plotly_chart(fig, use_column_width=True)
    
    with tab2:
        col_r1, col_r2 = st.columns(2)
        
        with col_r1:
            if not df_instances.empty and 'Motif' in df_instances.columns:
                motif_count = df_instances['Motif'].value_counts().head(10)
                fig = px.pie(
                    values=motif_count.values,
                    names=motif_count.index,
                    title="🥧 Top 10 Motifs Instances",
                    hole=0.4
                )
                st.plotly_chart(fig, use_column_width=True)
        
        with col_r2:
            if not df_derangements.empty and 'Type_Derangement' in df_derangements.columns:
                type_count = df_derangements['Type_Derangement'].value_counts()
                fig = px.bar(
                    x=type_count.values,
                    y=type_count.index,
                    orientation='h',
                    title="📊 Types de Dérangements"
                )use_column_widthuse_column_width=True)
    
    with tab3:
        # Analyse par secteur et agent
        if not df_instances.empty:
            col_a1, col_a2 = st.columns(2)
            
            with col_a1:
                if 'Secteur' in df_instances.columns:
                    secteur_count = df_instances['Secteur'].value_counts()
                    fig = px.bar(
                        x=secteur_count.index,
                        y=secteur_count.values,
                        title="🗺️ Instances par Secteur",
                        color=secteur_count.values,
                        color_continuous_scale='Blues'
                    )use_column_widthuse_column_width=True)
            
            with col_a2:
                if 'Agent' in df_instances.columns:
                    agent_count = df_instances['Agent'].value_counts()
                    fig = px.bar(
                        x=agent_count.index,
                        y=agent_count.values,
                        title="👷 Performance Agents",
                        color=agent_count.values,
                        color_continuous_scale='Viridis'
                    )
                    st.plotly_chart(fig, use_column_width=True)
    
    # Alertes et Notifications
    st.markdown("---")
    st.subheader("🚨 Alertes & Notifications")
    
    col_alert1, col_alert2, col_alert3 = st.columns(3)
    
    with col_alert1:
        if not df_derangements.empty and 'Priorite' in df_derangements.columns:
            urgents = len(df_derangements[df_derangements['Priorite'] == 'Urgente'])
            if urgents > 0:
                st.warning(f"⚠️ **{urgents} dérangements urgents** nécessitent une attention immédiate")
    
    with col_alert2:
        if not df_litiges.empty and 'Statut' in df_litiges.columns:
            escalades = len(df_litiges[df_litiges['Statut'] == 'Escaladé'])
            if escalades > 0:
                st.error(f"🔴 **{escalades} litiges escaladés** en attente")
    
    with col_alert3:
        if not df_instances.empty and 'Date' in df_instances.columns:
            df_instances['Date'] = pd.to_datetime(df_instances['Date'], errors='coerce')
            anciens = len(df_instances[df_instances['Date'] < datetime.now() - timedelta(days=30)])
            if anciens > 0:
                st.info(f"📅 **{anciens} instances** datent de plus de 30 jours")

# ====================== PAGE INSTANCES (Avec envoi email) ======================
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
        
        envoyer_notif = st.checkbox("📧 Envoyer une notification email", value=False)
        
        col_submit1, col_submit2 = st.columns([3, 1])
        with col_submit1:
            submitted = st.form_submit_button("✅ Valider et Enregistrer", type="primary", use_column_width=True)
        with col_submit2:
            clear = st.form_submit_button("🗑️ Effacer", use_column_width=True)

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
                    
                    # Envoi email si demandé
                    if envoyer_notif:
                        config = charger_config_email()
                        if config.get("emails_destinataires"):
                            sujet = f"Nouvelle Instance: {demande}"
                            corps = f"""
                            <h3>Nouvelle Instance Enregistrée</h3>
                            <ul>
                                <li><strong>Demande:</strong> {demande}</li>
                                <li><strong>Client:</strong> {nom}</li>
                                <li><strong>Secteur:</strong> {secteur}</li>
                                <li><strong>Motif:</strong> {motif}</li>
                                <li><strong>Agent:</strong> {agent}</li>
                                <li><strong>Statut:</strong> {statut}</li>
                            </ul>
                            """
                            succes, msg = envoyer_email(sujet, corps, config["emails_destinataires"])
                            if succes:
                                st.success("📧 Email envoyé avec succès")
                            else:
                                st.warning(f"⚠️ Email non envoyé: {msg}")
                        else:
                            st.warning("⚠️ Aucun destinataire configuré")
                else:
                    st.error("❌ Erreur lors de la sauvegarde")
            else:
                st.error("❌ Demande, Télécopie et Motif sont obligatoires")

    st.markdown("---")
    
    # Affichage avec filtres
    st.subheader("📋 Liste des Instances Enregistrées")
    
    if not df_instances.empty:
        col_filter1, col_filter2, col_filter3, col_filter4 = st.columns(4)
        
        with col_filter1:
            filtre_secteur = st.multiselect("Secteur", df_instances['Secteur'].unique() if 'Secteur' in df_instances.columns else [])
        with col_filter2:
            filtre_agent = st.multiselect("Agent", df_instances['Agent'].unique() if 'Agent' in df_instances.columns else [])
        with col_filter3:
            filtre_statut = st.multiselect("Statut", df_instances['Statut'].unique() if 'Statut' in df_instances.columns else [])
        with col_filter4:
            filtre_motif = st.multiselect("Motif", df_instances['Motif'].unique() if 'Motif' in df_instances.columns else [])

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
            use_column_width=True,
            height=400
        )
        
        # Export avec option d'envoi par email
        col_exp1, col_exp2, col_exp3, col_exp4 = st.columns(4)
        
        with col_exp2:
            excel_data = export_excel_multi_sheets({"Instances": df_filtre}, "instances.xlsx")
            st.download_button(
                "📥 Télécharger Excel",
                excel_data,
                f"instances_{datetime.now().strftime('%Y%m%d')}.xlsx",
                use_column_width=True
            )
        
        with col_exp3:
            csv = df_filtre.to_csv(index=False).encode('utf-8')
            st.download_button(
                "📥 Télécharger CSV",
                csv,
                f"instances_{datetime.now().strftime('%Y%m%d')}.csv",
                use_column_width=True
            )
        
        with col_exp4:
            if st.button("📧 Envoyer par Email", use_column_width=True):
                config = charger_config_email()
                if config.get("emails_destinataires"):
                    stats = {
                        "Total": len(df_filtre),
                        "En cours": len(df_filtre[df_filtre['Statut'] == 'En cours']),
                        "Résolus": len(df_filtre[df_filtre['Statut'] == 'Résolu'])
                    }
                    corps_html = generer_rapport_html("Rapport Instances", df_filtre, stats)
                    excel_data = export_excel_multi_sheets({"Instances": df_filtre}, "instances.xlsx")
                    
                    succes, msg = envoyer_email(
                        f"Rapport Instances - {datetime.now().strftime('%d/%m/%Y')}",
                        corps_html,
                        config["emails_destinataires"],
                        excel_data,
                        f"instances_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    )
                    
                    if succes:
                        st.success("📧 Email envoyé avec succès!")
                    else:
                        st.error(f"❌ {msg}")
                else:
                    st.warning("⚠️ Configurez d'abord les destinataires dans la section ALERTES & CONFIG")
    else:
        st.info("ℹ️ Aucune instance enregistrée")

# ====================== PAGE RAPPORTS AVANCÉS ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques Avancés")

    tab1, tab2, tab3, tab4 = st.tabs(["📈 Vue Générale", "📊 Analyse Motifs", "🗺️ Secteurs", "🎯 Analytics Avancés"])
    
    with tab1:
        try:
            etat_df = charger_excel(FICHIER_ETAT, sheet_name="SITUATION14.15")
            motif_df = charger_excel(FICHIER_MOTIF, sheet_name="MOTIF")

            if not etat_df.empty and not motif_df.empty:
                motif_col = find_column(motif_df, ['motif', 'detail', 'pc mauvais'])
                secteur_col = find_column(etat_df, ['secteur', 'sector'])
                etat_col = find_column(etat_df, ['etat', 'état', 'state'])
                delai_col = find_column(etat_df, ['délai', 'delai'])

                st.success("✅ Fichiers chargés avec succès")

                # KPIs
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("📦 Total Commandes", f"{len(etat_df):,}")
                
                with col2:
                    st.metric("📝 Total Motifs", f"{len(motif_df):,}")
                
                with col3:
                    if delai_col:
                        delai_moyen = round(etat_df[delai_col].mean(), 1)
                        st.metric("⏱️ Délai Moyen", f"{delai_moyen} j")
                
                with col4:
                    if etat_col:
                        nb_va = len(etat_df[etat_df[etat_col].astype(str).str.upper() == 'VA'])
                        st.metric("✅ Commandes VA", f"{nb_va:,}")

                st.markdown("---")

                # Graphiques avancés
                date_col = find_column(etat_df, ['date', 'Date'])
                if date_col:
                    st.subheader("📅 Analyse Temporelle")
                    etat_df[date_col] = pd.to_datetime(etat_df[date_col], errors='coerce')
                    df_temp = etat_df.dropna(subset=[date_col])
                    
                    # Choix de la période
                    col_p1, col_p2 = st.columns([1, 3])
                    with col_p1:
                        periode = st.selectbox("Période", ["Jour", "Semaine", "Mois", "Trimestre"])
                    
                    freq_map = {"Jour": "D", "Semaine": "W", "Mois": "M", "Trimestre": "Q"}
                    evolution = df_temp.set_index(date_col).resample(freq_map[periode]).size()
                    
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(
                        x=evolution.index,
                        y=evolution.values,
                        mode='lines+markers',
                        name='Évolution',
                        line=dict(color='#0E7CFF', width=3),
                        fill='tozeroy'
                    ))
                    fig.update_layout(
                        title=f"Évolution par {periode}",
                        xaxis_title=periode,
                        yaxis_title="Nombre de commandes",
                        hovermode='x unified'
                    )
                    st.plotly_chart(fig, use_column_width=True)
                    
                    # Statistiques de tendance
                    col_t1, col_t2, col_t3 = st.columns(3)
                    with col_t1:
                        moyenne = evolution.mean()
                        st.metric("📊 Moyenne", f"{moyenne:.1f}")
                    with col_t2:
                        max_val = evolution.max()
                        st.metric("📈 Maximum", f"{max_val}")
                    with col_t3:
                        min_val = evolution.min()
                        st.metric("📉 Minimum", f"{min_val}")

            else:
                st.warning("⚠️ Fichiers non trouvés")

        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")

    with tab2:
        st.subheader("🔍 Analyse Approfondie des Motifs")
        
        try:
            motif_df = charger_excel(FICHIER_MOTIF, sheet_name="MOTIF")
            
            if not motif_df.empty:
                motif_col = find_column(motif_df, ['motif', 'detail', 'pc mauvais'])
                
                if motif_col:
                    st.info(f"📌 Colonne analysée : **{motif_col}**")
                    
                    motif_series = motif_df[motif_col].astype(str).str.strip()
                    motif_series = motif_series[(motif_series != 'nan') & (motif_series != '') & (motif_series != 'None')]
                    
                    # Choix du nombre de motifs à afficher
                    col_choice1, col_choice2 = st.columns([1, 3])
                    with col_choice1:
                        top_n = st.slider("Nombre de motifs", 5, 30, 15)
                    
                    motif_count = motif_series.value_counts().head(top_n)

                    # Graphiques
                    col_chart1, col_chart2 = st.columns(2)
                    
                    with col_chart1:
                        fig1 = px.bar(
                            x=motif_count.index,
                            y=motif_count.values,
                            title=f"🏆 Top {top_n} des Motifs",
                            labels={'x': 'Motif', 'y': 'Nombre'},
                            color=motif_count.values,
                            color_continuous_scale='Blues'
                        )
                        fig1.update_layout(xaxis_tickangle=-45)
                        st.plotly_chart(fig1, use_column_width=True)

                    with col_chart2:
                        fig2 = px.pie(
                            values=motif_count.values,
                            names=motif_count.index,
                            title="📊 Répartition",
                            hole=0.4
                        )
                        st.plotly_chart(fig2, use_column_width=True)

                    # Tableau détaillé
                    st.subheader("📋 Détails des Motifs")
                    df_motifs_detail = pd.DataFrame({
                        'Motif': motif_count.index,
                        'Nombre': motif_count.values,
                        'Pourcentage': (motif_count.values / motif_count.sum() * 100).round(2),
                        'Cumul %': (motif_count.values / motif_count.sum() * 100).cumsum().round(2)
                    })
                    st.dataframe(df_motifs_detail, use_column_width=True)

                    # Export avec email
                    col_exp1, col_exp2, col_exp3 = st.columns(3)
                    with col_exp2:
                        excel_motifs = export_excel_multi_sheets({"Analyse_Motifs": df_motifs_detail}, "analyse_motifs.xlsx")
                        st.download_button(
                            "📥 Télécharger",
                            excel_motifs,
                            f"analyse_motifs_{datetime.now().strftime('%Y%m%d')}.xlsx",
                            use_column_width=True
                        )
                    with col_exp3:
                        if st.button("📧 Envoyer", key="email_motifs", use_column_width=True):
                            config = charger_config_email()
                            if config.get("emails_destinataires"):
                                stats = {
                                    "Total Motifs": len(motif_series),
                                    "Motifs Uniques": len(motif_count),
                                    "Motif Principal": motif_count.index[0]
                                }
                                corps = generer_rapport_html("Analyse des Motifs", df_motifs_detail, stats)
                                succes, msg = envoyer_email(
                                    f"Analyse Motifs - {datetime.now().strftime('%d/%m/%Y')}",
                                    corps,
                                    config["emails_destinataires"],
                                    excel_motifs,
                                    f"analyse_motifs_{datetime.now().strftime('%Y%m%d')}.xlsx"
                                )
                                if succes:
                                    st.success("📧 Email envoyé!")
                                else:
                                    st.error(f"❌ {msg}")
                else:
                    st.error("❌ Colonne motif introuvable")
        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")

    with tab3:
        st.subheader("🗺️ Analyse Géographique")
        
        try:
            etat_df = charger_excel(FICHIER_ETAT, sheet_name="SITUATION14.15")
            
            if not etat_df.empty:
                secteur_col = find_column(etat_df, ['secteur', 'sector'])
                
                if secteur_col:
                    secteur_count = etat_df[secteur_col].value_counts()
                    
                    # Graphiques
                    col_sect1, col_sect2 = st.columns(2)
                    
                    with col_sect1:
                        fig3 = px.bar(
                            x=secteur_count.index,
                            y=secteur_count.values,
                            title="📍 Commandes par Secteur",
                            color=secteur_count.values,
                            color_continuous_scale='Viridis'
                        )
                        st.plotly_chart(fig3, use_column_width=True)
                    
                    with col_sect2:
                        fig4 = px.pie(
                            values=secteur_count.values,
                            names=secteur_count.index,
                            title="🥧 Répartition Géographique"
                        )
                        st.plotly_chart(fig4, use_column_width=True)

                    # Analyse détaillée par secteur
                    st.subheader("📊 Analyse Détaillée par Secteur")
                    
                    secteur_selectionne = st.selectbox("Sélectionner un secteur", secteur_count.index)
                    
                    df_secteur = etat_df[etat_df[secteur_col] == secteur_selectionne]
                    
                    col_s1, col_s2, col_s3, col_s4 = st.columns(4)
                    
                    col_s1.metric("Total Commandes", len(df_secteur))
                    
                    delai_col = find_column(df_secteur, ['délai', 'delai'])
                    if delai_col:
                        col_s2.metric("Délai Moyen", f"{df_secteur[delai_col].mean():.1f}j")
                        col_s3.metric("Délai Min", f"{df_secteur[delai_col].min():.0f}j")
                        col_s4.metric("Délai Max", f"{df_secteur[delai_col].max():.0f}j")
                    
                    st.dataframe(df_secteur.head(100), use_column_width=True, height=300)
        
        except Exception as e:
            st.error(f"❌ Erreur: {str(e)}")

    with tab4:
        st.subheader("🎯 Analytics Avancés & Prédictions")
        
        df_instances = charger_excel(FICHIER_INSTANCES)
        
        if not df_instances.empty and 'Date' in df_instances.columns:
            df_instances['Date'] = pd.to_datetime(df_instances['Date'], errors='coerce')
            df_temp = df_instances.dropna(subset=['Date'])
            
            # Statistiques avancées
            col_stat1, col_stat2 = st.columns(2)
            
            with col_stat1:
                st.markdown("### 📊 Statistiques Descriptives")
                
                # Distribution temporelle
                df_temp['Jour_Semaine'] = df_temp['Date'].dt.day_name()
                jour_count = df_temp['Jour_Semaine'].value_counts()
                
                fig_jour = px.bar(
                    x=jour_count.index,
                    y=jour_count.values,
                    title="Distribution par Jour de la Semaine",
                    labels={'x': 'Jour', 'y': 'Nombre'}
                )
                st.plotly_chart(fig_jour, use_column_width=True)
            
            with col_stat2:
                st.markdown("### 📈 Tendances")
                
                # Calcul de la tendance
                df_temp['Mois'] = df_temp['Date'].dt.to_period('M')
                monthly = df_temp.groupby('Mois').size()
                
                if len(monthly) > 1:
                    trend = (monthly.iloc[-1] - monthly.iloc[0]) / len(monthly)
                    trend_pct = (trend / monthly.iloc[0] * 100) if monthly.iloc[0] > 0 else 0
                    
                    st.metric(
                        "Tendance Mensuelle",
                        f"{trend:+.1f} instances/mois",
                        f"{trend_pct:+.1f}%"
                    )
                    
                    # Graphique de tendance
                    fig_trend = go.Figure()
                    fig_trend.add_trace(go.Scatter(
                        x=monthly.index.astype(str),
                        y=monthly.values,
                        mode='lines+markers',
                        name='Réel'
                    ))
                    
                    # Ligne de tendance
                    z = np.polyfit(range(len(monthly)), monthly.values, 1)
                    p = np.poly1d(z)
                    fig_trend.add_trace(go.Scatter(
                        x=monthly.index.astype(str),
                        y=p(range(len(monthly))),
                        mode='lines',
                        name='Tendance',
                        line=dict(dash='dash')
                    ))
                    
                    st.plotly_chart(fig_trend, use_column_width=True)
            
            # Heatmap d'activité
            st.markdown("### 🔥 Heatmap d'Activité")
            df_temp['Jour'] = df_temp['Date'].dt.day_name()
            df_temp['Heure'] = df_temp['Date'].dt.hour
            
            heatmap_data = df_temp.groupby(['Jour', 'Heure']).size().reset_index(name='Count')
            
            if not heatmap_data.empty:
                pivot = heatmap_data.pivot(index='Jour', columns='Heure', values='Count').fillna(0)
                
                # Ordre des jours
                jours_ordre = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
                pivot = pivot.reindex([j for j in jours_ordre if j in pivot.index])
                
                fig_heat = go.Figure(data=go.Heatmap(
                    z=pivot.values,
                    x=pivot.columns,
                    y=pivot.index,
                    colorscale='Blues',
                    hoverongaps=False
                ))
                fig_heat.update_layout(
                    title="Activité par Jour et Heure",
                    xaxis_title="Heure",
                    yaxis_title="Jour"
                )
                st.plotly_chart(fig_heat, use_column_width=True)
        else:
            st.info("ℹ️ Pas assez de données pour les analytics avancés")

# ====================== PAGE ALERTES & CONFIG ======================
elif page == "📧 ALERTES & CONFIG":
    st.title("📧 Configuration des Alertes Email")
    
    tab1, tab2, tab3 = st.tabs(["⚙️ Configuration Email", "🔔 Règles d'Alertes", "✉️ Test & Envoi"])
    
    with tab1:
        st.subheader("⚙️ Configuration du Serveur Email")
        
        config = charger_config_email()
        
        with st.form("config_email_form"):
            st.info("💡 Pour Gmail, utilisez un mot de passe d'application (App Password)")
            
            smtp_server = st.text_input("Serveur SMTP", value=config.get("smtp_server", "smtp.gmail.com"))
            smtp_port = st.number_input("Port SMTP", value=config.get("smtp_port", 587), min_value=1, max_value=65535)
            email_expediteur = st.text_input("Email expéditeur", value=config.get("email_expediteur", ""))
            password = st.text_input("Mot de passe / App Password", value=config.get("password", ""), type="password")
            
            st.markdown("**Destinataires (un par ligne)**")
            emails_dest_text = st.text_area(
                "Emails destinataires",
                value="\n".join(config.get("emails_destinataires", [])),
                height=100
            )
            
            if st.form_submit_button("💾 Sauvegarder la Configuration", type="primary"):
                emails_list = [e.strip() for e in emails_dest_text.split("\n") if e.strip()]
                
                nouvelle_config = {
                    "smtp_server": smtp_server,
                    "smtp_port": smtp_port,
                    "email_expediteur": email_expediteur,
                    "password": password,
                    "emails_destinataires": emails_list
                }
                
                if sauvegarder_config_email(nouvelle_config):
                    st.success("✅ Configuration sauvegardée avec succès!")
                    st.cache_data.clear()
                else:
                    st.error("❌ Erreur de sauvegarde")
    
    with tab2:
        st.subheader("🔔 Configuration des Règles d'Alertes Automatiques")
        
        config_alertes = charger_config_alertes()
        
        with st.form("config_alertes_form"):
            st.markdown("### Alertes Automatiques")
            
            alerte_delai = st.checkbox(
                "⏱️ Alerter si délai dépassé",
                value=config_alertes.get("alerte_delai", True)
            )
            seuil_delai = st.number_input(
                "Seuil délai (jours)",
                value=config_alertes.get("seuil_delai_jours", 7),
                min_value=1,
                max_value=90
            )
            
            alerte_derangement = st.checkbox(
                "⚠️ Alerter pour dérangements urgents",
                value=config_alertes.get("alerte_derangement_urgent", True)
            )
            
            alerte_litige = st.checkbox(
                "⚖️ Alerter pour litiges escaladés",
                value=config_alertes.get("alerte_litige_escalade", True)
            )
            
            st.markdown("### Rapport Quotidien")
            
            alerte_quotidienne = st.checkbox(
                "📅 Activer le rapport quotidien automatique",
                value=config_alertes.get("alerte_quotidienne", False)
            )
            
            heure_alerte = st.time_input(
                "Heure d'envoi",
                value=datetime.strptime(config_alertes.get("heure_alerte_quotidienne", "08:00"), "%H:%M").time()
            )
            
            if st.form_submit_button("💾 Sauvegarder les Règles", type="primary"):
                nouvelles_alertes = {
                    "alerte_delai": alerte_delai,
                    "seuil_delai_jours": seuil_delai,
                    "alerte_derangement_urgent": alerte_derangement,
                    "alerte_litige_escalade": alerte_litige,
                    "alerte_quotidienne": alerte_quotidienne,
                    "heure_alerte_quotidienne": heure_alerte.strftime("%H:%M")
                }
                
                if sauvegarder_config_alertes(nouvelles_alertes):
                    st.success("✅ Règles d'alertes sauvegardées!")
                else:
                    st.error("❌ Erreur de sauvegarde")
        
        # Vérification des alertes actives
        st.markdown("---")
        st.subheader("🚨 Vérifier les Alertes Actives")
        
        if st.button("🔍 Analyser et Détecter les Alertes", type="primary"):
            alertes_detectees = []
            
            # Vérifier les instances avec délai dépassé
            if config_alertes.get("alerte_delai"):
                df_instances = charger_excel(FICHIER_INSTANCES)
                if not df_instances.empty and 'Date' in df_instances.columns:
                    df_instances['Date'] = pd.to_datetime(df_instances['Date'], errors='coerce')
                    seuil = datetime.now() - timedelta(days=config_alertes.get("seuil_delai_jours", 7))
                    instances_anciennes = df_instances[df_instances['Date'] < seuil]
                    
                    if len(instances_anciennes) > 0:
                        alertes_detectees.append({
                            "type": "⏱️ Délai dépassé",
                            "nombre": len(instances_anciennes),
                            "message": f"{len(instances_anciennes)} instances dépassent le délai de {config_alertes.get('seuil_delai_jours', 7)} jours"
                        })
            
            # Vérifier dérangements urgents
            if config_alertes.get("alerte_derangement_urgent"):
                df_derangements = charger_excel(FICHIER_DERANGEMENTS)
                if not df_derangements.empty and 'Priorite' in df_derangements.columns:
                    urgents = df_derangements[df_derangements['Priorite'] == 'Urgente']
                    if len(urgents) > 0:
                        alertes_detectees.append({
                            "type": "⚠️ Dérangements urgents",
                            "nombre": len(urgents),
                            "message": f"{len(urgents)} dérangements urgents nécessitent une attention immédiate"
                        })
            
            # Vérifier litiges escaladés
            if config_alertes.get("alerte_litige_escalade"):
                df_litiges = charger_excel(FICHIER_LITIGES)
                if not df_litiges.empty and 'Statut' in df_litiges.columns:
                    escalades = df_litiges[df_litiges['Statut'] == 'Escaladé']
                    if len(escalades) > 0:
                        alertes_detectees.append({
                            "type": "⚖️ Litiges escaladés",
                            "nombre": len(escalades),
                            "message": f"{len(escalades)} litiges ont été escaladés"
                        })
            
            # Afficher les alertes
            if alertes_detectees:
                st.warning(f"🚨 **{len(alertes_detectees)} type(s) d'alerte(s) détecté(s)**")
                for alerte in alertes_detectees:
                    st.error(f"{alerte['type']}: {alerte['message']}")
                
                # Option d'envoi immédiat
                if st.button("📧 Envoyer ces alertes par email maintenant"):
                    config = charger_config_email()
                    if config.get("emails_destinataires"):
                        corps_alertes = "<h2>🚨 Alertes Système MHAMID</h2><ul>"
                        for alerte in alertes_detectees:
                            corps_alertes += f"<li><strong>{alerte['type']}</strong>: {alerte['message']}</li>"
                        corps_alertes += "</ul>"
                        
                        succes, msg = envoyer_email(
                            f"🚨 Alertes MHAMID - {datetime.now().strftime('%d/%m/%Y %H:%M')}",
                            corps_alertes,
                            config["emails_destinataires"]
                        )
                        
                        if succes:
                            st.success("📧 Alertes envoyées par email!")
                        else:
                            st.error(f"❌ {msg}")
                    else:
                        st.warning("⚠️ Aucun destinataire configuré")
            else:
                st.success("✅ Aucune alerte active pour le moment")
    
    with tab3:
        st.subheader("✉️ Test et Envoi Manuel d'Emails")
        
        config = charger_config_email()
        
        if not config.get("email_expediteur") or not config.get("emails_destinataires"):
            st.warning("⚠️ Veuillez d'abord configurer vos paramètres email dans l'onglet Configuration")
        else:
            st.success(f"✅ Configuration active: {config['email_expediteur']}")
            st.info(f"📧 Destinataires: {', '.join(config['emails_destinataires'])}")
            
            # Test simple
            st.markdown("### 🧪 Test de Connexion")
            if st.button("🧪 Envoyer un Email de Test", type="primary"):
                sujet = "Test - Système MHAMID"
                corps = f"""
                <h2>✅ Email de Test</h2>
                <p>Ceci est un email de test envoyé depuis le système de gestion MHAMID.</p>
                <p><strong>Date:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
                <p><strong>Statut:</strong> Configuration fonctionnelle ✓</p>
                """
                
                succes, msg = envoyer_email(sujet, corps, config["emails_destinataires"])
                
                if succes:
                    st.success("✅ Email de test envoyé avec succès!")
                    st.balloons()
                else:
                    st.error(f"❌ Échec de l'envoi: {msg}")
                    st.info("💡 Vérifiez vos paramètres SMTP et mot de passe")
            
            st.markdown("---")
            
            # Envoi manuel de rapports
            st.markdown("### 📊 Envoi Manuel de Rapports")
            
            type_rapport = st.selectbox(
                "Choisir le rapport à envoyer",
                ["Instances", "Dérangements", "Litiges", "Fiabilisation", "Rapport Complet"]
            )
            
            if st.button(f"📧 Envoyer le rapport {type_rapport}", type="secondary"):
                dict_reports = {}
                stats_globales = {}
                
                if type_rapport == "Instances" or type_rapport == "Rapport Complet":
                    df = charger_excel(FICHIER_INSTANCES)
                    if not df.empty:
                        dict_reports["Instances"] = df
                        stats_globales["Instances"] = len(df)
                
                if type_rapport == "Dérangements" or type_rapport == "Rapport Complet":
                    df = charger_excel(FICHIER_DERANGEMENTS)
                    if not df.empty:
                        dict_reports["Derangements"] = df
                        stats_globales["Dérangements"] = len(df)
                
                if type_rapport == "Litiges" or type_rapport == "Rapport Complet":
                    df = charger_excel(FICHIER_LITIGES)
                    if not df.empty:
                        dict_reports["Litiges"] = df
                        stats_globales["Litiges"] = len(df)
                
                if type_rapport == "Fiabilisation" or type_rapport == "Rapport Complet":
                    df = charger_excel(FICHIER_FIABILISATION)
                    if not df.empty:
                        dict_reports["Fiabilisation"] = df
                        stats_globales["Fiabilisation"] = len(df)
                
                if dict_reports:
                    # Générer Excel multi-sheets
                    excel_data = export_excel_multi_sheets(dict_reports, f"rapport_{type_rapport}.xlsx")
                    
                    # Générer HTML
                    corps_html = f"<h2>📊 Rapport {type_rapport}</h2>"
                    corps_html += f"<p><strong>Date:</strong> {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>"
                    corps_html += "<div class='stats'>"
                    for key, value in stats_globales.items():
                        corps_html += f"<div class='metric'><div class='metric-value'>{value}</div><div class='metric-label'>{key}</div></div>"
                    corps_html += "</div>"
                    corps_html += "<p>Veuillez trouver le rapport complet en pièce jointe.</p>"
                    
                    succes, msg = envoyer_email(
                        f"Rapport {type_rapport} - {datetime.now().strftime('%d/%m/%Y')}",
                        corps_html,
                        config["emails_destinataires"],
                        excel_data,
                        f"rapport_{type_rapport}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                    )
                    
                    if succes:
                        st.success(f"✅ Rapport {type_rapport} envoyé avec succès!")
                    else:
                        st.error(f"❌ {msg}")
                else:
                    st.warning("⚠️ Aucune donnée disponible pour ce rapport")

# ====================== PAGE IMPORT/EXPORT ======================
elif page == "📥 IMPORT/EXPORT":
    st.title("📥📤 Import & Export de Données")
    
    tab1, tab2 = st.tabs(["📥 IMPORT", "📤 EXPORT"])
    
    with tab1:
        st.subheader("📥 Import de Données depuis Excel")
        
        st.info("💡 Importez des données en masse depuis un fichier Excel. Le fichier peut contenir plusieurs feuilles.")
        
        type_import = st.selectbox(
            "Type de données à importer",
            ["Instances", "Dérangements", "Litiges", "Fiabilisation"]
        )
        
        fichier_map = {
            "Instances": FICHIER_INSTANCES,
            "Dérangements": FICHIER_DERANGEMENTS,
            "Litiges": FICHIER_LITIGES,
            "Fiabilisation": FICHIER_FIABILISATION
        }
        
        uploaded_file = st.file_uploader(
            f"Choisir un fichier Excel pour {type_import}",
            type=['xlsx', 'xls'],
            key=f"upload_{type_import}"
        )
        
        if uploaded_file is not None:
            try:
                # Lire le fichier
                sheets = import_excel_vers_df(uploaded_file)
                
                if sheets:
                    st.success(f"✅ Fichier chargé avec {len(sheets)} feuille(s)")
                    
                    # Sélectionner la feuille
                    sheet_name = st.selectbox("Sélectionner la feuille à importer", list(sheets.keys()))
                    
                    df_import = sheets[sheet_name]
                    
                    st.markdown(f"**Aperçu des données ({len(df_import)} lignes)**")
                    st.dataframe(df_import.head(10), use_column_width=True)
                    
                    # Options d'import
                    col_opt1, col_opt2 = st.columns(2)
                    
                    with col_opt1:
                        mode_import = st.radio(
                            "Mode d'import",
                            ["Ajouter aux données existantes", "Remplacer les données existantes"]
                        )
                    
                    with col_opt2:
                        confirmer = st.checkbox("Je confirme l'import de ces données")
                    
                    if st.button("✅ Importer les Données", type="primary", disabled=not confirmer):
                        fichier_cible = fichier_map[type_import]
                        
                        if mode_import == "Remplacer les données existantes":
                            # Remplacer
                            if sauvegarder_excel(df_import, fichier_cible):
                                st.success(f"✅ {len(df_import)} lignes importées (remplacement)")
                                st.balloons()
                                st.cache_data.clear()
                            else:
                                st.error("❌ Erreur lors de l'import")
                        else:
                            # Ajouter
                            df_existant = charger_excel(fichier_cible)
                            df_combine = pd.concat([df_existant, df_import], ignore_index=True)
                            
                            if sauvegarder_excel(df_combine, fichier_cible):
                                st.success(f"✅ {len(df_import)} lignes ajoutées (total: {len(df_combine)})")
                                st.balloons()
                                st.cache_data.clear()
                            else:
                                st.error("❌ Erreur lors de l'import")
                else:
                    st.error("❌ Impossible de lire le fichier")
                    
            except Exception as e:
                st.error(f"❌ Erreur: {str(e)}")
        
        # Template de téléchargement
        st.markdown("---")
        st.subheader("📋 Télécharger un Template")
        
        templates = {
            "Instances": pd.DataFrame(columns=[
                "Date", "Demande", "Nom", "Contact", "Adresse", "Telecopie",
                "Date_Reception", "Secteur", "Agent", "Motif", "Statut"
            ]),
            "Dérangements": pd.DataFrame(columns=[
                "Date", "N_Ticket", "Client", "Adresse", "Type_Derangement",
                "Priorite", "Agent_Assigne", "Statut", "Date_Resolution", "Commentaire"
            ]),
            "Litiges": pd.DataFrame(columns=[
                "Date", "N_Litige", "Client", "Type_Litige", "Description",
                "Montant", "Statut", "Agent_Responsable", "Date_Resolution", "Commentaire"
            ]),
            "Fiabilisation": pd.DataFrame(columns=[
                "Date", "Zone", "Type_Intervention", "PC_Concerne",
                "Agent", "Probleme_Detecte", "Action_Corrective", "Statut", "Date_Planifiee"
            ])
        }
        
        col_t1, col_t2, col_t3, col_t4 = st.columns(4)
        
        for i, (nom, df_template) in enumerate(templates.items()):
            col = [col_t1, col_t2, col_t3, col_t4][i]
            with col:
                excel_template = export_excel_multi_sheets({nom: df_template}, f"template_{nom}.xlsx")
                st.download_button(
                    f"📥 {nom}",
                    excel_template,
                    f"template_{nom}.xlsx",
                    use_column_width=True
                )
    
    with tab2:
        st.subheader("📤 Export de Données vers Excel")
        
        # Export sélectif
        st.markdown("### 📊 Export Personnalisé")
        
        col_sel1, col_sel2 = st.columns(2)
        
        with col_sel1:
            tables_a_exporter = st.multiselect(
                "Sélectionner les données à exporter",
                ["Instances", "Dérangements", "Litiges", "Fiabilisation", "Toutes"]
            )
        
        with col_sel2:
            format_date = st.checkbox("Inclure la date dans le nom du fichier", value=True)
        
        if st.button("📤 Générer l'Export", type="primary"):
            if not tables_a_exporter:
                st.warning("⚠️ Veuillez sélectionner au moins une table")
            else:
                dict_export = {}
                
                if "Toutes" in tables_a_exporter:
                    tables_a_exporter = ["Instances", "Dérangements", "Litiges", "Fiabilisation"]
                
                for table in tables_a_exporter:
                    fichier = fichier_map.get(table)
                    if fichier:
                        df = charger_excel(fichier)
                        if not df.empty:
                            dict_export[table] = df
                
                if dict_export:
                    nom_fichier = "export_complet"
                    if format_date:
                        nom_fichier += f"_{datetime.now().strftime('%Y%m%d_%H%M')}"
                    nom_fichier += ".xlsx"
                    
                    excel_data = export_excel_multi_sheets(dict_export, nom_fichier)
                    
                    col_download1, col_download2 = st.columns([2, 1])
                    
                    with col_download1:
                        st.download_button(
                            "📥 Télécharger l'Export",
                            excel_data,
                            nom_fichier,
                            use_column_width=True
                        )
                    
                    with col_download2:
                        if st.button("📧 Envoyer par Email", use_column_width=True):
                            config = charger_config_email()
                            if config.get("emails_destinataires"):
                                corps = f"<h2>📤 Export de Données</h2><p>Veuillez trouver ci-joint l'export des données MHAMID.</p>"
                                succes, msg = envoyer_email(
                                    f"Export Données MHAMID - {datetime.now().strftime('%d/%m/%Y')}",
                                    corps,
                                    config["emails_destinataires"],
                                    excel_data,
                                    nom_fichier
                                )
                                if succes:
                                    st.success("📧 Export envoyé par email!")
                                else:
                                    st.error(f"❌ {msg}")
                            else:
                                st.warning("⚠️ Configurez d'abord les destinataires")
                    
                    st.success(f"✅ Export généré: {len(dict_export)} table(s), {nom_fichier}")
                else:
                    st.warning("⚠️ Aucune donnée à exporter")
        
        st.markdown("---")
        
        # Export rapide individuel
        st.markdown("### ⚡ Export Rapide")
        
        col_quick1, col_quick2, col_quick3, col_quick4 = st.columns(4)
        
        for i, (nom, fichier) in enumerate(fichier_map.items()):
            col = [col_quick1, col_quick2, col_quick3, col_quick4][i]
            with col:
                df = charger_excel(fichier)
                if not df.empty:
                    excel_data = export_excel_multi_sheets({nom: df}, f"{nom}.xlsx")
                    st.download_button(
                        f"📥 {nom}\n({len(df)} lignes)",
                        excel_data,
                        f"{nom}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                        use_column_width=True,
                        key=f"quick_export_{nom}"
                    )
                else:
                    st.info(f"Aucune donnée\n{nom}")

# ====================== AUTRES PAGES (Dérangements, Fiabilisation, Litiges) ======================
# [Le code des autres pages reste identique à la version précédente, mais ajoutez l'option email sur les exports]

# Pour les pages DÉRANGEMENTS, FIABILISATION, LITIGES, gardez le même code que la version précédente
# mais ajoutez systématiquement un bouton "📧 Envoyer par Email" à côté des exports

else:
    # Placeholder pour les autres pages (copier le code de la version précédente)
    st.info(f"Page {page} - Utilisez le code de la version précédente pour ces sections")

# ====================== FOOTER ======================
st.markdown("---")
col_f1, col_f2, col_f3 = st.columns([2, 1, 1])
with col_f1:
    st.caption("🏗️ Application MHAMID - Gestion Fibre & RTC")
with col_f2:
    st.caption(f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M')}")
with col_f3:
    st.caption("v3.0 - Ultimate Edition")
