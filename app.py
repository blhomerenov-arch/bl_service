import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
import io
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

# ====================== AUTHENTIFICATION ======================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.username = None

USERS = {
    "admin": "1234",
    "hamid": "hamid123",
    "shakhman": "shakhman123"
}

if not st.session_state.authenticated:
    st.title("🔐 Connexion à l'application")
    username = st.text_input("Nom d'utilisateur")
    password = st.text_input("Mot de passe", type="password")
   
    if st.button("Se connecter", type="primary"):
        if username in USERS and USERS[username] == password:
            st.session_state.authenticated = True
            st.session_state.username = username
            st.success(f"✅ Connexion réussie ! Bienvenue {username.upper()}")
            st.rerun()
        else:
            st.error("❌ Identifiants incorrects.")
    st.stop()

# ====================== CONFIGURATION EMAIL ======================
# ←←←←←←←←←←  CHANGE CES DEUX LIGNES  ←←←←←←←←←←
EMAIL_SENDER = "ton.email@gmail.com"                    # Ton adresse Gmail
EMAIL_PASSWORD = "ton_mot_de_passe_application"         # Mot de passe d'application Gmail (16 caractères)
EMAIL_RECIPIENT_DEFAULT = "superviseur.mhamid@gmail.com"  # Email du destinataire par défaut

# ====================== FONCTIONS ======================
def find_column(df, keywords):
    if df is None or df.empty:
        return None
    for col in df.columns:
        col_str = str(col).lower().strip()
        if any(k.lower().strip() in col_str for k in keywords):
            return col
    return None

def send_email(subject, body, recipient):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_SENDER
        msg['To'] = recipient
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_SENDER, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"❌ Erreur d'envoi email : {str(e)}")
        return False

# ====================== SAUVEGARDE AUTOMATIQUE ======================
INSTANCES_FILE = "instances_saved.xlsx"

if "instances" not in st.session_state:
    if os.path.exists(INSTANCES_FILE):
        try:
            st.session_state.instances = pd.read_excel(INSTANCES_FILE)
        except:
            st.session_state.instances = pd.DataFrame(columns=["Demande", "Nom", "Contact", "Adresse", "Téléscopie", "Date Réception", "Secteur", "Agent", "Motif", "Date Saisie"])
    else:
        st.session_state.instances = pd.DataFrame(columns=["Demande", "Nom", "Contact", "Adresse", "Téléscopie", "Date Réception", "Secteur", "Agent", "Motif", "Date Saisie"])

def save_instances():
    if not st.session_state.instances.empty:
        st.session_state.instances.to_excel(INSTANCES_FILE, index=False)

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
                    "Demande": demande, "Nom": nom, "Contact": contact, "Adresse": adresse,
                    "Téléscopie": telecopie, "Date Réception": date_reception,
                    "Secteur": secteur, "Agent": agent, "Motif": motif,
                    "Date Saisie": datetime.now()
                }])
                st.session_state.instances = pd.concat([st.session_state.instances, new_row], ignore_index=True)
                save_instances()
                st.success(f"✅ Motif enregistré pour **{demande}**")
                st.balloons()
            else:
                st.error("❌ Les champs Demande, Téléscopie et Motif sont obligatoires.")

    st.subheader("📋 Instances saisies")
    if not st.session_state.instances.empty:
        st.dataframe(st.session_state.instances, use_container_width=True)
    else:
        st.info("Aucune instance saisie pour le moment.")

# ====================== PAGE RAPPORTS ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Envoi par Email")

    # Bouton Email pour envoyer les saisies
    st.subheader("📧 Envoyer les Saisies par Email")

    recipient_email = st.text_input("Adresse email du destinataire", 
                                  value=EMAIL_RECIPIENT_DEFAULT, 
                                  help="L'email où envoyer le rapport")

    if st.button("📧 Envoyer les Instances Saisies par Email", type="primary", use_container_width=True):
        if st.session_state.instances.empty:
            st.warning("⚠️ Aucune instance à envoyer pour le moment.")
        else:
            subject = f"Rapport des Instances Saisies - MHAMID - {datetime.now().strftime('%d/%m/%Y %H:%M')}"

            body = f"""
Bonjour,

Voici le rapport des instances saisies aujourd'hui ({datetime.now().strftime('%d/%m/%Y')}) :

Nombre total d'instances : {len(st.session_state.instances)}

Détails des saisies :
{st.session_state.instances.to_string(index=False)}

Cordialement,
Application Gestion Chantier MHAMID
"""

            if send_email(subject, body, recipient_email):
                st.success(f"✅ Email envoyé avec succès à {recipient_email} !")
                st.balloons()
            else:
                st.error("❌ Échec de l'envoi de l'email. Vérifiez la configuration Gmail.")

    st.info("💡 Pense à configurer EMAIL_SENDER et EMAIL_PASSWORD dans le code avant d'envoyer.")

else:
    st.info(f"Page **{page}** est en cours de développement.")

st.caption(f"Application Gestion Chantier MHAMID | Connecté en tant que **{st.session_state.get('username', 'admin')}**")
