import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
import io
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Pour WhatsApp (Twilio) - installez avec : pip install twilio
from twilio.rest import Client

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

# ====================== CONFIGURATION EMAIL & WHATSAPP ======================
# === EMAIL (Gmail recommandé) ===
EMAIL_SENDER = "ton.email@gmail.com"                    # ← CHANGE
EMAIL_PASSWORD = "ton16caracteresapppassword"           # ← Mot de passe d'application Gmail
EMAIL_RECIPIENT_DEFAULT = "superviseur.mhamid@gmail.com"

# === WHATSAPP (Twilio) ===
TWILIO_ACCOUNT_SID = "ACxxxxxxxxxxxxxxxxxxxxxxxxxxxx"   # ← Change avec tes infos Twilio
TWILIO_AUTH_TOKEN = "xxxxxxxxxxxxxxxxxxxxxxxxxxxx"      # ← Change
TWILIO_WHATSAPP_FROM = "whatsapp:+14155238886"          # Numéro Twilio Sandbox ou ton numéro

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
        st.error(f"Erreur email : {str(e)}")
        return False

def send_whatsapp(message, to_number):
    try:
        client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)
        whatsapp_to = f"whatsapp:{to_number}" if not to_number.startswith("whatsapp:") else to_number
        
        client.messages.create(
            body=message,
            from_=TWILIO_WHATSAPP_FROM,
            to=whatsapp_to
        )
        return True
    except Exception as e:
        st.error(f"Erreur WhatsApp : {str(e)}")
        return False

# ====================== SAUVEGARDE INSTANCES ======================
INSTANCES_FILE = "instances_saved.xlsx"
if "instances" not in st.session_state:
    if os.path.exists(INSTANCES_FILE):
        st.session_state.instances = pd.read_excel(INSTANCES_FILE)
    else:
        st.session_state.instances = pd.DataFrame(columns=["Demande", "Nom", "Contact", "Adresse", "Téléscopie", "Date Réception", "Secteur", "Agent", "Motif", "Date Saisie"])

def save_instances():
    if not st.session_state.instances.empty:
        st.session_state.instances.to_excel(INSTANCES_FILE, index=False)

# ====================== NAVIGATION ======================
page = st.radio("Navigation", ["📝 INSTANCES", "📊 RAPPORTS", "⚠️ DÉRANGEMENTS", "🔧 FIABILISATION", "⚖️ LITIGES"], horizontal=True)

# ====================== PAGE INSTANCES ======================
if page == "📝 INSTANCES":
    st.subheader("📝 Saisie du Motif Journalier")
    # (Code de saisie identique à avant - je l'ai raccourci ici)
    with st.form("saisie_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            demande = st.text_input("Demande*", placeholder="000D740B")
            nom = st.text_input("Nom")
            contact = st.text_input("Contact")
            adresse = st.text_area("Adresse", height=80)
        with col2:
            telecopie = st.text_input("N° Téléscopie*", placeholder="525311326")
            date_reception = st.date_input("Date réception", datetime.now().date())
            secteur = st.selectbox("Secteur", ["MHAMID", "BOUAAKAZ", "Province M'HAMID"])
            agent = st.selectbox("Agent", ["hamid", "SHAKHMAN"])

        motif = st.selectbox("Motif", ["Adresse erronée", "Client refuse installation", "Transport saturé", "PC saturé", "INJOINABLE", "Local fermé + injoignable", "Création PC", "ETUDE CREATION PC", "MSAN saturé", "Autre"])
        if motif == "Autre":
            motif = st.text_input("Précisez le motif")

        if st.form_submit_button("✅ Enregistrer", type="primary"):
            if demande and telecopie and motif:
                new_row = pd.DataFrame([{"Demande": demande, "Nom": nom, "Contact": contact, "Adresse": adresse,
                                       "Téléscopie": telecopie, "Date Réception": date_reception, "Secteur": secteur,
                                       "Agent": agent, "Motif": motif, "Date Saisie": datetime.now()}])
                st.session_state.instances = pd.concat([st.session_state.instances, new_row], ignore_index=True)
                save_instances()
                st.success(f"✅ Enregistré pour {demande}")
                st.balloons()

    if not st.session_state.instances.empty:
        st.dataframe(st.session_state.instances, use_container_width=True)

# ====================== PAGE RAPPORTS ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Envoi Notifications")

    st.subheader("📧 Email & WhatsApp - Envoi Rapide")

    col1, col2 = st.columns(2)
    with col1:
        send_type = st.radio("Type d'envoi :", 
                           ["Installations Validées (VA)", 
                            "Litiges / Dérangements", 
                            "NA, RM, TL, TR (OP)", 
                            "Motifs RDV & INJ devenus joignables"])
    with col2:
        recipient_email = st.text_input("Email destinataire", value=EMAIL_RECIPIENT_DEFAULT)
        whatsapp_number = st.text_input("Numéro WhatsApp (avec +212...)", placeholder="+212612345678")

    if st.button("📤 Envoyer par Email + WhatsApp", type="primary", use_container_width=True):
        if send_type == "NA, RM, TL, TR (OP)":
            subject = f"OP en cours - MHAMID - {datetime.now().strftime('%d/%m/%Y')}"
            body = f"Bonjour,\n\nVoici les demandes avec OP : NA, RM, TL, TR au {datetime.now().strftime('%d/%m/%Y')}\n\nCordialement."
            message_whatsapp = "OP en cours aujourd'hui : NA, RM, TL, TR. Merci de vérifier."

        elif send_type == "Motifs RDV & INJ devenus joignables":
            subject = f"Motifs RDV & INJ Joignables - MHAMID"
            body = "Bonjour,\n\nLes clients avec motifs RDV et INJOINABLE sont maintenant joignables.\nVeuillez relancer l'installation."
            message_whatsapp = "Clients RDV & INJ devenus joignables. Relancez les installations."

        else:
            subject = f"Rapport {send_type} - MHAMID - {datetime.now().strftime('%d/%m/%Y')}"
            body = f"Bonjour,\n\nRapport {send_type} du {datetime.now().strftime('%d/%m/%Y')}.\n\nCordialement."
            message_whatsapp = f"Rapport {send_type} envoyé."

        email_ok = send_email(subject, body, recipient_email)
        whatsapp_ok = send_whatsapp(message_whatsapp, whatsapp_number) if whatsapp_number else False

        if email_ok:
            st.success("✅ Email envoyé avec succès !")
        if whatsapp_ok:
            st.success("✅ Message WhatsApp envoyé avec succès !")

    st.info("⚠️ Configure EMAIL_SENDER, EMAIL_PASSWORD et Twilio dans le code avant utilisation.")

else:
    st.info(f"Page **{page}** en cours de développement.")

st.caption(f"Application MHAMID | Connecté : **{st.session_state.get('username', 'admin')}**")
