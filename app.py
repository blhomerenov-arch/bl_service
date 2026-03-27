import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
import io
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Configuration de la page
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
# Les valeurs par défaut peuvent être modifiées via l'interface ou le code
INSTANCES_FILE = "instances_saved.xlsx"

# Création d'un expander pour configurer les emails sans toucher au code
with st.expander("⚙️ Configuration Email (Admin Only)", expanded=False):
    col_conf1, col_conf2 = st.columns(2)
    with col_conf1:
        EMAIL_SENDER = st.text_input("Email Expediteur (Gmail)", value="ton.email@gmail.com")
        EMAIL_PASSWORD = st.text_input("Mot de passe App Gmail", type="password", value="ton_mot_de_passe_app")
    with col_conf2:
        EMAIL_RECIPIENT_DEFAULT = st.text_input("Email Destinataire Par Défaut", value="superviseur.mhamid@gmail.com")

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
if "instances" not in st.session_state:
    if os.path.exists(INSTANCES_FILE):
        try:
            # Lecture du fichier Excel existant
            st.session_state.instances = pd.read_excel(INSTANCES_FILE)
            # Conversion des dates en string pour éviter les conflits de types
            if 'Date Réception' in st.session_state.instances.columns:
                st.session_state.instances['Date Réception'] = pd.to_datetime(st.session_state.instances['Date Réception']).dt.strftime('%Y-%m-%d')
            if 'Date Saisie' in st.session_state.instances.columns:
                st.session_state.instances['Date Saisie'] = pd.to_datetime(st.session_state.instances['Date Saisie']).dt.strftime('%Y-%m-%d %H:%M:%S')
        except Exception as e:
            st.warning(f"⚠️ Impossible de charger le fichier existant : {e}. Création d'une nouvelle base.")
            st.session_state.instances = pd.DataFrame(columns=["Demande", "Nom", "Contact", "Adresse", "Téléscopie", "Date Réception", "Secteur", "Agent", "Motif", "Date Saisie"])
    else:
        st.session_state.instances = pd.DataFrame(columns=["Demande", "Nom", "Contact", "Adresse", "Téléscopie", "Date Réception", "Secteur", "Agent", "Motif", "Date Saisie"])

def save_instances():
    # Conversion des dates en string avant sauvegarde Excel
    df_to_save = st.session_state.instances.copy()
    if 'Date Réception' in df_to_save.columns:
        df_to_save['Date Réception'] = df_to_save['Date Réception'].astype(str)
    if 'Date Saisie' in df_to_save.columns:
        df_to_save['Date Saisie'] = df_to_save['Date Saisie'].astype(str)
    
    try:
        df_to_save.to_excel(INSTANCES_FILE, index=False)
    except Exception as e:
        st.error(f"Erreur lors de la sauvegarde du fichier : {e}")

# ====================== NAVIGATION ======================
st.sidebar.title("Navigation")
page = st.sidebar.radio(
    "",
    ["📝 INSTANCES", "📊 RAPPORTS", "⚠️ DÉRANGEMENTS", "🔧 FIABILISATION", "⚖️ LITIGES"],
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
            # Validation : Si "Autre" est sélectionné, le motif_texte doit être rempli
            motif_final = motif
            if motif == "Autre":
                motif_final = st.session_state.get("motif_autre_temp", "") # Pas idéal dans un form, mais on gère via le text_input ci-dessus
                
            if demande and telecopie and motif_final:
                new_row = pd.DataFrame([{
                    "Demande": demande,
                    "Nom": nom,
                    "Contact": contact,
                    "Adresse": adresse,
                    "Téléscopie": telecopie,
                    "Date Réception": date_reception.strftime('%Y-%m-%d'), # Format string
                    "Secteur": secteur,
                    "Agent": agent,
                    "Motif": motif_final,
                    "Date Saisie": datetime.now().strftime('%Y-%m-%d %H:%M:%S') # Format string
                }])
                
                st.session_state.instances = pd.concat([st.session_state.instances, new_row], ignore_index=True)
                save_instances()
                st.success(f"✅ Motif enregistré pour **{demande}**")
                st.balloons()
            else:
                st.error("❌ Les champs Demande, Téléscopie et Motif sont obligatoires.")

    st.subheader("📋 Instances saisies")
    if not st.session_state.instances.empty:
        # Affichage inversé pour voir les dernières entrées en premier
        st.dataframe(st.session_state.instances.iloc[::-1], use_container_width=True)
    else:
        st.info("Aucune instance saisie pour le moment.")

# ====================== PAGE RAPPORTS ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Envoi par Email")

    # Statistiques simples
    if not st.session_state.instances.empty:
        st.metric("Total Instances enregistrées", len(st.session_state.instances))
        
        # Graphique simple des motifs (si plotly est installé)
        try:
            fig = px.histogram(st.session_state.instances, x="Motif", title="Répartition des Motifs")
            st.plotly_chart(fig, use_container_width=True)
        except:
            pass

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
            
            # Création du corps du mail en HTML pour une meilleure lisibilité
            body_html = """
            <html>
            <body>
            <p>Bonjour,</p>
            <p>Voici le rapport des instances saisies aujourd'hui (<strong>{date}</strong>) :</p>
            <p>Nombre total d'instances : <strong>{count}</strong></p>
            <br>
            <p><u>Détails des saisies :</u></p>
            {table}
            <br>
            <p>Cordialement,<br>
            Application Gestion Chantier MHAMID</p>
            </body>
            </html>
            """.format(
                date=datetime.now().strftime('%d/%m/%Y'),
                count=len(st.session_state.instances),
                table=st.session_state.instances.to_html(index=False, border=1)
            )

            if send_email(subject, body_html, recipient_email):
                st.success(f"✅ Email envoyé avec succès à {recipient_email} !")
                st.balloons()
            else:
                st.error("❌ Échec de l'envoi de l'email. Vérifiez la configuration Gmail.")

    st.info("💡 Configurez l'expéditeur et le mot de passe dans le menu déroulant en haut à droite.")

# ====================== PAGES EN DÉVELOPPEMENT ======================
else:
    st.info(f"🚧 Page **{page}** est en cours de développement.")

st.caption(f"Application Gestion Chantier MHAMID | Connecté en tant que **{st.session_state.get('username', 'admin')}**")
