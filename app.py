import json
import hashlib
import smtplib
import unicodedata
from email.message import EmailMessage
from io import BytesIO
from pathlib import Path
from datetime import datetime
from urllib.parse import quote

import pandas as pd
import plotly.express as px
import streamlit as st


# =========================================================
# CONFIG
# =========================================================
APP_NAME = "InstalPro"
st.set_page_config(page_title=APP_NAME, layout="wide")

BASE_DIR = Path(__file__).parent
ETAT_FILE = BASE_DIR / "ETAT FTTH RTC RTCL.xlsx"
MOTIF_FILE = BASE_DIR / "MOTIF TOTAL (1).xlsx"

SAISIES_FILE = BASE_DIR / "saisies_instances.csv"
SETTINGS_FILE = BASE_DIR / "parametres_app.json"
EMAIL_CONFIG_FILE = BASE_DIR / "email_config.json"
EMAIL_HISTORY_FILE = BASE_DIR / "email_history.csv"
LOGO_FILE = BASE_DIR / "logo_maroc_telecom.png"

ETAT_SHEET = "SITUATION14.15"
MOTIF_SHEET = "MOTIF"


# =========================================================
# STYLE ULTRA PROFESSIONNEL
# =========================================================
st.markdown(
    """
    <style>
    [data-testid="stAppViewContainer"] {
        background:
            radial-gradient(circle at top left, rgba(14,124,255,0.18), transparent 30%),
            radial-gradient(circle at top right, rgba(0,191,166,0.14), transparent 28%),
            linear-gradient(135deg, #f4f8ff 0%, #eef4fb 45%, #e8eef7 100%);
    }

    [data-testid="stHeader"] {
        background: rgba(255,255,255,0.55);
        backdrop-filter: blur(8px);
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0c2340 0%, #123b68 100%);
    }

    [data-testid="stSidebar"] * {
        color: white !important;
    }

    .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }

    .main-header {
        background: linear-gradient(120deg, #0E7CFF 0%, #1565C0 55%, #00A7B5 100%);
        color: white;
        padding: 22px;
        border-radius: 18px;
        text-align: center;
        margin-bottom: 16px;
        box-shadow: 0 10px 28px rgba(13, 71, 161, 0.20);
    }

    .glass-card {
        background: rgba(255,255,255,0.88);
        border: 1px solid rgba(255,255,255,0.55);
        backdrop-filter: blur(10px);
        border-radius: 16px;
        padding: 16px;
        margin-bottom: 14px;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.08);
    }

    .wa-button a {
        display: inline-block;
        width: 100%;
        text-align: center;
        padding: 11px 14px;
        background: linear-gradient(135deg, #25D366, #128C7E);
        color: white !important;
        text-decoration: none !important;
        border-radius: 12px;
        font-weight: 700;
        box-shadow: 0 8px 20px rgba(37, 211, 102, 0.28);
    }

    .info-chip {
        display: inline-block;
        padding: 6px 10px;
        margin: 2px 6px 2px 0;
        border-radius: 999px;
        background: #eef5ff;
        border: 1px solid #dbeafe;
        color: #0f172a;
        font-size: 12px;
        font-weight: 600;
    }

    .muted {
        color: #64748b;
        font-size: 13px;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# HELPERS
# =========================================================
def rerun_app():
    if hasattr(st, "rerun"):
        st.rerun()
    else:
        st.experimental_rerun()


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def generate_instance_id():
    return datetime.now().strftime("%Y%m%d%H%M%S%f")


def normalize_text(value):
    value = "" if value is None else str(value)
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("utf-8")
    return value.lower().strip()


def clean_phone(phone):
    if phone is None:
        return ""
    return "".join(ch for ch in str(phone) if ch.isdigit())


def to_excel_bytes(df, sheet_name="Data"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()


def global_search(df, query):
    if df.empty or not query:
        return df
    mask = df.astype(str).apply(lambda col: col.str.contains(query, case=False, na=False))
    return df[mask.any(axis=1)]


def safe_mean_numeric(series):
    s = pd.to_numeric(series, errors="coerce")
    if s.notna().any():
        return round(s.mean(), 1)
    return None


def get_secret(name, default=""):
    try:
        return st.secrets[name]
    except Exception:
        return default


@st.cache_data(show_spinner=False)
def load_excel(path_str, sheet_name):
    return pd.read_excel(path_str, sheet_name=sheet_name)


def safe_load_excel(path, sheet_name, label):
    if not path.exists():
        return pd.DataFrame()
    try:
        return load_excel(str(path), sheet_name)
    except Exception as e:
        st.warning(f"Erreur chargement {label} : {e}")
        return pd.DataFrame()


def find_column(df, keywords):
    if df is None or df.empty:
        return None

    normalized_keywords = [normalize_text(k) for k in keywords]

    for col in df.columns:
        col_name = normalize_text(col)
        if any(k in col_name for k in normalized_keywords):
            return col

    for col in df.columns:
        try:
            non_empty = (
                df[col].astype(str).str.strip().replace("nan", "").ne("").sum()
            )
            if non_empty > 10:
                return col
        except Exception:
            pass

    return None


def normalize_etat_agent(value):
    txt = normalize_text(value).replace(" ", "").replace("_", "").replace("-", "")
    mapping = {
        "na": "NA",
        "nouvelleinstallation": "NA",
        "rm": "RM",
        "remisedeservice": "RM",
        "remiseenservice": "RM",
        "tr": "TR",
        "transfert": "TR",
        "tl": "TL",
        "transfertlocal": "TL",
    }
    return mapping.get(txt, str(value).strip().upper())


# =========================================================
# SETTINGS
# =========================================================
def default_settings():
    return {
        "utilisateurs": ["admin"],
        "secteurs": ["MHAMID", "BOUAAKAZ", "Province M'HAMID"],
        "agents": ["NA", "RM", "TR", "TL"],
        "agent_contacts": {
            "NA": {"whatsapp": "", "email": ""},
            "RM": {"whatsapp": "", "email": ""},
            "TR": {"whatsapp": "", "email": ""},
            "TL": {"whatsapp": "", "email": ""}
        },
        "admin_username": "admin",
        "admin_password_hash": hash_password("admin123")
    }


def save_settings(data):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def sync_agent_contacts(settings):
    if "agent_contacts" not in settings or not isinstance(settings["agent_contacts"], dict):
        settings["agent_contacts"] = {}

    for agent in settings.get("agents", []):
        if agent not in settings["agent_contacts"]:
            settings["agent_contacts"][agent] = {"whatsapp": "", "email": ""}

    removed = [k for k in settings["agent_contacts"].keys() if k not in settings.get("agents", [])]
    for k in removed:
        del settings["agent_contacts"][k]

    return settings


def load_settings():
    defaults = default_settings()

    if not SETTINGS_FILE.exists():
        save_settings(defaults)
        return defaults

    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        data = defaults

    for key, value in defaults.items():
        if key not in data:
            data[key] = value

    data = sync_agent_contacts(data)
    save_settings(data)
    return data


def add_item_to_settings(settings, key, value):
    value = str(value).strip()
    if not value:
        return False, "Champ vide."

    existing = [x.lower() for x in settings.get(key, [])]
    if value.lower() in existing:
        return False, f"{value} existe déjà."

    settings[key].append(value)
    settings[key] = sorted(settings[key], key=lambda x: x.lower())

    if key == "agents":
        settings["agent_contacts"][value] = {"whatsapp": "", "email": ""}

    save_settings(settings)
    return True, f"{value} ajouté avec succès."


def update_item_in_settings(settings, key, old_value, new_value):
    old_value = str(old_value).strip()
    new_value = str(new_value).strip()

    if not old_value or not new_value:
        return False, "Valeur invalide."

    if old_value not in settings.get(key, []):
        return False, "Élément introuvable."

    existing_lower = [x.lower() for x in settings.get(key, []) if x != old_value]
    if new_value.lower() in existing_lower:
        return False, f"{new_value} existe déjà."

    index = settings[key].index(old_value)
    settings[key][index] = new_value
    settings[key] = sorted(settings[key], key=lambda x: x.lower())

    if key == "agents":
        settings["agent_contacts"][new_value] = settings["agent_contacts"].get(old_value, {"whatsapp": "", "email": ""})
        if old_value in settings["agent_contacts"]:
            del settings["agent_contacts"][old_value]

    save_settings(settings)
    return True, f"{old_value} modifié en {new_value}."


def delete_item_in_settings(settings, key, value):
    value = str(value).strip()

    if value not in settings.get(key, []):
        return False, "Élément introuvable."

    if len(settings[key]) <= 1:
        return False, "Impossible de supprimer le dernier élément."

    settings[key].remove(value)

    if key == "agents" and value in settings["agent_contacts"]:
        del settings["agent_contacts"][value]

    save_settings(settings)
    return True, f"{value} supprimé avec succès."


def get_agent_contact(settings, agent_name):
    contacts = settings.get("agent_contacts", {})
    return contacts.get(agent_name, {"whatsapp": "", "email": ""})


def update_agent_contact(settings, agent_name, whatsapp, email):
    settings = sync_agent_contacts(settings)
    settings["agent_contacts"][agent_name] = {
        "whatsapp": str(whatsapp).strip(),
        "email": str(email).strip()
    }
    save_settings(settings)
    return True


# =========================================================
# EMAIL CONFIG
# =========================================================
def default_email_config():
    return {
        "smtp_server": "smtp.office365.com",
        "smtp_port": 587,
        "sender_email": "",
        "default_recipient": ""
    }


def save_email_config(data):
    with open(EMAIL_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_email_config():
    defaults = default_email_config()

    if not EMAIL_CONFIG_FILE.exists():
        save_email_config(defaults)
        return defaults

    try:
        with open(EMAIL_CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        data = defaults

    for key, value in defaults.items():
        if key not in data:
            data[key] = value

    return data


# =========================================================
# AUTH ADMIN
# =========================================================
def init_auth_state():
    if "is_admin" not in st.session_state:
        st.session_state["is_admin"] = False
    if "admin_user" not in st.session_state:
        st.session_state["admin_user"] = ""


def admin_login(settings, username, password):
    ok_user = username == settings.get("admin_username", "admin")
    ok_pass = hash_password(password) == settings.get("admin_password_hash", "")
    if ok_user and ok_pass:
        st.session_state["is_admin"] = True
        st.session_state["admin_user"] = username
        return True
    return False


def admin_logout():
    st.session_state["is_admin"] = False
    st.session_state["admin_user"] = ""


# =========================================================
# DATA
# =========================================================
def load_saved_instances():
    if SAISIES_FILE.exists():
        try:
            return pd.read_csv(SAISIES_FILE)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()


def save_instance(record):
    new_df = pd.DataFrame([record])

    if SAISIES_FILE.exists():
        old_df = pd.read_csv(SAISIES_FILE)
        final_df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        final_df = new_df

    final_df.to_csv(SAISIES_FILE, index=False)


def update_instance(instance_id, updates):
    df = load_saved_instances()
    if df.empty or "instance_id" not in df.columns:
        return False

    mask = df["instance_id"].astype(str) == str(instance_id)
    if not mask.any():
        return False

    for key, value in updates.items():
        if key not in df.columns:
            df[key] = ""
        df.loc[mask, key] = value

    df.to_csv(SAISIES_FILE, index=False)
    return True


# =========================================================
# EMAIL HISTORY
# =========================================================
def load_email_history():
    if EMAIL_HISTORY_FILE.exists():
        try:
            return pd.read_csv(EMAIL_HISTORY_FILE)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()


def append_email_history(record):
    new_df = pd.DataFrame([record])

    if EMAIL_HISTORY_FILE.exists():
        old_df = pd.read_csv(EMAIL_HISTORY_FILE)
        final_df = pd.concat([old_df, new_df], ignore_index=True)
    else:
        final_df = new_df

    final_df.to_csv(EMAIL_HISTORY_FILE, index=False)


# =========================================================
# EMAIL SENDER
# =========================================================
def send_email_smtp(config, password, to_email, subject, body):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = config["sender_email"]
    msg["To"] = to_email
    msg.set_content(body)

    server = smtplib.SMTP(config["smtp_server"], int(config["smtp_port"]))
    server.starttls()
    server.login(config["sender_email"], password)
    server.send_message(msg)
    server.quit()


# =========================================================
# MESSAGE BUILDERS
# =========================================================
def build_instance_message(row):
    return f"""
Bonjour,

Merci de prendre en charge cette instance.

Demande : {row.get('demande', '')}
Nom : {row.get('nom', '')}
Contact : {row.get('contact', '')}
Adresse : {row.get('adresse', '')}
Télécopie : {row.get('telecopie', '')}
Date de réception : {row.get('date_reception', '')}
Secteur : {row.get('secteur', '')}
Agent : {row.get('agent', '')}
Motif : {row.get('motif', '')}
Utilisateur : {row.get('utilisateur', '')}

Cordialement,
InstalPro
""".strip()


def build_source_excel_message(row, etat_col):
    etat_value = row.get(etat_col, "") if etat_col else ""
    return f"""
Bonjour,

Merci de prendre en charge cette demande issue du fichier source Excel.

Demande / Référence : {row.get('Demande', row.get('demande', row.get('Commande', row.get('commande', ''))))}
Nom : {row.get('Nom', row.get('nom', ''))}
Contact : {row.get('Contact', row.get('contact', ''))}
Adresse : {row.get('Adresse', row.get('adresse', ''))}
Secteur : {row.get('Secteur', row.get('secteur', ''))}
État : {etat_value}

Cordialement,
InstalPro
""".strip()


def build_whatsapp_url(agent_whatsapp, row):
    phone = clean_phone(agent_whatsapp)
    if not phone:
        return ""
    text = build_instance_message(row)
    return f"https://wa.me/{phone}?text={quote(text)}"


def build_source_excel_whatsapp_url(agent_whatsapp, row, etat_col):
    phone = clean_phone(agent_whatsapp)
    if not phone:
        return ""
    text = build_source_excel_message(row, etat_col)
    return f"https://wa.me/{phone}?text={quote(text)}"


def render_whatsapp_button(url):
    return f"""
    <div class="wa-button">
        <a href="{url}" target="_blank">💬 Envoyer par WhatsApp</a>
    </div>
    """


# =========================================================
# HEADER
# =========================================================
def render_header():
    if LOGO_FILE.exists():
        c1, c2 = st.columns([1, 5])
        with c1:
            st.image(str(LOGO_FILE), width=160)
        with c2:
            st.markdown(
                f"""
                <div class="main-header">
                    <h2>{APP_NAME}</h2>
                </div>
                """,
                unsafe_allow_html=True
            )
    else:
        st.markdown(
            f"""
            <div class="main-header">
                <h2>{APP_NAME}</h2>
            </div>
            """,
            unsafe_allow_html=True
        )


# =========================================================
# ADMIN RENDER
# =========================================================
def render_manager_tab(settings, key, label):
    current_items = settings.get(key, [])

    st.markdown(f"### Gestion des {label.lower()}s")
    st.write(current_items)

    a1, a2, a3 = st.columns(3)

    with a1:
        with st.form(f"add_{key}_form"):
            new_value = st.text_input(f"Nouveau {label.lower()}", key=f"new_{key}")
            submit = st.form_submit_button("➕ Ajouter")
            if submit:
                ok, msg = add_item_to_settings(settings, key, new_value)
                if ok:
                    st.success(msg)
                    rerun_app()
                else:
                    st.error(msg)

    with a2:
        with st.form(f"edit_{key}_form"):
            old_value = st.selectbox(f"{label} à modifier", current_items, key=f"old_{key}")
            new_value = st.text_input("Nouveau nom", key=f"edit_{key}")
            submit = st.form_submit_button("✏️ Modifier")
            if submit:
                ok, msg = update_item_in_settings(settings, key, old_value, new_value)
                if ok:
                    st.success(msg)
                    rerun_app()
                else:
                    st.error(msg)

    with a3:
        with st.form(f"delete_{key}_form"):
            selected = st.selectbox(f"{label} à supprimer", current_items, key=f"del_{key}")
            submit = st.form_submit_button("🗑️ Supprimer")
            if submit:
                ok, msg = delete_item_in_settings(settings, key, selected)
                if ok:
                    st.success(msg)
                    rerun_app()
                else:
                    st.error(msg)


def render_agent_contacts_admin(settings):
    st.markdown("### Numéros WhatsApp / emails des agents")
    if not settings.get("agents"):
        st.info("Aucun agent disponible.")
        return

    selected_agent = st.selectbox("Agent", settings["agents"], key="contact_agent_select")
    current = get_agent_contact(settings, selected_agent)

    with st.form("agent_contact_form"):
        whatsapp_value = st.text_input("Numéro WhatsApp de l'agent", value=current.get("whatsapp", ""))
        email_value = st.text_input("Email de l'agent", value=current.get("email", ""))
        submit = st.form_submit_button("💾 Enregistrer contact agent")
        if submit:
            update_agent_contact(settings, selected_agent, whatsapp_value, email_value)
            st.success(f"Contact de l'agent {selected_agent} mis à jour.")
            rerun_app()

    st.markdown("#### Résumé")
    for agent in settings["agents"]:
        contact = get_agent_contact(settings, agent)
        st.markdown(
            f"- **{agent}** | WhatsApp : `{contact.get('whatsapp', '')}` | Email : `{contact.get('email', '')}`"
        )


# =========================================================
# INIT
# =========================================================
init_auth_state()
settings = load_settings()
email_config = load_email_config()

if "smtp_password" not in st.session_state:
    st.session_state["smtp_password"] = get_secret("smtp_password", "")

render_header()


# =========================================================
# SIDEBAR ADMIN
# =========================================================
with st.sidebar:
    st.header("Administration")

    if not st.session_state["is_admin"]:
        with st.form("admin_login_form"):
            username = st.text_input("Nom admin")
            password = st.text_input("Mot de passe admin", type="password")
            login_btn = st.form_submit_button("🔐 Se connecter")
            if login_btn:
                if admin_login(settings, username, password):
                    st.success("Connexion admin réussie.")
                    rerun_app()
                else:
                    st.error("Identifiants admin incorrects.")
        st.caption("Login par défaut : admin / admin123")
    else:
        st.success(f"Connecté : {st.session_state['admin_user']}")

        if st.button("🚪 Déconnexion admin"):
            admin_logout()
            rerun_app()

        st.markdown("---")
        st.subheader("Configuration e-mail Outlook")
        with st.form("email_config_form"):
            smtp_server = st.text_input(
                "Serveur SMTP",
                value=email_config.get("smtp_server", "smtp.office365.com")
            )
            smtp_port = st.number_input(
                "Port SMTP",
                min_value=1,
                max_value=9999,
                value=int(email_config.get("smtp_port", 587))
            )
            sender_email = st.text_input(
                "Email expéditeur",
                value=email_config.get("sender_email", "")
            )
            default_recipient = st.text_input(
                "Email destinataire par défaut",
                value=email_config.get("default_recipient", "")
            )
            smtp_password = st.text_input(
                "Mot de passe SMTP",
                type="password",
                value=st.session_state.get("smtp_password", "")
            )

            save_email_btn = st.form_submit_button("💾 Enregistrer e-mail")
            if save_email_btn:
                save_email_config({
                    "smtp_server": smtp_server,
                    "smtp_port": int(smtp_port),
                    "sender_email": sender_email,
                    "default_recipient": default_recipient
                })
                st.session_state["smtp_password"] = smtp_password
                st.success("Configuration e-mail Outlook enregistrée.")

        st.markdown("---")
        st.subheader("Logo")
        uploaded_logo = st.file_uploader("Uploader le logo Maroc Telecom", type=["png", "jpg", "jpeg"])
        lc1, lc2 = st.columns(2)

        with lc1:
            if st.button("💾 Sauvegarder logo"):
                if uploaded_logo is not None:
                    with open(LOGO_FILE, "wb") as f:
                        f.write(uploaded_logo.getbuffer())
                    st.success("Logo sauvegardé.")
                    rerun_app()
                else:
                    st.warning("Choisis un logo d'abord.")

        with lc2:
            if st.button("🗑️ Supprimer logo"):
                if LOGO_FILE.exists():
                    LOGO_FILE.unlink()
                    st.success("Logo supprimé.")
                    rerun_app()
                else:
                    st.info("Aucun logo à supprimer.")

        st.markdown("---")
        st.subheader("Sécurité admin")
        with st.form("change_admin_credentials"):
            new_admin_user = st.text_input("Nouveau nom admin", value=settings.get("admin_username", "admin"))
            new_admin_password = st.text_input("Nouveau mot de passe admin", type="password")
            save_admin_btn = st.form_submit_button("🔑 Mettre à jour")
            if save_admin_btn:
                if not new_admin_user.strip():
                    st.error("Le nom admin est obligatoire.")
                elif not new_admin_password.strip():
                    st.error("Le mot de passe admin est obligatoire.")
                else:
                    settings["admin_username"] = new_admin_user.strip()
                    settings["admin_password_hash"] = hash_password(new_admin_password.strip())
                    save_settings(settings)
                    st.success("Identifiants admin mis à jour.")


# =========================================================
# NAVIGATION
# =========================================================
page = st.radio(
    "Navigation",
    ["📝 INSTANCES", "📊 RAPPORTS", "⚠️ DÉRANGEMENTS", "🔧 FIABILISATION", "⚖️ LITIGES"],
    horizontal=True,
    label_visibility="collapsed"
)


# =========================================================
# PAGE INSTANCES
# =========================================================
if page == "📝 INSTANCES":
    st.subheader("Étape 1 - Saisie et enregistrement")

    if st.session_state["is_admin"]:
        with st.expander("⚙️ Administration complète", expanded=False):
            tabs = st.tabs(["Utilisateurs", "Secteurs", "Agents", "Contacts agents"])
            with tabs[0]:
                render_manager_tab(settings, "utilisateurs", "Utilisateur")
            with tabs[1]:
                render_manager_tab(settings, "secteurs", "Secteur")
            with tabs[2]:
                render_manager_tab(settings, "agents", "Agent")
            with tabs[3]:
                render_agent_contacts_admin(settings)

    with st.form("instance_form", clear_on_submit=True):
        col1, col2 = st.columns(2)

        with col1:
            utilisateur = st.selectbox("Utilisateur", settings["utilisateurs"])
            demande = st.text_input("Demande *", placeholder="000D740B")
            nom = st.text_input("Nom")
            contact = st.text_input("Contact")
            adresse = st.text_area("Adresse", height=90)

        with col2:
            telecopie = st.text_input("N° de télécopie *", placeholder="525311326")
            date_reception = st.date_input("Date de réception", datetime.now().date())
            secteur = st.selectbox("Secteur", settings["secteurs"])
            agent = st.selectbox("Agent destinataire", settings["agents"])

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
            "Autre"
        ]
        motif = st.selectbox("Motif", motif_options)
        if motif == "Autre":
            motif = st.text_input("Précisez le motif")

        submit_instance = st.form_submit_button("✅ Enregistrer l'instance")

        if submit_instance:
            if demande and telecopie and motif:
                agent_contact = get_agent_contact(settings, agent)

                record = {
                    "instance_id": generate_instance_id(),
                    "date_saisie": now_str(),
                    "utilisateur": utilisateur,
                    "demande": demande,
                    "nom": nom,
                    "contact": contact,
                    "adresse": adresse,
                    "telecopie": telecopie,
                    "date_reception": str(date_reception),
                    "secteur": secteur,
                    "agent": agent,
                    "agent_whatsapp": agent_contact.get("whatsapp", ""),
                    "agent_email": agent_contact.get("email", ""),
                    "motif": motif,
                    "statut_etape": "Étape 1 - enregistrée",
                    "statut_email": "Non envoyé",
                    "date_email": "",
                    "statut_whatsapp": "Non envoyé",
                    "date_whatsapp": ""
                }
                save_instance(record)
                st.success("Instance enregistrée avec succès. Passe à l'étape 2 pour l'envoi.")
            else:
                st.error("Les champs Demande, Télécopie et Motif sont obligatoires.")

    st.markdown("---")
    st.subheader("Étape 2 - Envoi des instances")

    saved_df = load_saved_instances()
    email_history_df = load_email_history()

    if saved_df.empty:
        st.info("Aucune instance enregistrée.")
    else:
        f1, f2, f3 = st.columns(3)
        with f1:
            search_saved = st.text_input("Recherche", placeholder="demande, agent, motif...")
        with f2:
            sector_choices = ["Tous"]
            if "secteur" in saved_df.columns:
                sector_choices += sorted(saved_df["secteur"].dropna().astype(str).unique().tolist())
            selected_sector = st.selectbox("Filtre secteur", sector_choices)
        with f3:
            selected_agent = st.selectbox("Filtre agent", ["Tous"] + settings["agents"])

        filtered_saved = global_search(saved_df, search_saved).copy()

        if selected_sector != "Tous" and "secteur" in filtered_saved.columns:
            filtered_saved = filtered_saved[filtered_saved["secteur"].astype(str) == selected_sector]

        if selected_agent != "Tous" and "agent" in filtered_saved.columns:
            filtered_saved = filtered_saved[filtered_saved["agent"].astype(str) == selected_agent]

        try:
            filtered_saved = filtered_saved.sort_values(by="date_saisie", ascending=False)
        except Exception:
            pass

        for _, row in filtered_saved.iterrows():
            instance_id = str(row.get("instance_id", ""))
            agent_name = str(row.get("agent", ""))
            agent_contact = get_agent_contact(settings, agent_name)
            agent_whatsapp = agent_contact.get("whatsapp", "") or row.get("agent_whatsapp", "")
            agent_email = agent_contact.get("email", "") or row.get("agent_email", "")

            message = build_instance_message(row)
            wa_url = build_whatsapp_url(agent_whatsapp, row)

            st.markdown('<div class="glass-card">', unsafe_allow_html=True)

            c1, c2, c3, c4 = st.columns([4.4, 1.5, 1.6, 1.3])

            with c1:
                st.markdown(
                    f"""
**Demande :** {row.get('demande', '')}  
**Agent :** {agent_name}  
**Secteur :** {row.get('secteur', '')}  
**Motif :** {row.get('motif', '')}  
**Statut global :** {row.get('statut_etape', '')}
"""
                )
                st.markdown(
                    f"""
<span class="info-chip">WhatsApp agent : {agent_whatsapp or 'Non configuré'}</span>
<span class="info-chip">Email agent : {agent_email or 'Non configuré'}</span>
                    """,
                    unsafe_allow_html=True
                )

            with c2:
                if st.session_state["is_admin"]:
                    if st.button("📧 Envoyer e-mail", key=f"mail_{instance_id}"):
                        current_config = load_email_config()
                        password = st.session_state.get("smtp_password", "") or get_secret("smtp_password", "")
                        recipient = agent_email.strip() or current_config.get("default_recipient", "").strip()
                        subject = f"Nouvelle instance - {row.get('demande', '')}"

                        if not current_config.get("sender_email"):
                            st.error("Configure l'email expéditeur dans la sidebar admin.")
                        elif not password:
                            st.error("Ajoute le mot de passe SMTP dans la sidebar admin.")
                        elif not recipient:
                            st.error("Aucun email agent ni email destinataire par défaut.")
                        else:
                            try:
                                send_email_smtp(current_config, password, recipient, subject, message)
                                update_instance(
                                    instance_id,
                                    {
                                        "statut_email": "Envoyé",
                                        "date_email": now_str(),
                                        "statut_etape": "Étape 2 - e-mail envoyé"
                                    }
                                )
                                append_email_history(
                                    {
                                        "timestamp": now_str(),
                                        "instance_id": instance_id,
                                        "demande": row.get("demande", ""),
                                        "agent": agent_name,
                                        "recipient": recipient,
                                        "subject": subject,
                                        "status": "SUCCESS",
                                        "error": "",
                                        "sent_by": st.session_state.get("admin_user", "")
                                    }
                                )
                                st.success(f"E-mail envoyé à {recipient}")
                                rerun_app()
                            except Exception as e:
                                append_email_history(
                                    {
                                        "timestamp": now_str(),
                                        "instance_id": instance_id,
                                        "demande": row.get("demande", ""),
                                        "agent": agent_name,
                                        "recipient": recipient,
                                        "subject": subject,
                                        "status": "ERROR",
                                        "error": str(e),
                                        "sent_by": st.session_state.get("admin_user", "")
                                    }
                                )
                                st.error(f"Erreur e-mail : {e}")
                else:
                    st.caption("Admin requis")

            with c3:
                if wa_url:
                    st.markdown(render_whatsapp_button(wa_url), unsafe_allow_html=True)
                else:
                    st.caption("WhatsApp non configuré")

            with c4:
                if st.session_state["is_admin"]:
                    if st.button("✅ Marquer WA", key=f"wa_mark_{instance_id}"):
                        update_instance(
                            instance_id,
                            {
                                "statut_whatsapp": "Envoyé",
                                "date_whatsapp": now_str(),
                                "statut_etape": "Étape 2 - WhatsApp envoyé"
                            }
                        )
                        st.success("WhatsApp marqué comme envoyé.")
                        rerun_app()
                else:
                    st.caption("Admin requis")

            with st.expander(f"Voir détail - {row.get('demande', '')}"):
                st.text_area(
                    "Message à envoyer",
                    value=message,
                    height=220,
                    key=f"msg_{instance_id}"
                )
                st.write(f"Email agent : {agent_email}")
                st.write(f"WhatsApp agent : {agent_whatsapp}")
                st.write(f"Statut email : {row.get('statut_email', 'Non envoyé')}")
                st.write(f"Statut WhatsApp : {row.get('statut_whatsapp', 'Non envoyé')}")

                if not email_history_df.empty and "instance_id" in email_history_df.columns:
                    hist_one = email_history_df[email_history_df["instance_id"].astype(str) == instance_id]
                    if not hist_one.empty:
                        st.markdown("**Historique e-mail de cette instance**")
                        st.dataframe(hist_one, use_container_width=True, height=180)

            st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")
    st.subheader("Données source Excel et dispatch par agent")

    etat_df = safe_load_excel(ETAT_FILE, ETAT_SHEET, "ETAT FTTH RTC RTCL.xlsx")

    if not etat_df.empty:
        search_excel = st.text_input(
            "Recherche dans le fichier source",
            placeholder="demande, secteur, état..."
        )
        filtered_etat = global_search(etat_df, search_excel).copy()

        etat_col = find_column(filtered_etat, ["etat", "état", "state"])
        secteur_col = find_column(filtered_etat, ["secteur", "sector"])
        demande_col = find_column(filtered_etat, ["demande", "commande", "reference", "référence"])

        if etat_col:
            filtered_etat["AGENT_CIBLE"] = filtered_etat[etat_col].apply(normalize_etat_agent)
            actionable_df = filtered_etat[
                filtered_etat["AGENT_CIBLE"].isin(["NA", "RM", "TR", "TL"])
            ].copy()

            st.markdown("### Lignes dispatchables selon la colonne État")

            if actionable_df.empty:
                st.info("Aucune ligne avec État = NA / RM / TR / TL.")
            else:
                max_rows = st.number_input(
                    "Nombre de lignes à afficher",
                    min_value=5,
                    max_value=200,
                    value=20,
                    step=5
                )

                preview_df = actionable_df.head(max_rows)

                for idx, row in preview_df.iterrows():
                    agent_code = row.get("AGENT_CIBLE", "")
                    agent_contact = get_agent_contact(settings, agent_code)
                    agent_whatsapp = agent_contact.get("whatsapp", "")
                    agent_email = agent_contact.get("email", "")
                    wa_url = build_source_excel_whatsapp_url(agent_whatsapp, row, etat_col)

                    st.markdown('<div class="glass-card">', unsafe_allow_html=True)

                    c1, c2, c3 = st.columns([5, 1.8, 1.6])

                    with c1:
                        demande_value = row.get(demande_col, "") if demande_col else ""
                        secteur_value = row.get(secteur_col, "") if secteur_col else ""
                        etat_value = row.get(etat_col, "")

                        st.markdown(
                            f"""
**Demande :** {demande_value}  
**Secteur :** {secteur_value}  
**État source :** {etat_value}  
**Agent cible :** {agent_code}
"""
                        )

                        st.markdown(
                            f"""
<span class="info-chip">WhatsApp agent : {agent_whatsapp or 'Non configuré'}</span>
<span class="info-chip">Email agent : {agent_email or 'Non configuré'}</span>
                            """,
                            unsafe_allow_html=True
                        )

                    with c2:
                        if st.session_state["is_admin"]:
                            if st.button("📧 Envoyer agent", key=f"excel_email_{idx}"):
                                current_config = load_email_config()
                                password = st.session_state.get("smtp_password", "") or get_secret("smtp_password", "")
                                recipient = agent_email.strip() or current_config.get("default_recipient", "").strip()
                                subject = f"Dispatch {agent_code} - {demande_value}"
                                message = build_source_excel_message(row, etat_col)

                                if not current_config.get("sender_email"):
                                    st.error("Configure l'email expéditeur dans la sidebar admin.")
                                elif not password:
                                    st.error("Ajoute le mot de passe SMTP dans la sidebar admin.")
                                elif not recipient:
                                    st.error("Aucun email agent ni email par défaut.")
                                else:
                                    try:
                                        send_email_smtp(current_config, password, recipient, subject, message)
                                        append_email_history(
                                            {
                                                "timestamp": now_str(),
                                                "instance_id": f"excel_{idx}",
                                                "demande": demande_value,
                                                "agent": agent_code,
                                                "recipient": recipient,
                                                "subject": subject,
                                                "status": "SUCCESS",
                                                "error": "",
                                                "sent_by": st.session_state.get("admin_user", "")
                                            }
                                        )
                                        st.success(f"E-mail envoyé à {recipient}")
                                    except Exception as e:
                                        append_email_history(
                                            {
                                                "timestamp": now_str(),
                                                "instance_id": f"excel_{idx}",
                                                "demande": demande_value,
                                                "agent": agent_code,
                                                "recipient": recipient,
                                                "subject": subject,
                                                "status": "ERROR",
                                                "error": str(e),
                                                "sent_by": st.session_state.get("admin_user", "")
                                            }
                                        )
                                        st.error(f"Erreur e-mail : {e}")
                        else:
                            st.caption("Admin requis")

                    with c3:
                        if wa_url:
                            st.markdown(render_whatsapp_button(wa_url), unsafe_allow_html=True)
                        else:
                            st.caption("WhatsApp non configuré")

                    with st.expander(f"Détail ligne Excel - {demande_value if demande_value else idx}"):
                        st.text_area(
                            "Message WhatsApp / Email",
                            value=build_source_excel_message(row, etat_col),
                            height=220,
                            key=f"excel_msg_{idx}"
                        )
                        st.dataframe(pd.DataFrame([row]), use_container_width=True)

                    st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("### Tableau source complet")
        st.dataframe(filtered_etat, use_container_width=True, height=420)

    else:
        st.info("Le fichier source ETAT FTTH RTC RTCL.xlsx n'a pas été chargé.")


# =========================================================
# PAGE RAPPORTS
# =========================================================
elif page == "📊 RAPPORTS":
    st.subheader("Rapports et statistiques avancées")

    saved_df = load_saved_instances()
    email_history_df = load_email_history()
    etat_df = safe_load_excel(ETAT_FILE, ETAT_SHEET, "ETAT FTTH RTC RTCL.xlsx")
    motif_df = safe_load_excel(MOTIF_FILE, MOTIF_SHEET, "MOTIF TOTAL (1).xlsx")

    tabs = st.tabs(["📈 Opérationnel", "📧 Historique e-mails", "📄 Source Excel"])

    with tabs[0]:
        if saved_df.empty:
            st.info("Aucune instance saisie pour le moment.")
        else:
            for col in ["secteur", "agent", "motif", "utilisateur", "statut_email", "statut_whatsapp"]:
                if col not in saved_df.columns:
                    saved_df[col] = ""

            total_instances = len(saved_df)
            emails_sent = int((saved_df["statut_email"].astype(str) == "Envoyé").sum())
            whatsapp_sent = int((saved_df["statut_whatsapp"].astype(str) == "Envoyé").sum())
            agents_count = saved_df["agent"].astype(str).replace("nan", "").replace("", pd.NA).dropna().nunique()
            secteurs_count = saved_df["secteur"].astype(str).replace("nan", "").replace("", pd.NA).dropna().nunique()

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("Instances", total_instances)
            k2.metric("Emails envoyés", emails_sent)
            k3.metric("WhatsApp envoyés", whatsapp_sent)
            k4.metric("Secteurs actifs", secteurs_count)

            st.markdown("---")

            c1, c2 = st.columns(2)

            with c1:
                secteur_count = (
                    saved_df["secteur"].fillna("Non renseigné").astype(str).value_counts().reset_index()
                )
                secteur_count.columns = ["Secteur", "Nombre"]
                fig1 = px.bar(secteur_count, x="Secteur", y="Nombre", color="Nombre", title="Instances par secteur")
                st.plotly_chart(fig1, use_container_width=True)

            with c2:
                agent_count = (
                    saved_df["agent"].fillna("Non renseigné").astype(str).value_counts().reset_index()
                )
                agent_count.columns = ["Agent", "Nombre"]
                fig2 = px.bar(agent_count, x="Agent", y="Nombre", color="Nombre", title="Instances par agent")
                st.plotly_chart(fig2, use_container_width=True)

            c3, c4 = st.columns(2)

            with c3:
                motif_count = (
                    saved_df["motif"].fillna("Non renseigné").astype(str).value_counts().head(15).reset_index()
                )
                motif_count.columns = ["Motif", "Nombre"]
                fig3 = px.bar(motif_count, x="Motif", y="Nombre", color="Nombre", title="Top 15 motifs")
                fig3.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig3, use_container_width=True)

            with c4:
                email_status = (
                    saved_df["statut_email"].fillna("Non renseigné").astype(str).value_counts().reset_index()
                )
                email_status.columns = ["Statut", "Nombre"]
                fig4 = px.pie(email_status, values="Nombre", names="Statut", title="Statut des e-mails")
                st.plotly_chart(fig4, use_container_width=True)

            st.markdown("---")
            st.dataframe(saved_df, use_container_width=True, height=340)

    with tabs[1]:
        st.subheader("Historique des e-mails envoyés")
        if email_history_df.empty:
            st.info("Aucun e-mail envoyé pour le moment.")
        else:
            search_mail = st.text_input("Recherche historique e-mail", placeholder="demande, destinataire, agent...")
            hist_filtered = global_search(email_history_df, search_mail)
            st.dataframe(hist_filtered, use_container_width=True, height=380)

            e1, e2 = st.columns(2)
            with e1:
                st.download_button(
                    "⬇️ Export historique CSV",
                    data=hist_filtered.to_csv(index=False).encode("utf-8"),
                    file_name="historique_emails.csv",
                    mime="text/csv"
                )
            with e2:
                st.download_button(
                    "⬇️ Export historique Excel",
                    data=to_excel_bytes(hist_filtered, "Emails"),
                    file_name="historique_emails.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    with tabs[2]:
        if etat_df.empty and motif_df.empty:
            st.info("Aucune source Excel chargée.")
        else:
            if not etat_df.empty:
                st.markdown("### Analyse source commandes")
                secteur_col = find_column(etat_df, ["secteur", "sector"])
                etat_col = find_column(etat_df, ["etat", "état", "state"])
                delai_col = find_column(etat_df, ["delai", "délai"])

                x1, x2, x3 = st.columns(3)
                x1.metric("Total commandes source", len(etat_df))
                x2.metric("Commandes VA", int((etat_df[etat_col].astype(str).str.upper() == "VA").sum()) if etat_col else 0)
                x3.metric("Délai moyen source", safe_mean_numeric(etat_df[delai_col]) if delai_col else "N/A")

                if secteur_col:
                    secteur_source = etat_df[secteur_col].fillna("Non renseigné").astype(str).value_counts().reset_index()
                    secteur_source.columns = ["Secteur", "Nombre"]
                    fig5 = px.bar(secteur_source, x="Secteur", y="Nombre", color="Nombre", title="Commandes source par secteur")
                    st.plotly_chart(fig5, use_container_width=True)

            if not motif_df.empty:
                st.markdown("### Analyse source motifs")
                motif_col = find_column(motif_df, ["motif", "detail", "pc mauvais"])
                if motif_col:
                    motif_source = motif_df[motif_col].fillna("Non renseigné").astype(str).value_counts().head(15).reset_index()
                    motif_source.columns = ["Motif", "Nombre"]
                    fig6 = px.bar(motif_source, x="Motif", y="Nombre", color="Nombre", title="Top 15 motifs source")
                    fig6.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig6, use_container_width=True)


# =========================================================
# PAGES SECONDAIRES
# =========================================================
elif page == "⚠️ DÉRANGEMENTS":
    st.subheader("Dérangements")
    st.info("Page prête pour évolution future.")

elif page == "🔧 FIABILISATION":
    st.subheader("Fiabilisation")
    st.info("Page prête pour évolution future.")

elif page == "⚖️ LITIGES":
    st.subheader("Litiges")
    st.info("Page prête pour évolution future.")


st.caption("InstalPro - Pilotage opérationnel des interventions Fibre & RTC")
