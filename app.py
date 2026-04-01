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
LOCAL_ETAT_FILE = BASE_DIR / "ETAT FTTH RTC RTCL.xlsx"
LOCAL_MOTIF_FILE = BASE_DIR / "MOTIF TOTAL (1).xlsx"

SETTINGS_FILE = BASE_DIR / "parametres_app.json"
LOGO_FILE = BASE_DIR / "logo_maroc_telecom.png"
ASSIGNMENTS_FILE = BASE_DIR / "affectations_agents.csv"
FEEDBACK_FILE = BASE_DIR / "retours_intervention.csv"
SMTP_CONFIG_FILE = BASE_DIR / "smtp_config.json"

ETAT_SHEET = "SITUATION14.15"
MOTIF_SHEET = "MOTIF"


# =========================================================
# STYLE
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

    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] .stMarkdown,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] h4,
    [data-testid="stSidebar"] h5,
    [data-testid="stSidebar"] h6,
    [data-testid="stSidebar"] span,
    [data-testid="stSidebar"] div {
        color: white !important;
    }

    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea,
    [data-testid="stSidebar"] [data-baseweb="input"] input,
    [data-testid="stSidebar"] [data-baseweb="base-input"] input,
    [data-testid="stSidebar"] [data-baseweb="select"] > div,
    [data-testid="stSidebar"] [data-baseweb="base-input"] > div,
    [data-testid="stSidebar"] [data-baseweb="textarea"] textarea {
        background: white !important;
        color: #0f172a !important;
        -webkit-text-fill-color: #0f172a !important;
        border-radius: 10px !important;
    }

    [data-testid="stSidebar"] input::placeholder,
    [data-testid="stSidebar"] textarea::placeholder {
        color: #64748b !important;
        -webkit-text-fill-color: #64748b !important;
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

    .assign-button {
        display: inline-block;
        width: 100%;
        text-align: center;
        padding: 11px 14px;
        background: linear-gradient(135deg, #2563eb, #1d4ed8);
        color: white !important;
        text-decoration: none !important;
        border-radius: 12px;
        font-weight: 700;
        box-shadow: 0 8px 20px rgba(37, 99, 235, 0.28);
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

    .section-title {
        font-size: 1.05rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.5rem;
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


def find_column(df, keywords):
    if df is None or df.empty:
        return None

    normalized_keywords = [normalize_text(k) for k in keywords]

    for col in df.columns:
        col_name = normalize_text(col)
        if any(k in col_name for k in normalized_keywords):
            return col

    return None


def normalize_intervention_code(value):
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


def normalize_product(value):
    txt = normalize_text(value).replace(" ", "").replace("_", "").replace("-", "")
    mapping = {
        "ftth": "FTTH",
        "ftthdfo": "FTTHDFO",
        "rtc": "RTC",
        "rtcdtl": "RTCDTL",
    }
    return mapping.get(txt, str(value).strip().upper())


def prepare_col_a_filter(df):
    if df is None or df.empty:
        return df, None

    df = df.copy()
    col_a = df.columns[0]
    parsed_dates = pd.to_datetime(df[col_a], errors="coerce", dayfirst=True)

    if parsed_dates.notna().sum() > 0:
        df["_col_a_filter_"] = parsed_dates.dt.date
    else:
        df["_col_a_filter_"] = df[col_a].astype(str).str.strip()

    return df, col_a


def collect_col_a_values(*dfs):
    values = []
    for df in dfs:
        if isinstance(df, pd.DataFrame) and not df.empty and "_col_a_filter_" in df.columns:
            values.extend(df["_col_a_filter_"].dropna().tolist())

    unique_values = []
    for v in values:
        if v not in unique_values:
            unique_values.append(v)

    return unique_values


def filter_by_col_a_value(df, selected_value):
    if df is None or df.empty:
        return df
    if "_col_a_filter_" not in df.columns:
        return df
    if selected_value in [None, ""]:
        return df
    return df[df["_col_a_filter_"] == selected_value]


def sanitize_row_dict(row_dict):
    clean = {}
    for k, v in row_dict.items():
        if str(k).startswith("_"):
            continue
        if pd.isna(v):
            clean[str(k)] = ""
        else:
            clean[str(k)] = str(v)
    return clean


def make_row_id(row_dict):
    payload = json.dumps(sanitize_row_dict(row_dict), ensure_ascii=False, sort_keys=True)
    return hashlib.md5(payload.encode("utf-8")).hexdigest()


def build_full_row_message(row_dict, title="Intervention terrain"):
    row_dict = sanitize_row_dict(row_dict)
    lines = [title, ""]
    for col, val in row_dict.items():
        if val == "":
            continue
        lines.append(f"{col} : {val}")
    lines.append("")
    lines.append("Merci d'intervenir et de faire le retour.")
    lines.append("InstalPro")
    return "\n".join(lines)


def build_whatsapp_url(phone, message):
    phone = clean_phone(phone)
    if not phone:
        return ""
    return f"https://wa.me/{phone}?text={quote(message)}"


def default_feedback_record():
    return {
        "sr": "",
        "tt": "",
        "pc": "",
        "port": "",
        "rosasse": "",
        "msan_port": "",
        "cable": "",
        "numero_validation": "",
        "msan_slot_port_sn": "",
        "metre_ftth": "",
        "autre_consomable": "",
        "commentaire": "",
        "email_status": "",
        "email_date": "",
    }


# =========================================================
# FILE LOADERS
# =========================================================
def load_excel_from_upload_or_local(uploaded_file, local_path, sheet_name, label):
    if uploaded_file is not None:
        try:
            uploaded_bytes = BytesIO(uploaded_file.getvalue())
            return pd.read_excel(uploaded_bytes, sheet_name=sheet_name)
        except Exception as e:
            st.warning(f"Erreur lecture fichier importé {label} : {e}")
            return pd.DataFrame()

    if local_path.exists():
        try:
            return pd.read_excel(local_path, sheet_name=sheet_name)
        except Exception as e:
            st.warning(f"Erreur lecture fichier local {label} : {e}")
            return pd.DataFrame()

    return pd.DataFrame()


# =========================================================
# JSON / CSV STORAGE
# =========================================================
def load_json(path, default_data):
    if not path.exists():
        save_json(path, default_data)
        return default_data
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        save_json(path, default_data)
        return default_data


def save_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_csv(path):
    if path.exists():
        try:
            return pd.read_csv(path)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()


def upsert_csv_record(path, key_col, record):
    df = load_csv(path)
    record_df = pd.DataFrame([record])

    if df.empty or key_col not in df.columns:
        record_df.to_csv(path, index=False)
        return

    mask = df[key_col].astype(str) == str(record[key_col])
    if mask.any():
        for col, val in record.items():
            if col not in df.columns:
                df[col] = ""
            df.loc[mask, col] = val
    else:
        df = pd.concat([df, record_df], ignore_index=True)

    df.to_csv(path, index=False)


# =========================================================
# SETTINGS
# =========================================================
def default_settings():
    return {
        "utilisateurs": ["admin"],
        "secteurs": ["MHAMID", "BOUAAKAZ", "Province M'HAMID"],
        "agents": ["Agent 1", "Agent 2"],
        "agent_contacts": {
            "Agent 1": {"whatsapp": ""},
            "Agent 2": {"whatsapp": ""}
        },
        "admin_username": "admin",
        "admin_password_hash": hash_password("admin123")
    }


def sync_agent_contacts(settings):
    if "agent_contacts" not in settings or not isinstance(settings["agent_contacts"], dict):
        settings["agent_contacts"] = {}

    for agent in settings.get("agents", []):
        if agent not in settings["agent_contacts"]:
            settings["agent_contacts"][agent] = {"whatsapp": ""}

    removed = [k for k in settings["agent_contacts"].keys() if k not in settings.get("agents", [])]
    for k in removed:
        del settings["agent_contacts"][k]

    return settings


def load_settings():
    settings = load_json(SETTINGS_FILE, default_settings())
    settings = sync_agent_contacts(settings)
    save_json(SETTINGS_FILE, settings)
    return settings


def save_settings(settings):
    save_json(SETTINGS_FILE, settings)


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
        settings["agent_contacts"][value] = {"whatsapp": ""}

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

    idx = settings[key].index(old_value)
    settings[key][idx] = new_value
    settings[key] = sorted(settings[key], key=lambda x: x.lower())

    if key == "agents":
        settings["agent_contacts"][new_value] = settings["agent_contacts"].get(old_value, {"whatsapp": ""})
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
    return settings.get("agent_contacts", {}).get(agent_name, {"whatsapp": ""})


def update_agent_contact(settings, agent_name, whatsapp):
    settings = sync_agent_contacts(settings)
    settings["agent_contacts"][agent_name] = {"whatsapp": str(whatsapp).strip()}
    save_settings(settings)


# =========================================================
# SMTP CONFIG
# =========================================================
def default_smtp_config():
    return {
        "smtp_server": "smtp.office365.com",
        "smtp_port": 587,
        "sender_email": "",
        "sender_password": "",
        "recipient_email": ""
    }


def load_smtp_config():
    return load_json(SMTP_CONFIG_FILE, default_smtp_config())


def save_smtp_config(config):
    save_json(SMTP_CONFIG_FILE, config)


def send_email_smtp(config, subject, body):
    sender = config.get("sender_email", "").strip()
    password = config.get("sender_password", "").strip()
    recipient = config.get("recipient_email", "").strip()
    smtp_server = config.get("smtp_server", "").strip()
    smtp_port = int(config.get("smtp_port", 587))

    if not sender or not password or not recipient or not smtp_server:
        raise ValueError("Configuration SMTP incomplète.")

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient
    msg.set_content(body)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)


def build_feedback_email_subject(feedback_record):
    return f"Retour intervention - {feedback_record.get('commande', '')} - {feedback_record.get('agent_name', '')}"


def build_feedback_email_body(feedback_record):
    lines = [
        "Retour intervention terrain",
        "",
    ]
    for k, v in feedback_record.items():
        if v is None or str(v).strip() == "":
            continue
        lines.append(f"{k} : {v}")
    return "\n".join(lines)


# =========================================================
# AUTH
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
# ADMIN UI HELPERS
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


def render_manager_tab(settings, key, label):
    current_items = settings.get(key, [])

    st.markdown(f"### Gestion des {label.lower()}s")
    st.write(current_items)

    c1, c2, c3 = st.columns(3)

    with c1:
        with st.form(f"add_{key}_form"):
            new_value = st.text_input(f"Nouveau {label.lower()}", key=f"new_{key}")
            submitted = st.form_submit_button("➕ Ajouter")
            if submitted:
                ok, msg = add_item_to_settings(settings, key, new_value)
                if ok:
                    st.success(msg)
                    rerun_app()
                else:
                    st.error(msg)

    with c2:
        with st.form(f"edit_{key}_form"):
            old_value = st.selectbox(f"{label} à modifier", current_items, key=f"old_{key}")
            new_value = st.text_input("Nouveau nom", key=f"edit_{key}")
            submitted = st.form_submit_button("✏️ Modifier")
            if submitted:
                ok, msg = update_item_in_settings(settings, key, old_value, new_value)
                if ok:
                    st.success(msg)
                    rerun_app()
                else:
                    st.error(msg)

    with c3:
        with st.form(f"delete_{key}_form"):
            selected = st.selectbox(f"{label} à supprimer", current_items, key=f"del_{key}")
            submitted = st.form_submit_button("🗑️ Supprimer")
            if submitted:
                ok, msg = delete_item_in_settings(settings, key, selected)
                if ok:
                    st.success(msg)
                    rerun_app()
                else:
                    st.error(msg)


def render_agent_contacts_admin(settings):
    st.markdown("### Vrais agents et numéros WhatsApp")
    if not settings.get("agents"):
        st.info("Aucun agent disponible.")
        return

    selected_agent = st.selectbox("Agent", settings["agents"], key="contact_agent_select")
    current = get_agent_contact(settings, selected_agent)

    with st.form("agent_contact_form"):
        whatsapp_value = st.text_input("Numéro WhatsApp de l'agent", value=current.get("whatsapp", ""))
        submitted = st.form_submit_button("💾 Enregistrer contact agent")
        if submitted:
            update_agent_contact(settings, selected_agent, whatsapp_value)
            st.success(f"WhatsApp de l'agent {selected_agent} mis à jour.")
            rerun_app()

    st.markdown("#### Résumé")
    for agent in settings["agents"]:
        contact = get_agent_contact(settings, agent)
        st.markdown(f"- **{agent}** | WhatsApp : `{contact.get('whatsapp', '')}`")


# =========================================================
# INIT
# =========================================================
init_auth_state()
settings = load_settings()
smtp_config = load_smtp_config()
render_header()


# =========================================================
# SIDEBAR
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
        st.subheader("Email Outlook pour retour terrain")
        with st.form("smtp_form"):
            smtp_server = st.text_input("Serveur SMTP", value=smtp_config.get("smtp_server", "smtp.office365.com"))
            smtp_port = st.number_input("Port SMTP", min_value=1, max_value=9999, value=int(smtp_config.get("smtp_port", 587)))
            sender_email = st.text_input("Email expéditeur", value=smtp_config.get("sender_email", ""))
            sender_password = st.text_input("Mot de passe SMTP", type="password", value=smtp_config.get("sender_password", ""))
            recipient_email = st.text_input("Email destinataire retour", value=smtp_config.get("recipient_email", ""))

            smtp_btn = st.form_submit_button("💾 Enregistrer email Outlook")
            if smtp_btn:
                smtp_config = {
                    "smtp_server": smtp_server,
                    "smtp_port": int(smtp_port),
                    "sender_email": sender_email,
                    "sender_password": sender_password,
                    "recipient_email": recipient_email,
                }
                save_smtp_config(smtp_config)
                st.success("Configuration email enregistrée.")

        st.markdown("---")
        st.subheader("Logo")
        uploaded_logo = st.file_uploader("Uploader le logo Maroc Telecom", type=["png", "jpg", "jpeg"])
        c1, c2 = st.columns(2)

        with c1:
            if st.button("💾 Sauvegarder logo"):
                if uploaded_logo is not None:
                    with open(LOGO_FILE, "wb") as f:
                        f.write(uploaded_logo.getbuffer())
                    st.success("Logo sauvegardé.")
                    rerun_app()
                else:
                    st.warning("Choisis un logo d'abord.")

        with c2:
            if st.button("🗑️ Supprimer logo"):
                if LOGO_FILE.exists():
                    LOGO_FILE.unlink()
                    st.success("Logo supprimé.")
                    rerun_app()
                else:
                    st.info("Aucun logo à supprimer.")

        st.markdown("---")
        st.subheader("Sécurité admin")
        with st.form("admin_security_form"):
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
    ["🗂️ INSTANCES", "📈 RAPPORTS", "🚨 DÉRANGEMENTS", "🛠️ FIABILISATION", "🧾 LITIGES"],
    horizontal=True,
    label_visibility="collapsed"
)


# =========================================================
# DATA FOR PAGES
# =========================================================
assignments_df = load_csv(ASSIGNMENTS_FILE)
feedback_df = load_csv(FEEDBACK_FILE)


# =========================================================
# PAGE INSTANCES
# =========================================================
if page == "🗂️ INSTANCES":
    st.subheader("Import et dispatch des instances")

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

    st.markdown("### Import des fichiers")
    up1, up2 = st.columns(2)
    with up1:
        uploaded_etat = st.file_uploader("Importer ETAT FTTH RTC", type=["xlsx"], key="upload_etat")
    with up2:
        uploaded_motif = st.file_uploader("Importer MOTIF TOTAL", type=["xlsx"], key="upload_motif")

    etat_df = load_excel_from_upload_or_local(uploaded_etat, LOCAL_ETAT_FILE, ETAT_SHEET, "ETAT")
    motif_df = load_excel_from_upload_or_local(uploaded_motif, LOCAL_MOTIF_FILE, MOTIF_SHEET, "MOTIF")

    if etat_df.empty:
        st.warning("Importe le fichier ETAT FTTH RTC pour commencer.")
    else:
        etat_df, _ = prepare_col_a_filter(etat_df)
        motif_df, _ = prepare_col_a_filter(motif_df)

        col_a_values = collect_col_a_values(etat_df, motif_df)
        produit_col = find_column(etat_df, ["s.produit", "produit", "s produit"])
        motif_produit_col = find_column(motif_df, ["s.produit", "produit", "s produit"])

        f1, f2, f3 = st.columns([2, 2, 3])

        with f1:
            selected_day = st.selectbox(
                "Filtre journalier colonne A",
                col_a_values if col_a_values else [""],
                key="day_filter_instances"
            ) if col_a_values else ""

        with f2:
            selected_product = st.selectbox(
                "Filtre s.produit",
                ["Tous", "FTTH", "FTTHDFO", "RTC", "RTCDTL"],
                key="product_filter_instances"
            )

        with f3:
            search_query = st.text_input(
                "Recherche globale",
                placeholder="commande, adresse, contact, secteur..."
            )

        if selected_day:
            etat_df = filter_by_col_a_value(etat_df, selected_day)
            motif_df = filter_by_col_a_value(motif_df, selected_day)

        if produit_col and selected_product != "Tous":
            etat_df = etat_df[etat_df[produit_col].astype(str).apply(normalize_product) == selected_product]

        if motif_produit_col and selected_product != "Tous":
            motif_df = motif_df[motif_df[motif_produit_col].astype(str).apply(normalize_product) == selected_product]

        filtered_etat = global_search(etat_df, search_query).copy()

        etat_col = find_column(filtered_etat, ["etat", "état", "state"])
        secteur_col = find_column(filtered_etat, ["secteur", "sector"])
        demande_col = find_column(filtered_etat, ["demande", "commande", "reference", "référence"])
        produit_col_filtered = find_column(filtered_etat, ["s.produit", "produit", "s produit"])

        if etat_col:
            filtered_etat["CODE_INTERVENTION"] = filtered_etat[etat_col].apply(normalize_intervention_code)
            actionable_df = filtered_etat[
                filtered_etat["CODE_INTERVENTION"].isin(["NA", "RM", "TR", "TL"])
            ].copy()

            st.markdown("### Lignes dispatchables")
            st.caption("NA / RM / TR / TL = code d’intervention. Choisis ensuite un vrai agent.")

            if actionable_df.empty:
                st.info("Aucune ligne dispatchable avec code intervention NA / RM / TR / TL.")
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
                    row_dict = row.to_dict()
                    row_id = make_row_id(row_dict)

                    existing_assignment = pd.DataFrame()
                    if not assignments_df.empty and "row_id" in assignments_df.columns:
                        existing_assignment = assignments_df[assignments_df["row_id"].astype(str) == row_id]

                    existing_agent = ""
                    existing_whatsapp = ""
                    existing_sent_status = "Non envoyé"
                    if not existing_assignment.empty:
                        existing_agent = str(existing_assignment.iloc[0].get("agent_name", ""))
                        existing_whatsapp = str(existing_assignment.iloc[0].get("agent_whatsapp", ""))
                        existing_sent_status = str(existing_assignment.iloc[0].get("whatsapp_status", "Non envoyé"))

                    code_intervention = row.get("CODE_INTERVENTION", "")
                    demande_value = row.get(demande_col, "") if demande_col else ""
                    secteur_value = row.get(secteur_col, "") if secteur_col else ""
                    produit_value = row.get(produit_col_filtered, "") if produit_col_filtered else ""
                    produit_norm = normalize_product(produit_value)

                    default_agent_index = 0
                    if existing_agent and existing_agent in settings["agents"]:
                        default_agent_index = settings["agents"].index(existing_agent)

                    st.markdown('<div class="glass-card">', unsafe_allow_html=True)

                    top1, top2 = st.columns([4.8, 2.2])

                    with top1:
                        st.markdown(
                            f"""
**Commande :** {demande_value}  
**Secteur :** {secteur_value}  
**Produit :** {produit_norm}  
**Code intervention :** {code_intervention}
"""
                        )

                    with top2:
                        selected_real_agent = st.selectbox(
                            f"Choisir l'agent - ligne {idx}",
                            settings["agents"],
                            index=default_agent_index,
                            key=f"agent_pick_{row_id}"
                        )

                    current_agent_contact = get_agent_contact(settings, selected_real_agent)
                    selected_agent_whatsapp = current_agent_contact.get("whatsapp", "")
                    full_message = build_full_row_message(
                        row_dict,
                        title=f"Intervention terrain - {code_intervention}"
                    )
                    wa_url = build_whatsapp_url(selected_agent_whatsapp, full_message)

                    st.markdown(
                        f"""
<span class="info-chip">Agent choisi : {selected_real_agent}</span>
<span class="info-chip">WhatsApp : {selected_agent_whatsapp or 'Non configuré'}</span>
<span class="info-chip">Statut WhatsApp : {existing_sent_status}</span>
                        """,
                        unsafe_allow_html=True
                    )

                    a1, a2, a3 = st.columns([2.1, 1.8, 1.2])

                    with a1:
                        if st.button("💾 Affecter et enregistrer l’agent choisi", key=f"assign_{row_id}"):
                            assignment_record = {
                                "row_id": row_id,
                                "date_assignment": now_str(),
                                "agent_name": selected_real_agent,
                                "agent_whatsapp": selected_agent_whatsapp,
                                "code_intervention": code_intervention,
                                "produit": produit_norm,
                                "commande": demande_value,
                                "secteur": secteur_value,
                                "whatsapp_status": existing_sent_status,
                                "row_payload_json": json.dumps(sanitize_row_dict(row_dict), ensure_ascii=False)
                            }
                            upsert_csv_record(ASSIGNMENTS_FILE, "row_id", assignment_record)
                            st.success(f"Agent {selected_real_agent} affecté.")
                            rerun_app()

                    with a2:
                        if wa_url:
                            st.markdown(
                                f"""
<div class="wa-button">
    <a href="{wa_url}" target="_blank">💬 Envoyer à l’agent</a>
</div>
                                """,
                                unsafe_allow_html=True
                            )
                        else:
                            st.caption("WhatsApp agent non configuré")

                    with a3:
                        if st.session_state["is_admin"]:
                            if st.button("✅ Marquer WA", key=f"mark_wa_{row_id}"):
                                assignment_record = {
                                    "row_id": row_id,
                                    "date_assignment": now_str(),
                                    "agent_name": selected_real_agent,
                                    "agent_whatsapp": selected_agent_whatsapp,
                                    "code_intervention": code_intervention,
                                    "produit": produit_norm,
                                    "commande": demande_value,
                                    "secteur": secteur_value,
                                    "whatsapp_status": "Envoyé",
                                    "row_payload_json": json.dumps(sanitize_row_dict(row_dict), ensure_ascii=False)
                                }
                                upsert_csv_record(ASSIGNMENTS_FILE, "row_id", assignment_record)
                                st.success("WhatsApp marqué comme envoyé.")
                                rerun_app()

                    existing_feedback = pd.DataFrame()
                    if not feedback_df.empty and "row_id" in feedback_df.columns:
                        existing_feedback = feedback_df[feedback_df["row_id"].astype(str) == row_id]

                    fb = default_feedback_record()
                    if not existing_feedback.empty:
                        for k in fb.keys():
                            fb[k] = str(existing_feedback.iloc[0].get(k, ""))

                    with st.expander("📝 Saisie retour agent + envoi email Outlook"):
                        st.text_area(
                            "Message complet envoyé à l’agent",
                            value=full_message,
                            height=220,
                            key=f"msg_full_{row_id}"
                        )

                        with st.form(f"feedback_form_{row_id}"):
                            st.markdown("#### Retour terrain")
                            cfb1, cfb2 = st.columns(2)

                            with cfb1:
                                commentaire = st.text_area("Commentaire agent", value=fb["commentaire"], height=80)

                            with cfb2:
                                st.write(f"Produit détecté : **{produit_norm}**")
                                st.write(f"Code intervention : **{code_intervention}**")
                                st.write(f"Agent : **{selected_real_agent}**")

                            rtc_mode = produit_norm in ["RTC", "RTCDTL"]
                            ftth_mode = produit_norm in ["FTTH", "FTTHDFO"]

                            if rtc_mode:
                                st.markdown("#### Champs RTC / RTCDTL")
                                r1, r2, r3 = st.columns(3)
                                with r1:
                                    sr = st.text_input("SR", value=fb["sr"])
                                    tt = st.text_input("TT", value=fb["tt"])
                                with r2:
                                    pc = st.text_input("PC", value=fb["pc"])
                                    port = st.text_input("Port", value=fb["port"])
                                with r3:
                                    rosasse = st.text_input("Rosasse", value=fb["rosasse"])
                                    msan_port = st.text_input("MSAN.port", value=fb["msan_port"])

                                cable = st.selectbox(
                                    "Câble",
                                    ["", "1/6", "5/9"],
                                    index=(["", "1/6", "5/9"].index(fb["cable"]) if fb["cable"] in ["", "1/6", "5/9"] else 0),
                                    key=f"cable_{row_id}"
                                )

                                numero_validation = ""
                                msan_slot_port_sn = ""
                                metre_ftth = ""
                                autre_consomable = ""

                            elif ftth_mode:
                                st.markdown("#### Champs FTTH / FTTHDFO")
                                r1, r2 = st.columns(2)
                                with r1:
                                    numero_validation = st.text_input("Numéro de validation", value=fb["numero_validation"])
                                    msan_slot_port_sn = st.text_input("MSAN.slot.port.sn", value=fb["msan_slot_port_sn"])
                                with r2:
                                    metre_ftth = st.text_input("Combien de mètre FTTH", value=fb["metre_ftth"])
                                    autre_consomable = st.text_input("Autre consommable", value=fb["autre_consomable"])

                                sr = ""
                                tt = ""
                                pc = ""
                                port = ""
                                rosasse = ""
                                msan_port = ""
                                cable = ""

                            else:
                                st.markdown("#### Champs libres")
                                sr = ""
                                tt = ""
                                pc = ""
                                port = ""
                                rosasse = ""
                                msan_port = ""
                                cable = ""
                                numero_validation = ""
                                msan_slot_port_sn = ""
                                metre_ftth = ""
                                autre_consomable = ""

                            submit_feedback = st.form_submit_button("📤 Enregistrer la saisie et envoyer l’email Outlook")

                            if submit_feedback:
                                feedback_record = {
                                    "row_id": row_id,
                                    "date_feedback": now_str(),
                                    "commande": demande_value,
                                    "secteur": secteur_value,
                                    "produit": produit_norm,
                                    "code_intervention": code_intervention,
                                    "agent_name": selected_real_agent,
                                    "agent_whatsapp": selected_agent_whatsapp,
                                    "sr": sr,
                                    "tt": tt,
                                    "pc": pc,
                                    "port": port,
                                    "rosasse": rosasse,
                                    "msan_port": msan_port,
                                    "cable": cable,
                                    "numero_validation": numero_validation,
                                    "msan_slot_port_sn": msan_slot_port_sn,
                                    "metre_ftth": metre_ftth,
                                    "autre_consomable": autre_consomable,
                                    "commentaire": commentaire,
                                    "email_status": "Non envoyé",
                                    "email_date": "",
                                }

                                try:
                                    subject = build_feedback_email_subject(feedback_record)
                                    body = build_feedback_email_body(feedback_record)
                                    send_email_smtp(load_smtp_config(), subject, body)
                                    feedback_record["email_status"] = "Envoyé"
                                    feedback_record["email_date"] = now_str()
                                except Exception as e:
                                    feedback_record["email_status"] = f"Erreur: {e}"
                                    feedback_record["email_date"] = now_str()

                                upsert_csv_record(FEEDBACK_FILE, "row_id", feedback_record)
                                st.success("Retour enregistré. Vérifie le statut email.")
                                rerun_app()

                    st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("### Tableau ETAT FTTH RTC filtré")
        st.dataframe(
            filtered_etat.drop(columns=["_col_a_filter_"], errors="ignore"),
            use_container_width=True,
            height=380
        )

        if not motif_df.empty:
            st.markdown("### Tableau MOTIF TOTAL filtré")
            st.dataframe(
                motif_df.drop(columns=["_col_a_filter_"], errors="ignore"),
                use_container_width=True,
                height=280
            )


# =========================================================
# PAGE RAPPORTS
# =========================================================
elif page == "📈 RAPPORTS":
    st.subheader("Rapports et statistiques")

    etat_df = load_excel_from_upload_or_local(None, LOCAL_ETAT_FILE, ETAT_SHEET, "ETAT")
    motif_df = load_excel_from_upload_or_local(None, LOCAL_MOTIF_FILE, MOTIF_SHEET, "MOTIF")
    assignments_df = load_csv(ASSIGNMENTS_FILE)
    feedback_df = load_csv(FEEDBACK_FILE)

    etat_df, _ = prepare_col_a_filter(etat_df)
    motif_df, _ = prepare_col_a_filter(motif_df)

    col_a_values = collect_col_a_values(etat_df, motif_df)
    if col_a_values:
        selected_col_a = st.selectbox(
            "Filtre journalier (colonne A)",
            col_a_values,
            key="report_day_filter"
        )
        etat_df = filter_by_col_a_value(etat_df, selected_col_a)
        motif_df = filter_by_col_a_value(motif_df, selected_col_a)

    etat_prod_col = find_column(etat_df, ["s.produit", "produit", "s produit"])
    motif_prod_col = find_column(motif_df, ["s.produit", "produit", "s produit"])

    selected_produit = st.selectbox(
        "Filtre s.produit",
        ["Tous", "FTTH", "FTTHDFO", "RTC", "RTCDTL"],
        key="report_product_filter"
    )

    if selected_produit != "Tous":
        if etat_prod_col:
            etat_df = etat_df[etat_df[etat_prod_col].astype(str).apply(normalize_product) == selected_produit]
        if motif_prod_col:
            motif_df = motif_df[motif_prod_col.astype(str).apply(normalize_product) == selected_produit]

    etat_df = etat_df.drop(columns=["_col_a_filter_"], errors="ignore")
    motif_df = motif_df.drop(columns=["_col_a_filter_"], errors="ignore")

    tabs = st.tabs(["📊 KPI", "👷 Affectations", "📝 Retours", "📄 Source Excel"])

    with tabs[0]:
        total_affectations = len(assignments_df) if not assignments_df.empty else 0
        wa_sent = 0
        if not assignments_df.empty and "whatsapp_status" in assignments_df.columns:
            wa_sent = int((assignments_df["whatsapp_status"].astype(str) == "Envoyé").sum())

        total_retours = len(feedback_df) if not feedback_df.empty else 0
        email_ok = 0
        if not feedback_df.empty and "email_status" in feedback_df.columns:
            email_ok = int((feedback_df["email_status"].astype(str) == "Envoyé").sum())

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Affectations", total_affectations)
        k2.metric("WhatsApp envoyés", wa_sent)
        k3.metric("Retours saisis", total_retours)
        k4.metric("Emails envoyés", email_ok)

        st.markdown("---")

        if not assignments_df.empty and "agent_name" in assignments_df.columns:
            agent_count = assignments_df["agent_name"].fillna("Non renseigné").astype(str).value_counts().reset_index()
            agent_count.columns = ["Agent", "Nombre"]
            fig1 = px.bar(agent_count, x="Agent", y="Nombre", color="Nombre", title="Affectations par agent")
            st.plotly_chart(fig1, use_container_width=True)

        if not feedback_df.empty and "produit" in feedback_df.columns:
            prod_count = feedback_df["produit"].fillna("Non renseigné").astype(str).value_counts().reset_index()
            prod_count.columns = ["Produit", "Nombre"]
            fig2 = px.pie(prod_count, values="Nombre", names="Produit", title="Retours par produit")
            st.plotly_chart(fig2, use_container_width=True)

    with tabs[1]:
        if assignments_df.empty:
            st.info("Aucune affectation enregistrée.")
        else:
            st.dataframe(assignments_df, use_container_width=True, height=380)
            st.download_button(
                "⬇️ Export affectations Excel",
                data=to_excel_bytes(assignments_df, "Affectations"),
                file_name="affectations_agents.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with tabs[2]:
        if feedback_df.empty:
            st.info("Aucun retour terrain enregistré.")
        else:
            st.dataframe(feedback_df, use_container_width=True, height=380)
            st.download_button(
                "⬇️ Export retours Excel",
                data=to_excel_bytes(feedback_df, "Retours"),
                file_name="retours_intervention.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with tabs[3]:
        if not etat_df.empty:
            st.markdown("### ETAT FTTH RTC")
            st.dataframe(etat_df, use_container_width=True, height=280)
        if not motif_df.empty:
            st.markdown("### MOTIF TOTAL")
            st.dataframe(motif_df, use_container_width=True, height=280)


# =========================================================
# AUTRES PAGES
# =========================================================
elif page == "🚨 DÉRANGEMENTS":
    st.subheader("Dérangements")
    st.info("Page prête pour évolution future.")

elif page == "🛠️ FIABILISATION":
    st.subheader("Fiabilisation")
    st.info("Page prête pour évolution future.")

elif page == "🧾 LITIGES":
    st.subheader("Litiges")
    st.info("Page prête pour évolution future.")


st.caption("InstalPro - Dispatch WhatsApp, affectation agent, retour terrain et email Outlook")
