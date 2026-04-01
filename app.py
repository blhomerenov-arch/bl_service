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
APP_NAME = "Système Centralisé de Gestion, Dispatching et Suivi des Interventions Fibre & RTC"
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
# STYLE
# =========================================================
st.markdown(
    """
    <style>
    .main-header {
        background: linear-gradient(90deg, #0E7CFF, #0A58CA);
        color: white;
        padding: 20px;
        border-radius: 14px;
        text-align: center;
        margin-bottom: 12px;
    }
    .card-box {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 12px;
        padding: 14px;
        margin-bottom: 12px;
    }
    .wa-button a {
        display: inline-block;
        width: 100%;
        text-align: center;
        padding: 10px 14px;
        background: #25D366;
        color: white !important;
        text-decoration: none !important;
        border-radius: 10px;
        font-weight: 700;
    }
    .small-muted {
        color: #6c757d;
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


def normalize_text(value):
    value = "" if value is None else str(value)
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("utf-8")
    return value.lower().strip()


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def generate_instance_id():
    return datetime.now().strftime("%Y%m%d%H%M%S%f")


def clean_phone_for_whatsapp(phone):
    if phone is None:
        return ""
    return "".join(ch for ch in str(phone) if ch.isdigit())


def to_excel_bytes(df, sheet_name="Data"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
    output.seek(0)
    return output.getvalue()


def safe_mean_numeric(series):
    s = pd.to_numeric(series, errors="coerce")
    if s.notna().any():
        return round(s.mean(), 1)
    return None


def global_search(df, query):
    if df.empty or not query:
        return df

    mask = df.astype(str).apply(lambda col: col.str.contains(query, case=False, na=False))
    return df[mask.any(axis=1)]


def find_column(df, keywords):
    if df is None or df.empty:
        return None

    keywords = [normalize_text(k) for k in keywords]

    for col in df.columns:
        col_name = normalize_text(col)
        if any(k in col_name for k in keywords):
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


def get_secret(name, default=""):
    try:
        return st.secrets[name]
    except Exception:
        return default


# =========================================================
# SETTINGS / CONFIG
# =========================================================
def default_settings():
    return {
        "utilisateurs": ["admin"],
        "agents": ["hamid", "SHAKHMAN"],
        "secteurs": ["MHAMID", "BOUAAKAZ", "Province M'HAMID"],
        "admin_username": "admin",
        "admin_password_hash": hash_password("admin123")
    }


def save_settings(data):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


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
    save_settings(settings)
    return True, f"{old_value} modifié en {new_value}."


def delete_item_in_settings(settings, key, value):
    value = str(value).strip()

    if value not in settings.get(key, []):
        return False, "Élément introuvable."

    if len(settings[key]) <= 1:
        return False, f"Impossible de supprimer le dernier {key[:-1] if key.endswith('s') else key}."

    settings[key].remove(value)
    save_settings(settings)
    return True, f"{value} supprimé avec succès."


def default_email_config():
    return {
        "smtp_server": "smtp.gmail.com",
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
    good_user = username == settings.get("admin_username", "admin")
    good_pass = hash_password(password) == settings.get("admin_password_hash", "")
    if good_user and good_pass:
        st.session_state["is_admin"] = True
        st.session_state["admin_user"] = username
        return True
    return False


def admin_logout():
    st.session_state["is_admin"] = False
    st.session_state["admin_user"] = ""


# =========================================================
# DATA SAVE / LOAD
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
# EMAIL
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
Technicien : {row.get('technicien_nom', '')}
Motif : {row.get('motif', '')}
Utilisateur saisie : {row.get('utilisateur', '')}

Cordialement,
{APP_NAME}
""".strip()


def build_whatsapp_url(row):
    phone = clean_phone_for_whatsapp(row.get("technicien_whatsapp", ""))
    if not phone:
        return ""
    text = build_instance_message(row)
    return f"https://wa.me/{phone}?text={quote(text)}"


def render_whatsapp_button(url):
    return f"""
    <div class="wa-button">
        <a href="{url}" target="_blank">💬 Envoyer par WhatsApp</a>
    </div>
    """


# =========================================================
# RENDER HEADER
# =========================================================
def render_header():
    if LOGO_FILE.exists():
        c1, c2 = st.columns([1, 4])
        with c1:
            st.image(str(LOGO_FILE), width=180)
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
# ADMIN PANELS
# =========================================================
def render_manager_tab(settings, key, label):
    current_items = settings.get(key, [])

    st.markdown(f"### Gestion des {label.lower()}s")
    if current_items:
        st.write(current_items)
    else:
        st.info(f"Aucun {label.lower()}.")

    a1, a2, a3 = st.columns(3)

    with a1:
        with st.form(f"add_{key}_form"):
            st.markdown(f"**Ajouter {label.lower()}**")
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
            st.markdown(f"**Modifier {label.lower()}**")
            old_value = st.selectbox(f"{label} à modifier", current_items, key=f"old_{key}")
            new_value = st.text_input(f"Nouveau nom", key=f"edit_{key}")
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
            st.markdown(f"**Supprimer {label.lower()}**")
            selected = st.selectbox(f"{label} à supprimer", current_items, key=f"del_{key}")
            submit = st.form_submit_button("🗑️ Supprimer")
            if submit:
                ok, msg = delete_item_in_settings(settings, key, selected)
                if ok:
                    st.success(msg)
                    rerun_app()
                else:
                    st.error(msg)


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
        st.caption("Identifiants par défaut : admin / admin123")
    else:
        st.success(f"Connecté en admin : {st.session_state['admin_user']}")
        if st.button("🚪 Déconnexion admin"):
            admin_logout()
            rerun_app()

        st.markdown("---")
        st.subheader("Configuration e-mail")
        with st.form("email_config_form"):
            smtp_server = st.text_input("Serveur SMTP", value=email_config.get("smtp_server", "smtp.gmail.com"))
            smtp_port = st.number_input("Port SMTP", min_value=1, max_value=9999, value=int(email_config.get("smtp_port", 587)))
            sender_email = st.text_input("Email expéditeur", value=email_config.get("sender_email", ""))
            smtp_password = st.text_input(
                "Mot de passe / mot de passe d'application",
                type="password",
                value=st.session_state.get("smtp_password", "")
            )
            default_recipient = st.text_input("Destinataire par défaut", value=email_config.get("default_recipient", ""))
            save_email_btn = st.form_submit_button("💾 Enregistrer config e-mail")

            if save_email_btn:
                save_email_config({
                    "smtp_server": smtp_server,
                    "smtp_port": int(smtp_port),
                    "sender_email": sender_email,
                    "default_recipient": default_recipient
                })
                st.session_state["smtp_password"] = smtp_password
                st.success("Configuration e-mail enregistrée.")

        st.markdown("---")
        st.subheader("Logo Maroc Telecom")
        uploaded_logo = st.file_uploader("Uploader le logo", type=["png", "jpg", "jpeg"])
        c_logo1, c_logo2 = st.columns(2)

        with c_logo1:
            if st.button("💾 Enregistrer le logo"):
                if uploaded_logo is not None:
                    with open(LOGO_FILE, "wb") as f:
                        f.write(uploaded_logo.getbuffer())
                    st.success("Logo enregistré.")
                    rerun_app()
                else:
                    st.warning("Choisis d'abord un fichier logo.")

        with c_logo2:
            if st.button("🗑️ Supprimer le logo"):
                if LOGO_FILE.exists():
                    LOGO_FILE.unlink()
                    st.success("Logo supprimé.")
                    rerun_app()
                else:
                    st.info("Aucun logo enregistré.")

        st.markdown("---")
        st.subheader("Sécurité admin")
        with st.form("change_admin_credentials"):
            new_admin_user = st.text_input("Nouveau nom admin", value=settings.get("admin_username", "admin"))
            new_admin_password = st.text_input("Nouveau mot de passe admin", type="password")
            save_admin_btn = st.form_submit_button("🔑 Mettre à jour l'admin")
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
        with st.expander("⚙️ Administration V2 : utilisateurs / agents / secteurs", expanded=False):
            tabs = st.tabs(["Utilisateurs", "Agents", "Secteurs"])
            with tabs[0]:
                render_manager_tab(settings, "utilisateurs", "Utilisateur")
            with tabs[1]:
                render_manager_tab(settings, "agents", "Agent")
            with tabs[2]:
                render_manager_tab(settings, "secteurs", "Secteur")

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
            agent = st.selectbox("Agent", settings["agents"])

        st.markdown("### Affectation technicien")
        t1, t2, t3 = st.columns(3)
        with t1:
            technicien_nom = st.text_input("Nom technicien")
        with t2:
            technicien_whatsapp = st.text_input("WhatsApp technicien", placeholder="2126XXXXXXXX")
        with t3:
            technicien_email = st.text_input("Email technicien")

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
                    "technicien_nom": technicien_nom,
                    "technicien_whatsapp": technicien_whatsapp,
                    "technicien_email": technicien_email,
                    "motif": motif,
                    "statut_etape": "Étape 1 - enregistrée",
                    "statut_email": "Non envoyé",
                    "date_email": "",
                    "statut_whatsapp": "Non envoyé",
                    "date_whatsapp": ""
                }
                save_instance(record)
                st.success("Instance enregistrée. Passe maintenant à l'étape 2 pour l'envoi.")
            else:
                st.error("Les champs Demande, Télécopie et Motif sont obligatoires.")

    st.markdown("---")
    st.subheader("Étape 2 - Envoi au technicien")

    saved_df = load_saved_instances()
    email_history_df = load_email_history()

    if saved_df.empty:
        st.info("Aucune instance enregistrée.")
    else:
        f1, f2, f3 = st.columns(3)
        with f1:
            search_saved = st.text_input("Recherche", placeholder="demande, technicien, motif...")
        with f2:
            selected_sector = st.selectbox(
                "Filtre secteur",
                ["Tous"] + sorted(saved_df["secteur"].dropna().astype(str).unique().tolist()) if "secteur" in saved_df.columns else ["Tous"]
            )
        with f3:
            selected_email_status = st.selectbox(
                "Filtre email",
                ["Tous", "Non envoyé", "Envoyé"]
            )

        filtered_saved = global_search(saved_df, search_saved).copy()

        if selected_sector != "Tous" and "secteur" in filtered_saved.columns:
            filtered_saved = filtered_saved[filtered_saved["secteur"].astype(str) == selected_sector]

        if selected_email_status != "Tous" and "statut_email" in filtered_saved.columns:
            filtered_saved = filtered_saved[filtered_saved["statut_email"].astype(str) == selected_email_status]

        try:
            filtered_saved = filtered_saved.sort_values(by="date_saisie", ascending=False)
        except Exception:
            pass

        for _, row in filtered_saved.iterrows():
            instance_id = str(row.get("instance_id", ""))
            message = build_instance_message(row)
            wa_url = build_whatsapp_url(row)

            st.markdown('<div class="card-box">', unsafe_allow_html=True)

            c1, c2, c3, c4 = st.columns([4.5, 1.4, 1.6, 1.4])

            with c1:
                st.markdown(
                    f"""
**Demande :** {row.get('demande', '')}  
**Technicien :** {row.get('technicien_nom', '')}  
**Secteur :** {row.get('secteur', '')}  
**Agent :** {row.get('agent', '')}  
**Motif :** {row.get('motif', '')}  
**Statut global :** {row.get('statut_etape', '')}
"""
                )

            with c2:
                if st.button("📧 Envoyer e-mail", key=f"mail_{instance_id}"):
                    current_config = load_email_config()
                    password = st.session_state.get("smtp_password", "") or get_secret("smtp_password", "")
                    recipient = str(row.get("technicien_email", "")).strip() or current_config.get("default_recipient", "").strip()
                    subject = f"Nouvelle instance - {row.get('demande', '')}"

                    if not current_config.get("sender_email"):
                        st.error("Configure d'abord l'email expéditeur dans la sidebar.")
                    elif not password:
                        st.error("Ajoute le mot de passe e-mail dans la sidebar ou dans st.secrets.")
                    elif not recipient:
                        st.error("Aucun email technicien ni destinataire par défaut.")
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
                                    "recipient": recipient,
                                    "subject": subject,
                                    "status": "ERROR",
                                    "error": str(e),
                                    "sent_by": st.session_state.get("admin_user", "")
                                }
                            )
                            st.error(f"Erreur e-mail : {e}")

            with c3:
                if wa_url:
                    st.markdown(render_whatsapp_button(wa_url), unsafe_allow_html=True)
                else:
                    st.caption("Pas de n° WhatsApp")

            with c4:
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

            with st.expander(f"Voir le détail - {row.get('demande', '')}"):
                st.text_area(
                    "Message",
                    value=message,
                    height=220,
                    key=f"msg_{instance_id}"
                )
                st.write(f"Email technicien : {row.get('technicien_email', '')}")
                st.write(f"WhatsApp technicien : {row.get('technicien_whatsapp', '')}")
                st.write(f"Statut email : {row.get('statut_email', 'Non envoyé')}")
                st.write(f"Statut WhatsApp : {row.get('statut_whatsapp', 'Non envoyé')}")

                if not email_history_df.empty and "instance_id" in email_history_df.columns:
                    hist_one = email_history_df[email_history_df["instance_id"].astype(str) == instance_id]
                    if not hist_one.empty:
                        st.markdown("**Historique e-mail de cette instance**")
                        st.dataframe(hist_one, use_container_width=True, height=180)

            st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("---")
        st.subheader("Tableau des instances")
        st.dataframe(filtered_saved, use_container_width=True, height=320)

        ex1, ex2 = st.columns(2)
        with ex1:
            st.download_button(
                "⬇️ Export CSV",
                data=filtered_saved.to_csv(index=False).encode("utf-8"),
                file_name="instances_enregistrees.csv",
                mime="text/csv"
            )
        with ex2:
            st.download_button(
                "⬇️ Export Excel",
                data=to_excel_bytes(filtered_saved, "Instances"),
                file_name="instances_enregistrees.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    st.markdown("---")
    st.subheader("Données source Excel")
    etat_df = safe_load_excel(ETAT_FILE, ETAT_SHEET, "ETAT FTTH RTC RTCL.xlsx")
    if not etat_df.empty:
        search_excel = st.text_input("Recherche dans le fichier source", placeholder="demande, secteur, état...")
        filtered_etat = global_search(etat_df, search_excel)
        st.dataframe(filtered_etat, use_container_width=True, height=400)
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
            for col in ["secteur", "agent", "motif", "utilisateur", "statut_email", "statut_whatsapp", "technicien_nom"]:
                if col not in saved_df.columns:
                    saved_df[col] = ""

            total_instances = len(saved_df)
            emails_sent = int((saved_df["statut_email"].astype(str) == "Envoyé").sum())
            whatsapp_sent = int((saved_df["statut_whatsapp"].astype(str) == "Envoyé").sum())
            agents_count = saved_df["agent"].astype(str).replace("nan", "").replace("", pd.NA).dropna().nunique()
            tech_count = saved_df["technicien_nom"].astype(str).replace("nan", "").replace("", pd.NA).dropna().nunique()
            secteurs_count = saved_df["secteur"].astype(str).replace("nan", "").replace("", pd.NA).dropna().nunique()

            k1, k2, k3, k4, k5, k6 = st.columns(6)
            k1.metric("Instances", total_instances)
            k2.metric("Emails envoyés", emails_sent)
            k3.metric("WhatsApp envoyés", whatsapp_sent)
            k4.metric("Agents actifs", agents_count)
            k5.metric("Techniciens", tech_count)
            k6.metric("Secteurs actifs", secteurs_count)

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
                fig3 = px.bar(motif_count, x="Motif", y="Nombre", color="Nombre", title="Top 15 des motifs")
                fig3.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig3, use_container_width=True)

            with c4:
                status_email = (
                    saved_df["statut_email"].fillna("Non renseigné").astype(str).value_counts().reset_index()
                )
                status_email.columns = ["Statut", "Nombre"]
                fig4 = px.pie(status_email, values="Nombre", names="Statut", title="Répartition des statuts e-mail")
                st.plotly_chart(fig4, use_container_width=True)

            c5, c6 = st.columns(2)

            with c5:
                status_wa = (
                    saved_df["statut_whatsapp"].fillna("Non renseigné").astype(str).value_counts().reset_index()
                )
                status_wa.columns = ["Statut", "Nombre"]
                fig5 = px.pie(status_wa, values="Nombre", names="Statut", title="Répartition des statuts WhatsApp")
                st.plotly_chart(fig5, use_container_width=True)

            with c6:
                user_count = (
                    saved_df["utilisateur"].fillna("Non renseigné").astype(str).value_counts().reset_index()
                )
                user_count.columns = ["Utilisateur", "Nombre"]
                fig6 = px.bar(user_count, x="Utilisateur", y="Nombre", color="Nombre", title="Saisies par utilisateur")
                st.plotly_chart(fig6, use_container_width=True)

            st.markdown("---")
            st.dataframe(saved_df, use_container_width=True, height=320)

    with tabs[1]:
        st.subheader("Historique des e-mails envoyés")
        if email_history_df.empty:
            st.info("Aucun e-mail envoyé pour le moment.")
        else:
            search_mail = st.text_input("Recherche dans l'historique e-mail", placeholder="demande, destinataire, statut...")
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

                kx1, kx2, kx3 = st.columns(3)
                kx1.metric("Total commandes source", len(etat_df))
                kx2.metric("Commandes VA", int((etat_df[etat_col].astype(str).str.upper() == "VA").sum()) if etat_col else 0)
                kx3.metric("Délai moyen source", safe_mean_numeric(etat_df[delai_col]) if delai_col else "N/A")

                if secteur_col:
                    secteur_source = etat_df[secteur_col].fillna("Non renseigné").astype(str).value_counts().reset_index()
                    secteur_source.columns = ["Secteur", "Nombre"]
                    fig7 = px.bar(secteur_source, x="Secteur", y="Nombre", color="Nombre", title="Commandes source par secteur")
                    st.plotly_chart(fig7, use_container_width=True)

            if not motif_df.empty:
                st.markdown("### Analyse source motifs")
                motif_col = find_column(motif_df, ["motif", "detail", "pc mauvais"])
                if motif_col:
                    mcount = motif_df[motif_col].fillna("Non renseigné").astype(str).value_counts().head(15).reset_index()
                    mcount.columns = ["Motif", "Nombre"]
                    fig8 = px.bar(mcount, x="Motif", y="Nombre", color="Nombre", title="Top 15 motifs source")
                    fig8.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig8, use_container_width=True)


# =========================================================
# PAGES SECONDAIRES
# =========================================================
elif page == "⚠️ DÉRANGEMENTS":
    st.subheader("Dérangements")
    st.info("Page prête pour la V3.")

elif page == "🔧 FIABILISATION":
    st.subheader("Fiabilisation")
    st.info("Page prête pour la V3.")

elif page == "⚖️ LITIGES":
    st.subheader("Litiges")
    st.info("Page prête pour la V3.")


st.caption("Pilotage opérationnel des interventions Fibre & RTC")
