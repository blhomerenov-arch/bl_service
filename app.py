import json
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
APP_NAME = "Système Centralisé de Gestion, Suivi et Dispatching des Interventions Fibre, RTC et Réclamations Terrain"

st.set_page_config(page_title=APP_NAME, layout="wide")

BASE_DIR = Path(__file__).parent
ETAT_FILE = BASE_DIR / "ETAT FTTH RTC RTCL.xlsx"
MOTIF_FILE = BASE_DIR / "MOTIF TOTAL (1).xlsx"

SAISIES_FILE = BASE_DIR / "saisies_instances.csv"
SETTINGS_FILE = BASE_DIR / "parametres_app.json"
EMAIL_CONFIG_FILE = BASE_DIR / "email_config.json"

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
        border-radius: 12px;
        text-align: center;
        margin-bottom: 15px;
    }
    .small-note {
        color: #6c757d;
        font-size: 13px;
    }
    .action-box {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        border-radius: 10px;
        padding: 12px;
        margin-bottom: 10px;
    }
    .wa-btn a {
        display: inline-block;
        padding: 0.45rem 0.75rem;
        background-color: #25D366;
        color: white !important;
        text-decoration: none;
        border-radius: 8px;
        font-weight: 600;
        text-align: center;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown(
    f"""
    <div class="main-header">
        <h2>{APP_NAME}</h2>
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# OUTILS
# =========================================================
def normalize_text(value):
    value = "" if value is None else str(value)
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("utf-8")
    return value.lower().strip()


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
                df[col]
                .astype(str)
                .str.strip()
                .replace("nan", "")
                .ne("")
                .sum()
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


def global_search(df, query):
    if df.empty or not query:
        return df

    mask = df.astype(str).apply(lambda col: col.str.contains(query, case=False, na=False))
    return df[mask.any(axis=1)]


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


def now_str():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def generate_instance_id():
    return datetime.now().strftime("%Y%m%d%H%M%S%f")


def clean_phone_for_whatsapp(phone):
    if phone is None:
        return ""
    phone = str(phone).strip()
    return "".join(ch for ch in phone if ch.isdigit())


# =========================================================
# SETTINGS
# =========================================================
def default_settings():
    return {
        "utilisateurs": ["admin"],
        "agents": ["hamid", "SHAKHMAN"],
        "secteurs": ["MHAMID", "BOUAAKAZ", "Province M'HAMID"]
    }


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
        if key not in data or not isinstance(data[key], list):
            data[key] = value

    return data


def save_settings(data):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


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


# =========================================================
# EMAIL CONFIG
# =========================================================
def default_email_config():
    return {
        "smtp_server": "smtp.gmail.com",
        "smtp_port": 587,
        "sender_email": "",
        "default_recipient": ""
    }


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


def save_email_config(data):
    with open(EMAIL_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_secret(name, default=""):
    try:
        return st.secrets[name]
    except Exception:
        return default


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
Utilisateur saisie : {row.get('utilisateur', '')}

Cordialement,
Application de gestion chantier
""".strip()


def build_whatsapp_url(row):
    phone = clean_phone_for_whatsapp(row.get("technicien_whatsapp", ""))
    if not phone:
        return ""

    text = build_instance_message(row)
    return f"https://wa.me/{phone}?text={quote(text)}"


# =========================================================
# LOAD GLOBAL CONFIG
# =========================================================
settings = load_settings()
email_config = load_email_config()

if "email_password" not in st.session_state:
    st.session_state["email_password"] = get_secret("smtp_password", "")


# =========================================================
# SIDEBAR
# =========================================================
with st.sidebar:
    st.header("Configuration e-mail")

    with st.form("email_config_form"):
        smtp_server = st.text_input("Serveur SMTP", value=email_config.get("smtp_server", "smtp.gmail.com"))
        smtp_port = st.number_input("Port SMTP", min_value=1, max_value=9999, value=int(email_config.get("smtp_port", 587)))
        sender_email = st.text_input("Email expéditeur", value=email_config.get("sender_email", ""))
        email_password = st.text_input(
            "Mot de passe / mot de passe d'application",
            type="password",
            value=st.session_state.get("email_password", "")
        )
        default_recipient = st.text_input("Destinataire par défaut", value=email_config.get("default_recipient", ""))

        save_email_btn = st.form_submit_button("💾 Enregistrer config email")

        if save_email_btn:
            config_to_save = {
                "smtp_server": smtp_server,
                "smtp_port": int(smtp_port),
                "sender_email": sender_email,
                "default_recipient": default_recipient
            }
            save_email_config(config_to_save)
            st.session_state["email_password"] = email_password
            st.success("Configuration e-mail enregistrée.")

    st.caption("Conseil : utilise un mot de passe d'application pour Gmail/Outlook.")


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
    st.subheader("Étape 1 - Saisie et enregistrement de l'instance")

    with st.expander("⚙️ Administration : utilisateurs / agents / secteurs", expanded=False):
        c1, c2, c3 = st.columns(3)

        with c1:
            with st.form("add_user_form"):
                st.markdown("**Ajouter un utilisateur**")
                new_user = st.text_input("Nom utilisateur")
                submit_user = st.form_submit_button("➕ Ajouter utilisateur")
                if submit_user:
                    ok, msg = add_item_to_settings(settings, "utilisateurs", new_user)
                    if ok:
                        st.success(msg)
                        st.experimental_rerun()
                    else:
                        st.error(msg)

        with c2:
            with st.form("add_agent_form"):
                st.markdown("**Ajouter un agent**")
                new_agent = st.text_input("Nom agent")
                submit_agent = st.form_submit_button("➕ Ajouter agent")
                if submit_agent:
                    ok, msg = add_item_to_settings(settings, "agents", new_agent)
                    if ok:
                        st.success(msg)
                        st.experimental_rerun()
                    else:
                        st.error(msg)

        with c3:
            with st.form("add_sector_form"):
                st.markdown("**Ajouter un secteur**")
                new_sector = st.text_input("Nom secteur")
                submit_sector = st.form_submit_button("➕ Ajouter secteur")
                if submit_sector:
                    ok, msg = add_item_to_settings(settings, "secteurs", new_sector)
                    if ok:
                        st.success(msg)
                        st.experimental_rerun()
                    else:
                        st.error(msg)

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

        st.markdown("**Affectation technicien**")
        t1, t2, t3 = st.columns(3)
        with t1:
            technicien_nom = st.text_input("Technicien")
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
                    "statut_whatsapp": "Non préparé",
                    "date_whatsapp": ""
                }
                save_instance(record)
                st.success("Instance enregistrée. Passe maintenant à l'étape 2 pour l'envoi.")
            else:
                st.error("Les champs Demande, Télécopie et Motif sont obligatoires.")

    st.markdown("---")
    st.subheader("Étape 2 - Envoi des instances enregistrées")

    saved_df = load_saved_instances()

    if saved_df.empty:
        st.info("Aucune instance enregistrée.")
    else:
        search_saved = st.text_input(
            "Recherche dans les instances enregistrées",
            placeholder="demande, secteur, agent, technicien, motif..."
        )

        filtered_saved = global_search(saved_df, search_saved).copy()

        try:
            filtered_saved = filtered_saved.sort_values(by="date_saisie", ascending=False)
        except Exception:
            pass

        st.markdown("### Actions par instance")

        for _, row in filtered_saved.iterrows():
            instance_id = str(row.get("instance_id", ""))
            wa_url = build_whatsapp_url(row)

            with st.container():
                st.markdown('<div class="action-box">', unsafe_allow_html=True)

                top1, top2, top3 = st.columns([5, 1.2, 1.5])

                with top1:
                    st.markdown(
                        f"""
**Demande :** {row.get('demande', '')}  
**Secteur :** {row.get('secteur', '')}  
**Agent :** {row.get('agent', '')}  
**Technicien :** {row.get('technicien_nom', '')}  
**Motif :** {row.get('motif', '')}  
**Statut :** {row.get('statut_etape', 'Étape 1 - enregistrée')}
"""
                    )

                with top2:
                    if st.button("📧 Envoyer e-mail", key=f"email_{instance_id}"):
                        current_config = load_email_config()
                        password = st.session_state.get("email_password", "") or get_secret("smtp_password", "")
                        recipient = str(row.get("technicien_email", "")).strip() or current_config.get("default_recipient", "").strip()

                        if not current_config.get("sender_email"):
                            st.error("Configure d'abord l'email expéditeur dans la barre latérale.")
                        elif not password:
                            st.error("Configure le mot de passe e-mail dans la barre latérale ou dans st.secrets.")
                        elif not recipient:
                            st.error("Aucun email technicien ou destinataire par défaut.")
                        else:
                            try:
                                subject = f"Nouvelle instance chantier - {row.get('demande', '')}"
                                body = build_instance_message(row)
                                send_email_smtp(current_config, password, recipient, subject, body)

                                update_instance(
                                    instance_id,
                                    {
                                        "statut_email": "Envoyé",
                                        "date_email": now_str(),
                                        "statut_etape": "Étape 2 - e-mail envoyé"
                                    }
                                )
                                st.success(f"E-mail envoyé à {recipient}")
                                st.experimental_rerun()
                            except Exception as e:
                                st.error(f"Erreur email : {e}")

                with top3:
                    if wa_url:
                        if st.button("💬 Préparer WhatsApp", key=f"wa_prepare_{instance_id}"):
                            update_instance(
                                instance_id,
                                {
                                    "statut_whatsapp": "Préparé",
                                    "date_whatsapp": now_str(),
                                    "statut_etape": "Étape 2 - WhatsApp préparé"
                                }
                            )
                            st.session_state[f"wa_link_{instance_id}"] = wa_url
                            st.experimental_rerun()
                    else:
                        st.caption("Pas de n° WhatsApp")

                if st.session_state.get(f"wa_link_{instance_id}"):
                    st.markdown(
                        f"""
<div class="wa-btn">
<a href="{st.session_state[f'wa_link_{instance_id}']}" target="_blank">Ouvrir WhatsApp pour cette instance</a>
</div>
                        """,
                        unsafe_allow_html=True
                    )

                with st.expander(f"Voir le détail / message - {row.get('demande', '')}"):
                    st.text_area(
                        "Message à envoyer",
                        value=build_instance_message(row),
                        height=220,
                        key=f"msg_{instance_id}"
                    )

                    st.write(f"Email technicien : {row.get('technicien_email', '')}")
                    st.write(f"WhatsApp technicien : {row.get('technicien_whatsapp', '')}")
                    st.write(f"Statut email : {row.get('statut_email', 'Non envoyé')}")
                    st.write(f"Statut WhatsApp : {row.get('statut_whatsapp', 'Non préparé')}")

                st.markdown("</div>", unsafe_allow_html=True)
                st.markdown("---")

        st.subheader("Tableau des instances enregistrées")
        st.dataframe(filtered_saved, use_container_width=True, height=350)

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
    etat_df = safe_load_excel(ETAT_FILE, ETAT_SHEET, "ETAT FTTH RTC RTCL.xlsx")
    motif_df = safe_load_excel(MOTIF_FILE, MOTIF_SHEET, "MOTIF TOTAL (1).xlsx")

    if saved_df.empty and etat_df.empty and motif_df.empty:
        st.warning("Aucune donnée disponible pour les rapports.")
    else:
        if not saved_df.empty:
            for col in [
                "utilisateur", "secteur", "agent", "motif", "statut_email",
                "technicien_nom", "date_reception", "date_saisie"
            ]:
                if col not in saved_df.columns:
                    saved_df[col] = ""

            total_instances = len(saved_df)
            emails_envoyes = int((saved_df["statut_email"].astype(str) == "Envoyé").sum()) if "statut_email" in saved_df.columns else 0
            taux_email = round((emails_envoyes / total_instances) * 100, 1) if total_instances > 0 else 0
            agents_actifs = saved_df["agent"].astype(str).replace("nan", "").replace("", pd.NA).dropna().nunique()
            secteurs_actifs = saved_df["secteur"].astype(str).replace("nan", "").replace("", pd.NA).dropna().nunique()
            techs_actifs = saved_df["technicien_nom"].astype(str).replace("nan", "").replace("", pd.NA).dropna().nunique()

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Instances saisies", total_instances)
            k2.metric("Emails envoyés", emails_envoyes)
            k3.metric("Taux envoi email", f"{taux_email}%")
            k4.metric("Agents actifs", agents_actifs)
            k5.metric("Techniciens affectés", techs_actifs)

            st.markdown("---")

            c1, c2 = st.columns(2)

            with c1:
                secteur_count = (
                    saved_df["secteur"]
                    .fillna("Non renseigné")
                    .astype(str)
                    .value_counts()
                    .reset_index()
                )
                secteur_count.columns = ["Secteur", "Nombre"]

                if not secteur_count.empty:
                    fig_secteur = px.bar(
                        secteur_count,
                        x="Secteur",
                        y="Nombre",
                        color="Nombre",
                        title="Instances par secteur"
                    )
                    st.plotly_chart(fig_secteur, use_container_width=True)

            with c2:
                agent_count = (
                    saved_df["agent"]
                    .fillna("Non renseigné")
                    .astype(str)
                    .value_counts()
                    .reset_index()
                )
                agent_count.columns = ["Agent", "Nombre"]

                if not agent_count.empty:
                    fig_agent = px.bar(
                        agent_count,
                        x="Agent",
                        y="Nombre",
                        color="Nombre",
                        title="Instances par agent"
                    )
                    st.plotly_chart(fig_agent, use_container_width=True)

            c3, c4 = st.columns(2)

            with c3:
                motif_count = (
                    saved_df["motif"]
                    .fillna("Non renseigné")
                    .astype(str)
                    .value_counts()
                    .head(15)
                    .reset_index()
                )
                motif_count.columns = ["Motif", "Nombre"]

                if not motif_count.empty:
                    fig_motif = px.bar(
                        motif_count,
                        x="Motif",
                        y="Nombre",
                        color="Nombre",
                        title="Top 15 des motifs"
                    )
                    fig_motif.update_layout(xaxis_tickangle=-45)
                    st.plotly_chart(fig_motif, use_container_width=True)

            with c4:
                email_status_count = (
                    saved_df["statut_email"]
                    .fillna("Non renseigné")
                    .astype(str)
                    .value_counts()
                    .reset_index()
                )
                email_status_count.columns = ["Statut", "Nombre"]

                if not email_status_count.empty:
                    fig_email = px.pie(
                        email_status_count,
                        values="Nombre",
                        names="Statut",
                        title="Répartition des statuts e-mail"
                    )
                    st.plotly_chart(fig_email, use_container_width=True)

            c5, c6 = st.columns(2)

            with c5:
                user_count = (
                    saved_df["utilisateur"]
                    .fillna("Non renseigné")
                    .astype(str)
                    .value_counts()
                    .reset_index()
                )
                user_count.columns = ["Utilisateur", "Nombre"]

                if not user_count.empty:
                    fig_user = px.bar(
                        user_count,
                        x="Utilisateur",
                        y="Nombre",
                        color="Nombre",
                        title="Saisies par utilisateur"
                    )
                    st.plotly_chart(fig_user, use_container_width=True)

            with c6:
                tech_count = (
                    saved_df["technicien_nom"]
                    .fillna("Non renseigné")
                    .astype(str)
                    .value_counts()
                    .reset_index()
                )
                tech_count.columns = ["Technicien", "Nombre"]

                if not tech_count.empty:
                    fig_tech = px.bar(
                        tech_count,
                        x="Technicien",
                        y="Nombre",
                        color="Nombre",
                        title="Instances par technicien"
                    )
                    st.plotly_chart(fig_tech, use_container_width=True)

            st.markdown("---")
            st.subheader("Données de reporting")

            st.dataframe(saved_df, use_container_width=True, height=350)

            rp1, rp2 = st.columns(2)
            with rp1:
                st.download_button(
                    "⬇️ Export rapport CSV",
                    data=saved_df.to_csv(index=False).encode("utf-8"),
                    file_name="rapport_instances.csv",
                    mime="text/csv"
                )
            with rp2:
                st.download_button(
                    "⬇️ Export rapport Excel",
                    data=to_excel_bytes(saved_df, "Rapports"),
                    file_name="rapport_instances.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # complément depuis ETAT source
        if not etat_df.empty:
            st.markdown("---")
            st.subheader("Analyse complémentaire du fichier source")

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
                fig_source = px.bar(secteur_source, x="Secteur", y="Nombre", color="Nombre", title="Commandes source par secteur")
                st.plotly_chart(fig_source, use_container_width=True)


# =========================================================
# PAGE DÉRANGEMENTS
# =========================================================
elif page == "⚠️ DÉRANGEMENTS":
    st.subheader("Dérangements")
    st.info("Page prête pour ajout futur du suivi des incidents.")


# =========================================================
# PAGE FIABILISATION
# =========================================================
elif page == "🔧 FIABILISATION":
    st.subheader("Fiabilisation")
    st.info("Page prête pour ajout futur du suivi de fiabilisation.")


# =========================================================
# PAGE LITIGES
# =========================================================
elif page == "⚖️ LITIGES":
    st.subheader("Litiges")
    st.info("Page prête pour ajout futur du suivi des litiges.")


st.caption("Pilotage opérationnel des interventions Fibre, RTC et suivi terrain")
