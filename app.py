import streamlit as st
import pandas as pd
from datetime import datetime
import plotly.express as px
import io
import os

st.set_page_config(page_title="Gestion Chantier MHAMID", layout="wide")

# ====================== AUTHENTIFICATION MULTI-UTILISATEURS ======================
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.username = None

USERS = {
    "admin": "1234",
    "hamid": "hamid123",
    "shakhman": "shakhman123"
    # ← Ajoute ici d'autres utilisateurs : "nom": "motdepasse"
}

if not st.session_state.authenticated:
    st.title("🔐 Connexion à l'application")
    username = st.text_input("Nom d'utilisateur")
    password = st.text_input("Mot de passe", type="password")
   
    if st.button("Se connecter", type="primary"):
        if username in USERS and USERS[username] == password:
            st.session_state.authenticated = True
            st.session_state.username = username
            st.success(f"✅ Connexion réussie ! Bienvenue {username}")
            st.rerun()
        else:
            st.error("❌ Identifiants incorrects. Vérifiez votre nom d'utilisateur et mot de passe.")
    st.stop()

# ====================== STYLE ======================
st.markdown("""
    <style>
    .header {background-color: #0E7CFF; color: white; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>🚧 Gestion Chantier Fibre & RTC - MHAMID</h2></div>', unsafe_allow_html=True)

# ====================== FONCTION DÉTECTION COLONNES ======================
def find_column(df, keywords):
    if df is None or df.empty:
        return None
    for col in df.columns:
        if any(k.lower() in str(col).lower() for k in keywords):
            return col
    return None

# ====================== CHARGEMENT / SAUVEGARDE AUTOMATIQUE INSTANCES ======================
INSTANCES_FILE = "instances_saved.xlsx"

if "instances" not in st.session_state:
    if os.path.exists(INSTANCES_FILE):
        try:
            st.session_state.instances = pd.read_excel(INSTANCES_FILE)
            st.success(f"✅ {len(st.session_state.instances)} instances chargées depuis le fichier sauvegardé.")
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
                    "Téléscopie": telecopie, "Date Réception": date_reception, "Secteur": secteur,
                    "Agent": agent, "Motif": motif, "Date Saisie": datetime.now()
                }])
                st.session_state.instances = pd.concat([st.session_state.instances, new_row], ignore_index=True)
                save_instances()                     # ← Sauvegarde automatique Excel
                st.success(f"✅ Motif enregistré pour **{demande}** et sauvegardé automatiquement")
                st.balloons()
            else:
                st.error("❌ Les champs **Demande**, **Téléscopie** et **Motif** sont obligatoires.")

    st.subheader("📋 Instances saisies")
    if not st.session_state.instances.empty:
        st.dataframe(st.session_state.instances, use_container_width=True)
        csv = st.session_state.instances.to_csv(index=False).encode('utf-8')
        st.download_button("⬇️ Télécharger Instances", csv, f"instances_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")
    else:
        st.info("Aucune instance saisie pour le moment.")

    try:
        df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        st.subheader("📂 Fichier de référence")
        st.dataframe(df, use_container_width=True, height=400)
    except:
        st.warning("Impossible de charger ETAT FTTH RTC RTCL.xlsx")

# ====================== PAGE RAPPORTS ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")

    try:
        etat_df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        motif_df = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF")

        # ====================== IMPORT + VALIDATION EXCEL ======================
        st.subheader("📥 Import et Validation Tableau Excel")
        uploaded_file = st.file_uploader("Importez votre fichier Excel (OP + Etat)", type=["xlsx", "xls"])

        if uploaded_file:
            imported_df = pd.read_excel(uploaded_file)
            st.write("Colonnes détectées :", list(imported_df.columns))

            op_col = find_column(imported_df, ['op', 'OP'])
            etat_col = find_column(imported_df, ['etat', 'Etat', 'état', 'state'])

            if op_col and etat_col:
                # Validation
                valid_op = ["NA", "RM", "TR", "TL"]
                valid_etat = ["VA", "RE"]

                imported_df[op_col] = imported_df[op_col].astype(str).str.strip().str.upper()
                imported_df[etat_col] = imported_df[etat_col].astype(str).str.strip().str.upper()

                valid_mask = (imported_df[op_col].isin(valid_op)) & (imported_df[etat_col].isin(valid_etat))
                valid_df = imported_df[valid_mask].copy()
                invalid_df = imported_df[~valid_mask].copy()

                st.success(f"✅ {len(valid_df)} lignes valides | {len(invalid_df)} lignes rejetées")

                col1, col2 = st.columns(2)
                col1.metric("Lignes valides (OP + Etat)", len(valid_df))
                col2.metric("Lignes rejetées", len(invalid_df))

                if not invalid_df.empty:
                    st.error("Lignes avec erreurs (OP ou Etat invalide) :")
                    st.dataframe(invalid_df[[op_col, etat_col]], use_container_width=True)

                st.dataframe(valid_df, use_container_width=True)

                # Option téléchargement du résultat validé
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    valid_df.to_excel(writer, sheet_name="Données_Validées", index=False)
                output.seek(0)
                st.download_button("⬇️ Télécharger données validées", output, "Données_Validées.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.error("❌ Impossible de trouver les colonnes **OP** et **Etat** dans le fichier importé.")

        # ====================== STATISTIQUES (Excel + Instances) ======================
        all_motifs = pd.Series(dtype=str)

        motif_col = find_column(motif_df, ['detail motif', 'détail motif', 'motif', 'pc mauvais'])
        if motif_col and motif_col in motif_df.columns:
            all_motifs = pd.concat([all_motifs, motif_df[motif_col].astype(str).str.strip()])

        if not st.session_state.instances.empty:
            all_motifs = pd.concat([all_motifs, st.session_state.instances["Motif"].astype(str).str.strip()])

        all_motifs = all_motifs[(all_motifs != "") & (all_motifs != "nan") & (all_motifs != "None")]
        motif_count = all_motifs.value_counts().head(15)

        # KPIs
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Commandes", len(etat_df))
        col2.metric("Total Motifs", len(all_motifs))
        col3.metric("Instances saisies", len(st.session_state.instances))
        col4.metric("Motifs uniques", len(motif_count))

        st.divider()

        if not motif_count.empty:
            st.subheader("📊 Top 15 des Motifs")
            fig_bar = px.bar(x=motif_count.index, y=motif_count.values,
                             title="Top 15 Motifs (Excel + Instances)",
                             labels={"x": "Motif", "y": "Nombre"}, text=motif_count.values)
            fig_bar.update_layout(xaxis_tickangle=-45, height=550, margin=dict(b=200))
            st.plotly_chart(fig_bar, use_container_width=True)

            st.subheader("🥧 Répartition des Motifs")
            top10 = all_motifs.value_counts().head(10)
            fig_pie = px.pie(names=top10.index, values=top10.values, title="Répartition en %")
            fig_pie.update_traces(textinfo='percent+label')
            st.plotly_chart(fig_pie, use_container_width=True)

            st.dataframe(motif_count.reset_index().rename(columns={"index": "Motif", "count": "Nombre"}), use_container_width=True)
        else:
            st.warning("Aucun motif disponible pour le moment.")

    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des données : {str(e)}")
        st.info("Vérifiez que les fichiers ETAT FTTH RTC RTCL.xlsx et MOTIF TOTAL (1).xlsx sont présents.")

# ====================== AUTRES PAGES ======================
else:
    st.info(f"Page **{page}** est en cours de développement.")

st.caption(f"Application Gestion Chantier MHAMID | Connecté en tant que **{st.session_state.username}**")
