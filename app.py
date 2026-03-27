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
            st.error("❌ Identifiants incorrects. Vérifiez nom d'utilisateur et mot de passe.")
    st.stop()

# ====================== STYLE ======================
st.markdown("""
    <style>
    .header {background-color: #0E7CFF; color: white; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px;}
    .success {background-color: #d4edda; padding: 10px; border-radius: 8px;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="header"><h2>🚧 Gestion Chantier Fibre & RTC - MHAMID</h2></div>', unsafe_allow_html=True)

# ====================== FONCTION RECHERCHE AUTOMATIQUE DES COLONNES ======================
def find_column(df, keywords):
    """
    Recherche intelligente une colonne dans un DataFrame selon une liste de mots-clés.
    
    Exemple : find_column(df, ['op', 'OP', 'operation']) → retourne la colonne qui contient "op"
    Fonctionne même si les noms de colonnes sont mal écrits ou en majuscules/minuscules.
    """
    if df is None or df.empty:
        return None
    
    for col in df.columns:
        col_str = str(col).lower().strip()
        for keyword in [k.lower().strip() for k in keywords]:
            if keyword in col_str:
                return col  # Retourne le vrai nom de la colonne tel qu'il est dans le fichier
    return None

# Explication détaillée de cette fonction (tu peux la supprimer plus tard) :
# - Elle parcourt toutes les colonnes du fichier
# - Elle transforme le nom de la colonne et les mots-clés en minuscules pour éviter les problèmes de casse
# - Elle cherche si le mot-clé est "contenu" dans le nom de la colonne (ex: "code op" contient "op")
# - Très utile quand les fichiers Excel ont des noms de colonnes différents (OP, Code OP, Opération, etc.)

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
                st.success(f"✅ Instance enregistrée et sauvegardée pour la demande **{demande}**")
                st.balloons()
            else:
                st.error("❌ Les champs Demande, Téléscopie et Motif sont obligatoires.")

    # Affichage instances
    st.subheader("📋 Instances saisies")
    if not st.session_state.instances.empty:
        st.dataframe(st.session_state.instances, use_container_width=True)
        csv = st.session_state.instances.to_csv(index=False).encode('utf-8')
        st.download_button("⬇️ Télécharger Instances", csv, f"instances_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")
    else:
        st.info("Aucune instance saisie pour le moment.")

# ====================== PAGE RAPPORTS - VALIDATION AVANCÉE ======================
elif page == "📊 RAPPORTS":
    st.subheader("📊 Rapports et Statistiques")

    try:
        etat_df = pd.read_excel("ETAT FTTH RTC RTCL.xlsx", sheet_name="SITUATION14.15")
        motif_df = pd.read_excel("MOTIF TOTAL (1).xlsx", sheet_name="MOTIF")

        # ====================== IMPORT EXCEL AVEC VALIDATION AVANCÉE ======================
        st.subheader("📥 Import et Validation Avancée du Tableau Excel")

        uploaded_file = st.file_uploader("Importez votre fichier Excel à valider (colonnes OP et Etat)", type=["xlsx", "xls"])

        if uploaded_file:
            df_import = pd.read_excel(uploaded_file)
            st.write("**Colonnes détectées dans le fichier :**", list(df_import.columns))

            # Recherche automatique des colonnes
            op_col = find_column(df_import, ['op', 'OP', 'operation', 'code op'])
            etat_col = find_column(df_import, ['etat', 'Etat', 'état', 'state', 'status'])

            if op_col and etat_col:
                st.success(f"✅ Colonnes trouvées → **OP** : `{op_col}` | **Etat** : `{etat_col}`")

                # Nettoyage avancé
                df_import[op_col] = df_import[op_col].astype(str).str.strip().str.upper()
                df_import[etat_col] = df_import[etat_col].astype(str).str.strip().str.upper()

                # ====================== VALIDATION AVANCÉE ======================
                valid_op_values = ["NA", "RM", "TR", "TL"]
                valid_etat_values = ["VA", "RE"]

                # Création des masques de validation
                op_valid = df_import[op_col].isin(valid_op_values)
                etat_valid = df_import[etat_col].isin(valid_etat_values)

                df_import['OP_Valide'] = op_valid
                df_import['Etat_Valide'] = etat_valid
                df_import['Ligne_Valide'] = op_valid & etat_valid

                valid_df = df_import[df_import['Ligne_Valide']].copy()
                invalid_df = df_import[~df_import['Ligne_Valide']].copy()

                # Affichage des résultats
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Lignes", len(df_import))
                col2.metric("✅ Lignes Valides", len(valid_df), delta=len(valid_df) - len(df_import))
                col3.metric("❌ Lignes Invalides", len(invalid_df))

                if not invalid_df.empty:
                    st.error("⚠️ Lignes avec erreurs détectées :")
                    st.dataframe(invalid_df[[op_col, etat_col, 'OP_Valide', 'Etat_Valide']], use_container_width=True)

                st.success(f"✅ {len(valid_df)} lignes validées avec succès")
                st.dataframe(valid_df.drop(columns=['OP_Valide', 'Etat_Valide', 'Ligne_Valide']), use_container_width=True)

                # Téléchargement des données validées
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    valid_df.drop(columns=['OP_Valide', 'Etat_Valide', 'Ligne_Valide']).to_excel(writer, index=False, sheet_name="Données_Validées")
                output.seek(0)
                st.download_button("⬇️ Télécharger les données validées", 
                                 output, 
                                 "Données_Validées.xlsx",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            else:
                st.error("❌ Impossible de trouver les colonnes OP et Etat.")
                st.info("Noms attendus pour OP : OP, op, operation, code op\nNoms attendus pour Etat : Etat, etat, état, state, status")

        # ====================== STATISTIQUES (Excel + Instances) ======================
        st.divider()
        all_motifs = pd.Series(dtype=str)

        motif_col = find_column(motif_df, ['detail motif', 'détail motif', 'motif', 'pc mauvais'])
        if motif_col:
            all_motifs = pd.concat([all_motifs, motif_df[motif_col].astype(str).str.strip()])

        if not st.session_state.instances.empty:
            all_motifs = pd.concat([all_motifs, st.session_state.instances["Motif"].astype(str).str.strip()])

        all_motifs = all_motifs[(all_motifs != "") & (all_motifs != "nan") & (all_motifs != "None")]
        motif_count = all_motifs.value_counts().head(15)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Commandes", len(etat_df))
        col2.metric("Total Motifs", len(all_motifs))
        col3.metric("Instances saisies", len(st.session_state.instances))
        col4.metric("Motifs uniques", len(motif_count))

        if not motif_count.empty:
            st.subheader("📊 Top 15 des Motifs")
            fig = px.bar(x=motif_count.index, y=motif_count.values, title="Top 15 Motifs", text=motif_count.values)
            fig.update_layout(xaxis_tickangle=-45, height=550)
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("🥧 Répartition des Motifs")
            fig_pie = px.pie(names=motif_count.index[:10], values=motif_count.values[:10], title="Répartition en %")
            st.plotly_chart(fig_pie, use_container_width=True)

    except Exception as e:
        st.error(f"❌ Erreur générale : {str(e)}")

else:
    st.info(f"Page **{page}** est en cours de développement.")

st.caption(f"Application Gestion Chantier MHAMID | Connecté en tant que **{st.session_state.get('username', 'admin')}**")
