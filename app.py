import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, timedelta
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
import socket

DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"
PASSWORD = "admin123"   # 👉 tu peux changer le mot de passe ici

st.set_page_config(page_title="Gestion des chantiers", layout="wide")

# -------------------------------------
# FONCTIONS UTILITAIRES
# -------------------------------------

def get_local_ip():
    """Retourne l'adresse IP locale pour partager l'application."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
    except:
        ip = "localhost"
    finally:
        s.close()
    return ip


# -------------------------------------
# GESTION DES OPTIONS (COULEURS)
# -------------------------------------

def charger_options():
    if not os.path.exists(OPTIONS_FILE):
        default = {"rouge": 2, "orange": 7, "jaune": 14}
        sauvegarder_options(default)
        return default
    with open(OPTIONS_FILE, "r") as f:
        return json.load(f)

def sauvegarder_options(opts):
    with open(OPTIONS_FILE, "w") as f:
        json.dump(opts, f, indent=2)


# -------------------------------------
# GESTION DES DONNÉES
# -------------------------------------

def charger_chantiers():
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame(columns=["nom", "ref", "date"])

    with open(DATA_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    df = pd.DataFrame(data)
    df["date"] = pd.to_datetime(df["date"])
    return df

def sauvegarder_chantiers(df):
    data = df.copy()
    data["date"] = data["date"].dt.strftime("%Y-%m-%d")

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data.to_dict(orient="records"), f, indent=2)


# -------------------------------------
# COULEUR SELON ÉCHÉANCE
# -------------------------------------

def couleur_date(date, opts):
    today = datetime.today()
    delta = (date - today).days

    if delta <= opts["rouge"]:
        return "#FF0000"
    if delta <= opts["orange"]:
        return "#FF6600"
    if delta <= opts["jaune"]:
        return "#FFCC00"
    return "white"


# -------------------------------------
# EXPORT PDF
# -------------------------------------

def export_pdf(df, dossier):
    filename = os.path.join(dossier, "chantiers.pdf")
    c = canvas.Canvas(filename, pagesize=A4)
    width, height = A4

    y = height - 40
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "Liste des chantiers")
    y -= 40

    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, y, "Nom")
    c.drawString(260, y, "Référence")
    c.drawString(400, y, "Date montage")
    y -= 20

    c.setFont("Helvetica", 11)
    for _, row in df.iterrows():
        if y < 50:
            c.showPage()
            y = height - 40

        c.drawString(50, y, str(row["nom"])[:35])
        c.drawString(260, y, str(row["ref"]))
        c.drawString(400, y, row["date"].strftime("%d/%m/%Y"))
        y -= 18

    c.save()
    return filename


# ======================================================================
#                           INTERFACE STREAMLIT
# ======================================================================

opts = charger_options()
df = charger_chantiers()

# --------------------------------------------------
# 1. MODE EDITION / MOT DE PASSE
# --------------------------------------------------

if "edit_mode" not in st.session_state:
    st.session_state.edit_mode = False

with st.sidebar.expander("🔐 Connexion administrateur"):
    if not st.session_state.edit_mode:
        pwd = st.text_input("Mot de passe :", type="password")
        if st.button("Se connecter"):
            if pwd == PASSWORD:
                st.session_state.edit_mode = True
                st.success("Mode édition activé !")
                st.experimental_rerun()
            else:
                st.error("Mot de passe incorrect.")
    else:
        st.success("🔓 Mode édition ACTIVÉ")
        if st.button("Se déconnecter"):
            st.session_state.edit_mode = False
            st.experimental_rerun()

# --------------------------------------------------
# 2. COPIER LE LIEN LOCAL
# --------------------------------------------------

with st.sidebar.expander("🔗 Partage mobile"):
    ip = get_local_ip()
    lien = f"http://{ip}:8501"
    st.code(lien)
    st.write("Cliquez pour copier :")
    st.button("📋 Copier le lien", on_click=lambda: st.write(lien))

# --------------------------------------------------
# 3. OPTIONS (uniquement si admin)
# --------------------------------------------------

if st.session_state.edit_mode:
    with st.sidebar.expander("⚙️ Options d'urgences"):
        rouge = st.number_input("Rouge : jusqu'à X jours", min_value=0, max_value=60, value=opts["rouge"])
        orange = st.number_input("Orange : jusqu'à X jours", min_value=0, max_value=60, value=opts["orange"])
        jaune = st.number_input("Jaune : jusqu'à X jours", min_value=0, max_value=60, value=opts["jaune"])

        if st.button("Enregistrer les options"):
            opts = {"rouge": rouge, "orange": orange, "jaune": jaune}
            sauvegarder_options(opts)
            st.success("Options enregistrées !")
            st.experimental_rerun()


# --------------------------------------------------
# TITRE PRINCIPAL
# --------------------------------------------------

st.title("📋 Gestion des chantiers – Version atelier")


# --------------------------------------------------
# 4. AJOUT / MODIFICATION (admin uniquement)
# --------------------------------------------------

if st.session_state.edit_mode:
    st.subheader("➕ Ajouter un chantier")

    col1, col2, col3 = st.columns(3)
    with col1:
        nom = st.text_input("Nom du chantier")
    with col2:
        ref = st.text_input("Référence")
    with col3:
        date_montage = st.date_input("Date de montage")

    if st.button("Ajouter"):
        if nom and ref:
            new_row = pd.DataFrame({"nom": [nom], "ref": [ref], "date": [pd.to_datetime(date_montage)]})
            df = pd.concat([df, new_row], ignore_index=True)
            df = df.sort_values("date")
            sauvegarder_chantiers(df)
            st.success("Chantier ajouté !")
            st.experimental_rerun()
        else:
            st.error("Veuillez remplir tous les champs.")


# --------------------------------------------------
# 5. AFFICHAGE DU TABLEAU AVEC COULEURS
# --------------------------------------------------

st.subheader("📆 Liste des chantiers triée par date")

styled_df = df.style.apply(
    lambda row: [f"background-color: {couleur_date(row['date'], opts)}"] * len(row),
    axis=1
)

st.dataframe(styled_df, height=500)


# --------------------------------------------------
# 6. SUPPRESSION (admin)
# --------------------------------------------------

if st.session_state.edit_mode:

    st.subheader("🗑️ Supprimer un chantier")

    idx_to_delete = st.selectbox(
        "Choisir un chantier à supprimer",
        df.index,
        format_func=lambda x: f"{df.loc[x, 'nom']} - {df.loc[x, 'ref']}",
    )

    if st.button("Supprimer"):
        df = df.drop(idx_to_delete).reset_index(drop=True)
        sauvegarder_chantiers(df)
        st.success("Chantier supprimé.")
        st.experimental_rerun()


# --------------------------------------------------
# 7. EXPORTS (admin)
# --------------------------------------------------

if st.session_state.edit_mode:

    st.subheader("📤 Exporter")

    dossier = st.text_input("Dossier d'export :", value=os.getcwd())

    colA, colB = st.columns(2)

    with colA:
        if st.button("Exporter PDF"):
            fichier = export_pdf(df, dossier)
            st.success(f"PDF exporté : {fichier}")

    with colB:
        if st.button("Exporter Excel"):
            path = os.path.join(dossier, "chantiers.xlsx")
            df.to_excel(path, index=False)
            st.success(f"Excel exporté : {path}")
