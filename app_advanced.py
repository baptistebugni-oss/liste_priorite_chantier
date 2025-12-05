import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import qrcode

# ================== CONFIG ==================

DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"
ADMIN_PASSWORD = "admin123"

DEFAULT_OPTIONS = {
    "rouge": 2,
    "orange": 7,
    "jaune": 14,
    "show_qr": True,
}

COLUMNS = ["nom", "ref", "date", "commentaire", "statut", "priorite"]


# ================== UTILITAIRES ==================

def ensure_columns(df: pd.DataFrame, columns):
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    return df


def charger_chantiers():
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame(columns=COLUMNS)

    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except:
        data = []

    df = pd.DataFrame(data)
    df = ensure_columns(df, COLUMNS)

    if not df.empty:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    return df


def sauvegarder_chantiers(df):
    df2 = df.copy()
    df2["date"] = df2["date"].astype(str)
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(df2.to_dict(orient="records"), f, ensure_ascii=False, indent=4)


def charger_options():
    if not os.path.exists(OPTIONS_FILE):
        sauvegarder_options(DEFAULT_OPTIONS)
        return DEFAULT_OPTIONS.copy()

    try:
        with open(OPTIONS_FILE, "r", encoding="utf-8") as f:
            opts = json.load(f)
    except:
        opts = {}

    for k, v in DEFAULT_OPTIONS.items():
        opts.setdefault(k, v)

    return opts


def sauvegarder_options(opts):
    with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(opts, f, indent=4)


# ================== PRIORITÉS OPTION D ==================

def get_row_color(row, opts):
    # 1) Priorité manuelle
    prio = row["priorite"].lower()
    if prio in ("rouge", "orange", "jaune", "gris"):
        return {
            "rouge": "#FF4B4B",
            "orange": "#FFA500",
            "jaune": "#FFD966",
            "gris": "#DDDDDD",
        }[prio]

    # 2) Statut
    statut = row["statut"].lower()
    statut_map = {
        "prévu": "#E5F6FF",
        "en cours": "#FFE4B5",
        "en attente": "#F0F0F0",
        "terminé": "",
    }

    if statut in statut_map and statut_map[statut]:
        return statut_map[statut]

    # 3) Automatique (date)
    d = row["date"]
    if pd.isna(d):
        return ""

    delta = (d.date() - datetime.today().date()).days

    if delta <= opts["rouge"]:
        return "#FF4B4B"
    if delta <= opts["orange"]:
        return "#FFA500"
    if delta <= opts["jaune"]:
        return "#FFD966"

    return ""


def style_urgence(row, opts):
    color = get_row_color(row, opts)
    return [f"background-color: {color}"] * len(row) if color else [""] * len(row)


# ================== EXPORT PDF & EXCEL ==================

def build_excel(df):
    buf = BytesIO()
    df2 = df.copy()
    df2["date"] = df2["date"].dt.strftime("%d/%m/%Y")
    df2.to_excel(buf, index=False)
    buf.seek(0)
    return buf


def build_pdf(df):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    y = height - 40
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "Liste des chantiers")
    y -= 30

    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Nom")
    c.drawString(200, y, "Code")
    c.drawString(300, y, "Date")
    c.drawString(380, y, "Statut")
    y -= 20

    c.setFont("Helvetica", 10)
    for _, row in df.iterrows():
        if y < 40:
            c.showPage()
            y = height - 40

        c.drawString(40, y, str(row["nom"])[:25])
        c.drawString(200, y, str(row["ref"])[:12])
        c.drawString(300, y, row["date"].strftime("%d/%m/%Y") if not pd.isna(row["date"]) else "-")
        c.drawString(380, y, str(row["statut"]))
        y -= 15

    c.save()
    buf.seek(0)
    return buf


# ================== QR CODE ==================

def build_qr(url):
    qr = qrcode.QRCode(box_size=8, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image()
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


# ================== IMPORT EXCEL ULTRA-ROBUSTE ==================

def importer_excel(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip().str.lower()

    # Variantes reconnues
    possible_nom_cols = [
        "nom de l’affaire",
        "nom de l'affaire",
        "nom de laffaire",
        "nom affaire",
        "nom d’affaire",
        "nom d'affaire",
    ]

    possible_code_cols = [
        "code de l’affaire",
        "code de l'affaire",
        "code de laffaire",
        "code affaire",
        "codedaffaire",
    ]

    possible_debut_cols = ["début", "debut"]

    def find_col(possible):
        for col in df.columns:
            if col in possible:
                return col
        return None

    col_comp = "compétence"
    col_nom = find_col(possible_nom_cols)
    col_code = find_col(possible_code_cols)
    col_debut = find_col(possible_debut_cols)

    # Vérification
    missing = []
    if col_nom is None: missing.append("Nom de l’affaire")
    if col_code is None: missing.append("Code de l’affaire")
    if col_debut is None: missing.append("Début")
    if col_comp not in df.columns: missing.append("Compétence")

    if missing:
        st.error("⚠️ Colonnes manquantes : " + ", ".join(missing))
        return pd.DataFrame()

    # Filtre Montage
    df_filtre = df[df[col_comp].astype(str).str.lower() == "montage"]

    if df_filtre.empty:
        st.warning("Aucune ligne 'Montage' trouvée.")
        return pd.DataFrame()

    df_res = df_filtre[[col_nom, col_code, col_debut]].copy()
    df_res.columns = ["nom", "ref", "date"]
    df_res["date"] = pd.to_datetime(df_res["date"], errors="coerce")

    return df_res


# ================== INTERFACE STREAMLIT ==================

st.set_page_config(page_title="Liste Priorité Chantier", layout="wide")
st.title("📋 Gestion des priorités chantier")

mode = st.sidebar.selectbox("Mode", ["Lecture seule", "Administrateur"])
admin = False

if mode == "Administrateur":
    mdp = st.sidebar.text_input("Mot de passe", type="password")
    if mdp == ADMIN_PASSWORD:
        admin = True
    elif mdp:
        st.sidebar.error("Mot de passe incorrect")

df = charger_chantiers()
opts = charger_options()


# ================== OPTIONS ADMIN ==================
if admin:
    st.sidebar.markdown("### ⚙️ Options d'urgence")
    rouge = st.sidebar.number_input("Jours max ROUGE", 0, 90, opts["rouge"], key="rouge_opt")
    orange = st.sidebar.number_input("Jours max ORANGE", 0, 90, opts["orange"], key="orange_opt")
    jaune = st.sidebar.number_input("Jours max JAUNE", 0, 90, opts["jaune"], key="jaune_opt")
    show_qr = st.sidebar.checkbox("Afficher QR Code atelier", value=opts["show_qr"], key="qr_opt")

    if st.sidebar.button("Enregistrer options", key="save_options"):
        opts.update({"rouge": rouge, "orange": orange, "jaune": jaune, "show_qr": show_qr})
        sauvegarder_options(opts)
        st.sidebar.success("Options enregistrées ✔")


# ================== IMPORT EXCEL ==================
if admin:
    st.subheader("📥 Import Excel (filtré sur Compétence = 'Montage')")

    fichier_excel = st.file_uploader("Importer un fichier Excel", type=["xlsx"], key="import_excel")

    df_excel = pd.DataFrame()

    if fichier_excel:
        df_excel = importer_excel(fichier_excel)

        if not df_excel.empty:
            st.success("Fichier importé ✔")
            st.dataframe(df_excel)

            choix = st.selectbox(
                "Sélectionner une ligne à ajouter",
                df_excel.index,
                key="choix_excel",
                format_func=lambda i: f"{df_excel.loc[i,'nom']} — {df_excel.loc[i,'ref']}"
            )

            if st.button("➕ Ajouter depuis Excel", key="add_from_excel_btn"):
                ligne = df_excel.loc[choix]

                new = {
                    "nom": ligne["nom"],
                    "ref": ligne["ref"],
                    "date": ligne["date"],
                    "commentaire": "",
                    "statut": "Prévu",
                    "priorite": "auto",
                }

                df = pd.concat([df, pd.DataFrame([new])], ignore_index=True)
                sauvegarder_chantiers(df)
                st.success("Chantier ajouté ✔")
                st.rerun()


# ================== AJOUT MANUEL ==================
if admin:
    st.subheader("✏️ Ajouter manuellement un chantier")

    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom du chantier", key="nom_manual")
        date_montage = st.date_input("Date de montage", value=date.today(), key="date_manual")
    with col2:
        ref = st.text_input("Référence", key="ref_manual")
        statut = st.selectbox("Statut", ["Prévu", "En cours", "En attente", "Terminé"], key="statut_manual")

    commentaire = st.text_area("Commentaire", key="comment_manual")

    priorite_man = st.selectbox(
        "Priorité manuelle",
        ["Automatique", "Rouge", "Orange", "Jaune", "Gris"],
        key="prio_manual"
    )

    if st.button("Ajouter ce chantier", key="add_manual"):
        if nom and ref:
            new = {
                "nom": nom,
                "ref": ref,
                "date": str(date_montage),
                "commentaire": commentaire,
                "statut": statut,
                "priorite": priorite_man.lower() if priorite_man != "Automatique" else "auto",
            }
            df = pd.concat([df, pd.DataFrame([new])], ignore_index=True)
            sauvegarder_chantiers(df)
            st.success("Chantier ajouté ✔")
            st.rerun()
        else:
            st.error("Entrez au moins un nom et une référence.")


# ================== LISTE DES CHANTIERS ==================

st.subheader("📌 Liste des chantiers")

if df.empty:
    st.info("Aucun chantier enregistré.")
else:
    df_sorted = df.sort_values("date")
    styled = df_sorted.style.apply(lambda row: style_urgence(row, opts), axis=1)
    st.dataframe(styled, use_container_width=True)

    if admin:
        st.subheader("🛠 Modifier ou supprimer un chantier")

        idx = st.selectbox(
            "Sélectionner un chantier",
            df_sorted.index,
            key="edit_select",
            format_func=lambda i: f"{df_sorted.loc[i,'nom']} — {df_sorted.loc[i,'ref']}"
        )

        col1, col2 = st.columns(2)
        with col1:
            new_nom = st.text_input("Nom", df.loc[idx, "nom"], key=f"edit_nom_{idx}")
            current_date = df.loc[idx, "date"]
            new_date = st.date_input(
                "Date",
                value=current_date.date() if not pd.isna(current_date) else date.today(),
                key=f"edit_date_{idx}"
            )

        with col2:
            new_ref = st.text_input("Référence", df.loc[idx, "ref"], key=f"edit_ref_{idx}")
            new_statut = st.selectbox(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                index=["Prévu", "En cours", "En attente", "Terminé"].index(df.loc[idx, "statut"]),
                key=f"edit_statut_{idx}"     # 🔥 Correction du bug DuplicateElementId
            )

        new_commentaire = st.text_area("Commentaire", df.loc[idx, "commentaire"], key=f"edit_comment_{idx}")

        prio_list = ["Automatique", "Rouge", "Orange", "Jaune", "Gris"]
        current_prio = df.loc[idx, "priorite"]
        display_prio = current_prio.capitalize() if current_prio != "auto" else "Automatique"

        new_prio = st.selectbox(
            "Priorité manuelle",
            prio_list,
            index=prio_list.index(display_prio),
            key=f"edit_prio_{idx}"
        )

        colA, colB = st.columns(2)

        with colA:
            if st.button("💾 Enregistrer modifications", key=f"save_edit_{idx}"):
                df.loc[idx, "nom"] = new_nom
                df.loc[idx, "ref"] = new_ref
                df.loc[idx, "date"] = str(new_date)
                df.loc[idx, "statut"] = new_statut
                df.loc[idx, "commentaire"] = new_commentaire
                df.loc[idx, "priorite"] = new_prio.lower() if new_prio != "Automatique" else "auto"

                sauvegarder_chantiers(df)
                st.success("Modifié ✔")
                st.rerun()

        with colB:
            if st.button("🗑 Supprimer ce chantier", key=f"delete_{idx}"):
                df = df.drop(idx).reset_index(drop=True)
                sauvegarder_chantiers(df)
                st.success("Chantier supprimé ✔")
                st.rerun()


# ================== EXPORTS ==================

st.subheader("📤 Export")

if df.empty:
    st.info("Ajoutez un chantier pour activer l’export.")
else:
    col1, col2 = st.columns(2)
    with col1:
        st.download_button("📄 Export PDF", build_pdf(df), "chantiers.pdf", "application/pdf")
    with col2:
        st.download_button("📊 Export Excel", build_excel(df), "chantiers.xlsx")


# ================== QR CODE ==================

if opts["show_qr"]:
    st.subheader("📱 QR Code atelier")
    url = st.text_input("URL de l'application", key="qr_url")
    if url:
        st.image(build_qr(url), width=200)
