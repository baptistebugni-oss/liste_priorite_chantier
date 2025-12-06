import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import qrcode
import plotly.express as px

# ==============================
# CONFIGURATION
# ==============================

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


# ==============================
# OUTILS DE DONNÉES
# ==============================

def ensure_columns(df, cols):
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    return df


def charger_chantiers():
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame(columns=COLUMNS)

    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        data = []

    df = pd.DataFrame(data)
    df = ensure_columns(df, COLUMNS)

    if not df.empty:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    return df


def sauvegarder_chantiers(df):
    df2 = df.copy()
    if not df2.empty:
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
    except Exception:
        opts = {}

    for k, v in DEFAULT_OPTIONS.items():
        opts.setdefault(k, v)

    return opts


def sauvegarder_options(opts):
    with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(opts, f, indent=4)


# ==============================
# COULEURS & PASTILLES
# ==============================

def get_urgence_color(row, opts):
    d = row.get("date")
    if pd.isna(d):
        return "#DDDDDD"

    delta = (d.date() - datetime.today().date()).days
    if delta <= opts["rouge"]:
        return "#FF4B4B"
    if delta <= opts["orange"]:
        return "#FFA500"
    if delta <= opts["jaune"]:
        return "#FFD966"
    return "#DDDDDD"


def get_statut_color(statut):
    statut = str(statut).lower().strip()
    mapping = {
        "prévu": "#F5EEDC",
        "en cours": "#D9C8FF",
        "en attente": "#FFD6E7",
        "terminé": "#D9F8C4",
    }
    return mapping.get(statut, "#DDDDDD")


def urgence_emoji(row, opts):
    d = row.get("date")
    if pd.isna(d):
        return "⚪"
    delta = (d.date() - datetime.today().date()).days
    if delta <= opts["rouge"]:
        return "🔴"
    if delta <= opts["orange"]:
        return "🟠"
    if delta <= opts["jaune"]:
        return "🟡"
    return "⚪"


def statut_emoji(statut):
    statut = str(statut).lower().strip()
    mapping = {
        "prévu": "⚪",
        "en cours": "🟣",
        "en attente": "🔵",
        "terminé": "🟢",
    }
    return mapping.get(statut, "⚪")


# ==============================
# IMPORT EXCEL
# ==============================

def importer_excel(file):
    df_x = pd.read_excel(file)

    df_x.columns = (
        df_x.columns
        .str.strip()
        .str.lower()
        .str.replace("’", "'", regex=False)
    )

    col_nom = "nom de l'affaire"
    col_ref = "code de l'affaire"
    col_comp = "compétence"
    col_debut = "début"

    missing = []
    mapping = {
        col_nom: "Nom de l'affaire",
        col_ref: "Code de l'affaire",
        col_comp: "Compétence",
        col_debut: "Début",
    }

    for col, label in mapping.items():
        if col not in df_x.columns:
            missing.append(label)

    if missing:
        st.error("❌ Colonnes manquantes dans l’Excel : " + ", ".join(missing))
        return pd.DataFrame()

    df_montage = df_x[df_x[col_comp].astype(str).str.lower() == "montage"]

    if df_montage.empty:
        st.warning("Aucune ligne 'Montage' trouvée.")
        return pd.DataFrame()

    df_out = df_montage[[col_nom, col_ref, col_debut]].copy()
    df_out.columns = ["nom", "ref", "date"]
    df_out["date"] = pd.to_datetime(df_out["date"], errors="coerce")

    return df_out


# ==============================
# EXPORT EXCEL
# ==============================

def build_excel(df):
    buf = BytesIO()
    df2 = df.copy()

    if not df2.empty and pd.api.types.is_datetime64_any_dtype(df2["date"]):
        df2["date"] = df2["date"].dt.strftime("%d/%m/%Y")

    df2.rename(columns={
        "nom": "Nom de l'affaire",
        "ref": "Code de l'affaire",
        "date": "Date",
        "commentaire": "Commentaire",
        "statut": "Statut",
        "priorite": "Priorité",
    }, inplace=True)

    df2.to_excel(buf, index=False)
    buf.seek(0)
    return buf


# ==============================
# EXPORT PDF (AVEC LÉGENDE ALIGNÉE + QR)
# ==============================

def build_pdf(df, opts, qr_url=None):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    # En-tête
    c.setFillColor(colors.HexColor("#2F3C7E"))
    c.rect(0, height - 60, width, 60, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(width / 2, height - 35, "Gestion Projet Priorités")
    c.setFont("Helvetica", 10)
    c.drawString(40, height - 52, f"Export du {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    # Légende
    y = height - 95
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, y, "Légende :")
    y -= 20
    c.setFont("Helvetica", 10)

    legend_items = [
        ("#FF4B4B", f"Urgence forte (≤ {opts['rouge']} j)"),
        ("#FFA500", f"Urgence moyenne (≤ {opts['orange']} j)"),
        ("#FFD966", f"Approche (≤ {opts['jaune']} j)"),
        ("#DDDDDD", "Pas urgent"),
        ("#F5EEDC", "Prévu"),
        ("#D9C8FF", "En cours"),
        ("#FFD6E7", "En attente"),
        ("#D9F8C4", "Terminé"),
    ]

    square_size = 10

    for color_hex, label in legend_items:
        square_y = y - (square_size / 2) - 3
        c.setFillColor(colors.HexColor(color_hex))
        c.rect(40, square_y, square_size, square_size, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.drawString(40 + square_size + 8, y - 3, label)
        y -= 18

    y -= 10

    # Tableau
    headers = ["Urg.", "Nom", "Code", "Date", "Statut", "État"]
    col_widths = [30, 210, 60, 60, 70, 40]
    row_h = 18
    left = 40

    def draw_header(ypos):
        c.setFillColor(colors.HexColor("#EFEFEF"))
        c.rect(left, ypos - row_h, sum(col_widths), row_h, fill=1)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)
        x = left + 4
        for i, h in enumerate(headers):
            c.drawString(x, ypos - 13, h)
            x += col_widths[i]
        return ypos - row_h - 4

    y = draw_header(y)
    c.setFont("Helvetica", 9)

    df_sorted = df.sort_values("date")

    for _, row in df_sorted.iterrows():
        if y < 80:
            c.showPage()
            width, height = A4
            y = height - 60
            y = draw_header(y)
            c.setFont("Helvetica", 9)

        nom = str(row["nom"])
        code = str(row["ref"])
        d = row["date"]
        date_str = "-" if pd.isna(d) else d.strftime("%d/%m/%Y")
        statut_txt = str(row["statut"])

        urg_color = get_urgence_color(row, opts)
        stat_color = get_statut_color(statut_txt)

        x = left

        # Pastille urgence
        c.setFillColor(colors.HexColor(urg_color))
        c.rect(x + 8, y - row_h + 4, 8, 8, fill=1)
        x += col_widths[0]

        # Nom
        c.setFillColor(colors.black)
        c.drawString(x + 2, y - row_h + 5, nom[:40])
        x += col_widths[1]

        # Code
        c.drawString(x + 2, y - row_h + 5, code[:15])
        x += col_widths[2]

        # Date
        c.drawString(x + 2, y - row_h + 5, date_str)
        x += col_widths[3]

        # Statut texte
        c.drawString(x + 2, y - row_h + 5, statut_txt[:15])
        x += col_widths[4]

        # Pastille statut
        c.setFillColor(colors.HexColor(stat_color))
        c.rect(x + 10, y - row_h + 4, 8, 8, fill=1)

        y -= row_h

    # QR code
    if qr_url:
        qr = qrcode.QRCode(box_size=3, border=1)
        qr.add_data(qr_url)
        qr.make(fit=True)
        img = qr.make_image()

        qr_buf = BytesIO()
        img.save(qr_buf, format="PNG")
        qr_buf.seek(0)

        c.drawImage(ImageReader(qr_buf), width - 45 * mm, 15 * mm, width=30 * mm)
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.black)
        c.drawRightString(width - 10 * mm, 12 * mm, "Accès application")

    c.save()
    buf.seek(0)
    return buf


# ==============================
# QR CODE DISPLAY
# ==============================

def build_qr_image(url):
    qr = qrcode.QRCode(box_size=8, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image()
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


# ==============================
# INTERFACE PRINCIPALE
# ==============================

st.set_page_config(page_title="Gestion Priorité Chantier", layout="wide")
st.title("📋 Gestion des priorités chantier")

# ------------------------------
# MODE : Lecture seule / Admin
# ------------------------------

mode = st.sidebar.selectbox("Mode", ["Lecture seule", "Administrateur"])
is_admin = False

if mode == "Administrateur":
    mdp = st.sidebar.text_input("Mot de passe", type="password")
    if mdp == ADMIN_PASSWORD:
        is_admin = True
    else:
        st.sidebar.error("Mot de passe incorrect")

# ------------------------------
# Chargement des données
# ------------------------------

df = charger_chantiers()
opts = charger_options()

# ===============================================================
#   BLOC GANTT PLEIN ÉCRAN (déjà dans Partie 1, NE PAS DUPLIQUER)
# ===============================================================

# (Ce bloc est déjà inclus dans la Partie 1 / 3 juste avant le stop()
#  donc rien à ajouter ici — on continue avec l’interface normale)


# ==============================
# PANNEAU DE FILTRE (collapsible)
# ==============================

with st.expander("🔎 Filtres de recherche", expanded=False):
    col1, col2, col3 = st.columns(3)

    with col1:
        filtre_nom = st.text_input("Rechercher un nom")

    with col2:
        filtre_ref = st.text_input("Rechercher un code affaire")

    with col3:
        filtre_statut = st.multiselect(
            "Filtrer par statut",
            ["Prévu", "En cours", "En attente", "Terminé"]
        )

# Application des filtres
df_filtered = df.copy()

if filtre_nom:
    df_filtered = df_filtered[df_filtered["nom"].str.contains(filtre_nom, case=False, na=False)]

if filtre_ref:
    df_filtered = df_filtered[df_filtered["ref"].str.contains(filtre_ref, case=False, na=False)]

if filtre_statut:
    df_filtered = df_filtered[df_filtered["statut"].isin(filtre_statut)]


# ==============================
# BOUTON AFFICHER GANTT PLEIN ÉCRAN
# ==============================

st.markdown("---")

if st.button("📊 Afficher le tableau Gantt (plein écran)", key="btn_show_gantt"):
    st.session_state["gantt_fullscreen"] = True
    st.rerun()


# ==============================
# AFFICHAGE TABLEAU PRINCIPAL
# ==============================

st.subheader("📌 Liste des chantiers")

if df_filtered.empty:
    st.info("Aucun chantier ne correspond aux filtres.")
else:
    # Ajout des pastilles urgence + statut
    df_display = df_filtered.copy()
    df_display["urgence"] = df_display.apply(lambda r: urgence_emoji(r, opts), axis=1)
    df_display["etat"] = df_display["statut"].apply(statut_emoji)

    # Nettoyage date affichée
    if not df_display.empty and pd.api.types.is_datetime64_any_dtype(df_display["date"]):
        df_display["date"] = df_display["date"].dt.strftime("%d/%m/%Y")

    df_display = df_display[[
        "urgence",
        "nom",
        "ref",
        "date",
        "statut",
        "etat",
        "commentaire",
    ]]

    df_display.rename(columns={
        "urgence": "Urgence",
        "nom": "Nom de l'affaire",
        "ref": "Code affaire",
        "date": "Date",
        "statut": "Statut",
        "etat": "État",
        "commentaire": "Commentaire"
    }, inplace=True)

    st.dataframe(df_display, use_container_width=True)


# ============================================================
#  SECTION ADMIN — OPTIONS, AJOUT, MODIFICATION, SUPPRESSION
# ============================================================

if is_admin:
    st.subheader("⚙️ Options administrateur")

    with st.expander("🎨 Options générales", expanded=False):
        st.markdown("### Seuils d'urgence (en jours)")
        rouge = st.number_input("Urgence forte (rouge)", 1, 30, opts["rouge"])
        orange = st.number_input("Urgence moyenne (orange)", 1, 30, opts["orange"])
        jaune = st.number_input("Approche (jaune)", 1, 60, opts["jaune"])

        show_qr = st.checkbox("Afficher QR-code dans l'application", value=opts["show_qr"])

        if st.button("💾 Sauvegarder les options"):
            opts["rouge"] = rouge
            opts["orange"] = orange
            opts["jaune"] = jaune
            opts["show_qr"] = show_qr
            sauvegarder_options(opts)
            st.success("Options sauvegardées !")
            st.rerun()

    # ----------------------------------------------------------
    # AJOUT MANUEL D’UN CHANTIER
    # ----------------------------------------------------------

    with st.expander("➕ Ajouter manuellement un chantier", expanded=False):

        new_nom = st.text_input("Nom de l'affaire")
        new_ref = st.text_input("Code affaire")
        new_date = st.date_input("Date prévue")
        new_statut = st.selectbox("Statut", ["Prévu", "En cours", "En attente", "Terminé"])
        new_comment = st.text_area("Commentaire")

        if st.button("Ajouter", key="btn_add_manual"):
            new_row = {
                "nom": new_nom,
                "ref": new_ref,
                "date": pd.to_datetime(new_date),
                "statut": new_statut,
                "commentaire": new_comment,
                "priorite": "",
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            sauvegarder_chantiers(df)
            st.success("Chantier ajouté !")
            st.rerun()

    # ----------------------------------------------------------
    # IMPORT EXCEL (Montage uniquement)
    # ----------------------------------------------------------

    with st.expander("📥 Importer depuis Excel (Montage)", expanded=False):
        file = st.file_uploader("Choisir un fichier Excel", type=["xlsx"])

        if file:
            df_import = importer_excel(file)
            if not df_import.empty:
                st.success("Données détectées !")
                st.dataframe(df_import)

                if st.button("📌 Importer ces chantiers", key="btn_import_excel"):
                    df_import["statut"] = "Prévu"
                    df_import["commentaire"] = ""
                    df_import["priorite"] = ""
                    df = pd.concat([df, df_import], ignore_index=True)
                    sauvegarder_chantiers(df)
                    st.success("Importation terminée !")
                    st.rerun()

    # ----------------------------------------------------------
    # MODIFIER OU SUPPRIMER UN CHANTIER
    # ----------------------------------------------------------

    with st.expander("✏️ Modifier ou supprimer un chantier", expanded=False):

        if df.empty:
            st.info("Aucun chantier à modifier.")
        else:
            idx = st.selectbox(
                "Sélectionner un chantier",
                df.index,
                format_func=lambda i: f"{df.loc[i, 'nom']} - {df.loc[i, 'ref']}",
            )

            old = df.loc[idx]

            new_nom2 = st.text_input("Nom", old["nom"])
            new_ref2 = st.text_input("Code", old["ref"])
            new_date2 = st.date_input("Date", old["date"].date() if not pd.isna(old["date"]) else date.today())
            new_statut2 = st.selectbox("Statut", ["Prévu", "En cours", "En attente", "Terminé"],
                                       index=["Prévu", "En cours", "En attente", "Terminé"].index(old["statut"]))
            new_comment2 = st.text_area("Commentaire", old["commentaire"])

            colA, colB = st.columns(2)

            with colA:
                if st.button("💾 Sauvegarder modifications", key="btn_edit_save"):
                    df.loc[idx, "nom"] = new_nom2
                    df.loc[idx, "ref"] = new_ref2
                    df.loc[idx, "date"] = pd.to_datetime(new_date2)
                    df.loc[idx, "statut"] = new_statut2
                    df.loc[idx, "commentaire"] = new_comment2
                    sauvegarder_chantiers(df)
                    st.success("Modifications enregistrées !")
                    st.rerun()

            with colB:
                if st.button("🗑 Supprimer ce chantier", key="btn_edit_delete"):
                    df = df.drop(idx).reset_index(drop=True)
                    sauvegarder_chantiers(df)
                    st.success("Chantier supprimé !")
                    st.rerun()


# ============================================================
# EXPORTS (PDF + Excel)
# ============================================================

st.subheader("📤 Exporter les données")

col_pdf, col_excel = st.columns(2)

with col_pdf:
    if st.button("📄 Exporter en PDF", key="btn_export_pdf"):
        qr_url_pdf = st.session_state.get("qr_url", None) if opts["show_qr"] else None
        pdf_bytes = build_pdf(df, opts, qr_url=qr_url_pdf)
        st.download_button(
            "📥 Télécharger PDF",
            data=pdf_bytes,
            file_name=f"Gestion_Projet_Priorites_{date.today()}.pdf",
            mime="application/pdf",
        )

with col_excel:
    if st.button("📊 Exporter en Excel", key="btn_export_excel"):
        excel_bytes = build_excel(df)
        st.download_button(
            "📥 Télécharger Excel",
            data=excel_bytes,
            file_name=f"Gestion_Projet_Priorites_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ============================================================
# QR CODE D’AFFICHAGE (OPTIONNEL)
# ============================================================

if opts.get("show_qr", True):
    st.markdown("---")
    st.subheader("📱 Accès mobile à l'application")
    st.write("Scannez ce QR-code pour accéder directement à l'application depuis votre téléphone.")

    app_url = st.experimental_get_query_params().get("url", [""])[0]
    if not app_url:
        app_url = st.request.url if hasattr(st, "request") else ""

    if app_url:
        st.session_state["qr_url"] = app_url
        qr_buf = build_qr_image(app_url)
        st.image(qr_buf, width=160)
    else:
        st.info("Impossible de détecter l’URL de l’application.")
