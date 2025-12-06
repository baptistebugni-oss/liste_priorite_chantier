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
    except:
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
    except:
        opts = {}

    for k, v in DEFAULT_OPTIONS.items():
        opts.setdefault(k, v)

    return opts


def sauvegarder_options(opts):
    with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(opts, f, indent=4)


# ==============================
# COULEURS DES PASTILLES
# ==============================

def get_urgence_color(row, opts):
    d = row.get("date")
    if pd.isna(d):
        return "#DDDDDD"

    delta = (d.date() - datetime.today().date()).days

    if delta <= opts["rouge"]:
        return "#FF4B4B"   # rouge
    if delta <= opts["orange"]:
        return "#FFA500"   # orange
    if delta <= opts["jaune"]:
        return "#FFD966"   # jaune

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
    for col, label in [
        (col_nom, "Nom de l'affaire"),
        (col_ref, "Code de l'affaire"),
        (col_comp, "Compétence"),
        (col_debut, "Début"),
    ]:
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
# EXPORT PDF PREMIUM
# ==============================

def build_pdf(df, opts, qr_url=None):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    # ------- TITRE -------
    c.setFillColor(colors.HexColor("#2F3C7E"))
    c.rect(0, height - 60, width, 60, fill=1)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(width / 2, height - 35, "Gestion Projet Priorités")

    c.setFont("Helvetica", 10)
    c.drawString(40, height - 52, f"Export du {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    # ------- LÉGENDE -------
    y = height - 95
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, y, "Légende :")

    legend = [
        ("#FF4B4B", f"Urgence forte (≤ {opts['rouge']} j)"),
        ("#FFA500", f"Urgence moyenne (≤ {opts['orange']} j)"),
        ("#FFD966", f"Approche (≤ {opts['jaune']} j)"),
        ("#DDDDDD", "Pas urgent"),
        ("#F5EEDC", "Prévu"),
        ("#D9C8FF", "En cours"),
        ("#FFD6E7", "En attente"),
        ("#D9F8C4", "Terminé"),
    ]

    y -= 18
    c.setFont("Helvetica", 10)

    for color_hex, label in legend:
        c.setFillColor(colors.HexColor(color_hex))
        c.rect(40, y - 8, 8, 8, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.drawString(56, y - 2, label)
        y -= 14

    y -= 10

    # ------- TABLEAU -------
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

        # Colonne urgence
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

    if qr_url:
        qr = qrcode.QRCode(box_size=3, border=1)
        qr.add_data(qr_url)
        qr.make(fit=True)
        img = qr.make_image()

        qr_buf = BytesIO()
        img.save(qr_buf, format="PNG")
        qr_buf.seek(0)

        c.drawImage(ImageReader(qr_buf), width - 45 * mm, 15 * mm, width=30 * mm)
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 9)
        c.drawRightString(width - 10 * mm, 12 * mm, "Accès application")

    c.save()
    buf.seek(0)
    return buf


# ==============================
# QR CODE DANS L’APP
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
# INTERFACE STREAMLIT
# ==============================

st.set_page_config(page_title="Gestion Chantier", layout="wide")
st.title("📋 Gestion des priorités chantier")

# Mode admin
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

if "show_filters" not in st.session_state:
    st.session_state["show_filters"] = False
if "qr_url" not in st.session_state:
    st.session_state["qr_url"] = ""

st.divider()

# ===================================================================
# OPTIONS ADMIN
# ===================================================================

if admin:
    with st.expander("⚙️ Options générales (admin)", expanded=False):
        col1, col2 = st.columns(2)

        with col1:
            r = st.number_input("Jours ROUGE", 0, 90, opts["rouge"])
            o = st.number_input("Jours ORANGE", 0, 90, opts["orange"])
            j = st.number_input("Jours JAUNE", 0, 90, opts["jaune"])

        with col2:
            qr_enabled = st.checkbox("Inclure QR dans PDF", value=opts["show_qr"])

        if st.button("💾 Sauvegarder options"):
            opts.update({
                "rouge": r,
                "orange": o,
                "jaune": j,
                "show_qr": qr_enabled
            })
            sauvegarder_options(opts)
            st.success("✔ Options enregistrées")

st.divider()

# ===================================================================
# IMPORT EXCEL
# ===================================================================

if admin:
    with st.expander("📥 Import Excel (Montage)", expanded=False):
        excel_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

        if excel_file:
            df_import = importer_excel(excel_file)
            if not df_import.empty:
                st.success("Import réussi ✔")
                st.dataframe(df_import)

                row_id = st.selectbox(
                    "Sélectionner un chantier",
                    df_import.index,
                    format_func=lambda i: df_import.loc[i, "nom"]
                )

                if st.button("➕ Ajouter depuis Excel"):
                    row = df_import.loc[row_id]
                    new_row = {
                        "nom": row["nom"],
                        "ref": row["ref"],
                        "date": row["date"],
                        "commentaire": "",
                        "statut": "Prévu",
                        "priorite": "auto",
                    }
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    sauvegarder_chantiers(df)
                    st.success("Chantier ajouté ✔")
                    st.rerun()

st.divider()

# ===================================================================
# AJOUT MANUEL
# ===================================================================

if admin:
    with st.expander("✏️ Ajouter un chantier", expanded=False):
        colA, colB = st.columns(2)

        with colA:
            nom = st.text_input("Nom de l'affaire")
            date_m = st.date_input("Date de montage", value=date.today())

        with colB:
            ref = st.text_input("Code de l'affaire")
            statut = st.selectbox("Statut", ["Prévu", "En cours", "En attente", "Terminé"])

        commentaire = st.text_area("Commentaire")

        prio = st.selectbox("Priorité", ["Automatique", "Rouge", "Orange", "Jaune", "Gris"])

        if st.button("➕ Ajouter"):
            if nom and ref:
                df = pd.concat([df, pd.DataFrame([{
                    "nom": nom,
                    "ref": ref,
                    "date": str(date_m),
                    "commentaire": commentaire,
                    "statut": statut,
                    "priorite": prio.lower() if prio != "Automatique" else "auto",
                }])], ignore_index=True)

                sauvegarder_chantiers(df)
                st.success("Chantier ajouté ✔")
                st.rerun()
            else:
                st.error("Nom et Code obligatoires.")

# ===================================================================
# TABLEAU AVEC PASTILLES
# ===================================================================

st.subheader("📌 Liste des chantiers")

if df.empty:
    st.info("Aucun chantier.")
else:
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("🔎 Filtres"):
            st.session_state["show_filters"] = not st.session_state["show_filters"]

    if st.session_state["show_filters"]:
        col_filters, col_tab = st.columns([1, 3])
    else:
        col_filters, col_tab = st.columns([0.0001, 1])

    # ---- FILTRES ----
    if st.session_state["show_filters"]:
        stat_list = sorted(df["statut"].dropna().unique().tolist())

        search = st.text_input("Recherche", key="FILTER_TEXT")
        selected_stats = st.multiselect("Statuts", stat_list, default=stat_list)

        dates_valides = df["date"].dropna()
        if not dates_valides.empty:
            dmin = dates_valides.min().date()
            dmax = dates_valides.max().date()
            d1 = st.date_input("Date min", value=dmin)
            d2 = st.date_input("Date max", value=dmax)
        else:
            d1 = d2 = None

        if st.button("Réinitialiser"):
            st.session_state["FILTER_TEXT"] = ""
            st.rerun()

    # ---- TABLEAU ----
    with col_tab:
        df2 = df.copy()

        if st.session_state.get("FILTER_TEXT"):
            txt = st.session_state["FILTER_TEXT"].lower()
            mask = (
                df2["nom"].astype(str).str.lower().str.contains(txt) |
                df2["ref"].astype(str).str.lower().str.contains(txt) |
                df2["commentaire"].astype(str).str.lower().str.contains(txt)
            )
            df2 = df2[mask]

        if st.session_state.get("show_filters") and selected_stats:
            df2 = df2[df2["statut"].isin(selected_stats)]

        if d1 and d2:
            mask = df2["date"].notna()
            mask &= df2["date"].dt.date.between(d1, d2)
            df2 = df2[mask]

        df2_disp = df2.copy()

        df2_disp["Date"] = df2_disp["date"].dt.strftime("%d/%m/%Y").fillna("-")
        df2_disp["Urgence"] = df2.apply(lambda row: "●", axis=1)
        df2_disp["État"] = df2.apply(lambda row: "●", axis=1)

        df2_disp.rename(columns={
            "nom": "Nom de l'affaire",
            "ref": "Code de l'affaire",
            "statut": "Statut",
            "commentaire": "Commentaire",
        }, inplace=True)

        st.dataframe(df2_disp, use_container_width=True)

# ===================================================================
# MODIFIER / SUPPRIMER
# ===================================================================

if admin and not df.empty:
    with st.expander("🛠 Modifier / Supprimer", expanded=False):
        idx = st.selectbox(
            "Sélectionner un chantier",
            df.index,
            format_func=lambda i: f"{df.loc[i,'nom']} — {df.loc[i,'ref']}"
        )

        col1, col2 = st.columns(2)
        with col1:
            new_nom = st.text_input("Nom", df.loc[idx, "nom"])
            dval = df.loc[idx, "date"]
            new_date = st.date_input("Date", value=dval.date() if not pd.isna(dval) else date.today())
        with col2:
            new_ref = st.text_input("Code", df.loc[idx, "ref"])
            new_statut = st.selectbox("Statut", ["Prévu", "En cours", "En attente", "Terminé"],
                                      index=["Prévu", "En cours", "En attente", "Terminé"].index(df.loc[idx, "statut"]))

        new_com = st.text_area("Commentaire", df.loc[idx, "commentaire"])

        new_prio = st.selectbox("Priorité", ["Automatique", "Rouge", "Orange", "Jaune", "Gris"])

        cc1, cc2 = st.columns(2)

        with cc1:
            if st.button("💾 Enregistrer"):
                df.loc[idx, "nom"] = new_nom
                df.loc[idx, "ref"] = new_ref
                df.loc[idx, "date"] = str(new_date)
                df.loc[idx, "statut"] = new_statut
                df.loc[idx, "commentaire"] = new_com
                df.loc[idx, "priorite"] = new_prio.lower() if new_prio != "Automatique" else "auto"
                sauvegarder_chantiers(df)
                st.success("Modifié ✔")
                st.rerun()

        with cc2:
            if st.button("🗑 Supprimer"):
                df = df.drop(idx).reset_index(drop=True)
                sauvegarder_chantiers(df)
                st.success("Supprimé ✔")
                st.rerun()

# ===================================================================
# EXPORTS
# ===================================================================

st.subheader("📤 Exporter")

colP, colE = st.columns(2)

with colP:
    st.download_button(
        "📄 Export PDF",
        build_pdf(df, opts, qr_url=st.session_state.get("qr_url") or None),
        file_name="Gestion_Projet_Priorites.pdf",
        mime="application/pdf"
    )

with colE:
    st.download_button(
        "📊 Export Excel",
        build_excel(df),
        file_name="chantiers.xlsx"
    )

# ===================================================================
# QR CODE AFFICHAGE
# ===================================================================

if opts["show_qr"]:
    st.subheader("📱 QR Code atelier")

    qr_url = st.text_input("URL de l'application", value=st.session_state["qr_url"])
    st.session_state["qr_url"] = qr_url

    if qr_url:
        st.image(build_qr_image(qr_url), width=180)
        st.caption("Scannez ce QR code pour accéder à l'application")
