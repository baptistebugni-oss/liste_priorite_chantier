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
# CHARGEMENT / SAUVEGARDE
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
    mapping = {
        "Prévu": "#F5EEDC",
        "En cours": "#D9C8FF",
        "En attente": "#FFD6E7",
        "Terminé": "#D9F8C4",
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
    mapping = {
        "Prévu": "⚪",
        "En cours": "🟣",
        "En attente": "🔵",
        "Terminé": "🟢",
    }
    return mapping.get(statut, "⚪")


# ==============================
# IMPORT EXCEL
# ==============================

def importer_excel(file):
    df_x = pd.read_excel(file)
    df_x.columns = df_x.columns.str.strip().str.lower()

    required = ["nom de l'affaire", "code de l'affaire", "compétence", "début"]

    missing = [c for c in required if c not in df_x.columns]
    if missing:
        st.error("Colonnes manquantes : " + ", ".join(missing))
        return pd.DataFrame()

    df_m = df_x[df_x["compétence"].astype(str).str.lower() == "montage"]
    if df_m.empty:
        st.warning("Aucune ligne 'Montage'.")
        return pd.DataFrame()

    df_out = df_m[["nom de l'affaire", "code de l'affaire", "début"]].copy()
    df_out.columns = ["nom", "ref", "date"]
    df_out["date"] = pd.to_datetime(df_out["date"], errors="coerce")
    return df_out


# ==============================
# EXPORT EXCEL
# ==============================

def build_excel(df):
    buf = BytesIO()
    df2 = df.copy()
    df2["date"] = df2["date"].dt.strftime("%d/%m/%Y")
    df2.to_excel(buf, index=False)
    buf.seek(0)
    return buf


# ==============================
# EXPORT PDF (Légende alignée + QR)
# ==============================

def build_pdf(df, opts, qr_url=None):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    # Bandeau
    c.setFillColor(colors.HexColor("#2F3C7E"))
    c.rect(0, height - 60, width, 60, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(width / 2, height - 35, "Gestion Projet Priorités")

    # Légende
    y = height - 95
    legend = [
        ("#FF4B4B", f"Urgence forte (≤ {opts['rouge']} j)"),
        ("#FFA500", f"Urgence moyenne (≤ {opts['orange']} j)"),
        ("#FFD966", f"Approche (≤ {opts['jaune']} j)"),
        ("#F5EEDC", "Prévu"),
        ("#D9C8FF", "En cours"),
        ("#FFD6E7", "En attente"),
        ("#D9F8C4", "Terminé"),
    ]

    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.black)
    c.drawString(40, y, "Légende :")
    y -= 20

    c.setFont("Helvetica", 10)
    for color_hex, label in legend:
        c.setFillColor(colors.HexColor(color_hex))
        c.rect(40, y - 8, 10, 10, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.drawString(55, y - 4, label)
        y -= 16

    y -= 10

    # Tableau PDF
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

    for _, r in df.sort_values("date").iterrows():
        if y < 80:
            c.showPage()
            width, height = A4
            y = height - 60
            y = draw_header(y)
            c.setFont("Helvetica", 9)

        urg = urgence_emoji(r, opts)
        stc = statut_emoji(r["statut"])
        d = "-" if pd.isna(r["date"]) else r["date"].strftime("%d/%m/%Y")

        vals = [
            urg,
            r["nom"],
            r["ref"],
            d,
            r["statut"],
            stc
        ]

        x = left + 4
        for i, v in enumerate(vals):
            c.drawString(x, y - row_h + 5, str(v))
            x += col_widths[i]

        y -= row_h

    # QR
    if qr_url:
        qr = qrcode.QRCode(box_size=3, border=1)
        qr.add_data(qr_url)
        qr.make(fit=True)
        img = qr.make_image()
        buf_qr = BytesIO()
        img.save(buf_qr, format="PNG")
        buf_qr.seek(0)
        c.drawImage(ImageReader(buf_qr), width - 45*mm, 20*mm, width=30*mm)

    c.save()
    buf.seek(0)
    return buf


# =======================================================
#         INTERFACE PRINCIPALE STREAMLIT
# =======================================================

st.set_page_config(page_title="Gestion Priorité Chantier", layout="wide")
st.title("📋 Gestion des priorités chantier")

mode = st.sidebar.selectbox("Mode", ["Lecture seule", "Administrateur"])
is_admin = False

if mode == "Administrateur":
    pwd = st.sidebar.text_input("Mot de passe", type="password")
    if pwd == ADMIN_PASSWORD:
        is_admin = True
    else:
        st.sidebar.error("Mot de passe incorrect")


# Chargement
df = charger_chantiers()
opts = charger_options()


# =============================
# GANTT PLEIN ÉCRAN
# =============================

if "gantt_fullscreen" in st.session_state and st.session_state["gantt_fullscreen"]:

    st.button("⬅️ Retour", on_click=lambda: st.session_state.pop("gantt_fullscreen"), key="btn_back")
    st.subheader("📊 Vue Gantt — Plein écran")

    if df.empty:
        st.info("Aucun chantier")
        st.stop()

    df_g = df.copy()
    df_g = df_g[df_g["date"].notna()]
    df_g["Start"] = df_g["date"]
    df_g["Finish"] = df_g["date"] + pd.Timedelta(days=1)
    df_g["label"] = df_g["nom"] + " (" + df_g["ref"] + ")"

    fig = px.timeline(
        df_g,
        x_start="Start",
        x_end="Finish",
        y="label",
        color="statut",
        color_discrete_map={
            "Prévu": "#F5EEDC",
            "En cours": "#D9C8FF",
            "En attente": "#FFD6E7",
            "Terminé": "#D9F8C4",
        }
    )

    fig.update_yaxes(autorange="reversed")
    fig.update_layout(height=900)

    st.plotly_chart(fig, use_container_width=True)
    st.stop()


# =============================
# PANNEAU FILTRE
# =============================

with st.expander("🔎 Filtres de recherche", expanded=False):
    col1, col2, col3 = st.columns(3)

    filtre_nom = col1.text_input("Rechercher un nom")
    filtre_ref = col2.text_input("Rechercher un code")
    filtre_statut = col3.multiselect("Filtrer par statut",
                                     ["Prévu", "En cours", "En attente", "Terminé"])

df_filtered = df.copy()

if filtre_nom:
    df_filtered = df_filtered[df_filtered["nom"].str.contains(filtre_nom, case=False, na=False)]

if filtre_ref:
    df_filtered = df_filtered[df_filtered["ref"].str.contains(filtre_ref, case=False, na=False)]

if filtre_statut:
    df_filtered = df_filtered[df_filtered["statut"].isin(filtre_statut)]


# =============================
# BOUTON GANTT
# =============================

if st.button("📊 Afficher le tableau Gantt (plein écran)", key="btn_gantt"):
    st.session_state["gantt_fullscreen"] = True
    st.rerun()


# =============================
# TABLEAU PRINCIPAL
# =============================

st.subheader("📌 Liste des chantiers")

if df_filtered.empty:
    st.info("Aucun chantier ne correspond aux filtres.")
else:
    df_disp = df_filtered.copy()

    df_disp["Urgence"] = df_disp.apply(lambda r: urgence_emoji(r, opts), axis=1)
    df_disp["État"] = df_disp["statut"].apply(statut_emoji)

    df_disp["date"] = df_disp["date"].dt.strftime("%d/%m/%Y")

    df_disp = df_disp[["Urgence", "nom", "ref", "date", "statut", "État", "commentaire"]]

    df_disp.rename(columns={
        "nom": "Nom de l'affaire",
        "ref": "Code affaire",
        "statut": "Statut",
        "commentaire": "Commentaire"
    }, inplace=True)

    st.dataframe(df_disp, use_container_width=True)


# =======================================================
#                 ADMINISTRATION
# =======================================================

if is_admin:

    # -------------------------------------
    # IMPORT EXCEL
    # -------------------------------------
    with st.expander("📥 Import Excel (Montage uniquement)", expanded=False):
        fexcel = st.file_uploader("Importer un fichier Excel", type=["xlsx"])
        if fexcel:
            df_new = importer_excel(fexcel)
            if not df_new.empty:
                st.success(f"{len(df_new)} chantiers chargés")
                if st.button("Ajouter à la base", key="btn_import_add"):
                    df = pd.concat([df, df_new], ignore_index=True)
                    sauvegarder_chantiers(df)
                    st.success("Import terminé !")
                    st.rerun()

    # -------------------------------------
    # AJOUT MANUEL
    # -------------------------------------
    with st.expander("➕ Ajouter manuellement un chantier", expanded=False):

        new_nom = st.text_input("Nom de l'affaire", key="add_nom")
        new_ref = st.text_input("Code affaire", key="add_ref")
        new_date = st.date_input("Date", key="add_date")
        new_stat = st.selectbox("Statut", ["Prévu", "En cours", "En attente", "Terminé"], key="add_stat")
        new_comment = st.text_area("Commentaire", key="add_comment")

        if st.button("Ajouter", key="btn_add_manual"):
            df.loc[len(df)] = [
                new_nom, new_ref,
                pd.to_datetime(new_date),
                new_comment, new_stat, ""
            ]
            sauvegarder_chantiers(df)
            st.success("Chantier ajouté !")
            st.rerun()

    # -------------------------------------
    # MODIFIER / SUPPRIMER
    # -------------------------------------
    with st.expander("✏️ Modifier ou supprimer un chantier", expanded=False):

        if df.empty:
            st.info("Aucun chantier.")
        else:
            idx = st.selectbox(
                "Sélectionner",
                df.index,
                format_func=lambda i: f"{df.loc[i,'nom']} - {df.loc[i,'ref']}",
                key="edit_selector"
            )

            row = df.loc[idx]

            e_nom = st.text_input("Nom", row["nom"], key=f"e_nom_{idx}")
            e_ref = st.text_input("Code", row["ref"], key=f"e_ref_{idx}")
            e_date = st.date_input("Date",
                                   row["date"].date() if not pd.isna(row["date"]) else date.today(),
                                   key=f"e_date_{idx}")

            e_statut = st.selectbox(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                index=["Prévu", "En cours", "En attente", "Terminé"].index(row["statut"]),
                key=f"e_statut_{idx}"
            )

            e_comment = st.text_area("Commentaire", row["commentaire"], key=f"e_comment_{idx}")

            c1, c2 = st.columns(2)

            if c1.button("💾 Enregistrer", key=f"btn_save_{idx}"):
                df.loc[idx, "nom"] = e_nom
                df.loc[idx, "ref"] = e_ref
                df.loc[idx, "date"] = pd.to_datetime(e_date)
                df.loc[idx, "statut"] = e_statut
                df.loc[idx, "commentaire"] = e_comment
                sauvegarder_chantiers(df)
                st.success("Modifications enregistrées")
                st.rerun()

            if c2.button("🗑 Supprimer", key=f"btn_del_{idx}"):
                df = df.drop(idx).reset_index(drop=True)
                sauvegarder_chantiers(df)
                st.success("Chantier supprimé")
                st.rerun()

    # -------------------------------------
    # OPTIONS
    # -------------------------------------
    with st.expander("⚙️ Options générales", expanded=False):
        rouge = st.number_input("Urgence forte (rouge) ≤ jours", min_value=0, value=opts["rouge"])
        orange = st.number_input("Urgence moyenne (orange) ≤ jours", min_value=0, value=opts["orange"])
        jaune = st.number_input("Approche (jaune) ≤ jours", min_value=0, value=opts["jaune"])

        show_qr = st.checkbox("Afficher QR sur PDF", value=opts["show_qr"])

        if st.button("💾 Sauvegarder options", key="btn_opt_save"):
            opts["rouge"] = rouge
            opts["orange"] = orange
            opts["jaune"] = jaune
            opts["show_qr"] = show_qr
            sauvegarder_options(opts)
            st.success("Options enregistrées !")
            st.rerun()

# =======================================================
#                EXPORTS
# =======================================================

st.markdown("---")

colA, colB = st.columns(2)

with colA:
    if st.button("📄 Export PDF", key="btn_export_pdf"):
        qr_url = None
        if opts["show_qr"]:
            qr_url = st.experimental_get_query_params().get("share", [""])[0]

        pdf_file = build_pdf(df, opts, qr_url=qr_url)
        st.download_button("Télécharger PDF",
                           data=pdf_file,
                           file_name=f"Gestion_Projet_Priorites_{date.today()}.pdf",
                           mime="application/pdf")

with colB:
    if st.button("📊 Export Excel", key="btn_export_xlsx"):
        excel = build_excel(df)
        st.download_button("Télécharger Excel",
                           data=excel,
                           file_name=f"Chantiers_{date.today()}.xlsx",
                           mime="application/vnd.ms-excel")
