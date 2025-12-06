import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
import qrcode
from reportlab.lib.utils import ImageReader
from reportlab.lib.units import mm
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

    delta = (d.date() - date.today()).days
    if delta <= opts["rouge"]:
        return "#FF4B4B"
    if delta <= opts["orange"]:
        return "#FFA500"
    if delta <= opts["jaune"]:
        return "#FFD966"
    return "#DDDDDD"


def urgence_emoji(row, opts):
    d = row.get("date")
    if pd.isna(d):
        return "⚪"
    delta = (d.date() - date.today()).days
    if delta <= opts["rouge"]:
        return "🔴"
    if delta <= opts["orange"]:
        return "🟠"
    if delta <= opts["jaune"]:
        return "🟡"
    return "⚪"


def statut_emoji(statut):
    return {
        "Prévu": "⚪",
        "En cours": "🟣",
        "En attente": "🔵",
        "Terminé": "🟢",
    }.get(statut, "⚪")


# ==============================
# IMPORT EXCEL — OPTION A (Sélection)
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
# EXPORT PDF
# ==============================

def build_pdf(df, opts, qr_url=None):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    c.setFillColor(colors.HexColor("#2F3C7E"))
    c.rect(0, h - 60, w, 60, fill=1)
    c.setFont("Helvetica-Bold", 24)
    c.setFillColor(colors.white)
    c.drawCentredString(w/2, h - 35, "Gestion Projet Priorités")

    y = h - 90
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, y, "Légende :")
    y -= 20

    legend = [
        ("#FF4B4B", f"Urgence forte ≤ {opts['rouge']} j"),
        ("#FFA500", f"Urgence moyenne ≤ {opts['orange']} j"),
        ("#FFD966", f"Approche ≤ {opts['jaune']} j"),
        ("#F5EEDC", "Prévu"),
        ("#D9C8FF", "En cours"),
        ("#FFD6E7", "En attente"),
        ("#D9F8C4", "Terminé"),
    ]

    c.setFont("Helvetica", 10)
    for color_hex, label in legend:
        c.setFillColor(colors.HexColor(color_hex))
        c.rect(40, y - 8, 10, 10, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.drawString(55, y - 4, label)
        y -= 16

    y -= 20

    col_w = [30, 200, 60, 60, 70, 40]
    headers = ["Urg.", "Nom", "Code", "Date", "Statut", "État"]

    def draw_header(yy):
        c.setFillColor(colors.HexColor("#EFEFEF"))
        c.rect(40, yy - 18, sum(col_w), 18, fill=1)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)
        x = 44
        for i, htxt in enumerate(headers):
            c.drawString(x, yy - 13, htxt)
            x += col_w[i]
        return yy - 22

    y = draw_header(y)
    c.setFont("Helvetica", 9)

    for _, r in df.sort_values("date").iterrows():

        if y < 80:
            c.showPage()
            w, h = A4
            y = h - 60
            y = draw_header(y)
            c.setFont("Helvetica", 9)

        d = "-" if pd.isna(r["date"]) else r["date"].strftime("%d/%m/%Y")

        vals = [
            urgence_emoji(r, opts),
            r["nom"],
            r["ref"],
            d,
            r["statut"],
            statut_emoji(r["statut"])
        ]

        x = 44
        for i, v in enumerate(vals):
            c.drawString(x, y, str(v))
            x += col_w[i]

        y -= 18

    if qr_url:
        qr = qrcode.QRCode(box_size=3, border=1)
        qr.add_data(qr_url)
        qr.make()
        qr_img = qr.make_image()
        buf_qr = BytesIO()
        qr_img.save(buf_qr, format="PNG")
        buf_qr.seek(0)
        c.drawImage(ImageReader(buf_qr), w - 45*mm, 20*mm, width=30*mm)

    c.save()
    buf.seek(0)
    return buf


# =======================================================
# INTERFACE PRINCIPALE
# =======================================================

st.set_page_config(page_title="Gestion Priorité chantier", layout="wide")
st.title("📋 Gestion des priorités chantier")

mode = st.sidebar.selectbox("Mode", ["Lecture seule", "Administrateur"])
is_admin = False

if mode == "Administrateur":
    mdp = st.sidebar.text_input("Mot de passe", type="password")
    if mdp == ADMIN_PASSWORD:
        is_admin = True
    else:
        st.sidebar.error("Mot de passe incorrect")


df = charger_chantiers()
opts = charger_options()


# =============================
# GANTT FULLSCREEN
# =============================

if "gantt_fullscreen" in st.session_state and st.session_state["gantt_fullscreen"]:

    st.button("⬅️ Retour", on_click=lambda: st.session_state.pop("gantt_fullscreen"), key="gback")
    st.subheader("📊 Vue Gantt — Plein écran")

    if df.empty:
        st.info("Aucun chantier")
        st.stop()

    df_g = df[df["date"].notna()].copy()

    df_g["Start"] = df_g["date"]
    df_g["Finish"] = df_g["date"] + pd.Timedelta(days=1)

    df_g["label"] = (
        df_g["nom"].fillna("").astype(str)
        + " ("
        + df_g["ref"].fillna("").astype(str)
        + ")"
    )

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
# FILTRES
# =============================

with st.expander("🔎 Filtres de recherche"):
    col1, col2, col3 = st.columns(3)

    f_nom = col1.text_input("Rechercher un nom")
    f_ref = col2.text_input("Rechercher un code")
    f_statut = col3.multiselect("Filtrer par statut",
                                ["Prévu", "En cours", "En attente", "Terminé"])

df_filtered = df.copy()

if f_nom:
    df_filtered = df_filtered[df_filtered["nom"].str.contains(f_nom, case=False, na=False)]
if f_ref:
    df_filtered = df_filtered[df_filtered["ref"].str.contains(f_ref, case=False, na=False)]
if f_statut:
    df_filtered = df_filtered[df_filtered["statut"].isin(f_statut)]


# =============================
# GANTT BTN
# =============================

if st.button("📊 Afficher Gantt (plein écran)", key="ganttshow"):
    st.session_state["gantt_fullscreen"] = True
    st.rerun()


# =============================
# TABLEAU
# =============================

st.subheader("📌 Liste des chantiers")

if df_filtered.empty:
    st.info("Aucun chantier.")
else:
    disp = df_filtered.copy()
    disp["Urgence"] = disp.apply(lambda r: urgence_emoji(r, opts), axis=1)
    disp["État"] = disp["statut"].apply(statut_emoji)
    disp["date"] = disp["date"].dt.strftime("%d/%m/%Y")

    disp = disp[["Urgence", "nom", "ref", "date", "statut", "État", "commentaire"]]

    disp.rename(columns={
        "nom": "Nom de l'affaire",
        "ref": "Code affaire",
        "statut": "Statut",
        "commentaire": "Commentaire"
    }, inplace=True)

    st.dataframe(disp, use_container_width=True)


# =======================================================
# ADMIN
# =======================================================

if is_admin:

    # ------------------------
    # IMPORT EXCEL (OPTION A)
    # ------------------------
    with st.expander("📥 Import Excel — Sélection par chantier"):

        file = st.file_uploader("Importer un fichier Excel", type=["xlsx"], key="import_excel")

        if file:
            imported = importer_excel(file)

            if not imported.empty:
                st.success(f"{len(imported)} chantiers trouvés (Montage)")

                names = imported["nom"].fillna("Sans nom").tolist()
                idx_choice = st.selectbox("Sélectionner un chantier", range(len(imported)),
                                          format_func=lambda i: names[i],
                                          key="excel_select")

                row = imported.loc[idx_choice]

                st.info(f"""
                **Nom :** {row['nom']}  
                **Code :** {row['ref']}  
                **Date :** {row['date'].strftime('%d/%m/%Y') if not pd.isna(row['date']) else '-'}
                """)

                if st.button("➕ Ajouter ce chantier", key="excel_add"):
                    df.loc[len(df)] = [
                        row["nom"],
                        row["ref"],
                        row["date"],
                        "",
                        "Prévu",
                        ""
                    ]
                    sauvegarder_chantiers(df)
                    st.success("Chantier ajouté !")
                    st.rerun()

    # ------------------------
    # AJOUT MANUEL
    # ------------------------
    with st.expander("➕ Ajouter manuellement un chantier"):

        n_nom = st.text_input("Nom", key="m_nom")
        n_ref = st.text_input("Code", key="m_ref")
        n_date = st.date_input("Date", key="m_date")
        n_statut = st.selectbox("Statut", ["Prévu", "En cours", "En attente", "Terminé"],
                                key="m_statut")
        n_comment = st.text_area("Commentaire", key="m_comm")

        if st.button("Ajouter chantier", key="m_add"):
            df.loc[len(df)] = [n_nom, n_ref, pd.to_datetime(n_date), n_comment, n_statut, ""]
            sauvegarder_chantiers(df)
            st.success("Chantier ajouté")
            st.rerun()

    # ------------------------
    # MODIFIER / SUPPRIMER
    # ------------------------
    with st.expander("✏️ Modifier / Supprimer un chantier"):

        if df.empty:
            st.info("Aucun chantier.")
        else:
            idx = st.selectbox(
                "Sélectionner un chantier",
                df.index,
                format_func=lambda i: f"{df.loc[i,'nom']} - {df.loc[i,'ref']}",
                key="edit_select"
            )

            row = df.loc[idx]

            e_nom = st.text_input("Nom", row["nom"], key=f"edit_nom_{idx}")
            e_ref = st.text_input("Code", row["ref"], key=f"edit_ref_{idx}")

            e_date = st.date_input("Date",
                                   row["date"].date() if not pd.isna(row["date"]) else date.today(),
                                   key=f"edit_date_{idx}")

            e_statut = st.selectbox(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                index=["Prévu", "En cours", "En attente", "Terminé"].index(row["statut"]),
                key=f"edit_statut_{idx}"
            )

            e_comment = st.text_area("Commentaire", row["commentaire"], key=f"edit_comment_{idx}")

            c1, c2 = st.columns(2)

            if c1.button("💾 Enregistrer modifications", key=f"save_{idx}"):
                df.loc[idx] = [
                    e_nom, e_ref, pd.to_datetime(e_date), e_comment, e_statut, ""
                ]
                sauvegarder_chantiers(df)
                st.success("Modifié !")
                st.rerun()

            if c2.button("🗑 Supprimer", key=f"delete_{idx}"):
                df = df.drop(idx).reset_index(drop=True)
                sauvegarder_chantiers(df)
                st.success("Supprimé !")
                st.rerun()

    # ------------------------
    # OPTIONS
    # ------------------------
    with st.expander("⚙️ Options générales"):

        o1 = st.number_input("Urgence rouge ≤ jours", value=opts["rouge"], min_value=0)
        o2 = st.number_input("Urgence orange ≤ jours", value=opts["orange"], min_value=0)
        o3 = st.number_input("Urgence jaune ≤ jours", value=opts["jaune"], min_value=0)
        o4 = st.checkbox("Afficher QR dans PDF", value=opts["show_qr"])

        if st.button("Sauvegarder options", key="opt_save"):
            opts["rouge"] = o1
            opts["orange"] = o2
            opts["jaune"] = o3
            opts["show_qr"] = o4
            sauvegarder_options(opts)
            st.success("Options enregistrées !")
            st.rerun()


# =======================================================
# EXPORTS
# =======================================================

st.markdown("---")

cA, cB = st.columns(2)

with cA:
    if st.button("📄 Export PDF", key="pdf_btn"):
        qr_url = None
        if opts["show_qr"]:
            # Génération d’un lien partageable
            qr_url = st.experimental_get_query_params().get("share", [""])[0]

        pdf = build_pdf(df, opts, qr_url)
        st.download_button("Télécharger PDF",
                           data=pdf,
                           file_name=f"Gestion_Projet_Priorites_{date.today()}.pdf",
                           mime="application/pdf")

with cB:
    if st.button("📊 Export Excel", key="xlsx_btn"):
        excel = build_excel(df)
        st.download_button("Télécharger Excel",
                           data=excel,
                           file_name=f"Chantiers_{date.today()}.xlsx",
                           mime="application/vnd.ms-excel")
