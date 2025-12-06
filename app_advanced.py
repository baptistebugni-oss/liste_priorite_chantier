import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date
from io import BytesIO
import calendar

import matplotlib.pyplot as plt

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
    "rouge": 2,       # seuil en jours pour urgence rouge
    "orange": 7,      # seuil en jours pour urgence orange
    "jaune": 14,      # seuil en jours pour urgence jaune
    "show_qr": True,  # afficher un QR-code dans le PDF de liste
    "horizon": 60,    # horizon de calendrier en jours
    "qr_url": "",     # URL utilisée pour le QR-code
}

COLUMNS = ["nom", "ref", "date", "commentaire", "statut", "priorite"]


# ==============================
# OUTILS DONNÉES
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
        json.dump(opts, f, ensure_ascii=False, indent=4)


# ==============================
# COULEURS / ÉMOTICÔNES
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
# IMPORT EXCEL (Montage uniquement, sélection unitaire)
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
    if not df2.empty:
        df2["date"] = df2["date"].dt.strftime("%d/%m/%Y")
    df2.to_excel(buf, index=False)
    buf.seek(0)
    return buf


# ==============================
# EXPORT PDF — LISTE DES CHANTIERS
# ==============================

def build_pdf_liste(df, opts, qr_url=None):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4

    # En-tête
    c.setFillColor(colors.HexColor("#2F3C7E"))
    c.rect(0, h - 60, w, 60, fill=1)
    c.setFont("Helvetica-Bold", 24)
    c.setFillColor(colors.white)
    c.drawCentredString(w / 2, h - 35, "Gestion Projet Priorités")
    c.setFont("Helvetica", 10)
    c.drawString(40, h - 52, f"Export du {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    # Légende
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
            statut_emoji(r["statut"]),
        ]

        x = 44
        for i, v in enumerate(vals):
            c.drawString(x, y, str(v))
            x += col_w[i]

        y -= 18

    # QR-code éventuel
    if qr_url:
        qr = qrcode.QRCode(box_size=3, border=1)
        qr.add_data(qr_url)
        qr.make()
        img = qr.make_image()
        buf_qr = BytesIO()
        img.save(buf_qr, format="PNG")
        buf_qr.seek(0)
        c.drawImage(ImageReader(buf_qr), w - 45 * mm, 20 * mm, width=30 * mm)
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.black)
        c.drawRightString(w - 10 * mm, 15 * mm, "Accès application")

    c.save()
    buf.seek(0)
    return buf


# ==============================
# CALENDRIER — FIGURE MATPLOTLIB
# ==============================

def truncate(text, length=14):
    """Coupe le nom de l'affaire proprement pour éviter les débordements."""
    text = str(text)
    return text if len(text) <= length else text[:length] + "…"


def build_calendar_figure(df, opts):
    """Construit la figure matplotlib du calendrier à partir d'un DataFrame complet df."""
    df_cal = df.copy()
    df_cal = df_cal[df_cal["date"].notna()]

    if df_cal.empty:
        return None

    horizon = opts.get("horizon", 60)
    today = date.today()
    end_date = today + pd.Timedelta(days=horizon)

    df_cal = df_cal[
        (df_cal["date"].dt.date >= today) &
        (df_cal["date"].dt.date <= end_date)
    ]

    if df_cal.empty:
        return None

    df_cal["jour"] = df_cal["date"].dt.day
    df_cal["mois_num"] = df_cal["date"].dt.month
    df_cal["annee"] = df_cal["date"].dt.year

    df_cal["mois_label"] = df_cal.apply(
        lambda r: f"{calendar.month_name[r['mois_num']]} {r['annee']}",
        axis=1
    )

    mois_uniques = df_cal["mois_label"].unique().tolist()

    statut_colors = {
        "Prévu": "#B39B6B",
        "En cours": "#7F3FBF",
        "En attente": "#E75480",
        "Terminé": "#2E8B57",
    }

    fig, ax = plt.subplots(figsize=(10, 14))
    offsets = {}

    for _, r in df_cal.iterrows():
        x = mois_uniques.index(r["mois_label"])
        y = int(r["jour"])

        key = (x, y)
        offset = offsets.get(key, 0)
        offsets[key] = offset + 1

        y_text = y - (offset * 0.22)

        color = statut_colors.get(r["statut"], "black")

        ax.text(
            x,
            y_text,
            truncate(r["nom"], 14),
            ha="center",
            va="center",
            fontsize=8,
            color=color
        )

    ax.set_xticks(range(len(mois_uniques)))
    ax.set_xticklabels(mois_uniques, rotation=30, ha="right")

    ax.set_yticks(range(1, 32))
    ax.set_ylim(32, 0)
    ax.set_ylabel("Jour du mois")
    ax.set_xlabel("Mois")

    ax.grid(True, linestyle="--", linewidth=0.3, alpha=0.5)

    ax.set_title(f"Chantiers sur {horizon} jours à partir d'aujourd'hui")

    plt.tight_layout()
    return fig


def build_calendar_pdf(df, opts):
    """Construit un PDF du calendrier via matplotlib."""
    fig = build_calendar_figure(df, opts)
    if fig is None:
        return None

    buf = BytesIO()
    fig.savefig(buf, format="pdf", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


# =======================================================
# INTERFACE PRINCIPALE STREAMLIT
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
# FILTRES
# =============================

with st.expander("🔎 Filtres de recherche", expanded=False):
    col1, col2, col3 = st.columns(3)

    f_nom = col1.text_input("Rechercher un nom")
    f_ref = col2.text_input("Rechercher un code")
    f_statut = col3.multiselect(
        "Filtrer par statut",
        ["Prévu", "En cours", "En attente", "Terminé"],
    )

df_filtered = df.copy()

if f_nom:
    df_filtered = df_filtered[df_filtered["nom"].str.contains(f_nom, case=False, na=False)]
if f_ref:
    df_filtered = df_filtered[df_filtered["ref"].str.contains(f_ref, case=False, na=False)]
if f_statut:
    df_filtered = df_filtered[df_filtered["statut"].isin(f_statut)]


# =============================
# TABLEAU PRINCIPAL
# =============================

st.subheader("📌 Liste des chantiers")

if df_filtered.empty:
    st.info("Aucun chantier.")
else:
    disp = df_filtered.copy()
    disp["Urgence"] = disp.apply(lambda r: urgence_emoji(r, opts), axis=1)
    disp["État"] = disp["statut"].apply(statut_emoji)

    if not disp.empty:
        disp["date"] = disp["date"].dt.strftime("%d/%m/%Y")

    disp = disp[["Urgence", "nom", "ref", "date", "statut", "État", "commentaire"]]

    disp.rename(columns={
        "nom": "Nom de l'affaire",
        "ref": "Code affaire",
        "statut": "Statut",
        "commentaire": "Commentaire",
    }, inplace=True)

    st.dataframe(disp, use_container_width=True)


# =============================
# CALENDRIER — EXPANDER
# =============================

with st.expander("📆 Calendrier des chantiers", expanded=False):
    st.markdown("### Calendrier des chantiers")

    if df_filtered.empty:
        st.info("Aucun chantier à afficher dans le calendrier.")
    else:
        fig = build_calendar_figure(df_filtered, opts)
        if fig is None:
            st.info("Aucune date valide ou aucun chantier dans l'horizon défini.")
        else:
            st.pyplot(fig, use_container_width=True)


# =======================================================
# ADMINISTRATION
# =======================================================

if is_admin:

    # ------------------------
    # IMPORT EXCEL
    # ------------------------
    with st.expander("📥 Import Excel — Sélection unitaire", expanded=False):
        file = st.file_uploader("Importer un fichier Excel", type=["xlsx"], key="excel_file")

        if file:
            imported = importer_excel(file)

            if not imported.empty:
                st.success(f"{len(imported)} chantiers trouvés (Montage)")

                names = imported["nom"].fillna("Sans nom").tolist()

                idx_choice = st.selectbox(
                    "Sélectionner un chantier",
                    range(len(imported)),
                    format_func=lambda i: names[i],
                    key="excel_select",
                )

                row = imported.iloc[idx_choice]

                st.info(
                    f"**Nom :** {row['nom']}\n\n"
                    f"**Code :** {row['ref']}\n\n"
                    f"**Date :** "
                    f"{row['date'].strftime('%d/%m/%Y') if not pd.isna(row['date']) else '-'}"
                )

                if st.button("➕ Ajouter ce chantier", key="excel_add"):
                    df.loc[len(df)] = [
                        row["nom"],
                        row["ref"],
                        row["date"],
                        "",
                        "Prévu",
                        "",
                    ]
                    sauvegarder_chantiers(df)
                    st.success("Chantier ajouté !")
                    st.rerun()

    # ------------------------
    # AJOUT MANUEL
    # ------------------------
    with st.expander("➕ Ajouter manuellement un chantier", expanded=False):

        n_nom = st.text_input("Nom", key="m_nom")
        n_ref = st.text_input("Code", key="m_ref")
        n_date = st.date_input("Date", key="m_date")
        n_statut = st.selectbox(
            "Statut",
            ["Prévu", "En cours", "En attente", "Terminé"],
            key="m_statut",
        )
        n_comment = st.text_area("Commentaire", key="m_comm")

        if st.button("Ajouter chantier", key="m_add"):
            df.loc[len(df)] = [
                n_nom,
                n_ref,
                pd.to_datetime(n_date),
                n_comment,
                n_statut,
                "",
            ]
            sauvegarder_chantiers(df)
            st.success("Chantier ajouté")
            st.rerun()

    # ------------------------
    # MODIFIER / SUPPRIMER
    # ------------------------
    with st.expander("✏️ Modifier / Supprimer un chantier", expanded=False):
        if df.empty:
            st.info("Aucun chantier.")
        else:
            idx = st.selectbox(
                "Sélectionner",
                df.index,
                format_func=lambda i: f"{df.loc[i, 'nom']} - {df.loc[i, 'ref']}",
                key="edit_select",
            )

            row = df.loc[idx]

            e_nom = st.text_input("Nom", row["nom"], key=f"e_nom_{idx}")
            e_ref = st.text_input("Code", row["ref"], key=f"e_ref_{idx}")
            e_date = st.date_input(
                "Date",
                row["date"].date() if not pd.isna(row["date"]) else date.today(),
                key=f"e_date_{idx}",
            )
            e_statut = st.selectbox(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                index=["Prévu", "En cours", "En attente", "Terminé"].index(row["statut"]),
                key=f"e_statut_{idx}",
            )
            e_comment = st.text_area("Commentaire", row["commentaire"], key=f"e_comment_{idx}")

            c1, c2 = st.columns(2)

            if c1.button("💾 Sauvegarder", key=f"save_{idx}"):
                df.loc[idx] = [
                    e_nom,
                    e_ref,
                    pd.to_datetime(e_date),
                    e_comment,
                    e_statut,
                    "",
                ]
                sauvegarder_chantiers(df)
                st.success("Modifié !")
                st.rerun()

            if c2.button("🗑 Supprimer", key=f"del_{idx}"):
                df = df.drop(idx).reset_index(drop=True)
                sauvegarder_chantiers(df)
                st.success("Supprimé !")
                st.rerun()

    # ------------------------
    # OPTIONS
    # ------------------------
    with st.expander("⚙️ Options générales", expanded=False):

        o1 = st.number_input("Urgence rouge ≤ jours", value=opts["rouge"], min_value=0)
        o2 = st.number_input("Urgence orange ≤ jours", value=opts["orange"], min_value=0)
        o3 = st.number_input("Urgence jaune ≤ jours", value=opts["jaune"], min_value=0)
        o4 = st.checkbox("Afficher QR dans le PDF de liste", value=opts["show_qr"])
        o5 = st.number_input(
            "Horizon calendrier (en jours)",
            value=opts.get("horizon", 60),
            min_value=7,
            max_value=365,
        )
        o6 = st.text_input(
            "URL à encoder dans le QR-code (PDF liste)",
            value=opts.get("qr_url", ""),
        )

        if st.button("Sauvegarder options", key="opt_save"):
            opts["rouge"] = o1
            opts["orange"] = o2
            opts["jaune"] = o3
            opts["show_qr"] = o4
            opts["horizon"] = o5
            opts["qr_url"] = o6
            sauvegarder_options(opts)
            st.success("Options enregistrées !")
            st.rerun()


# =======================================================
# EXPORTS
# =======================================================

st.markdown("---")

col1, col2, col3 = st.columns(3)

with col1:
    if st.button("📄 Export PDF (liste)", key="pdf_liste_btn"):
        qr_url = opts["qr_url"] if opts.get("show_qr") and opts.get("qr_url") else None
        pdf = build_pdf_liste(df, opts, qr_url)
        st.download_button(
            "Télécharger PDF liste",
            data=pdf,
            file_name=f"Gestion_Projet_Priorites_{date.today()}.pdf",
            mime="application/pdf",
        )

with col2:
    if st.button("📊 Export Excel", key="xlsx_btn"):
        excel = build_excel(df)
        st.download_button(
            "Télécharger Excel",
            data=excel,
            file_name=f"Chantiers_{date.today()}.xlsx",
            mime="application/vnd.ms-excel",
        )

with col3:
    if st.button("📄 Export PDF (calendrier)", key="pdf_cal_btn"):
        pdf_cal = build_calendar_pdf(df, opts)
        if pdf_cal is None:
            st.warning("Impossible de générer le calendrier : aucune date dans l'horizon.")
        else:
            st.download_button(
                "Télécharger PDF calendrier",
                data=pdf_cal,
                file_name=f"Calendrier_Chantiers_{date.today()}.pdf",
                mime="application/pdf",
            )
