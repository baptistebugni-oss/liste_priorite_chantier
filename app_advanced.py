import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date
from io import BytesIO
import calendar

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
    "horizon": 60,    # horizon de la vue "par semaine" en jours
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
    c.setFillColor(colors.white)
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
# UTILITAIRE TEXTE
# ==============================

def truncate_nom(text, length=30):
    """Coupe le nom de l'affaire proprement pour éviter les débordements."""
    text = str(text)
    return text if len(text) <= length else text[:length] + "…"


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
# VUE PAR SEMAINE (LIGNE DU TEMPS)
# =============================

with st.expander("📅 Vue par semaine (ligne du temps)", expanded=False):
    if df_filtered.empty:
        st.info("Aucun chantier à afficher.")
    else:
        df_week = df_filtered.copy()
        df_week = df_week[df_week["date"].notna()]

        if df_week.empty:
            st.info("Aucune date valide pour cette vue.")
        else:
            horizon = opts.get("horizon", 60)
            today = date.today()
            end_date = today + pd.Timedelta(days=horizon)

            df_week = df_week[
                (df_week["date"].dt.date >= today) &
                (df_week["date"].dt.date <= end_date)
            ]

            if df_week.empty:
                st.warning("Aucun chantier dans l'horizon défini.")
            else:
                iso = df_week["date"].dt.isocalendar()
                df_week["week"] = iso.week
                df_week["year"] = iso.year

                df_week = df_week.sort_values("date")

                # Mapping des noms de mois en français
                mois_fr = {
                    1: "janvier",
                    2: "février",
                    3: "mars",
                    4: "avril",
                    5: "mai",
                    6: "juin",
                    7: "juillet",
                    8: "août",
                    9: "septembre",
                    10: "octobre",
                    11: "novembre",
                    12: "décembre",
                }

                for (year, week), g in df_week.groupby(["year", "week"]):
                    # Début et fin de semaine
                    start = date.fromisocalendar(year, int(week), 1)
                    endw = date.fromisocalendar(year, int(week), 7)

                    # Ajuster à l'horizon
                    if start < today:
                        start = today
                    if endw > end_date.date():
                        endw = end_date.date()

                    if start.month == endw.month:
                        range_str = f"du {start.day} au {endw.day} {mois_fr[start.month]} {year}"
                    else:
                        range_str = (
                            f"du {start.day} {mois_fr[start.month]} "
                            f"au {endw.day} {mois_fr[endw.month]} {year}"
                        )

                    st.markdown(
                        f"**──── Semaine {int(week):02d} ({range_str}) ────**"
                    )

                    for _, row in g.iterrows():
                        urg = urgence_emoji(row, opts)
                        stat_e = statut_emoji(row["statut"])
                        dstr = row["date"].strftime("%d/%m")
                        nom_tr = truncate_nom(row["nom"], 35)
                        ref = row["ref"]

                        c1, c2, c3, c4 = st.columns([1, 2, 5, 3])
                        c1.write(urg)
                        c2.write(dstr)
                        c3.write(f"{nom_tr} ({ref})")
                        c4.write(f"{stat_e} {row['statut']}")


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

                row_imp = imported.iloc[idx_choice]

                st.info(
                    f"**Nom :** {row_imp['nom']}\n\n"
                    f"**Code :** {row_imp['ref']}\n\n"
                    f"**Date :** "
                    f"{row_imp['date'].strftime('%d/%m/%Y') if not pd.isna(row_imp['date']) else '-'}"
                )

                if st.button("➕ Ajouter ce chantier", key="excel_add"):
                    df.loc[len(df)] = [
                        row_imp["nom"],
                        row_imp["ref"],
                        row_imp["date"],
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

            row_ed = df.loc[idx]

            e_nom = st.text_input("Nom", row_ed["nom"], key=f"e_nom_{idx}")
            e_ref = st.text_input("Code", row_ed["ref"], key=f"e_ref_{idx}")
            e_date = st.date_input(
                "Date",
                row_ed["date"].date() if not pd.isna(row_ed["date"]) else date.today(),
                key=f"e_date_{idx}",
            )
            e_statut = st.selectbox(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                index=["Prévu", "En cours", "En attente", "Terminé"].index(row_ed["statut"]),
                key=f"e_statut_{idx}",
            )
            e_comment = st.text_area("Commentaire", row_ed["commentaire"], key=f"e_comment_{idx}")

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
            "Horizon vue 'par semaine' (en jours)",
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

col1, col2 = st.columns(2)

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
