import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date, timedelta
from io import BytesIO

# ==== PDF ====
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import qrcode

# ==============================
# CONFIG
# ==============================

DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"
ADMIN_PASSWORD = "admin123"

DEFAULT_OPTIONS = {
    "rouge": 2,
    "orange": 7,
    "jaune": 14,
    "show_qr": True,
    "horizon": 60,
    "qr_url": "",
}

COLUMNS = ["nom", "ref", "date", "commentaire", "statut", "priorite"]

# ==============================
# COULEURS PDF
# ==============================

COLOR_URGENCE = {
    "rouge": "#FF4B4B",
    "orange": "#FFA500",
    "jaune": "#FFD966",
    "normal": "#DDDDDD",
}

COLOR_STATUT = {
    "Prévu": "#F5EEDC",
    "En cours": "#A066FF",
    "En attente": "#FF80B5",
    "Terminé": "#4CD964",
}

def draw_circle(c, x, y, size, color_hex):
    """Dessine une pastille ronde de couleur."""
    c.setFillColor(colors.HexColor(color_hex))
    c.circle(x, y, size, fill=1, stroke=0)


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
# URGENCE / STATUT LOGIQUE
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
# IMPORT EXCEL (Montage)
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
# UTILITAIRE TEXTE
# ==============================

def truncate_nom(text, length=30):
    text = str(text)
    return text if len(text) <= length else text[:length] + "…"


# ==============================
# EXPORT PDF (SIMPLE, LISIBLE)
# ==============================

def build_pdf_liste(df, opts, qr_url=None):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    page_w, page_h = A4
    margin_left = 30

    # =========================
    # EN-TÊTE
    # =========================
    c.setFillColor(colors.HexColor("#2F3C7E"))
    c.rect(0, page_h - 70, page_w, 70, fill=1)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(page_w / 2, page_h - 38, "Gestion Projet Priorités")

    c.setFont("Helvetica", 10)
    c.drawString(20, page_h - 60, f"Export du {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    y = page_h - 110

    # =========================
    # LÉGENDE
    # =========================
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.black)
    c.drawString(margin_left, y, "Légende :")
    y -= 15

    legend_items = [
        ("#FF4B4B", f"Urgence forte (≤ {opts['rouge']} j)"),
        ("#FFA500", f"Urgence moyenne (≤ {opts['orange']} j)"),
        ("#FFD966", f"Approche (≤ {opts['jaune']} j)"),
        ("#F5EEDC", "Prévu"),
        ("#A066FF", "En cours"),
        ("#FF80B5", "En attente"),
        ("#4CD964", "Terminé"),
    ]

    for color_hex, text_label in legend_items:
        draw_circle(c, margin_left + 4, y - 3, 4, color_hex)
        c.setFillColor(colors.black)
        c.drawString(margin_left + 15, y - 6, text_label)
        y -= 15

    y -= 20

    # =========================
    # TABLEAU PRINCIPAL
    # =========================
    col_widths = [22, 190, 60, 60, 70, 22]
    headers = ["", "Nom de l'affaire", "Code", "Date", "Statut", ""]

    def draw_header(ypos):
        c.setFillColor(colors.HexColor("#EFEFEF"))
        c.rect(margin_left, ypos - 18, sum(col_widths), 18, fill=1)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)

        xx = margin_left + 5
        for h, w in zip(headers, col_widths):
            c.drawString(xx, ypos - 12, h)
            xx += w

        return ypos - 35  # marges améliorées

    y = draw_header(y)
    c.setFont("Helvetica", 9)

    df_sorted = df.sort_values(by="date")

    for _, row in df_sorted.iterrows():

        if y < 80:
            c.showPage()
            c.setFont("Helvetica", 10)
            y = page_h - 60
            y = draw_header(y)
            c.setFont("Helvetica", 9)

        # --- Pastille urgence (gauche)
        urg_hex = get_urgence_color(row, opts)
        draw_circle(
            c,
            margin_left + col_widths[0] / 2,
            y + 6,
            4,
            urg_hex,
        )

        # --- Pastille statut (droite)
        stat_hex = COLOR_STATUT.get(row["statut"], "#F5EEDC")
        draw_circle(
            c,
            margin_left + sum(col_widths) - col_widths[-1] / 2,
            y + 6,
            4,
            stat_hex,
        )

        # --- Texte
        c.setFillColor(colors.black)
        dstr = "-" if pd.isna(row["date"]) else row["date"].strftime("%d/%m/%Y")

        values = ["", row["nom"], row["ref"], dstr, row["statut"], ""]

        xx = margin_left + 5
        for val, w in zip(values, col_widths):
            c.drawString(xx, y, str(val))
            xx += w

        y -= 18

    # ===============================
    # SÉPARATEUR AVANT VUE PAR SEMAINE
    # ===============================
    c.setStrokeColor(colors.HexColor("#BBBBBB"))
    c.setLineWidth(1)
    c.line(margin_left, y + 12, page_w - margin_left, y + 12)
    y -= 35

    # ===============================
    # TITRE : VUE PAR SEMAINE
    # ===============================
    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(colors.black)
    c.drawString(margin_left, y, "📅 Vue par semaine")
    y -= 25

    # ===============================
    # CONTENU : VUE PAR SEMAINE
    # ===============================
    df_week = df[df["date"].notna()].copy()

    if not df_week.empty:

        iso = df_week["date"].dt.isocalendar()
        df_week["week"] = iso.week
        df_week["year"] = iso.year
        df_week = df_week.sort_values("date")

        mois_fr = {
            1: "janvier", 2: "février", 3: "mars", 4: "avril",
            5: "mai", 6: "juin", 7: "juillet", 8: "août",
            9: "septembre", 10: "octobre", 11: "novembre", 12: "décembre"
        }

        for (year, week), g in df_week.groupby(["year", "week"]):

            if y < 120:
                c.showPage()
                c.setFont("Helvetica", 12)
                y = page_h - 60
                c.setFillColor(colors.black)

            start = date.fromisocalendar(int(year), int(week), 1)
            endw = date.fromisocalendar(int(year), int(week), 7)

            if start.month == endw.month:
                peri = f"du {start.day} au {endw.day} {mois_fr[start.month]} {year}"
            else:
                peri = (
                    f"du {start.day} {mois_fr[start.month]} "
                    f"au {endw.day} {mois_fr[endw.month]} {year}"
                )

            # --- Titre semaine
            c.setFont("Helvetica-Bold", 12)
            c.setFillColor(colors.black)
            c.drawString(margin_left, y, f"Semaine {int(week):02d} — {peri}")
            y -= 18

            # --- Lignes des chantiers
            for _, row in g.iterrows():

                if y < 100:
                    c.showPage()
                    c.setFont("Helvetica", 12)
                    y = page_h - 60
                    c.setFillColor(colors.black)

                # Pastille urgence
                urg_hex = get_urgence_color(row, opts)
                draw_circle(c, margin_left + 5, y + 1, 4, urg_hex)

                # Texte
                c.setFont("Helvetica", 10)
                c.setFillColor(colors.black)
                dstr = row["date"].strftime("%d/%m")
                txt = f"{dstr} — {row['nom']} ({row['ref']})"
                c.drawString(margin_left + 18, y, txt)

                # Pastille statut
                stat_hex = COLOR_STATUT.get(row["statut"], "#F5EEDC")
                draw_circle(c, margin_left + 300, y + 1, 4, stat_hex)

                y -= 16

    else:
        c.setFont("Helvetica", 10)
        c.setFillColor(colors.black)
        c.drawString(margin_left, y, "Aucune date disponible pour la vue par semaine.")
        y -= 20

    c.showPage()
    c.save()
    buf.seek(0)
    return buf


# =======================================================
# INTERFACE STREAMLIT — CONFIG DE BASE
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
# 📅 VUE PAR SEMAINE (TIMELINE)
# =============================
with st.expander("📅 Vue par semaine (ligne du temps)", expanded=False):

    df_week = df.copy()
    df_week = df_week[df_week["date"].notna()]

    if df_week.empty:
        st.info("Aucun chantier avec une date valide.")
    else:
        horizon = opts.get("horizon", 60)
        today = date.today()
        end_date = today + timedelta(days=horizon)

        # Filtrer l'horizon
        df_week = df_week[
            (df_week["date"].dt.date >= today) &
            (df_week["date"].dt.date <= end_date)
        ]

        if df_week.empty:
            st.warning("Aucun chantier dans l’horizon sélectionné.")
        else:
            # Extraire semaine + année
            iso = df_week["date"].dt.isocalendar()
            df_week["week"] = iso.week
            df_week["year"] = iso.year

            df_week = df_week.sort_values("date")

            # Mois en français
            mois_fr = {
                1: "janvier", 2: "février", 3: "mars", 4: "avril",
                5: "mai", 6: "juin", 7: "juillet", 8: "août",
                9: "septembre", 10: "octobre", 11: "novembre", 12: "décembre"
            }

            # Affichage par semaine
            for (year, week), g in df_week.groupby(["year", "week"]):

                start = date.fromisocalendar(int(year), int(week), 1)
                endw = date.fromisocalendar(int(year), int(week), 7)

                # Contraintes horizon
                if start < today:
                    start = today
                if endw > end_date:
                    endw = end_date

                # Texte de période
                if start.month == endw.month:
                    peri = f"du {start.day} au {endw.day} {mois_fr[start.month]} {year}"
                else:
                    peri = (
                        f"du {start.day} {mois_fr[start.month]} "
                        f"au {endw.day} {mois_fr[endw.month]} {year}"
                    )

                st.markdown(f"### 🗓️ Semaine {int(week):02d} — {peri}")

                # Affichage des chantiers
                for _, row in g.iterrows():
                    urg = urgence_emoji(row, opts)
                    stat = statut_emoji(row["statut"])
                    dstr = row["date"].strftime("%d/%m")
                    nom = truncate_nom(row["nom"], 40)

                    c1, c2, c3, c4 = st.columns([1, 2, 6, 3])
                    c1.write(urg)
                    c2.write(dstr)
                    c3.write(f"{nom} ({row['ref']})")
                    c4.write(f"{stat} {row['statut']}")


# =======================================================
# ADMINISTRATION
# =======================================================

if is_admin:

    # IMPORT EXCEL
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

    # AJOUT MANUEL
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

    # MODIFIER / SUPPRIMER
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

    # OPTIONS
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
# EXPORTS (PDF + EXCEL)
# =======================================================

st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    if st.button("📄 Export PDF (liste)", key="pdf_liste_btn"):
        qr_url_for_pdf = opts.get("qr_url") if opts.get("show_qr") and opts.get("qr_url") else None
        pdf = build_pdf_liste(df, opts, qr_url_for_pdf)
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
