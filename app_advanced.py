import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date, timedelta
from io import BytesIO

st.set_page_config(
    page_title="Gestion Projet Priorités",
    page_icon="app_icon.ico",    
)

st.write("Icon exists:", os.path.exists("app_icon.ico"))

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
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 12)
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

    for color_hex, label in legend_items:
        draw_circle(c, margin_left + 4, y - 3, 4, color_hex)
        c.setFillColor(colors.black)
        c.drawString(margin_left + 15, y - 6, label)
        y -= 15

    y -= 20

    # ===============================
    # CALCUL AUTOMATIQUE DES COLONNES
    # ===============================
    headers = ["", "Nom de l'affaire", "Code", "Date", "Statut", ""]

    columns = [
        [""] * len(df),
        df["nom"].astype(str).tolist(),
        df["ref"].astype(str).tolist(),
        df["date"].dt.strftime("%d/%m/%Y").fillna("-").tolist(),
        df["statut"].astype(str).tolist(),
        [""] * len(df),
    ]

    c.setFont("Helvetica", 10)
    max_widths = []

    for header, col in zip(headers, columns):
        header_w = c.stringWidth(header, "Helvetica-Bold", 10)
        cell_w = max(c.stringWidth(str(x), "Helvetica", 9) for x in col)
        max_widths.append(max(header_w, cell_w) + 12)

    total_target = 22 + 190 + 60 + 60 + 70 + 22
    sum_widths = sum(max_widths)

    if sum_widths > total_target:
        ratio = total_target / sum_widths
        col_widths = [w * ratio for w in max_widths]
    else:
        col_widths = max_widths

    # =========================
    # TABLEAU
    # =========================
    def draw_header(ypos):
        c.setFillColor(colors.HexColor("#EFEFEF"))
        c.rect(margin_left, ypos - 18, sum(col_widths), 18, fill=1)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)

        xx = margin_left + 5
        for h, w in zip(headers, col_widths):
            c.drawString(xx, ypos - 12, h)
            xx += w

        return ypos - 35

    y = draw_header(y)
    c.setFont("Helvetica", 9)

    df_sorted = df.sort_values("date")

    for _, row in df_sorted.iterrows():

        if y < 80:
            c.showPage()
            y = page_h - 60
            y = draw_header(y)
            c.setFont("Helvetica", 9)

        # Pastille urgence
        urg = get_urgence_color(row, opts)
        draw_circle(
            c,
            margin_left + col_widths[0] / 2,
            y + 4,
            4,
            urg,
        )

        # Pastille statut
        stat = COLOR_STATUT.get(row["statut"], "#F5EEDC")
        draw_circle(
            c,
            margin_left + sum(col_widths) - col_widths[-1] / 2,
            y + 5,
            4,
            stat,
        )

        # Texte
        c.setFillColor(colors.black)
        dd = "-" if pd.isna(row["date"]) else row["date"].strftime("%d/%m/%Y")

        values = ["", row["nom"], row["ref"], dd, row["statut"], ""]

        xx = margin_left + 5
        for val, w in zip(values, col_widths):
            c.drawString(xx, y, str(val))
            xx += w

        y -= 18

    # ===============================
    # SÉPARATEUR
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

        for (year, week), grp in df_week.groupby(["year", "week"]):

            if y < 120:
                c.showPage()
                y = page_h - 60
                c.setFont("Helvetica-Bold", 14)
                c.setFillColor(colors.black)
                c.drawString(margin_left, y, "📅 Vue par semaine")
                y -= 25

            start = date.fromisocalendar(year, week, 1)
            end = date.fromisocalendar(year, week, 7)

            if start.month == end.month:
                peri = f"du {start.day} au {end.day} {mois_fr[start.month]} {year}"
            else:
                peri = (
                    f"du {start.day} {mois_fr[start.month]} "
                    f"au {end.day} {mois_fr[end.month]} {year}"
                )

            c.setFont("Helvetica-Bold", 12)
            c.setFillColor(colors.black)
            c.drawString(margin_left, y, f"Semaine {week:02d} — {peri}")
            y -= 18

            for _, row in grp.iterrows():

                if y < 100:
                    c.showPage()
                    y = page_h - 60

                urg = get_urgence_color(row, opts)
                draw_circle(c, margin_left + 5, y + 2, 4, urg)

                c.setFillColor(colors.black)
                c.setFont("Helvetica", 10)

                # ======= DATE FRANÇAISE LONGUE =======
                jours_fr = {
                    0: "Lundi",
                    1: "Mardi",
                    2: "Mercredi",
                    3: "Jeudi",
                    4: "Vendredi",
                    5: "Samedi",
                    6: "Dimanche"
                }

                mois_fr_long = {
                    1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril",
                    5: "Mai", 6: "Juin", 7: "Juillet", 8: "Août",
                    9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
                }

                d = row["date"]

                jour_txt = jours_fr[d.weekday()]
                mois_txt = mois_fr_long[d.month]

                date_longue = f"{jour_txt} {d.day:02d} {mois_txt}"

                txt = f"{date_longue} — {row['nom']} ({row['ref']})"
                
                # Limite de largeur avant la pastille statut
                max_text_width = 300 - 40  # point fixe avant la pastille, marge de sécurité

                # Mesure de la largeur du texte
                text_width = c.stringWidth(txt, "Helvetica", 10)

                # Si le texte dépasse, on le coupe proprement et ajoute "..."
                if text_width > max_text_width:
                    cutoff = len(txt)
                    while c.stringWidth(txt[:cutoff] + "…", "Helvetica", 10) > max_text_width and cutoff > 10:
                        cutoff -= 1
                    txt_to_print = txt[:cutoff] + "…"
                else:
                    txt_to_print = txt

                # Impression du texte sécurisé
                c.drawString(margin_left + 18, y, txt_to_print)

                stat = COLOR_STATUT.get(row["statut"], "#F5EEDC")
                draw_circle(c, margin_left + 300, y + 2, 4, stat)

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
