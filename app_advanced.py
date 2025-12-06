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

import altair as alt

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
# EXPORT PDF (AVEC LÉGENDE ALIGNÉE)
# ==============================

def build_pdf(df, opts, qr_url=None):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    # ----------- TITRE ----------- #
    c.setFillColor(colors.HexColor("#2F3C7E"))
    c.rect(0, height - 60, width, 60, fill=1)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(width / 2, height - 35, "Gestion Projet Priorités")
    c.setFont("Helvetica", 10)
    c.drawString(40, height - 52, f"Export du {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    # ----------- LÉGENDE ----------- #
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

    square_size = 10  # Taille uniforme des pastilles

    for color_hex, label in legend_items:

        # Centrage vertical du carré avec le texte
        square_y = y - (square_size / 2) - 3

        # Carré
        c.setFillColor(colors.HexColor(color_hex))
        c.rect(40, square_y, square_size, square_size, fill=1, stroke=0)

        # Texte, parfaitement aligné
        c.setFillColor(colors.black)
        c.drawString(40 + square_size + 8, y - 3, label)

        y -= 18  # Espacement propre

    y -= 10

    # ----------- TABLEAU ----------- #
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

        # Urgence carré
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

        # Statut carré
        c.setFillColor(colors.HexColor(stat_color))
        c.rect(x + 10, y - row_h + 4, 8, 8, fill=1)

        y -= row_h

    # ----------- QR CODE ----------- #
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
# INTERFACE STREAMLIT
# ==============================

st.set_page_config(page_title="Gestion Chantier", layout="wide")
st.title("📋 Gestion des priorités chantier")

# Mode admin
mode = st.sidebar.selectbox("Mode", ["Lecture seule", "Administrateur"], key="mode_selector")
admin = False
if mode == "Administrateur":
    mdp = st.sidebar.text_input("Mot de passe", type="password", key="password_admin")
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

# ==============================
# MODE PLEIN ÉCRAN POUR LE GANTT
# ==============================

if "gantt_fullscreen" not in st.session_state:
    st.session_state["gantt_fullscreen"] = False

# Si on est en mode plein écran, on n'affiche QUE le Gantt
if st.session_state["gantt_fullscreen"]:
    st.subheader("📈 Tableau Gantt – Vue plein écran (60 jours)")

    if df.empty:
        st.info("Aucun chantier à afficher.")
    else:
        df_graph = df.dropna(subset=["date"]).copy()

        if df_graph.empty:
            st.info("Aucune date valide pour afficher le Gantt.")
        else:
            today = datetime.today().date()
            horizon = today + pd.Timedelta(days=60)

            # Filtre sur 60 jours
            df_graph = df_graph[
                (df_graph["date"].dt.date >= today) &
                (df_graph["date"].dt.date <= horizon)
            ]

            if df_graph.empty:
                st.warning("Aucun chantier dans les 60 prochains jours.")
            else:
                # Préparation des colonnes pour le Gantt
                df_graph["start"] = df_graph["date"]
                df_graph["end"] = df_graph["date"] + pd.Timedelta(days=1)

                # Mapping de couleurs pour les statuts
                color_map = {
                    "Prévu": "#F5EEDC",
                    "En cours": "#D9C8FF",
                    "En attente": "#FFD6E7",
                    "Terminé": "#D9F8C4",
                }

                # On force les statuts connus, sinon "Prévu" par défaut
                df_graph["statut"] = df_graph["statut"].fillna("Prévu")
                df_graph.loc[~df_graph["statut"].isin(color_map.keys()), "statut"] = "Prévu"

                fig = px.timeline(
                    df_graph,
                    x_start="start",
                    x_end="end",
                    y="nom",
                    color="statut",
                    color_discrete_map=color_map,
                    hover_data={
                        "ref": True,
                        "commentaire": True,
                        "date": True,
                        "start": False,
                        "end": False,
                    },
                )

                # Inverser l'ordre des chantiers (plus lisible)
                fig.update_yaxes(autorange="reversed", title_text="Chantiers")

                # Format de date à la française
                fig.update_xaxes(
                    title_text="Date",
                    range=[pd.to_datetime(today), pd.to_datetime(horizon)],
                    tickformat="%d/%m/%Y"
                )

                # Ligne "aujourd'hui"
                fig.add_vline(
                    x=pd.to_datetime(today),
                    line_color="red",
                    line_dash="dash",
                    annotation_text="Aujourd'hui",
                    annotation_position="top left"
                )

                fig.update_layout(
                    height=600,
                    margin=dict(l=60, r=20, t=60, b=40),
                    legend_title_text="Statut",
                )

                st.plotly_chart(fig, use_container_width=True)

    # Bouton retour AVANT le st.stop()
if st.button("🔙 Retour à la vue principale", key="btn_back_gantt"):
    st.session_state["gantt_fullscreen"] = False
    st.rerun()   # nouvelle méthode

# Et seulement APRES, on coupe le reste
st.stop()
    

# ==============================
# OPTIONS ADMIN
# ==============================

if admin:
    with st.expander("⚙️ Options générales (admin)", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            r = st.number_input("Jours ROUGE", min_value=0, max_value=90, value=opts["rouge"], key="opt_red")
            o = st.number_input("Jours ORANGE", min_value=0, max_value=90, value=opts["orange"], key="opt_orange")
            j = st.number_input("Jours JAUNE", min_value=0, max_value=90, value=opts["jaune"], key="opt_yellow")
        with col2:
            show_qr = st.checkbox("Inclure QR dans le PDF", value=opts["show_qr"], key="opt_show_qr")

        if st.button("💾 Sauvegarder options", key="save_options"):
            opts.update({"rouge": r, "orange": o, "jaune": j, "show_qr": show_qr})
            sauvegarder_options(opts)
            st.success("Options enregistrées ✔")

st.divider()

# ==============================
# IMPORT EXCEL (MONTAGE)
# ==============================

if admin:
    with st.expander("📥 Import Excel (Montage)", expanded=False):
        excel_file = st.file_uploader("Importer un Excel (.xlsx)", type=["xlsx"], key="excel_upload")
        if excel_file:
            df_imp = importer_excel(excel_file)
            if not df_imp.empty:
                st.success("Lignes 'Montage' détectées ✔")
                st.dataframe(df_imp, use_container_width=True)

                choix = st.selectbox(
                    "Sélectionner une ligne à ajouter",
                    df_imp.index,
                    format_func=lambda i: f"{df_imp.loc[i, 'nom']} — {df_imp.loc[i, 'ref']}",
                    key="excel_line_choice"
                )

                if st.button("➕ Ajouter ce chantier", key="add_excel_row"):
                    ligne = df_imp.loc[choix]
                    new_row = {
                        "nom": ligne["nom"],
                        "ref": ligne["ref"],
                        "date": ligne["date"],
                        "commentaire": "",
                        "statut": "Prévu",
                        "priorite": "auto",
                    }
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    sauvegarder_chantiers(df)
                    st.success("Chantier ajouté ✔")
                    st.rerun()

st.divider()

# ==============================
# AJOUT MANUEL D’UN CHANTIER
# ==============================

if admin:
    with st.expander("✏️ Ajouter manuellement un chantier", expanded=False):
        colA, colB = st.columns(2)
        with colA:
            nom = st.text_input("Nom de l'affaire", key="add_nom_input")
            d_montage = st.date_input("Date de montage", value=date.today(), key="add_date_input")
        with colB:
            ref = st.text_input("Code de l'affaire", key="add_ref_input")
            statut = st.selectbox(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                key="add_statut_input"
            )

        commentaire = st.text_area("Commentaire", key="add_comment_input")
        prio = st.selectbox(
            "Priorité manuelle",
            ["Automatique", "Rouge", "Orange", "Jaune", "Gris"],
            key="add_prio_input"
        )

        if st.button("➕ Ajouter ce chantier", key="add_manual_button"):
            if nom and ref:
                new_row = {
                    "nom": nom,
                    "ref": ref,
                    "date": str(d_montage),
                    "commentaire": commentaire,
                    "statut": statut,
                    "priorite": prio.lower() if prio != "Automatique" else "auto",
                }
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                sauvegarder_chantiers(df)
                st.success("Ajouté ✔")
                st.rerun()
            else:
                st.error("Nom et Code sont obligatoires.")


# ==============================
# BOUTON FILTRES
# ==============================

if st.button("🔎 Filtres", key="toggle_filters_button"):
    st.session_state["show_filters"] = not st.session_state["show_filters"]

if st.session_state["show_filters"]:
    with st.expander("🎚️ Panneau de filtres", expanded=True):

        colF1, colF2 = st.columns(2)

        with colF1:
            search = st.text_input("Recherche (nom, code…)", key="filter_search")

            statut_filter = st.multiselect(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                key="filter_statut"
            )

        with colF2:
            date_min = st.date_input("Depuis la date :", value=None, key="filter_date_min")
            date_max = st.date_input("Jusqu'à la date :", value=None, key="filter_date_max")

        # Appliquer filtres
        df_filtered = df.copy()

        if search:
            s = search.lower()
            df_filtered = df_filtered[
                df_filtered["nom"].str.lower().str.contains(s, na=False)
                | df_filtered["ref"].str.lower().str.contains(s, na=False)
            ]

        if statut_filter:
            df_filtered = df_filtered[df_filtered["statut"].isin(statut_filter)]

        if date_min:
            df_filtered = df_filtered[df_filtered["date"].dt.date >= date_min]
        if date_max:
            df_filtered = df_filtered[df_filtered["date"].dt.date <= date_max]

else:
    df_filtered = df.copy()


# ==============================
# TABLEAU PRINCIPAL
# ==============================

st.subheader("📌 Liste des chantiers")

if df_filtered.empty:
    st.info("Aucun chantier trouvé.")
else:

    # Ajout des pastilles dans le DF affiché
    df_show = df_filtered.copy()

    df_show["Urgence"] = df_show.apply(lambda row: urgence_emoji(row, opts), axis=1)
    df_show["État"] = df_show["statut"].apply(statut_emoji)

    df_show["Date"] = df_show["date"].dt.strftime("%d/%m/%Y")

    df_show = df_show.rename(columns={
        "nom": "Nom de l'affaire",
        "ref": "Code de l'affaire",
        "statut": "Statut",
        "commentaire": "Commentaire",
    })

    df_show = df_show[["Urgence", "Nom de l'affaire", "Code de l'affaire",
                       "Date", "Statut", "État", "Commentaire"]]

    st.dataframe(df_show, use_container_width=True)

st.markdown("---")
if st.button("📊 Afficher le tableau Gantt (plein écran)", key="btn_show_gantt"):
    st.session_state["gantt_fullscreen"] = True
    st.experimental_rerun()


# ==============================
# MODIFIER / SUPPRIMER (ADMIN)
# ==============================

if admin:
    st.divider()
    st.subheader("🛠️ Modifier ou supprimer un chantier")

    if df.empty:
        st.info("Aucun chantier à modifier.")
    else:
        choix_modif = st.selectbox(
            "Sélectionner une ligne",
            df.index,
            format_func=lambda i: f"{df.loc[i,'nom']} — {df.loc[i,'ref']}",
            key="edit_selectbox"
        )

        row = df.loc[choix_modif]

        col1, col2 = st.columns(2)

        with col1:
            new_nom = st.text_input("Nom de l'affaire", row["nom"], key=f"edit_nom_{choix_modif}")
            new_ref = st.text_input("Code de l'affaire", row["ref"], key=f"edit_ref_{choix_modif}")
            new_date = st.date_input(
                "Date de montage",
                value=row["date"].date() if not pd.isna(row["date"]) else date.today(),
                key=f"edit_date_{choix_modif}"
            )

        with col2:
            new_comment = st.text_area("Commentaire", row["commentaire"], key=f"edit_com_{choix_modif}")

            new_statut = st.selectbox(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                index=["Prévu","En cours","En attente","Terminé"].index(row["statut"]),
                key=f"edit_statut_{choix_modif}"
            )

            new_prio = st.selectbox(
                "Priorité manuelle",
                ["Automatique", "Rouge", "Orange", "Jaune", "Gris"],
                index=["auto","rouge","orange","jaune","gris"].index(row["priorite"]),
                key=f"edit_prio_{choix_modif}"
            )

        colA, colB = st.columns(2)

        with colA:
            if st.button("💾 Enregistrer modifications", key=f"save_edit_{choix_modif}"):
                df.at[choix_modif, "nom"] = new_nom
                df.at[choix_modif, "ref"] = new_ref
                df.at[choix_modif, "date"] = str(new_date)
                df.at[choix_modif, "commentaire"] = new_comment
                df.at[choix_modif, "statut"] = new_statut
                df.at[choix_modif, "priorite"] = new_prio.lower() if new_prio != "Automatique" else "auto"

                sauvegarder_chantiers(df)
                st.success("Modifications enregistrées ✔")
                st.rerun()

        with colB:
            if st.button("🗑️ Supprimer ce chantier", key=f"delete_{choix_modif}"):
                df = df.drop(choix_modif).reset_index(drop=True)
                sauvegarder_chantiers(df)
                st.success("Chantier supprimé ✔")
                st.rerun()


# ==============================
# VUE GANTT : POSITION DES CHANTIERS DANS LE TEMPS (60 JOURS)
# ==============================

st.divider()
st.subheader("📈 Vue temporelle (Gantt) des 60 prochains jours")

if df.empty:
    st.info("Aucun chantier à afficher dans le graphique.")
else:
    df_graph = df.dropna(subset=["date"]).copy()

    if df_graph.empty:
        st.info("Aucune date valide pour afficher le graphique.")
    else:
        import altair as alt

        # --------------- CONFIG TEMPS --------------- #

        today = datetime.today().date()
        horizon = today + pd.Timedelta(days=60)

        # On garde uniquement le domaine visuel
        df_graph = df_graph[(df_graph["date"].dt.date >= today) &
                            (df_graph["date"].dt.date <= horizon)]

        if df_graph.empty:
            st.warning("Aucun chantier dans les 60 prochains jours.")
        else:
            # Préparation du DataFrame
            df_graph["statut"] = df_graph["statut"].fillna("Prévu")
            df_graph["date_only"] = df_graph["date"].dt.date

            # Ajout de l’urgence (couleur pastille)
            df_graph["urgence_color"] = df_graph.apply(lambda r: get_urgence_color(r, opts), axis=1)

            # --------------- BARRE DU GANTT --------------- #

            # Chaque chantier sera dessiné comme un segment bref (un jour)
            gantt = alt.Chart(df_graph).mark_bar(height=12).encode(
                x=alt.X("date_only:T",
                        title="Date",
                        scale=alt.Scale(domain=[pd.to_datetime(today), pd.to_datetime(horizon)])),
                y=alt.Y("nom:N",
                        title="Chantiers",
                        sort="-x"),
                color=alt.Color(
                    "statut:N",
                    legend=alt.Legend(title="Statut"),
                    scale=alt.Scale(
                        domain=["Prévu", "En cours", "En attente", "Terminé"],
                        range=["#F5EEDC", "#D9C8FF", "#FFD6E7", "#D9F8C4"],
                    )
                ),
                tooltip=[
                    alt.Tooltip("nom:N", title="Nom de l'affaire"),
                    alt.Tooltip("ref:N", title="Code de l'affaire"),
                    alt.Tooltip("date_only:T", title="Date"),
                    alt.Tooltip("statut:N", title="Statut"),
                ]
            )

            # --------------- PASTILLE D’URGENCE --------------- #

            urgence_points = alt.Chart(df_graph).mark_circle(size=200).encode(
                x="date_only:T",
                y="nom:N",
                color=alt.Color("urgence_color:N", scale=None, legend=None),
            )

            # --------------- LIGNE AUJOURD’HUI --------------- #

            today_df = pd.DataFrame({"today": [pd.to_datetime(today)]})

            today_line = alt.Chart(today_df).mark_rule(
                color="red",
                strokeWidth=2
            ).encode(
                x="today:T"
            )

            # --------------- LIGNES DE SEMAINE (CHAQUE LUNDI) --------------- #

            mondays = pd.date_range(start=today, end=horizon, freq="W-MON")
            df_mondays = pd.DataFrame({"monday": mondays})

            week_lines = alt.Chart(df_mondays).mark_rule(
                color="#CCCCCC",
                strokeDash=[4, 4]
            ).encode(
                x="monday:T"
            )

            # --------------- TITRE & MISE EN PAGE --------------- #

            chart = (
                gantt
                + urgence_points
                + today_line
                + week_lines
            ).properties(
                height=450,
                width="container",
                title="Position des chantiers sur 60 jours"
            ).interactive()

            st.altair_chart(chart, use_container_width=True)


# ==============================
# EXPORT
# ==============================

st.divider()
st.subheader("📤 Exporter les données")

colE1, colE2, colE3 = st.columns(3)

with colE1:
    if st.button("📄 Export PDF", key="export_pdf_button"):
        qr_url = st.session_state.get("qr_url", "")

        pdf = build_pdf(df_filtered, opts, qr_url=qr_url if opts["show_qr"] else None)
        st.download_button(
            label="📥 Télécharger PDF",
            data=pdf,
            file_name=f"Gestion_Projet_Priorites_{datetime.today().date()}.pdf",
            mime="application/pdf",
            key="download_pdf"
        )

with colE2:
    if st.button("📊 Export Excel", key="export_excel_button"):
        excel = build_excel(df_filtered)
        st.download_button(
            label="📥 Télécharger Excel",
            data=excel,
            file_name=f"Gestion_Projet_Priorites_{datetime.today().date()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_excel"
        )

with colE3:
    if admin:
        st.subheader("QR-Code")
        app_url = st.text_input("URL de l'application", value=st.session_state.get("qr_url",""), key="qr_url_input")

        if st.button("🔄 Générer QR", key="qr_generate_button"):
            st.session_state["qr_url"] = app_url
            st.success("QR mis à jour ✔")

        if st.session_state["qr_url"]:
            img = build_qr_image(st.session_state["qr_url"])
            st.image(img, width=160, caption="QR-Code d'accès")
