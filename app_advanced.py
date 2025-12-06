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
    """Couleur hex pour le PDF (bande urgence)."""
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
    """Couleur hex pour le PDF (bande statut)."""
    statut = str(statut).lower().strip()
    mapping = {
        "prévu": "#F5EEDC",
        "en cours": "#D9C8FF",
        "en attente": "#FFD6E7",
        "terminé": "#D9F8C4",
    }
    return mapping.get(statut, "#DDDDDD")


def urgence_emoji(row, opts):
    """Pastille d’urgence (UI)."""
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
    """Pastille de statut (UI)."""
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

    # Normalisation
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
# EXPORT PDF
# ==============================

def build_pdf(df, opts, qr_url=None):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    # Titre
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
    y -= 18
    c.setFont("Helvetica", 10)

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
    for color_hex, label in legend:
        c.setFillColor(colors.HexColor(color_hex))
        c.rect(40, y - 8, 8, 8, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.drawString(56, y - 2, label)
        y -= 14

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

        # Urgence (petit carré)
        c.setFillColor(colors.HexColor(urg_color))
        c.rect(x + 8, y - row_h + 4, 8, 8, fill=1, stroke=0)
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

        # Statut (carré)
        c.setFillColor(colors.HexColor(stat_color))
        c.rect(x + 10, y - row_h + 4, 8, 8, fill=1, stroke=0)

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
        c.drawImage(ImageReader(qr_buf), width - 45 * mm, 15 * mm, width=30 * mm, preserveAspectRatio=True)
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 9)
        c.drawRightString(width - 10 * mm, 12 * mm, "Accès application")

    c.save()
    buf.seek(0)
    return buf


# ==============================
# QR CODE POUR L’APP
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

# ==============================
# OPTIONS ADMIN
# ==============================

if admin:
    with st.expander("⚙️ Options générales (admin)", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            r = st.number_input("Jours ROUGE", 0, 90, opts["rouge"])
            o = st.number_input("Jours ORANGE", 0, 90, opts["orange"])
            j = st.number_input("Jours JAUNE", 0, 90, opts["jaune"])
        with col2:
            show_qr = st.checkbox("Inclure QR dans le PDF", value=opts["show_qr"])

        if st.button("💾 Sauvegarder options"):
            opts.update({"rouge": r, "orange": o, "jaune": j, "show_qr": show_qr})
            sauvegarder_options(opts)
            st.success("Options enregistrées ✔")

st.divider()

# ==============================
# IMPORT EXCEL
# ==============================

if admin:
    with st.expander("📥 Import Excel (Montage)", expanded=False):
        excel_file = st.file_uploader("Importer un Excel (.xlsx)", type=["xlsx"])
        if excel_file:
            df_imp = importer_excel(excel_file)
            if not df_imp.empty:
                st.success("Lignes 'Montage' détectées ✔")
                st.dataframe(df_imp, use_container_width=True)
                choix = st.selectbox(
                    "Sélectionner une ligne à ajouter",
                    df_imp.index,
                    format_func=lambda i: f"{df_imp.loc[i, 'nom']} — {df_imp.loc[i, 'ref']}",
                )
                if st.button("➕ Ajouter ce chantier"):
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
# AJOUT MANUEL
# ==============================

if admin:
    with st.expander("✏️ Ajouter manuellement un chantier", expanded=False):
        colA, colB = st.columns(2)
        with colA:
            nom = st.text_input("Nom de l'affaire")
            d_montage = st.date_input("Date de montage", value=date.today())
        with colB:
            ref = st.text_input("Code de l'affaire")
            statut = st.selectbox("Statut", ["Prévu", "En cours", "En attente", "Terminé"])

        commentaire = st.text_area("Commentaire")
        prio = st.selectbox("Priorité manuelle", ["Automatique", "Rouge", "Orange", "Jaune", "Gris"])

        if st.button("➕ Ajouter ce chantier"):
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

st.divider()

# ==============================
# LISTE + FILTRES + PASTILLES
# ==============================

st.subheader("📌 Liste des chantiers")

if df.empty:
    st.info("Aucun chantier.")
else:
    # bouton filtres
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("🔎 Filtres"):
            st.session_state["show_filters"] = not st.session_state["show_filters"]

    d1 = d2 = None
    search = ""
    selected_stats = []

    if st.session_state["show_filters"]:
        col_filters, col_tab = st.columns([1, 3])
    else:
        col_filters, col_tab = st.columns([0.0001, 1])

    # Filtres
    with col_filters:
        if st.session_state["show_filters"]:
            search = st.text_input("Recherche (nom, code, commentaire)", key="search_global")

            stat_uniques = sorted(df["statut"].dropna().unique().tolist())
            selected_stats = st.multiselect("Filtre statuts", stat_uniques, default=stat_uniques)

            dates_valides = df["date"].dropna()
            if not dates_valides.empty:
                dmin = dates_valides.min().date()
                dmax = dates_valides.max().date()
                d1 = st.date_input("Date min", value=dmin, key="filter_d1")
                d2 = st.date_input("Date max", value=dmax, key="filter_d2")

            if st.button("Réinitialiser les filtres"):
                st.session_state["search_global"] = ""
                st.rerun()

    # Tableau
    with col_tab:
        df2 = df.copy().sort_values("date")

        if search:
            txt = search.lower()
            mask = (
                df2["nom"].astype(str).str.lower().str.contains(txt) |
                df2["ref"].astype(str).str.lower().str.contains(txt) |
                df2["commentaire"].astype(str).str.lower().str.contains(txt)
            )
            df2 = df2[mask]

        if selected_stats:
            df2 = df2[df2["statut"].isin(selected_stats)]

        if d1 is not None and d2 is not None:
            mask_date = df2["date"].notna()
            mask_date &= df2["date"].dt.date.between(d1, d2)
            df2 = df2[mask_date]

        df_disp = df2.copy()

        # colonnes date formatées
        if not df_disp.empty and pd.api.types.is_datetime64_any_dtype(df_disp["date"]):
            df_disp["Date"] = df_disp["date"].dt.strftime("%d/%m/%Y")
        else:
            df_disp["Date"] = "-"

        # pastilles
        df_disp["Urgence"] = df2.apply(lambda r: urgence_emoji(r, opts), axis=1)
        df_disp["État"] = df2["statut"].apply(statut_emoji)

        df_disp.rename(columns={
            "nom": "Nom de l'affaire",
            "ref": "Code de l'affaire",
            "commentaire": "Commentaire",
            "statut": "Statut",
        }, inplace=True)

        # réorganisation colonnes
        cols_order = ["Urgence", "Nom de l'affaire", "Code de l'affaire", "Date", "Statut", "État", "Commentaire"]
        for c in cols_order:
            if c not in df_disp.columns:
                df_disp[c] = ""
        df_disp = df_disp[cols_order]

        st.caption(f"{len(df_disp)} chantier(s) affiché(s)")
        st.dataframe(
            df_disp,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Urgence": st.column_config.TextColumn(
                    "Urgence",
                    help="🔴 très urgent · 🟠 moyen · 🟡 approche · ⚪ pas urgent",
                ),
                "État": st.column_config.TextColumn(
                    "État",
                    help="⚪ prévu · 🟣 en cours · 🔵 en attente · 🟢 terminé",
                ),
            }
        )

# ==============================
# MODIFIER / SUPPRIMER
# ==============================

if admin and not df.empty:
    st.divider()
    with st.expander("🛠 Modifier ou supprimer un chantier", expanded=False):
        idx = st.selectbox(
            "Sélectionner un chantier",
            df.index,
            format_func=lambda i: f"{df.loc[i, 'nom']} — {df.loc[i, 'ref']}",
        )

        col1, col2 = st.columns(2)
        with col1:
            new_nom = st.text_input("Nom de l'affaire", df.loc[idx, "nom"])
            dval = df.loc[idx, "date"]
            new_date = st.date_input(
                "Date de montage",
                value=dval.date() if not pd.isna(dval) else date.today()
            )
        with col2:
            new_ref = st.text_input("Code de l'affaire", df.loc[idx, "ref"])
            new_statut = st.selectbox(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                index=["Prévu", "En cours", "En attente", "Terminé"].index(df.loc[idx, "statut"])
            )

        new_com = st.text_area("Commentaire", df.loc[idx, "commentaire"])
        prio_list = ["Automatique", "Rouge", "Orange", "Jaune", "Gris"]
        cur_prio = df.loc[idx, "priorite"]
        prio_display = "Automatique" if cur_prio == "auto" else cur_prio.capitalize()
        new_prio = st.selectbox("Priorité manuelle", prio_list, index=prio_list.index(prio_display))

        colA, colB = st.columns(2)
        with colA:
            if st.button("💾 Enregistrer les modifications"):
                df.loc[idx, "nom"] = new_nom
                df.loc[idx, "ref"] = new_ref
                df.loc[idx, "date"] = str(new_date)
                df.loc[idx, "statut"] = new_statut
                df.loc[idx, "commentaire"] = new_com
                df.loc[idx, "priorite"] = new_prio.lower() if new_prio != "Automatique" else "auto"
                sauvegarder_chantiers(df)
                st.success("Modifications enregistrées ✔")
                st.rerun()

        with colB:
            if st.button("🗑 Supprimer ce chantier"):
                df = df.drop(idx).reset_index(drop=True)
                sauvegarder_chantiers(df)
                st.success("Chantier supprimé ✔")
                st.rerun()

# ==============================
# EXPORTS
# ==============================

st.divider()
st.subheader("📤 Exports")

colP, colE = st.columns(2)
with colP:
    if df.empty:
        st.button("📄 Export PDF", disabled=True)
    else:
        qr_for_pdf = st.session_state.get("qr_url") if opts["show_qr"] else None
        st.download_button(
            "📄 Export PDF",
            build_pdf(df, opts, qr_url=qr_for_pdf),
            file_name=f"Gestion_Projet_Priorites_{datetime.now().strftime('%Y-%m-%d')}.pdf",
            mime="application/pdf"
        )

with colE:
    if df.empty:
        st.button("📊 Export Excel", disabled=True)
    else:
        st.download_button(
            "📊 Export Excel",
            build_excel(df),
            file_name="chantiers.xlsx"
        )

# ==============================
# QR CODE AFFICHAGE
# ==============================

if opts["show_qr"]:
    st.divider()
    st.subheader("📱 QR Code atelier")

    qr_url = st.text_input(
        "URL de l'application (pour le QR code)",
        value=st.session_state.get("qr_url", "")
    )
    st.session_state["qr_url"] = qr_url

    if qr_url:
        st.image(build_qr_image(qr_url), width=180)
        st.caption("Scannez ce QR code pour accéder à l'application")
