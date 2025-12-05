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

# ================== CONFIGURATION ==================

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


# ================== UTILITAIRES ==================

def ensure_columns(df: pd.DataFrame, columns):
    """Ajoute les colonnes manquantes."""
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    return df


def charger_chantiers():
    """Charge les données JSON + assure cohérence des colonnes."""
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
    """Sauvegarde les données dans JSON."""
    df2 = df.copy()
    if not df2.empty:
        df2["date"] = df2["date"].astype(str)

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(df2.to_dict(orient="records"), f, ensure_ascii=False, indent=4)


def charger_options():
    """Charge les options admin ou crée celles par défaut."""
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
    """Sauvegarde les options admin."""
    with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(opts, f, indent=4)


# ================== COULEURS ==================

def get_statut_color(statut):
    """Retourne la couleur associée au statut."""
    statut = str(statut).lower()
    mapping = {
        "prévu": "#F5EEDC",
        "en cours": "#D9C8FF",
        "en attente": "#FFD6E7",
        "terminé": "#D9F8C4",
    }
    return mapping.get(statut, "")


def get_urgence_color(row, opts):
    """Retourne la couleur d'urgence rouge/orange/jaune."""
    d = row.get("date")
    if pd.isna(d):
        return ""

    delta = (d.date() - datetime.today().date()).days

    if delta <= opts["rouge"]:
        return "#FF4B4B"
    if delta <= opts["orange"]:
        return "#FFA500"
    if delta <= opts["jaune"]:
        return "#FFD966"

    return ""


# ================== STYLE DU TABLEAU STREAMLIT ==================

def make_style_table(df_display, df_filtered, opts):
    """
    Applique les couleurs :
    - Urgence sur Nom / Code / Date
    - Statut uniquement sur Statut
    - Autres cellules neutres
    """
    df_f = df_filtered.loc[df_display.index]  # réalignement sur les mêmes index

    styles = pd.DataFrame("", index=df_display.index, columns=df_display.columns)

    for idx in df_display.index:
        r = df_f.loc[idx]

        urg = get_urgence_color(r, opts)
        stat_col = get_statut_color(r["statut"])

        # Urgence sur colonnes principales
        if urg:
            for col in ["Nom de l'affaire", "Code de l'affaire", "Date"]:
                if col in styles.columns:
                    styles.loc[idx, col] = f"background-color: {urg}"

        # Couleur statut
        if stat_col and "Statut" in styles.columns:
            styles.loc[idx, "Statut"] = f"background-color: {stat_col}"

    return styles


# ================== EXPORT EXCEL ==================

def build_excel(df):
    buf = BytesIO()
    df2 = df.copy()
    if not df2.empty and pd.api.types.is_datetime64_any_dtype(df2["date"]):
        df2["date"] = df2["date"].dt.strftime("%d/%m/%Y")
    df2.rename(
        columns={
            "nom": "Nom de l'affaire",
            "ref": "Code de l'affaire",
            "date": "Date",
            "statut": "Statut",
            "priorite": "Priorité",
            "commentaire": "Commentaire",
        },
        inplace=True,
    )
    df2.to_excel(buf, index=False)
    buf.seek(0)
    return buf


# ================== EXPORT PDF PREMIUM ==================

def build_pdf(df, opts, qr_url=None):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    # ----- Bandeau titre -----
    c.setFillColor(colors.HexColor("#2F3C7E"))
    c.rect(0, height - 70, width, 70, fill=1, stroke=0)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(width / 2, height - 40, "Gestion Projet Priorités")

    c.setFont("Helvetica", 10)
    c.drawString(40, height - 58, f"Export du {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    # ----- Légende -----
    y = height - 110
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(colors.black)
    c.drawString(40, y, "Légende des couleurs :")

    y -= 20
    legend_items = [
        ("#FF4B4B", f"Urgence forte (≤ {opts['rouge']} jours)"),
        ("#FFA500", f"Urgence moyenne (≤ {opts['orange']} jours)"),
        ("#FFD966", f"Approche (≤ {opts['jaune']} jours)"),
        ("#F5EEDC", "Statut : Prévu"),
        ("#D9C8FF", "Statut : En cours"),
        ("#FFD6E7", "Statut : En attente"),
        ("#D9F8C4", "Statut : Terminé"),
    ]

    for color_hex, label in legend_items:
        c.setFillColor(colors.HexColor(color_hex))
        c.rect(40, y - 10, 8, 8, fill=1)
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 10)
        c.drawString(55, y - 9, label)
        y -= 14

    y -= 10

    # ----- Tableau : colonnes -----
    headers = ["Nom de l'affaire", "Code", "Date", "Statut"]
    col_widths = [220, 70, 60, 80]
    left = 40

    # En-têtes
    c.setFillColor(colors.HexColor("#EFEFEF"))
    c.rect(left, y - 18, sum(col_widths), 18, fill=1, stroke=0)

    c.setFont("Helvetica-Bold", 10)
    c.setFillColor(colors.black)
    x = left + 4
    for i, h in enumerate(headers):
        c.drawString(x, y - 13, h)
        x += col_widths[i]

    y -= 24
    c.setFont("Helvetica", 9)

    # Tri par date pour le PDF
    df_sorted = df.sort_values("date")

    # ----- Lignes -----
    for _, row in df_sorted.iterrows():
        # Saut de page si nécessaire
        if y < 80:
            c.showPage()
            width, height = A4
            y = height - 60

            c.setFillColor(colors.HexColor("#EFEFEF"))
            c.rect(left, y - 18, sum(col_widths), 18, fill=1)
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 10)
            x = left + 4
            for i, h in enumerate(headers):
                c.drawString(x, y - 13, h)
                x += col_widths[i]

            y -= 24
            c.setFont("Helvetica", 9)

        # Bande verticale d'urgence
        urg = get_urgence_color(row, opts)
        if urg:
            c.setFillColor(colors.HexColor(urg))
            c.rect(left - 8, y - 14, 6, 14, fill=1, stroke=0)

        # Données
        nom = str(row["nom"]) if not pd.isna(row["nom"]) else ""
        code = str(row["ref"]) if not pd.isna(row["ref"]) else ""
        if not pd.isna(row["date"]):
            date_str = row["date"].strftime("%d/%m/%Y")
        else:
            date_str = "-"
        statut = str(row["statut"]) if not pd.isna(row["statut"]) else ""

        x = left + 4

        # Nom
        c.setFillColor(colors.black)
        c.drawString(x, y, nom[:40])
        x += col_widths[0]

        # Code
        c.drawString(x, y, code[:15])
        x += col_widths[1]

        # Date
        c.drawString(x, y, date_str)
        x += col_widths[2]

        # Statut avec fond coloré
        statut_bg = get_statut_color(statut)
        if statut_bg:
            c.setFillColor(colors.HexColor(statut_bg))
            c.rect(x, y - 12, col_widths[3] - 4, 14, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.drawString(x + 4, y, statut[:15])

        y -= 18

    # ----- QR CODE en bas -----
    if qr_url:
        qr = qrcode.QRCode(box_size=3, border=1)
        qr.add_data(qr_url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")

        qr_buf = BytesIO()
        img.save(qr_buf, format="PNG")
        qr_buf.seek(0)

        qr_img = ImageReader(qr_buf)
        c.drawImage(qr_img, width - 45 * mm, 15 * mm, width=30 * mm, preserveAspectRatio=True)

        c.setFont("Helvetica", 9)
        c.setFillColor(colors.black)
        c.drawRightString(width - 10 * mm, 12 * mm, "Accès application")

    c.save()
    buf.seek(0)
    return buf


# ================== QR CODE (IMAGE APP) ==================

def build_qr_image(url):
    qr = qrcode.QRCode(box_size=8, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image()
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


# ================== IMPORT EXCEL ULTRA-ROBUSTE ==================

def importer_excel(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip().str.lower()

    possible_nom_cols = [
        "nom de l’affaire",
        "nom de l'affaire",
        "nom de laffaire",
        "nom affaire",
        "nom d’affaire",
        "nom d'affaire",
    ]

    possible_code_cols = [
        "code de l’affaire",
        "code de l'affaire",
        "code de laffaire",
        "code affaire",
        "codedaffaire",
    ]

    possible_debut_cols = ["début", "debut"]

    def find_col(possible):
        for col in df.columns:
            if col in possible:
                return col
        return None

    col_comp = "compétence"
    col_nom = find_col(possible_nom_cols)
    col_code = find_col(possible_code_cols)
    col_debut = find_col(possible_debut_cols)

    missing = []
    if col_nom is None:
        missing.append("Nom de l’affaire")
    if col_code is None:
        missing.append("Code de l’affaire")
    if col_debut is None:
        missing.append("Début")
    if col_comp not in df.columns:
        missing.append("Compétence")

    if missing:
        st.error("⚠️ Colonnes manquantes : " + ", ".join(missing))
        return pd.DataFrame()

    df_filtre = df[df[col_comp].astype(str).str.lower() == "montage"]

    if df_filtre.empty:
        st.warning("Aucune ligne 'Montage' trouvée.")
        return pd.DataFrame()

    df_res = df_filtre[[col_nom, col_code, col_debut]].copy()
    df_res.columns = ["nom", "ref", "date"]
    df_res["date"] = pd.to_datetime(df_res["date"], errors="coerce")

    return df_res


# ================== INTERFACE STREAMLIT ==================

st.set_page_config(page_title="Liste Priorité Chantier", layout="wide")
st.title("📋 Gestion des priorités chantier")

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

# ================== OPTIONS ADMIN ==================
if admin:
    with st.expander("⚙️ Options générales (admin)", expanded=False):
        col_opt1, col_opt2 = st.columns(2)
        with col_opt1:
            rouge = st.number_input("Jours max ROUGE", 0, 90, opts["rouge"], key="rouge_opt")
            orange = st.number_input("Jours max ORANGE", 0, 90, opts["orange"], key="orange_opt")
            jaune = st.number_input("Jours max JAUNE", 0, 90, opts["jaune"], key="jaune_opt")
        with col_opt2:
            show_qr = st.checkbox("Afficher QR Code atelier", value=opts["show_qr"], key="qr_opt")

        if st.button("Enregistrer options", key="save_options"):
            opts.update({"rouge": rouge, "orange": orange, "jaune": jaune, "show_qr": show_qr})
            sauvegarder_options(opts)
            st.success("Options enregistrées ✔")

st.divider()

# ================== IMPORT EXCEL ==================
if admin:
    with st.expander("📥 Importer depuis un export Excel", expanded=False):
        st.caption("Seules les lignes où Compétence = 'Montage' seront prises en compte.")
        fichier_excel = st.file_uploader("Importer un export Excel (.xlsx)", type=["xlsx"], key="import_excel")

        df_excel = pd.DataFrame()

        if fichier_excel:
            df_excel = importer_excel(fichier_excel)

            if not df_excel.empty:
                st.success("Fichier importé ✔ — lignes 'Montage' détectées.")
                st.dataframe(df_excel, use_container_width=True)

                choix = st.selectbox(
                    "Sélectionner une ligne à ajouter",
                    df_excel.index,
                    key="choix_excel",
                    format_func=lambda i: f"{df_excel.loc[i, 'nom']} — {df_excel.loc[i, 'ref']}"
                )

                if st.button("➕ Ajouter ce chantier", key="add_from_excel_btn"):
                    ligne = df_excel.loc[choix]

                    new = {
                        "nom": ligne["nom"],
                        "ref": ligne["ref"],
                        "date": ligne["date"],
                        "commentaire": "",
                        "statut": "Prévu",
                        "priorite": "auto",
                    }

                    df = pd.concat([df, pd.DataFrame([new])], ignore_index=True)
                    sauvegarder_chantiers(df)
                    st.success("Chantier ajouté ✔")
                    st.rerun()

st.divider()

# ================== AJOUT MANUEL ==================
if admin:
    with st.expander("✏️ Ajouter manuellement un chantier", expanded=False):
        col1, col2 = st.columns(2)
        with col1:
            nom = st.text_input("Nom de l'affaire", key="nom_manual")
            date_montage = st.date_input("Date de montage", value=date.today(), key="date_manual")
        with col2:
            ref = st.text_input("Code de l'affaire", key="ref_manual")
            statut = st.selectbox("Statut", ["Prévu", "En cours", "En attente", "Terminé"], key="statut_manual")

        commentaire = st.text_area("Commentaire", key="comment_manual")

        priorite_man = st.selectbox(
            "Priorité manuelle",
            ["Automatique", "Rouge", "Orange", "Jaune", "Gris"],
            key="prio_manual"
        )

        if st.button("Ajouter ce chantier", key="add_manual"):
            if nom and ref:
                new = {
                    "nom": nom,
                    "ref": ref,
                    "date": str(date_montage),
                    "commentaire": commentaire,
                    "statut": statut,
                    "priorite": priorite_man.lower() if priorite_man != "Automatique" else "auto",
                }
                df = pd.concat([df, pd.DataFrame([new])], ignore_index=True)
                sauvegarder_chantiers(df)
                st.success("Chantier ajouté ✔")
                st.rerun()
            else:
                st.error("Entrez au moins un nom et un code.")

st.divider()

# ================== LISTE + RECHERCHE + FILTRES ==================

st.subheader("📌 Liste des chantiers")

if df.empty:
    st.info("Aucun chantier enregistré.")
else:
    df_sorted = df.sort_values("date")

    # Bouton Filtres
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("🔎 Filtres", key="toggle_filters"):
            st.session_state["show_filters"] = not st.session_state["show_filters"]

    statuts_sel = []
    prio_sel_internal = []
    start_date = end_date = None
    search = ""

    if st.session_state["show_filters"]:
        st.markdown("#### 🎛 Panneau de filtres")
        col_filters, col_table = st.columns([1, 3])
    else:
        col_filters, col_table = st.columns([0.0001, 1])

    with col_filters:
        if st.session_state["show_filters"]:
            st.markdown("##### Filtres")

            search = st.text_input(
                "Recherche texte",
                placeholder="Nom, code, statut, commentaire...",
                key="search_global"
            )

            statuts_uniques = sorted(df_sorted["statut"].dropna().unique().tolist())
            statuts_sel = st.multiselect(
                "Statut",
                statuts_uniques,
                default=statuts_uniques,
                key="filter_statut"
            )

            prio_map_display = {
                "auto": "Automatique",
                "rouge": "Rouge",
                "orange": "Orange",
                "jaune": "Jaune",
                "gris": "Gris",
            }
            prio_uniques = sorted(df_sorted["priorite"].dropna().unique().tolist())
            prio_display_options = [prio_map_display.get(p, p) for p in prio_uniques]

            prio_sel_display = st.multiselect(
                "Priorité",
                prio_display_options,
                default=prio_display_options,
                key="filter_prio"
            )

            inv_prio_map = {v: k for k, v in prio_map_display.items()}
            prio_sel_internal = [inv_prio_map.get(p, p) for p in prio_sel_display]

            dates_valides = df_sorted["date"].dropna()
            if not dates_valides.empty:
                min_date = dates_valides.min().date()
                max_date = dates_valides.max().date()
                start_date = st.date_input(
                    "Date min",
                    value=min_date,
                    key="filter_date_min"
                )
                end_date = st.date_input(
                    "Date max",
                    value=max_date,
                    key="filter_date_max"
                )
            else:
                st.caption("Aucune date valide pour filtrer.")

            if st.button("Réinitialiser les filtres", key="reset_filters"):
                st.session_state["search_global"] = ""
                st.session_state["filter_statut"] = statuts_uniques
                st.session_state["filter_prio"] = prio_display_options
                if "filter_date_min" in st.session_state:
                    del st.session_state["filter_date_min"]
                if "filter_date_max" in st.session_state:
                    del st.session_state["filter_date_max"]
                st.rerun()

    with col_table:
        df_filtered = df_sorted.copy()

        if statuts_sel:
            df_filtered = df_filtered[df_filtered["statut"].isin(statuts_sel)]

        if prio_sel_internal:
            df_filtered = df_filtered[df_filtered["priorite"].isin(prio_sel_internal)]

        if start_date and end_date:
            mask_date = df_filtered["date"].notna()
            mask_date &= df_filtered["date"].dt.date.between(start_date, end_date)
            df_filtered = df_filtered[mask_date]

        if search:
            mask = pd.Series(False, index=df_filtered.index)
            for col in ["nom", "ref", "commentaire", "statut"]:
                mask |= df_filtered[col].astype(str).str.contains(search, case=False, na=False)
            df_filtered = df_filtered[mask]

        df_display = df_filtered.copy()
        if not df_display.empty and pd.api.types.is_datetime64_any_dtype(df_display["date"]):
            df_display["date"] = df_display["date"].dt.strftime("%d/%m/%Y")

        df_display = df_display.rename(
            columns={
                "nom": "Nom de l'affaire",
                "ref": "Code de l'affaire",
                "date": "Date",
                "statut": "Statut",
                "priorite": "Priorité",
                "commentaire": "Commentaire",
            }
        )

        styles = make_style_table(df_display, df_filtered, opts)
        styled = df_display.style.apply(lambda _: styles, axis=None)

        st.caption(f"{len(df_filtered)} chantier(s) affiché(s)")
        st.dataframe(styled, use_container_width=True)

    # ================== MODIFIER / SUPPRIMER (ADMIN) ==================
    if admin and not df_sorted.empty:
        st.divider()
        with st.expander("🛠 Modifier ou supprimer un chantier", expanded=False):
            idx = st.selectbox(
                "Sélectionner un chantier",
                df_sorted.index,
                key="edit_select",
                format_func=lambda i: f"{df_sorted.loc[i, 'nom']} — {df_sorted.loc[i, 'ref']}"
            )

            col1, col2 = st.columns(2)
            with col1:
                new_nom = st.text_input("Nom de l'affaire", df.loc[idx, "nom"], key=f"edit_nom_{idx}")
                current_date = df.loc[idx, "date"]
                new_date = st.date_input(
                    "Date de montage",
                    value=current_date.date() if not pd.isna(current_date) else date.today(),
                    key=f"edit_date_{idx}"
                )
            with col2:
                new_ref = st.text_input("Code de l'affaire", df.loc[idx, "ref"], key=f"edit_ref_{idx}")
                new_statut = st.selectbox(
                    "Statut",
                    ["Prévu", "En cours", "En attente", "Terminé"],
                    index=["Prévu", "En cours", "En attente", "Terminé"].index(df.loc[idx, "statut"]),
                    key=f"edit_statut_{idx}"
                )

            new_commentaire = st.text_area(
                "Commentaire",
                df.loc[idx, "commentaire"],
                key=f"edit_comment_{idx}"
            )

            prio_list = ["Automatique", "Rouge", "Orange", "Jaune", "Gris"]
            current_prio = df.loc[idx, "priorite"]
            display_prio = current_prio.capitalize() if current_prio != "auto" else "Automatique"

            new_prio = st.selectbox(
                "Priorité manuelle",
                prio_list,
                index=prio_list.index(display_prio),
                key=f"edit_prio_{idx}"
            )

            colA, colB = st.columns(2)

            with colA:
                if st.button("💾 Enregistrer modifications", key=f"save_edit_{idx}"):
                    df.loc[idx, "nom"] = new_nom
                    df.loc[idx, "ref"] = new_ref
                    df.loc[idx, "date"] = str(new_date)
                    df.loc[idx, "statut"] = new_statut
                    df.loc[idx, "commentaire"] = new_commentaire
                    df.loc[idx, "priorite"] = new_prio.lower() if new_prio != "Automatique" else "auto"

                    sauvegarder_chantiers(df)
                    st.success("Modifié ✔")
                    st.rerun()

            with colB:
                if st.button("🗑 Supprimer ce chantier", key=f"delete_{idx}"):
                    df = df.drop(idx).reset_index(drop=True)
                    sauvegarder_chantiers(df)
                    st.success("Chantier supprimé ✔")
                    st.rerun()

st.divider()

# ================== EXPORTS ==================

st.subheader("📤 Export")

if df.empty:
    st.info("Ajoutez un chantier pour activer l’export.")
else:
    col1, col2 = st.columns(2)
    with col1:
        qr_url_for_pdf = st.session_state.get("qr_url") or None
        st.download_button(
            "📄 Export PDF",
            build_pdf(df, opts, qr_url=qr_url_for_pdf),
            file_name=f"Gestion_Projet_Priorites_{datetime.now().strftime('%Y-%m-%d')}.pdf",
            mime="application/pdf"
        )
    with col2:
        st.download_button("📊 Export Excel", build_excel(df), "chantiers.xlsx")

st.divider()

# ================== QR CODE ==================

if opts["show_qr"]:
    st.subheader("📱 QR Code atelier")

    qr_url = st.text_input(
        "URL de l'application",
        key="qr_url",
        placeholder="https://..."
    )

    if qr_url:
        st.image(build_qr_image(qr_url), width=200)
