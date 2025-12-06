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
import base64

# ================== CONFIGURATION GÉNÉRALE ==================

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


# ================== COULEURS DES PASTILLES ==================

def urgence_emoji(row, opts):
    """Retourne l'emoji de priorité (pastille ronde)."""
    d = row.get("date")
    if pd.isna(d):
        return "⚪"  # pas urgent si pas de date

    delta = (d.date() - datetime.today().date()).days

    if delta <= opts["rouge"]:
        return "🔴"  # urgence forte
    if delta <= opts["orange"]:
        return "🟠"  # urgence moyenne
    if delta <= opts["jaune"]:
        return "🟡"  # approche

    return "⚪"  # neutre / pas urgent


def statut_emoji(statut):
    """Retourne l'emoji de statut (pastille ronde)."""
    statut = str(statut).lower()

    mapping = {
        "prévu": "⚪",
        "en cours": "🟣",
        "en attente": "🔵",
        "terminé": "🟢",
    }
    return mapping.get(statut, "⚪")


# ================== EXPORT EXCEL ==================

def build_excel(df):
    """Construit le fichier Excel exportable."""
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
    """Construit un PDF propre avec couleurs et légende."""
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    # ---------- Bandeau titre ----------
    c.setFillColor(colors.HexColor("#2F3C7E"))
    c.rect(0, height - 60, width, 60, fill=1, stroke=0)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 22)
    c.drawCentredString(width / 2, height - 35, "Gestion Projet Priorités")

    c.setFont("Helvetica", 10)
    c.drawString(40, height - 52, f"Export du {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    # ---------- Légende ----------
    left_margin = 40
    y = height - 90
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left_margin, y, "Légende :")

    y -= 18
    c.setFont("Helvetica", 10)

    legend_items = [
        ("#FF4B4B", f"Urgence forte (≤ {opts['rouge']} jours)"),
        ("#FFA500", f"Urgence moyenne (≤ {opts['orange']} jours)"),
        ("#FFD966", f"Approche (≤ {opts['jaune']} jours)"),
        ("#DDDDDD", "Pas urgent"),
        ("#F5EEDC", "Statut : Prévu"),
        ("#D9C8FF", "Statut : En cours"),
        ("#FFD6E7", "Statut : En attente"),
        ("#D9F8C4", "Statut : Terminé"),
    ]

    for color_hex, label in legend_items:
        # petit carré
        c.setFillColor(colors.HexColor(color_hex))
        c.rect(left_margin, y - 8, 8, 8, fill=1, stroke=0)
        # texte
        c.setFillColor(colors.black)
        c.drawString(left_margin + 14, y - 2, label)
        y -= 14

    y -= 10  # petit espace avant le tableau

    # ---------- Tableau : colonnes ----------
    headers = ["Urgence", "Nom de l'affaire", "Code", "Date", "Statut", "État"]
    col_widths = [30, 210, 60, 60, 70, 40]
    table_left = left_margin
    row_height = 18

    def draw_table_header(y_pos):
        c.setFillColor(colors.HexColor("#EFEFEF"))
        c.rect(table_left, y_pos - row_height, sum(col_widths), row_height, fill=1, stroke=0)
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)
        x = table_left + 4
        for i, h in enumerate(headers):
            c.drawString(x, y_pos - row_height + 5, h)
            x += col_widths[i]
        return y_pos - row_height - 4  # nouvelle position y (avec marge)

    # Dessiner l'entête une première fois
    y = draw_table_header(y)

    # ---------- Lignes du tableau ----------
    c.setFont("Helvetica", 9)
    df_sorted = df.sort_values("date")

    for _, row in df_sorted.iterrows():
        # saut de page si on arrive en bas
        if y < 80:
            c.showPage()
            width, height = A4
            y = height - 60
            # redessiner l'entête sur la nouvelle page
            y = draw_table_header(y)
            c.setFont("Helvetica", 9)

        # Données de la ligne
        nom = str(row["nom"]) if not pd.isna(row["nom"]) else ""
        code = str(row["ref"]) if not pd.isna(row["ref"]) else ""
        d_val = row.get("date", None)
        if isinstance(d_val, str):
            try:
                d_val = pd.to_datetime(d_val, errors="coerce")
            except Exception:
                d_val = None
        if pd.isna(d_val):
            date_str = "-"
        else:
            date_str = d_val.strftime("%d/%m/%Y")

        statut_txt = str(row["statut"]) if not pd.isna(row["statut"]) else ""

        # Couleur d'urgence (rectangle)
        urg_color = get_urgence_color(row, opts) or "#DDDDDD"

        # Couleur de statut (rectangle)
        statut_color = get_statut_color(statut_txt) or "#DDDDDD"

        # Dessin ligne
        x = table_left

        # 1) Colonne Urgence : petit carré coloré
        c.setFillColor(colors.HexColor(urg_color))
        c.rect(x + 8, y - row_height + 4, 8, 8, fill=1, stroke=0)  # carré au centre de la cellule
        x += col_widths[0]

        # 2) Nom
        c.setFillColor(colors.black)
        c.drawString(x + 2, y - row_height + 5, nom[:40])
        x += col_widths[1]

        # 3) Code
        c.drawString(x + 2, y - row_height + 5, code[:15])
        x += col_widths[2]

        # 4) Date
        c.drawString(x + 2, y - row_height + 5, date_str)
        x += col_widths[3]

        # 5) Statut (texte)
        c.drawString(x + 2, y - row_height + 5, statut_txt[:15])
        x += col_widths[4]

        # 6) Colonne État : petit carré couleur statut
        c.setFillColor(colors.HexColor(statut_color))
        c.rect(x + 10, y - row_height + 4, 8, 8, fill=1, stroke=0)

        # passer à la ligne suivante
        y -= row_height

    # ---------- QR CODE en bas de page ----------
    if qr_url:
        qr = qrcode.QRCode(box_size=3, border=1)
        qr.add_data(qr_url)
        qr.make(fit=True)
        img = qr.make_image(fill_color="black", back_color="white")

        qr_buf = BytesIO()
        img.save(qr_buf, format="PNG")
        qr_buf.seek(0)

        c.drawImage(ImageReader(qr_buf), width - 45 * mm, 15 * mm, width=30 * mm, preserveAspectRatio=True)
        c.setFont("Helvetica", 9)
        c.setFillColor(colors.black)
        c.drawRightString(width - 10 * mm, 12 * mm, "Accès application")

    c.save()
    buf.seek(0)
    return buf


# ================== QR CODE SIMPLE POUR L'APP ==================

def build_qr_image(url):
    qr = qrcode.QRCode(box_size=8, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image()
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


# ================== INTERFACE STREAMLIT ==================

st.set_page_config(page_title="Liste Priorité Chantier", layout="wide")
st.title("📋 Gestion des priorités chantier")

# Choix du mode
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

# États internes
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
            rouge = st.number_input("Délai ROUGE (jours)", 0, 90, opts["rouge"], key="opt_r")
            orange = st.number_input("Délai ORANGE (jours)", 0, 90, opts["orange"], key="opt_o")
            jaune = st.number_input("Délai JAUNE (jours)", 0, 90, opts["jaune"], key="opt_j")

        with col_opt2:
            show_qr = st.checkbox("Afficher QR Code dans le PDF", value=opts["show_qr"], key="opt_qr")

        if st.button("💾 Enregistrer options", key="save_opts"):
            opts.update({
                "rouge": rouge,
                "orange": orange,
                "jaune": jaune,
                "show_qr": show_qr
            })
            sauvegarder_options(opts)
            st.success("Options sauvegardées ✔")


st.divider()


# ================== IMPORT EXCEL “MONTAGE” ==================

def importer_excel(file):
    df_x = pd.read_excel(file)

    # Normalisation des colonnes
    # - supprime les espaces
    # - met tout en minuscules
    # - remplace apostrophes fantômes “ ’ ” par "'"
    df_x.columns = (
        df_x.columns
        .str.strip()
        .str.lower()
        .str.replace("’", "'", regex=False)
    )

    # Noms attendus normalisés
    required_nom = "nom de l'affaire"
    required_code = "code de l'affaire"
    required_comp = "compétence"
    required_debut = "début"

    # Vérification colonne compétence
    if required_comp not in df_x.columns:
        st.error("❌ Colonne manquante dans l’Excel : Compétence")
        return pd.DataFrame()

    # Filtrer uniquement Montage
    df_montage = df_x[df_x[required_comp].astype(str).str.lower() == "montage"]

    if df_montage.empty:
        st.warning("Aucune ligne 'Montage' trouvée.")
        return pd.DataFrame()

    # Vérification colonnes principales
    missing = []
    if required_nom not in df_x.columns:
        missing.append("Nom de l'affaire")
    if required_code not in df_x.columns:
        missing.append("Code de l'affaire")
    if required_debut not in df_x.columns:
        missing.append("Début")

    if missing:
        st.error("❌ Colonnes manquantes dans l'Excel : " + ", ".join(missing))
        return pd.DataFrame()

    # Extraction propre
    df_out = df_montage[[required_nom, required_code, required_debut]].copy()

    # Renommage normalisé
    df_out.columns = ["nom", "ref", "date"]

    # Conversion date
    df_out["date"] = pd.to_datetime(df_out["date"], errors="coerce")

    return df_out


if admin:
    with st.expander("📥 Importer depuis un export Excel (Montage)", expanded=False):

        st.caption("Rappel : seules les lignes avec Compétence = 'Montage' seront importées.")

        excel_file = st.file_uploader("Importer un fichier Excel", type=["xlsx"], key="import_xls")

        df_excel = pd.DataFrame()

        if excel_file:
            df_excel = importer_excel(excel_file)

            if not df_excel.empty:
                st.success("Extraction réussie ✔")

                st.dataframe(df_excel, use_container_width=True)

                choix = st.selectbox(
                    "Sélectionner une ligne à ajouter dans la liste",
                    df_excel.index,
                    key="xl_select",
                    format_func=lambda i: f"{df_excel.loc[i,'nom']} — {df_excel.loc[i,'ref']}"
                )

                if st.button("➕ Ajouter ce chantier", key="xl_add"):
                    row = df_excel.loc[choix]

                    new_row = {
                        "nom": row["nom"],
                        "ref": row["ref"],
                        "date": row["date"],
                        "commentaire": "",
                        "statut": "Prévu",
                        "priorite": "auto"
                    }

                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    sauvegarder_chantiers(df)
                    st.success("Chantier ajouté ✔")
                    st.rerun()


st.divider()


# ================== AJOUT MANUEL D’UN CHANTIER ==================

if admin:
    with st.expander("✏️ Ajouter manuellement un chantier", expanded=False):

        col1, col2 = st.columns(2)

        with col1:
            nom_manual = st.text_input("Nom de l'affaire", key="nom_manu")
            date_manual = st.date_input("Date de montage", value=date.today(), key="date_manu")

        with col2:
            ref_manual = st.text_input("Code de l'affaire", key="ref_manu")
            statut_manual = st.selectbox(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                key="statut_manu"
            )

        comm_manual = st.text_area("Commentaire", key="comm_manu")

        prio_manual = st.selectbox(
            "Priorité manuelle",
            ["Automatique", "Rouge", "Orange", "Jaune", "Gris"],
            key="prio_manu"
        )

        if st.button("➕ Ajouter ce chantier", key="add_manu_btn"):
            if nom_manual and ref_manual:
                df = pd.concat([
                    df,
                    pd.DataFrame([{
                        "nom": nom_manual,
                        "ref": ref_manual,
                        "date": str(date_manual),
                        "commentaire": comm_manual,
                        "statut": statut_manual,
                        "priorite": prio_manual.lower() if prio_manual != "Automatique" else "auto",
                    }])
                ], ignore_index=True)

                sauvegarder_chantiers(df)
                st.success("Ajouté ✔")
                st.rerun()
            else:
                st.error("Veuillez renseigner au minimum Nom et Code.")


# ================== LISTE + RECHERCHE + FILTRES ==================

st.subheader("📌 Liste des chantiers")

if df.empty:
    st.info("Aucun chantier enregistré.")
else:
    # Tri de base
    df_sorted = df.sort_values("date")

    # Bouton d'affichage des filtres
    col_btn, _ = st.columns([1, 3])
    with col_btn:
        if st.button("🔎 Filtres", key="toggle_filters"):
            st.session_state["show_filters"] = not st.session_state["show_filters"]

    # Variables de filtres
    statuts_sel = []
    start_date = end_date = None
    search = ""

    if st.session_state["show_filters"]:
        st.markdown("#### 🎛 Panneau de filtres")
        col_filters, col_table = st.columns([1, 3])
    else:
        col_filters, col_table = st.columns([0.0001, 1])

    # -------- Panneau de filtres --------
    with col_filters:
        if st.session_state["show_filters"]:
            st.markdown("##### Filtres")

            # Recherche texte globale
            search = st.text_input(
                "Recherche",
                placeholder="Nom, code, statut, commentaire...",
                key="search_global"
            )

            # Filtre sur les statuts
            statuts_uniques = sorted(df_sorted["statut"].dropna().unique().tolist())
            statuts_sel = st.multiselect(
                "Statut",
                statuts_uniques,
                default=statuts_uniques,
                key="filter_statut"
            )

            # Filtre sur les dates
            dates_valides = df_sorted["date"].dropna()
            if not dates_valides.empty:
                min_date = dates_valides.min().date()
                max_date = dates_valides.max().date()
                start_date = st.date_input("Date min", value=min_date, key="filter_date_min")
                end_date = st.date_input("Date max", value=max_date, key="filter_date_max")
            else:
                st.caption("Aucune date valide pour filtrer.")

            # Bouton de réinitialisation
            if st.button("Réinitialiser les filtres", key="reset_filters"):
                st.session_state["search_global"] = ""
                st.session_state["filter_statut"] = statuts_uniques
                if "filter_date_min" in st.session_state:
                    del st.session_state["filter_date_min"]
                if "filter_date_max" in st.session_state:
                    del st.session_state["filter_date_max"]
                st.rerun()

    # -------- Tableau principal --------
    with col_table:
        df_filtered = df_sorted.copy()

        # Filtre statut
        if statuts_sel:
            df_filtered = df_filtered[df_filtered["statut"].isin(statuts_sel)]

        # Filtre dates
        if start_date and end_date:
            mask_date = df_filtered["date"].notna()
            mask_date &= df_filtered["date"].dt.date.between(start_date, end_date)
            df_filtered = df_filtered[mask_date]

        # Filtre texte global
        if search:
            mask = pd.Series(False, index=df_filtered.index)
            for col in ["nom", "ref", "commentaire", "statut"]:
                mask |= df_filtered[col].astype(str).str.contains(search, case=False, na=False)
            df_filtered = df_filtered[mask]

        # Construction du tableau affiché
        df_display = df_filtered.copy()

        # Gestion des dates en texte
        if not df_display.empty and pd.api.types.is_datetime64_any_dtype(df_display["date"]):
            df_display["date"] = df_display["date"].dt.strftime("%d/%m/%Y")

        # Renommage des colonnes
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

        # Ajout des colonnes pastilles (urgence + statut)
        # On repart de df_filtered (version avec vraie date) pour calculer les pastilles
        df_filtered_for_emoji = df_filtered.copy()

        # Urgence (tout à gauche)
        urgence_col = []
        for idx in df_filtered_for_emoji.index:
            urgence_col.append(urgence_emoji(df_filtered_for_emoji.loc[idx], opts))

        # Statut (tout à droite)
        statut_col = []
        for idx in df_filtered_for_emoji.index:
            statut_col.append(statut_emoji(df_filtered_for_emoji.loc[idx, "statut"]))

        df_display.insert(0, "Urgence", urgence_col)
        df_display["État"] = statut_col

        # Affichage
        st.caption(f"{len(df_filtered)} chantier(s) affiché(s)")
        st.dataframe(
            df_display,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Urgence": st.column_config.TextColumn(
                    "Urgence",
                    help="🔴 très urgent · 🟠 moyen · 🟡 approche · ⚪ pas urgent"
                ),
                "État": st.column_config.TextColumn(
                    "État",
                    help="⚪ prévu · 🟣 en cours · 🔵 en attente · 🟢 terminé"
                ),
            },
        )


# ================== MODIFIER / SUPPRIMER (ADMIN) ==================

st.divider()

if admin and not df.empty:
    with st.expander("🛠 Modifier ou supprimer un chantier", expanded=False):

        choix_modif = st.selectbox(
            "Sélectionner un chantier",
            df.index,
            key="edit_select",
            format_func=lambda i: f"{df.loc[i, 'nom']} — {df.loc[i, 'ref']}"
        )

        col1, col2 = st.columns(2)

        with col1:
            new_nom = st.text_input("Nom de l'affaire", df.loc[choix_modif, "nom"], key=f"edit_nom_{choix_modif}")
            current_date = df.loc[choix_modif, "date"]
            new_date = st.date_input(
                "Date de montage",
                value=current_date.date() if not pd.isna(current_date) else date.today(),
                key=f"edit_date_{choix_modif}"
            )

        with col2:
            new_ref = st.text_input("Code de l'affaire", df.loc[choix_modif, "ref"], key=f"edit_ref_{choix_modif}")
            new_statut = st.selectbox(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                index=["Prévu", "En cours", "En attente", "Terminé"].index(df.loc[choix_modif, "statut"]),
                key=f"edit_statut_{choix_modif}"
            )

        new_commentaire = st.text_area(
            "Commentaire",
            df.loc[choix_modif, "commentaire"],
            key=f"edit_comm_{choix_modif}"
        )

        prio_list = ["Automatique", "Rouge", "Orange", "Jaune", "Gris"]
        current_prio = df.loc[choix_modif, "priorite"]
        prio_display = current_prio.capitalize() if current_prio != "auto" else "Automatique"

        new_prio = st.selectbox(
            "Priorité manuelle",
            prio_list,
            index=prio_list.index(prio_display),
            key=f"edit_prio_{choix_modif}"
        )

        colA, colB = st.columns(2)

        with colA:
            if st.button("💾 Enregistrer modifications", key=f"save_edit_{choix_modif}"):
                df.loc[choix_modif, "nom"] = new_nom
                df.loc[choix_modif, "ref"] = new_ref
                df.loc[choix_modif, "date"] = str(new_date)
                df.loc[choix_modif, "statut"] = new_statut
                df.loc[choix_modif, "commentaire"] = new_commentaire
                df.loc[choix_modif, "priorite"] = new_prio.lower() if new_prio != "Automatique" else "auto"

                sauvegarder_chantiers(df)
                st.success("Modifications enregistrées ✔")
                st.rerun()

        with colB:
            if st.button("🗑 Supprimer ce chantier", key=f"delete_{choix_modif}"):
                df = df.drop(choix_modif).reset_index(drop=True)
                sauvegarder_chantiers(df)
                st.success("Chantier supprimé ✔")
                st.rerun()


# ================== EXPORTS ==================

st.divider()
st.subheader("📤 Exports")

if df.empty:
    st.info("Ajoutez un chantier pour activer les exports.")
else:
    col_pdf, col_excel = st.columns(2)

    with col_pdf:
        qr_url_for_pdf = st.session_state.get("qr_url") or None

        st.download_button(
            "📄 Export PDF",
            build_pdf(df, opts, qr_url=qr_url_for_pdf),
            file_name=f"Gestion_Projet_Priorites_{datetime.now().strftime('%Y-%m-%d')}.pdf",
            mime="application/pdf"
        )

    with col_excel:
        st.download_button(
            "📊 Export Excel",
            build_excel(df),
            file_name="chantiers.xlsx"
        )


# ================== QR CODE DANS L’APPLICATION ==================

if opts["show_qr"]:
    st.divider()
    st.subheader("📱 QR Code atelier")

    qr_url = st.text_input(
        "URL de l'application (pour générer un QR code)",
        value=st.session_state["qr_url"],
        key="qr_url"
    )

    if qr_url:
        st.image(build_qr_image(qr_url), width=200)
        st.caption("Scannez ce QR Code pour accéder à l’application")
