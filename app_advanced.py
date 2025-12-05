import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import qrcode

# ================== FICHIERS & CONSTANTES ==================

DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"
ADMIN_PASSWORD = "admin123"  # change-le si tu veux

# Seuils par défaut pour les couleurs automatiques
DEFAULT_OPTIONS = {
    "rouge": 2,       # jours avant la date -> rouge
    "orange": 7,      # jours -> orange
    "jaune": 14,      # jours -> jaune
    "show_qr": True,
}

# Colonnes pour la base de données chantiers
COLUMNS = ["nom", "ref", "date", "commentaire", "statut", "priorite"]


# ================== UTILITAIRES ==================

def ensure_columns(df: pd.DataFrame, columns) -> pd.DataFrame:
    """S'assure que toutes les colonnes existent dans le DataFrame."""
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    return df


def find_column(df: pd.DataFrame, candidates) -> str | None:
    """
    Cherche une colonne dont le nom contient l'un des 'candidates'
    (en minuscules, espaces ignorés).
    """
    normalized = {col.lower().strip(): col for col in df.columns}
    for key, original in normalized.items():
        for cand in candidates:
            if cand in key.replace(" ", ""):
                return original
    return None


# ================== GESTION DES DONNÉES ==================

def charger_chantiers() -> pd.DataFrame:
    """Charge le fichier JSON et renvoie un DataFrame propre."""
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame(columns=COLUMNS)

    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, FileNotFoundError):
        data = []

    df = pd.DataFrame(data)
    df = ensure_columns(df, COLUMNS)

    # Conversion de la date
    if not df.empty:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    return df


def sauvegarder_chantiers(df: pd.DataFrame) -> None:
    """Sauvegarde le DataFrame dans le JSON."""
    df_to_save = df.copy()
    if not df_to_save.empty:
        df_to_save["date"] = df_to_save["date"].astype(str)

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(df_to_save.to_dict(orient="records"), f, ensure_ascii=False, indent=4)


# ================== OPTIONS D’URGENCE ==================

def charger_options() -> dict:
    """Charge les options ou crée les valeurs par défaut."""
    if not os.path.exists(OPTIONS_FILE):
        sauvegarder_options(DEFAULT_OPTIONS)
        return DEFAULT_OPTIONS.copy()

    try:
        with open(OPTIONS_FILE, "r", encoding="utf-8") as f:
            opts = json.load(f)
    except json.JSONDecodeError:
        opts = {}

    for k, v in DEFAULT_OPTIONS.items():
        opts.setdefault(k, v)

    return opts


def sauvegarder_options(opts: dict) -> None:
    with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(opts, f, indent=4)


def get_row_color(row: pd.Series, opts: dict) -> str:
    """
    Système pro de priorité (Option D) :
    1) priorite manuelle > 2) statut > 3) automatique (date)
    Retourne une couleur hexadécimale ou "".
    """

    # 1. Priorité manuelle
    prio = str(row.get("priorite", "auto")).lower()
    if prio in ("rouge", "orange", "jaune", "gris"):
        mapping = {
            "rouge": "#FF4B4B",
            "orange": "#FFA500",
            "jaune": "#FFD966",
            "gris": "#DDDDDD",
        }
        return mapping.get(prio, "")

    # 2. Statut -> couleur (si pas de priorité manuelle)
    statut = str(row.get("statut", "")).lower()
    if statut:
        statut_map = {
            "prévu": "#E5F6FF",    # bleu très clair
            "en cours": "#FFE4B5", # beige/orangé
            "en attente": "#F0F0F0",
            "terminé": "",         # pas de couleur
        }
        if statut in statut_map and statut_map[statut]:
            return statut_map[statut]

    # 3. Automatique selon la date (si pas de couleur déjà définie)
    d = row.get("date")
    if pd.isna(d):
        return ""

    today = datetime.today().date()
    delta = (d.date() - today).days

    if delta <= opts["rouge"]:
        return "#FF4B4B"   # rouge
    elif delta <= opts["orange"]:
        return "#FFA500"   # orange
    elif delta <= opts["jaune"]:
        return "#FFD966"   # jaune

    return ""


def style_urgence(row: pd.Series, opts: dict):
    """Applique la couleur de fond à toute la ligne."""
    color = get_row_color(row, opts)
    if not color:
        return [""] * len(row)
    return [f"background-color: {color}"] * len(row)


# ================== EXPORTS ==================

def build_excel(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    df_to_export = df.copy()
    if not df_to_export.empty and isinstance(df_to_export["date"].iloc[0], (pd.Timestamp, datetime)):
        df_to_export["date"] = df_to_export["date"].dt.strftime("%d/%m/%Y")
    df_to_export.to_excel(buf, index=False)
    buf.seek(0)
    return buf


def build_pdf(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    y = height - 40
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "Liste des chantiers")
    y -= 20
    c.setFont("Helvetica", 10)
    c.drawString(50, y, f"Généré le {datetime.now().strftime('%d/%m/%Y %H:%M')}")
    y -= 30

    # En-têtes
    c.setFont("Helvetica-Bold", 11)
    c.drawString(30, y, "Nom")
    c.drawString(210, y, "Référence")
    c.drawString(320, y, "Date")
    c.drawString(400, y, "Statut")
    y -= 18
    c.setFont("Helvetica", 9)

    for _, row in df.iterrows():
        if y < 40:
            c.showPage()
            y = height - 40

        c.drawString(30, y, str(row["nom"])[:30])
        c.drawString(210, y, str(row["ref"])[:18])
        if not pd.isna(row["date"]):
            c.drawString(320, y, row["date"].strftime("%d/%m/%Y"))
        else:
            c.drawString(320, y, "-")
        c.drawString(400, y, str(row["statut"])[:18])
        y -= 16

    c.save()
    buf.seek(0)
    return buf


# ================== QR CODE ==================

def build_qr(url: str) -> BytesIO:
    qr = qrcode.QRCode(box_size=8, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


# ================== IMPORT EXCEL ==================

def importer_excel(file) -> pd.DataFrame:
    """
    Lit un fichier Excel, filtre les lignes où Compétence = 'Montage',
    et retourne un DF avec nom, ref, date (à partir de la colonne 'Début').
    """

    df = pd.read_excel(file)

    # Normalisation des en-têtes
    df.columns = df.columns.str.strip()

    # Trouver les colonnes importantes par nom approché
    col_comp = find_column(df, ["compétence", "competence"])
    col_nom = find_column(df, ["nom d'affaire", "nomdaffaire"])
    col_code = find_column(df, ["code d'affaire", "codedaffaire"])
    col_debut = find_column(df, ["début", "debut"])
    # col_fin possible plus tard si besoin
    # col_fin = find_column(df, ["fin"])

    required = {
        "Compétence": col_comp,
        "Nom d'affaire": col_nom,
        "Code d'affaire": col_code,
        "Début": col_debut,
    }

    missing = [name for name, col in required.items() if col is None]
    if missing:
        st.error("⚠️ Colonnes manquantes dans l'Excel : " + ", ".join(missing))
        return pd.DataFrame()

    # Filtrer sur Compétence = Montage
    df_filtre = df[df[col_comp].astype(str).str.lower() == "montage"].copy()
    if df_filtre.empty:
        st.warning("Aucune ligne avec Compétence = 'Montage' trouvée.")
        return pd.DataFrame()

    # Extraire les colonnes utiles
    df_result = df_filtre[[col_nom, col_code, col_debut]].copy()
    df_result.columns = ["nom", "ref", "date"]

    df_result["date"] = pd.to_datetime(df_result["date"], errors="coerce")

    return df_result


# ================== INTERFACE STREAMLIT ==================

st.set_page_config(page_title="Liste Priorité Chantier", layout="centered")
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

# Charger données et options
df = charger_chantiers()
opts = charger_options()

# -------- OPTIONS ADMIN (SEUILS & QR) --------
if admin:
    st.sidebar.markdown("### ⚙️ Options d'urgence")
    rouge = st.sidebar.number_input("Jours max ROUGE", 0, 60, int(opts["rouge"]))
    orange = st.sidebar.number_input("Jours max ORANGE", 0, 60, int(opts["orange"]))
    jaune = st.sidebar.number_input("Jours max JAUNE", 0, 60, int(opts["jaune"]))
    show_qr = st.sidebar.checkbox("Afficher la section QR", value=bool(opts.get("show_qr", True)))

    if st.sidebar.button("Enregistrer les options"):
        opts.update({"rouge": rouge, "orange": orange, "jaune": jaune, "show_qr": show_qr})
        sauvegarder_options(opts)
        st.sidebar.success("Options enregistrées")
else:
    show_qr = bool(opts.get("show_qr", True))

# -------- IMPORT EXCEL (ADMIN) --------
if admin:
    st.subheader("📥 Importer un fichier Excel")

    fichier_excel = st.file_uploader("Importer l’export Excel (fichier .xlsx)", type=["xlsx"])
    df_excel = pd.DataFrame()

    if fichier_excel:
        df_excel = importer_excel(fichier_excel)
        if not df_excel.empty:
            st.success("Fichier importé et filtré avec succès (Compétence = Montage).")
            st.dataframe(df_excel, use_container_width=True)

            st.markdown("#### ➕ Ajouter un chantier depuis l’Excel")

            choix = st.selectbox(
                "Sélectionner un chantier à ajouter",
                df_excel.index,
                format_func=lambda i: f"{df_excel.loc[i, 'nom']} — {df_excel.loc[i, 'ref']}",
            )

            if st.button("Ajouter ce chantier"):
                ligne = df_excel.loc[choix]
                df = charger_chantiers()

                new_row = {
                    "nom": ligne["nom"],
                    "ref": ligne["ref"],
                    "date": ligne["date"] if not pd.isna(ligne["date"]) else "",
                    "commentaire": "",
                    "statut": "Prévu",
                    "priorite": "auto",  # priorité automatique par défaut
                }

                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                sauvegarder_chantiers(df)
                st.success("Chantier ajouté automatiquement ✔")
                st.rerun()

# -------- AJOUT MANUEL (ADMIN) --------
if admin:
    st.subheader("✏️ Ajouter un chantier manuellement")

    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom du chantier")
        date_montage = st.date_input("Date de montage", value=date.today())
    with col2:
        ref = st.text_input("Code / Référence")
        statut = st.selectbox("Statut", ["Prévu", "En cours", "Terminé", "En attente"])

    commentaire = st.text_area("Commentaire (optionnel)")
    priorite_man = st.selectbox(
        "Priorité manuelle",
        ["Automatique", "Rouge", "Orange", "Jaune", "Gris"],
    )

    if st.button("Ajouter ce chantier manuellement"):
        if nom and ref:
            prio_value = priorite_man.lower() if priorite_man != "Automatique" else "auto"
            new_row = {
                "nom": nom,
                "ref": ref,
                "date": str(date_montage),
                "commentaire": commentaire,
                "statut": statut,
                "priorite": prio_value,
            }
            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            sauvegarder_chantiers(df)
            st.success("Chantier ajouté ✔")
            st.rerun()
        else:
            st.error("Nom et Référence sont obligatoires.")

# -------- LISTE + COULEURS D’URGENCE --------
st.subheader("📌 Liste des chantiers")

if df.empty:
    st.info("Aucun chantier pour le moment.")
else:
    df_sorted = df.sort_values(by="date", na_position="last")
    styled = df_sorted.style.apply(lambda row: style_urgence(row, opts), axis=1)
    st.dataframe(styled, use_container_width=True)

    # -------- MODIFIER / SUPPRIMER (ADMIN) --------
    if admin:
        st.subheader("🛠 Modifier ou supprimer un chantier")

        idx_options = df_sorted.index.tolist()
        if idx_options:
            idx_sel = st.selectbox(
                "Choisir un chantier",
                idx_options,
                format_func=lambda x: f"{df_sorted.loc[x, 'nom']} — {df_sorted.loc[x, 'ref']}",
            )

            real_idx = idx_sel

            col1, col2 = st.columns(2)
            with col1:
                new_nom = st.text_input("Nom", value=str(df.loc[real_idx, "nom"]))
                current_date = df.loc[real_idx, "date"]
                new_date = st.date_input(
                    "Date de montage",
                    value=current_date.date() if not pd.isna(current_date) else date.today(),
                )
            with col2:
                new_ref = st.text_input("Référence", value=str(df.loc[real_idx, "ref"]))
                statut_list = ["Prévu", "En cours", "Terminé", "En attente"]
                current_statut = df.loc[real_idx, "statut"]
                idx_statut = statut_list.index(current_statut) if current_statut in statut_list else 0
                new_statut = st.selectbox("Statut", statut_list, index=idx_statut)

            new_commentaire = st.text_area(
                "Commentaire",
                value=str(df.loc[real_idx, "commentaire"]),
            )

            priorite_list = ["Automatique", "Rouge", "Orange", "Jaune", "Gris"]
            current_prio = str(df.loc[real_idx, "priorite"])
            prio_display = "Automatique"
            for p in priorite_list[1:]:
                if current_prio == p.lower():
                    prio_display = p
                    break

            new_prio_display = st.selectbox("Priorité manuelle", priorite_list, index=priorite_list.index(prio_display))

            col_mod, col_suppr = st.columns(2)
            with col_mod:
                if st.button("Enregistrer les modifications"):
                    df.loc[real_idx, "nom"] = new_nom
                    df.loc[real_idx, "ref"] = new_ref
                    df.loc[real_idx, "date"] = str(new_date)
                    df.loc[real_idx, "statut"] = new_statut
                    df.loc[real_idx, "commentaire"] = new_commentaire
                    df.loc[real_idx, "priorite"] = new_prio_display.lower() if new_prio_display != "Automatique" else "auto"
                    sauvegarder_chantiers(df)
                    st.success("Chantier modifié ✔")
                    st.rerun()

            with col_suppr:
                if st.button("🗑 Supprimer ce chantier"):
                    df = df.drop(real_idx).reset_index(drop=True)
                    sauvegarder_chantiers(df)
                    st.success("Chantier supprimé ✔")
                    st.rerun()

# -------- EXPORTS --------
st.subheader("📤 Export des données")

if df.empty:
    st.info("Ajoutez au moins un chantier pour activer les exports.")
else:
    df_sorted = df.sort_values(by="date", na_position="last")

    col_pdf, col_xlsx = st.columns(2)
    with col_pdf:
        pdf_buf = build_pdf(df_sorted)
        st.download_button(
            "📄 Télécharger en PDF",
            data=pdf_buf,
            file_name="chantiers.pdf",
            mime="application/pdf",
        )

    with col_xlsx:
        xlsx_buf = build_excel(df_sorted)
        st.download_button(
            "📊 Télécharger en Excel",
            data=xlsx_buf,
            file_name="chantiers.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# -------- QR CODE --------
if show_qr:
    st.subheader("📱 Accès mobile (QR code)")
    st.caption("Colle l'URL de cette application pour générer un QR code à afficher dans l'atelier.")

    url = st.text_input("URL de l'application", placeholder="https://...")
    if url:
        qr_buf = build_qr(url)
        st.image(qr_buf, caption="Scannez avec votre téléphone", width=200)
    else:
        st.info("Renseigne l'URL de l'application pour générer le QR code.")
