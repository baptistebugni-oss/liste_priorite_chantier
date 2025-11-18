import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import qrcode

# ================== FICHIERS & CONSTANTES ==================

DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"
ADMIN_PASSWORD = "admin123"  # change-le si tu veux

DEFAULT_OPTIONS = {
    "rouge": 2,       # jours avant la date -> rouge
    "orange": 7,      # jours -> orange
    "jaune": 14,      # jours -> jaune
    "show_qr": True,
}

COLUMNS = ["nom", "ref", "date", "commentaire", "statut"]


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

    # S'assurer que toutes les colonnes existent
    for col in COLUMNS:
        if col not in df.columns:
            df[col] = ""

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
    """Charge les options de couleur / QR ou crée des valeurs par défaut."""
    if not os.path.exists(OPTIONS_FILE):
        sauvegarder_options(DEFAULT_OPTIONS)
        return DEFAULT_OPTIONS.copy()

    try:
        with open(OPTIONS_FILE, "r", encoding="utf-8") as f:
            opts = json.load(f)
    except json.JSONDecodeError:
        opts = {}

    # Compléter avec les valeurs par défaut manquantes
    for k, v in DEFAULT_OPTIONS.items():
        if k not in opts:
            opts[k] = v

    return opts


def sauvegarder_options(opts: dict) -> None:
    with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(opts, f, indent=4)


def style_urgence(row: pd.Series, opts: dict):
    """Renvoie un style de ligne (couleur de fond) selon la date et les seuils."""
    d = row.get("date")
    if pd.isna(d):
        return [""] * len(row)

    today = datetime.today().date()
    delta = (d.date() - today).days

    if delta <= opts["rouge"]:
        color = "#FF4B4B"   # rouge
    elif delta <= opts["orange"]:
        color = "#FFA500"   # orange
    elif delta <= opts["jaune"]:
        color = "#FFD966"   # jaune
    else:
        return [""] * len(row)

    return [f"background-color: {color}"] * len(row)


# ================== EXPORTS ==================

def build_excel(df: pd.DataFrame) -> BytesIO:
    buf = BytesIO()
    df.to_excel(buf, index=False)
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

    # en-têtes
    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Nom")
    c.drawString(200, y, "Référence")
    c.drawString(320, y, "Date")
    c.drawString(400, y, "Statut")
    y -= 18
    c.setFont("Helvetica", 9)

    for _, row in df.iterrows():
        if y < 40:
            c.showPage()
            y = height - 40

        c.drawString(40, y, str(row["nom"])[:30])
        c.drawString(200, y, str(row["ref"])[:20])
        if not pd.isna(row["date"]):
            c.drawString(320, y, row["date"].strftime("%d/%m/%Y"))
        else:
            c.drawString(320, y, "-")
        c.drawString(400, y, str(row["statut"])[:20])
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

# -------- OPTIONS ADMIN AVANCÉES --------
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

# -------- AJOUT DE CHANTIER (ADMIN) --------
if admin:
    st.subheader("➕ Ajouter un chantier")

    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom du chantier")
        date_montage = st.date_input("Date de montage")
    with col2:
        ref = st.text_input("Référence")
        statut = st.selectbox("Statut", ["Prévu", "En cours", "Terminé", "En attente"])

    commentaire = st.text_area("Commentaire (optionnel)")

    if st.button("Ajouter le chantier"):
        if nom and ref:
            new_row = {
                "nom": nom,
                "ref": ref,
                "date": str(date_montage),
                "commentaire": commentaire,
                "statut": statut,
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
        st.subheader("✏️ Modifier un chantier")
        idx_options = df_sorted.index.tolist()

        if idx_options:
            idx_sel = st.selectbox(
                "Choisir un chantier à modifier",
                idx_options,
                format_func=lambda x: f"{df_sorted.loc[x, 'nom']} - {df_sorted.loc[x, 'ref']}",
            )

            # Récupérer l'index réel dans df
            real_idx = idx_sel

            col1, col2 = st.columns(2)
            with col1:
                new_nom = st.text_input("Nom", value=str(df.loc[real_idx, "nom"]))
                new_date = st.date_input(
                    "Date de montage",
                    value=df.loc[real_idx, "date"].date() if not pd.isna(df.loc[real_idx, "date"]) else datetime.today().date(),
                    key="edit_date",
                )
            with col2:
                new_ref = st.text_input("Référence", value=str(df.loc[real_idx, "ref"]))
                new_statut = st.selectbox(
                    "Statut",
                    ["Prévu", "En cours", "Terminé", "En attente"],
                    index=["Prévu", "En cours", "Terminé", "En attente"].index(str(df.loc[real_idx, "statut"])) if df.loc[real_idx, "statut"] in ["Prévu", "En cours", "Terminé", "En attente"] else 0,
                    key="edit_statut",
                )

            new_commentaire = st.text_area(
                "Commentaire",
                value=str(df.loc[real_idx, "commentaire"]),
                key="edit_commentaire",
            )

            col_mod, col_suppr = st.columns(2)
            with col_mod:
                if st.button("Enregistrer les modifications"):
                    df.loc[real_idx, "nom"] = new_nom
                    df.loc[real_idx, "ref"] = new_ref
                    df.loc[real_idx, "date"] = str(new_date)
                    df.loc[real_idx, "statut"] = new_statut
                    df.loc[real_idx, "commentaire"] = new_commentaire
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
