import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import qrcode

# ================== CONFIG ==================

DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"
ADMIN_PASSWORD = "admin123"  # à modifier si tu veux

DEFAULT_OPTIONS = {
    "rouge": 2,
    "orange": 7,
    "jaune": 14,
    "show_qr": True,
}

COLUMNS = ["nom", "ref", "date", "commentaire", "statut", "priorite"]


# ================== UTILITAIRES ==================

def ensure_columns(df: pd.DataFrame, columns):
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    return df


def charger_chantiers():
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
    except:
        opts = {}
    for k, v in DEFAULT_OPTIONS.items():
        opts.setdefault(k, v)
    return opts


def sauvegarder_options(opts):
    with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(opts, f, indent=4)


# ================== SYSTEME DE PRIORITÉ (OPTION D) ==================

def get_row_color(row, opts):
    # 1) Priorité manuelle
    prio = row["priorite"].lower()
    if prio in ("rouge", "orange", "jaune", "gris"):
        return {
            "rouge": "#FF4B4B",
            "orange": "#FFA500",
            "jaune": "#FFD966",
            "gris": "#DDDDDD",
        }[prio]

    # 2) Statut
    statut = row["statut"].lower()
    statut_mapping = {
        "prévu": "#E5F6FF",
        "en cours": "#FFE4B5",
        "en attente": "#F0F0F0",
        "terminé": "",
    }
    if statut in statut_mapping and statut_mapping[statut]:
        return statut_mapping[statut]

    # 3) Automatique par date
    d = row["date"]
    if pd.isna(d):
        return ""
    delta = (d.date() - datetime.today().date()).days

    if delta <= opts["rouge"]:
        return "#FF4B4B"
    elif delta <= opts["orange"]:
        return "#FFA500"
    elif delta <= opts["jaune"]:
        return "#FFD966"
    return ""


def style_urgence(row, opts):
    color = get_row_color(row, opts)
    return [f"background-color: {color}"] * len(row) if color else [""] * len(row)


# ================== EXPORT PDF & EXCEL ==================

def build_excel(df):
    buf = BytesIO()
    df2 = df.copy()
    df2["date"] = df2["date"].dt.strftime("%d/%m/%Y")
    df2.to_excel(buf, index=False)
    buf.seek(0)
    return buf


def build_pdf(df):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    width, height = A4

    y = height - 40
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "Liste des chantiers")
    y -= 30

    c.setFont("Helvetica-Bold", 11)
    c.drawString(40, y, "Nom")
    c.drawString(200, y, "Code")
    c.drawString(300, y, "Date")
    c.drawString(380, y, "Statut")
    y -= 20
    c.setFont("Helvetica", 10)

    for _, row in df.iterrows():
        if y < 40:
            c.showPage()
            y = height - 40
        c.drawString(40, y, str(row["nom"])[:25])
        c.drawString(200, y, str(row["ref"])[:12])
        c.drawString(300, y, row["date"].strftime("%d/%m/%Y") if not pd.isna(row["date"]) else "-")
        c.drawString(380, y, str(row["statut"]))
        y -= 15

    c.save()
    buf.seek(0)
    return buf


# ================== QR CODE ==================

def build_qr(url):
    qr = qrcode.QRCode(box_size=8, border=2)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image()
    buf = BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf


# ================== IMPORT EXCEL (CORRIGÉ POUR TON FICHIER) ==================

def importer_excel(file):
    df = pd.read_excel(file)
    df.columns = df.columns.str.strip().str.lower()

    col_comp = "compétence"
    col_nom = "nom de l’affaire"
    col_code = "code de l’affaire"
    col_debut = "début"

    for col in [col_comp, col_nom, col_code, col_debut]:
        if col not in df.columns:
            st.error(f"⚠️ Colonne manquante : {col}")
            return pd.DataFrame()

    df_filtre = df[df[col_comp].astype(str).str.lower() == "montage"].copy()

    if df_filtre.empty:
        st.warning("⚠️ Aucune ligne 'Montage' trouvée.")
        return pd.DataFrame()

    df_result = df_filtre[[col_nom, col_code, col_debut]].copy()
    df_result.columns = ["nom", "ref", "date"]
    df_result["date"] = pd.to_datetime(df_result["date"], errors="coerce")

    return df_result
