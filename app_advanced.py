import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from io import BytesIO

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import qrcode

DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"
ADMIN_PASSWORD = "admin123"

DEFAULT_OPTIONS = {
    "rouge": 2,
    "orange": 7,
    "jaune": 14,
    "show_qr": True,
}

def charger_chantiers() -> pd.DataFrame:
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame(columns=["nom", "ref", "date"])
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
    except:
        data = []
    df = pd.DataFrame(data)
    for col in ["nom", "ref", "date"]:
    if col not in df.columns:
        df[col] = ""
    if not df.empty:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    return df

def sauvegarder_chantiers(df: pd.DataFrame):
    df = df.copy()
    if not df.empty:
        df["date"] = df["date"].astype(str)
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"), f, ensure_ascii=False, indent=4)

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

def style_urgence(row, opts):
    d = row.get("date")
    if pd.isna(d):
        return ["" for _ in row]
    delta = (d.date() - datetime.today().date()).days
    if delta <= opts["rouge"]:
        c = "#FF4B4B"
    elif delta <= opts["orange"]:
        c = "#FFA500"
    elif delta <= opts["jaune"]:
        c = "#FFD966"
    else:
        return ["" for _ in row]
    return [f"background-color: {c}" for _ in row]

def build_excel(df):
    buf = BytesIO()
    df.to_excel(buf, index=False)
    buf.seek(0)
    return buf

def build_pdf(df):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    y = h - 40
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, y, "Liste des chantiers")
    y -= 30
    for _, row in df.iterrows():
        if y < 40:
            c.showPage()
            y = h - 40
        c.setFont("Helvetica", 12)
        c.drawString(50, y, str(row["nom"]))
        c.drawString(250, y, str(row["ref"]))
        c.drawString(400, y, row["date"].strftime("%d/%m/%Y") if not pd.isna(row["date"]) else "-")
        y -= 20
    c.save()
    buf.seek(0)
    return buf

def build_qr(url):
    import qrcode
    buf = BytesIO()
    qrcode.make(url).save(buf, format="PNG")
    buf.seek(0)
    return buf

st.set_page_config(page_title="Liste Priorité Chantier", layout="centered")
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

if admin:
    st.sidebar.subheader("⚙️ Options avancées")
    opts["rouge"] = st.sidebar.number_input("Rouge (jours)", 0, 60, opts["rouge"])
    opts["orange"] = st.sidebar.number_input("Orange (jours)", 0, 60, opts["orange"])
    opts["jaune"] = st.sidebar.number_input("Jaune (jours)", 0, 60, opts["jaune"])
    opts["show_qr"] = st.sidebar.checkbox("Afficher QR", opts["show_qr"])
    if st.sidebar.button("Enregistrer"):
        sauvegarder_options(opts)
        st.sidebar.success("Options enregistrées")

if admin:
    st.subheader("➕ Ajouter un chantier")
    nom = st.text_input("Nom")
    ref = st.text_input("Référence")
    date = st.date_input("Date montage")
    if st.button("Ajouter"):
        if nom and ref:
            df = pd.concat([df, pd.DataFrame([{"nom": nom, "ref": ref, "date": str(date)}])], ignore_index=True)
            sauvegarder_chantiers(df)
            st.rerun()
        else:
            st.error("Nom + Référence obligatoires.")

st.subheader("📌 Liste des chantiers")
if df.empty:
    st.info("Aucun chantier.")
else:
    df_sorted = df.sort_values(by="date", na_position="last")
    st.dataframe(df_sorted.style.apply(lambda r: style_urgence(r, opts), axis=1), use_container_width=True)

if admin and not df.empty:
    st.subheader("🗑 Supprimer un chantier")
    idx = st.selectbox("Choisir", df.index, format_func=lambda x: df.loc[x, "nom"])
    if st.button("Supprimer"):
        df = df.drop(idx).reset_index(drop=True)
        sauvegarder_chantiers(df)
        st.rerun()

if not df.empty:
    st.subheader("📤 Export")
    st.download_button("📄 PDF", build_pdf(df_sorted), "chantiers.pdf")
    st.download_button("📊 Excel", build_excel(df_sorted), "chantiers.xlsx")

if opts["show_qr"]:
    st.subheader("📱 QR Code")
    url = st.text_input("URL Streamlit")
    if url:
        st.image(build_qr(url), width=200)
