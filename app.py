import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime
from io import BytesIO


from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import qrcode


# ----------------------- FICHIERS & CONST -----------------------
DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"
ADMIN_PASSWORD = "admin123" # À changer si besoin


DEFAULT_OPTIONS = {
"rouge": 2,
"orange": 7,
"jaune": 14,
"show_qr": True,
}


# ----------------------- DONNÉES -----------------------


def charger_chantiers() -> pd.DataFrame:
"""Charge le fichier JSON et renvoie un DataFrame propre."""
if not os.path.exists(DATA_FILE):
return pd.DataFrame(columns=["nom", "ref", "date"])


try:
with open(DATA_FILE, "r", encoding="utf-8") as f:
data = json.load(f)
except (json.JSONDecodeError, FileNotFoundError):
data = []


df = pd.DataFrame(data)
for col in ["nom", "ref", "date"]:
if col not in df.columns:
df[col] = ""


if not df.empty:
df["date"] = pd.to_datetime(df["date"], errors="coerce")


return df




def sauvegarder_chantiers(df: pd.DataFrame) -> None:
df_to_save = df.copy()
if not df_to_save.empty:
df_to_save["date"] = df_to_save["date"].astype(str)


with open(DATA_FILE, "w", encoding="utf-8") as f:
json.dump(df_to_save.to_dict(orient="records"), f, ensure_ascii=False, indent=4)




# ----------------------- OPTIONS -----------------------


def charger_options() -> dict:
if not os.path.exists(OPTIONS_FILE):
sauvegarder_options(DEFAULT_OPTIONS)
return DEFAULT_OPTIONS.copy()


try:
with open(OPTIONS_FILE, "r", encoding="utf-8") as f:
opts = json.load(f)
except json.JSONDecodeError:
opts = {}


# complèter avec les valeurs par défaut
for k, v in DEFAULT_OPTIONS.items():
opts.setdefault(k, v)


return opts




def sauvegarder_options(opts: dict) -> None:
with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
json.dump(opts, f, indent=4)




# ----------------------- COULEURS D'URGENCE -----------------------


def style_urgence(row, opts: dict):
"""Retourne un style de ligne en fonction de la date et des seuils."""
d = row.get("date")
if pd.isna(d):
return [""] * len(row)


today = datetime.today().date()
delta = (d.date() - today).days


if delta <= opts["rouge"]:
color = "#FF4B4B" # rouge vif
elif delta <= opts["orange"]:
color = "#FFA500" # orange
elif delta <= opts["jaune"]:
color = "#FFD966" # jaune
else:
color = ""


if not color:
return [""] * len(row)


return [f"background-color: {color}"] * len(row)




# ----------------------- EXPORTS -----------------------


def build_excel(df: pd.DataFrame) -> BytesIO:
st.info("Renseignez l'URL de l'application pour générer un QR code.")
