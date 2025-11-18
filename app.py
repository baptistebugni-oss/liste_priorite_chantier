import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime


DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"


# ------------------- CHARGEMENT DES DONNÉES -------------------
# --- FIXED VERSION ---
def charger_chantiers():
if not os.path.exists(DATA_FILE):
return pd.DataFrame(columns=["nom", "ref", "date"])


with open(DATA_FILE, "r", encoding="utf-8") as f:
data = json.load(f)


df = pd.DataFrame(data)


# S'assurer que les colonnes existent
for col in ["nom", "ref", "date"]:
if col not in df.columns:
df[col] = ""


# Convertir la date
if not df.empty:
df["date"] = pd.to_datetime(df["date"], errors="coerce")


return df


# ------------------- SAUVEGARDE DES DONNÉES -------------------
def sauvegarder_chantiers(df):
df_to_save = df.copy()
df_to_save["date"] = df_to_save["date"].astype(str)


with open(DATA_FILE, "w", encoding="utf-8") as f:
json.dump(df_to_save.to_dict(orient="records"), f, ensure_ascii=False, indent=4)


# ------------------- CHARGEMENT OPTIONS -------------------
def charger_options():
if not os.path.exists(OPTIONS_FILE):
opts = {"rouge": 2, "orange": 7, "jaune": 14}
with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
json.dump(opts, f, indent=4)
return opts


with open(OPTIONS_FILE, "r", encoding="utf-8") as f:
return json.load(f)


# ------------------- SAUVEGARDE OPTIONS -------------------
def sauvegarder_options(opts):
with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
json.dump(opts, f, indent=4)


# ------------------- APPLICATION -------------------
st.set_page_config(page_title="Liste Priorité Chantier", layout="wide")
st.title("📋 Gestion des priorités chantier")


mdp_admin = "admin123" # À personnaliser
mode = st.sidebar.selectbox("Mode", ["Lecture seule", "Administrateur"])


admin = False
if mode == "Administrateur":
mdp = st.sidebar.text_input("Mot de passe", type="password")
if mdp == mdp_admin:
admin = True
else:
st.sidebar.error("Mot de passe incorrect")


# Charger les données
df = charger_chantiers()
options = charger_options()


# ------------------- ADMIN : AJOUT -------------------
if admin:
st.subheader("➕ Ajouter un chantier")
nom = st.text_input("Nom du chantier")
ref = st.text_input("Référence")
date = st.date_input("Date de montage")


if st.button("Ajouter"):
new_row = {"nom": nom, "ref": ref, "date": str(date)}
df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
sauvegarder_chantiers(df)
st.success("Chantier ajouté !")


# ------------------- ADMIN : SUPPRESSION -------------------
if admin:
st.subheader("🗑 Supprimer un chantier")
if df.empty:
st.info("Aucun chantier à supprimer")
else:
idx_to_delete = st.selectbox(
"Choisir un chantier", df.index,
format_func=lambda x: f"{df.loc[x, 'nom']} - {df.loc[x, 'ref']}"
)
if st.button("Supprimer"):
df = df.drop(idx_to_delete).reset_index(drop=True)
sauvegarder_chantiers(df)
st.success("Chantier supprimé !")
st.experimental_rerun()


# ------------------- LISTE ORDONNÉE -------------------
st.subheader("📌 Liste des chantiers")
if df.empty:
st.info("Aucun chantier enregistré")
else:
df_sorted = df.sort_values(by="date")
st.dataframe(df_sorted)
