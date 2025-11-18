import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime

# ----------------------- FICHIERS -----------------------
DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"
ADMIN_PASSWORD = "admin123"  # À modifier selon ton choix

# ----------------------- CHARGEMENT -----------------------
def charger_chantiers():
    """Charge le fichier JSON et assure un DataFrame propre et complet."""

    # Si fichier inexistant → créer DataFrame vide
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame(columns=["nom", "ref", "date"])

    # Lire JSON
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        try:
            data = json.load(f)
        except json.JSONDecodeError:
            data = []

    # Transformer en DataFrame
    df = pd.DataFrame(data)

    # Colonnes obligatoires
    colonnes = ["nom", "ref", "date"]
    for col in colonnes:
        if col not in df.columns:
            df[col] = ""

    # Conversion date sécurisée
    if not df.empty:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    return df

# ----------------------- SAUVEGARDE -----------------------
def sauvegarder_chantiers(df):
    """Sauvegarde le DataFrame dans le fichier JSON."""
    df_to_save = df.copy()
    df_to_save["date"] = df_to_save["date"].astype(str)

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(df_to_save.to_dict(orient="records"), f, ensure_ascii=False, indent=4)

# ----------------------- OPTIONS -----------------------
def charger_options():
    """Charge les options ou crée les valeurs par défaut."""
    if not os.path.exists(OPTIONS_FILE):
        opts = {"rouge": 2, "orange": 7, "jaune": 14}
        sauvegarder_options(opts)
        return opts

    with open(OPTIONS_FILE, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError:
            return {"rouge": 2, "orange": 7, "jaune": 14}


def sauvegarder_options(opts):
    with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(opts, f, indent=4)

# ============================================================
#                         INTERFACE APP
# ============================================================
st.set_page_config(page_title="Liste Priorité Chantier", layout="wide")
st.title("📋 Gestion des priorités chantier")

# Mode utilisateur / admin
mode = st.sidebar.selectbox("Mode", ["Lecture seule", "Administrateur"])
admin = False

if mode == "Administrateur":
    mdp = st.sidebar.text_input("Mot de passe", type="password")
    if mdp == ADMIN_PASSWORD:
        admin = True
    else:
        st.sidebar.error("Mot de passe incorrect")

# Charger données
df = charger_chantiers()
opts = charger_options()

# ----------------------- AJOUT -----------------------
if admin:
    st.subheader("➕ Ajouter un chantier")

    nom = st.text_input("Nom du chantier")
    ref = st.text_input("Référence")
    date = st.date_input("Date de montage")

    if st.button("Ajouter le chantier"):
        new_row = {"nom": nom, "ref": ref, "date": str(date)}
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        sauvegarder_chantiers(df)
        st.success("Chantier ajouté !")
        st.rerun()

# ----------------------- SUPPRESSION -----------------------
if admin:
    st.subheader("🗑 Supprimer un chantier")

    if df.empty:
        st.info("Aucun chantier à supprimer.")
    else:
        idx = st.selectbox(
            "Sélectionner un chantier",
            df.index,
            format_func=lambda x: f"{df.loc[x, 'nom']} - {df.loc[x, 'ref']}"
        )

        if st.button("Supprimer"):
            df = df.drop(idx).reset_index(drop=True)
            sauvegarder_chantiers(df)
            st.success("Chantier supprimé !")
            st.rerun()

# ----------------------- AFFICHAGE -----------------------
st.subheader("📌 Liste des chantiers")

if df.empty:
    st.info("Aucun chantier enregistré.")
else:
    df_aff = df.sort_values(by="date")
    st.dataframe(df_aff)