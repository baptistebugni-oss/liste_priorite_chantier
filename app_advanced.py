import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime

DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"
ADMIN_PASSWORD = "admin123"

DEFAULT_OPTIONS = {"rouge": 2, "orange": 7, "jaune": 14, "show_qr": True}

def charger_chantiers():
    if not os.path.exists(DATA_FILE):
        return pd.DataFrame(columns=["nom","ref","date"])
    try:
        with open(DATA_FILE,"r",encoding="utf-8") as f:
            data=json.load(f)
    except:
        data=[]
    df=pd.DataFrame(data)
    for col in ["nom","ref","date"]:
        if col not in df.columns:
            df[col]=""
    if not df.empty:
        df["date"]=pd.to_datetime(df["date"],errors="coerce")
    return df

def sauvegarder_chantiers(df):
    df=df.copy()
    if not df.empty:
        df["date"]=df["date"].astype(str)
    with open(DATA_FILE,"w",encoding="utf-8") as f:
        json.dump(df.to_dict(orient="records"),f,ensure_ascii=False,indent=4)

def charger_options():
    if not os.path.exists(OPTIONS_FILE):
        sauvegarder_options(DEFAULT_OPTIONS)
        return DEFAULT_OPTIONS.copy()
    try:
        with open(OPTIONS_FILE,"r",encoding="utf-8") as f:
            opts=json.load(f)
    except:
        opts={}
    for k,v in DEFAULT_OPTIONS.items():
        if k not in opts:
            opts[k]=v
    return opts

def sauvegarder_options(opts):
    with open(OPTIONS_FILE,"w",encoding="utf-8") as f:
        json.dump(opts,f,indent=4)

def style_urgence(row,opts):
    d=row.get("date")
    if pd.isna(d):
        return ["" for _ in row]
    delta=(d.date()-datetime.today().date()).days
    if delta<=opts["rouge"]:
        c="#FF4B4B"
    elif delta<=opts["orange"]:
        c="#FFA500"
    elif delta<=opts["jaune"]:
        c="#FFD966"
    else:
        return ["" for _ in row]
    return [f"background-color:{c}" for _ in row]

st.set_page_config(page_title="Liste Priorité Chantier",layout="centered")
st.title("📋 Gestion des priorités chantier")

mode=st.sidebar.selectbox("Mode",["Lecture seule","Administrateur"])
admin=False

if mode=="Administrateur":
    mdp=st.sidebar.text_input("Mot de passe",type="password")
    if mdp==ADMIN_PASSWORD:
        admin=True
    elif mdp:
        st.sidebar.error("Mot de passe incorrect")

df=charger_chantiers()
opts=charger_options()

if admin:
    st.sidebar.subheader("⚙️ Options avancées")
    opts["rouge"]=st.sidebar.number_input("Rouge (jours)",0,60,opts["rouge"])
    opts["orange"]=st.sidebar.number_input("Orange (jours)",0,60,opts["orange"])
    opts["jaune"]=st.sidebar.number_input("Jaune (jours)",0,60,opts["jaune"])
    if st.sidebar.button("Enregistrer"):
        sauvegarder_options(opts)
        st.sidebar.success("Options enregistrées")

if admin:
    st.subheader("Ajouter un chantier")
    nom=st.text_input("Nom")
    ref=st.text_input("Référence")
    date=st.date_input("Date de montage")
    if st.button("Ajouter"):
        if nom and ref:
            df=pd.concat([df,pd.DataFrame([{"nom":nom,"ref":ref,"date":str(date)}])],ignore_index=True)
            sauvegarder_chantiers(df)
            st.rerun()
        else:
            st.error("Nom + Référence obligatoires.")

st.subheader("Liste des chantiers")
if df.empty:
    st.info("Aucun chantier.")
else:
    df_sorted=df.sort_values(by="date",na_position="last")
    st.dataframe(df_sorted.style.apply(lambda r:style_urgence(r,opts),axis=1),use_container_width=True)
