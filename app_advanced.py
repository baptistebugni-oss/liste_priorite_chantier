import streamlit as st
import pandas as pd
import json
import os
from datetime import datetime, date, timedelta
from io import BytesIO

st.set_page_config(
    page_title="Gestion Projet Priorités",
    page_icon="https://raw.githubusercontent.com/baptistebugni-oss/liste_priorite_chantier/refs/heads/main/app_icon.ico",
)

# ==== PDF ====
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import qrcode

import base64
import requests


def safe_json_dump(data):
    """Empêche l’écriture d’un JSON vide, ce qui casserait GitHub."""
    if data is None:
        return "[]"
    if isinstance(data, list) and len(data) == 0:
        return "[]"
    try:
        return json.dumps(data, ensure_ascii=False, indent=4)
    except:
        return "[]"


# ==============================
# CONFIG
# ==============================

DATA_FILE = "chantiers.json"
OPTIONS_FILE = "options.json"
ADMIN_PASSWORD = "admin123"

DEFAULT_OPTIONS = {
    "rouge": 2,
    "orange": 7,
    "jaune": 14,
    "show_qr": True,
    "horizon": 60,
    "qr_url": "",
}

COLUMNS = ["nom", "ref", "date", "commentaire", "statut", "priorite"]

# ==============================
# COULEURS PDF
# ==============================

COLOR_URGENCE = {
    "rouge": "#FF4B4B",
    "orange": "#FFA500",
    "jaune": "#FFD966",
    "normal": "#DDDDDD",
}

COLOR_STATUT = {
    "Prévu": "#F5EEDC",
    "En cours": "#A066FF",
    "En attente": "#FF80B5",
    "Terminé": "#4CD964",
}


def draw_circle(c, x, y, size, color_hex):
    """Dessine une pastille ronde de couleur."""
    c.setFillColor(colors.HexColor(color_hex))
    c.circle(x, y, size, fill=1, stroke=0)


# ==============================
# OUTILS DONNÉES
# ==============================

def ensure_columns(df, cols):
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    return df


# ==============================
# GESTION SAUVEGARDE GITHUB
# ==============================

def safe_json_dump(data):
    """Empêche d'envoyer un fichier JSON vide sur GitHub."""
    if data is None:
        return "[]"
    if isinstance(data, list) and len(data) == 0:
        return "[]"
    try:
        return json.dumps(data, ensure_ascii=False, indent=4)
    except:
        return "[]"


def get_github_cfg():
    """Récupère la config GitHub depuis st.secrets. Retourne None si incomplet."""
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo = st.secrets["GITHUB_REPO"]
        branch = st.secrets.get("GITHUB_BRANCH", "main")
        return token, repo, branch
    except Exception:
        return None


def github_get_file_sha(path_repo):
    """Récupère le SHA du fichier sur GitHub (pour mise à jour). Retourne None si inexistant."""
    cfg = get_github_cfg()
    if cfg is None:
        return None

    token, repo, branch = cfg
    url = f"https://api.github.com/repos/{repo}/contents/{path_repo}"
    headers = {"Authorization": f"Bearer {token}"}

    try:
        r = requests.get(url, headers=headers, params={"ref": branch})
        if r.status_code == 200:
            data = r.json()
            return data.get("sha")
        else:
            return None
    except Exception as e:
        print("Erreur get_file_sha GitHub:", e)
        return None


def github_save_file(path_repo, content_str, message="Mise à jour via app Streamlit"):
    """Sauvegarde sécurisée : n'écrase JAMAIS un fichier par un contenu vide."""
    
    cfg = get_github_cfg()
    if cfg is None:
        print("⚠️ Pas de config GitHub → sauvegarde locale seulement.")
        return

    token, repo, branch = cfg

    # 1️⃣ Anti-corruption : jamais enregistrer un fichier vide
    if content_str.strip() == "":
        print("⛔ Refus d’envoyer un fichier vide sur GitHub !")
        return

    url = f"https://api.github.com/repos/{repo}/contents/{path_repo}"
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
    }

    # 2️⃣ Encodage en base64
    b64_content = base64.b64encode(content_str.encode("utf-8")).decode("utf-8")

    # 3️⃣ Récupération SHA actuel
    sha = github_get_file_sha(path_repo)

    payload = {
        "message": message,
        "content": b64_content,
        "branch": branch,
    }
    if sha:
        payload["sha"] = sha

    # 4️⃣ Appel API protégé
    try:
        r = requests.put(url, headers=headers, json=payload)
        if r.status_code in (200, 201):
            print(f"✔ GitHub sauvegardé : {path_repo}")
        else:
            print("❌ Erreur GitHub :", r.status_code, r.text)
    except Exception as e:
        print("❌ Exception GitHub :", e)


def github_fetch_file(path_repo):
    """Essaie de récupérer un fichier texte depuis GitHub. Retourne le texte ou None."""
    cfg = get_github_cfg()
    if cfg is None:
        return None

    token, repo, branch = cfg
    url = f"https://api.github.com/repos/{repo}/contents/{path_repo}"
    headers = {"Authorization": f"Bearer {token}"}

    try:
        r = requests.get(url, headers=headers, params={"ref": branch})
        if r.status_code == 200:
            data = r.json()
            content_b64 = data.get("content", "")
            return base64.b64decode(content_b64).decode("utf-8")
        else:
            return None
    except Exception as e:
        print("Erreur fetch_file GitHub:", e)
        return None


def charger_chantiers():
    """Charge le fichier JSON local ou depuis GitHub, puis renvoie un DataFrame propre."""
    # 1) Si fichier local existe -> on le lit normalement
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            data = []
    else:
        # 2) Sinon, on tente de récupérer depuis GitHub
        txt = github_fetch_file(DATA_FILE)
        if txt is not None:
            try:
                data = json.loads(txt)
                # On le recrée aussi en local
                with open(DATA_FILE, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
            except Exception:
                data = []
        else:
            data = []

    df = pd.DataFrame(data)
    # s'assurer qu'on a toutes les colonnes nécessaires
    for col in ["nom", "ref", "date", "commentaire", "statut", "priorite"]:
        if col not in df.columns:
            df[col] = ""

    if not df.empty and "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")

    return df


def sauvegarder_chantiers(df):
    """Sauvegarde en local + envoie sur GitHub (sécurisé)."""
    df_to_save = df.copy()

    # Convertir les dates en string
    if "date" in df_to_save.columns:
        df_to_save["date"] = df_to_save["date"].astype(str)

    data_list = df_to_save.to_dict(orient="records")

    # --- 1) Sauvegarde locale ---
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data_list, f, ensure_ascii=False, indent=4)

    # --- 2) Sauvegarde GitHub sécurisée ---
    json_str = safe_json_dump(data_list)

    # Ne JAMAIS envoyer un contenu vide à GitHub
    if json_str.strip() == "" or json_str.strip() == "[]":
        print("⛔ Annulé : GitHub ne recevra pas un fichier JSON vide.")
        return

    github_save_file(
        DATA_FILE,
        json_str,
        message="Mise à jour chantiers.json via app"
    )


def charger_options():
    """Charge les options (seuils, QR, etc.) depuis local ou GitHub."""
    if os.path.exists(OPTIONS_FILE):
        try:
            with open(OPTIONS_FILE, "r", encoding="utf-8") as f:
                opts = json.load(f)
        except Exception:
            opts = {}
    else:
        txt = github_fetch_file(OPTIONS_FILE)
        if txt is not None:
            try:
                opts = json.loads(txt)
                with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
                    json.dump(opts, f, ensure_ascii=False, indent=4)
            except Exception:
                opts = {}
        else:
            opts = {}

    # Valeurs par défaut si manquantes
    if "rouge" not in opts:
        opts["rouge"] = 2
    if "orange" not in opts:
        opts["orange"] = 7
    if "jaune" not in opts:
        opts["jaune"] = 14
    if "show_qr" not in opts:
        opts["show_qr"] = False
    if "qr_url" not in opts:
        opts["qr_url"] = ""
    if "horizon" not in opts:
        opts["horizon"] = 60

    return opts


def sauvegarder_options(opts):
    """Sauvegarde des options en local + GitHub."""
    # 1) Local
    with open(OPTIONS_FILE, "w", encoding="utf-8") as f:
        json.dump(opts, f, ensure_ascii=False, indent=4)

    # 2) GitHub
    try:
        json_str = json.dumps(opts, ensure_ascii=False, indent=4)
        github_save_file(OPTIONS_FILE, json_str, message="Mise à jour options.json via app")
    except Exception as e:
        print("Erreur sauvegarder_options -> GitHub:", e)


# ==============================
# URGENCE / STATUT LOGIQUE
# ==============================

def get_urgence_color(row, opts):
    d = row.get("date")
    if pd.isna(d):
        return "#DDDDDD"

    delta = (d.date() - date.today()).days
    if delta <= opts["rouge"]:
        return "#FF4B4B"
    if delta <= opts["orange"]:
        return "#FFA500"
    if delta <= opts["jaune"]:
        return "#FFD966"
    return "#DDDDDD"


def urgence_emoji(row, opts):
    d = row.get("date")
    if pd.isna(d):
        return "⚪"
    delta = (d.date() - date.today()).days
    if delta <= opts["rouge"]:
        return "🔴"
    if delta <= opts["orange"]:
        return "🟠"
    if delta <= opts["jaune"]:
        return "🟡"
    return "⚪"


def statut_emoji(statut):
    return {
        "Prévu": "⚪",
        "En cours": "🟣",
        "En attente": "🔵",
        "Terminé": "🟢",
    }.get(statut, "⚪")


# ==============================
# IMPORT EXCEL (Montage)
# ==============================

def importer_excel(file):
    df_x = pd.read_excel(file)
    df_x.columns = df_x.columns.str.strip().str.lower()

    required = ["nom de l'affaire", "code de l'affaire", "compétence", "début"]
    missing = [c for c in required if c not in df_x.columns]

    if missing:
        st.error("Colonnes manquantes : " + ", ".join(missing))
        return pd.DataFrame()

    df_m = df_x[df_x["compétence"].astype(str).str.lower() == "montage"]

    if df_m.empty:
        st.warning("Aucune ligne 'Montage'.")
        return pd.DataFrame()

    df_out = df_m[["nom de l'affaire", "code de l'affaire", "début"]].copy()
    df_out.columns = ["nom", "ref", "date"]
    df_out["date"] = pd.to_datetime(df_out["date"], errors="coerce")

    return df_out


# ==============================
# EXPORT EXCEL
# ==============================

def build_excel(df):
    buf = BytesIO()
    df2 = df.copy()
    if not df2.empty:
        df2["date"] = df2["date"].dt.strftime("%d/%m/%Y")
    df2.to_excel(buf, index=False)
    buf.seek(0)
    return buf


# ==============================
# UTILITAIRE TEXTE
# ==============================

def truncate_nom(text, length=30):
    text = str(text)
    return text if len(text) <= length else text[:length] + "…"


# ==============================
# EXPORT PDF (SIMPLE, LISIBLE)
# ==============================

def build_pdf_liste(df, opts, qr_url=None):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)

    page_w, page_h = A4
    margin_left = 30

    # =========================
    # EN-TÊTE
    # =========================
    c.setFillColor(colors.HexColor("#2F3C7E"))
    c.rect(0, page_h - 70, page_w, 70, fill=1)

    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(page_w / 2, page_h - 38, "Gestion Projet Priorités")

    c.setFont("Helvetica", 10)
    c.drawString(20, page_h - 60, f"Export du {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    y = page_h - 110

    # =========================
    # LÉGENDE
    # =========================
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin_left, y, "Légende :")
    y -= 15

    legend_items = [
        ("#FF4B4B", f"Urgence forte (≤ {opts['rouge']} j)"),
        ("#FFA500", f"Urgence moyenne (≤ {opts['orange']} j)"),
        ("#FFD966", f"Approche (≤ {opts['jaune']} j)"),
        ("#F5EEDC", "Prévu"),
        ("#A066FF", "En cours"),
        ("#FF80B5", "En attente"),
        ("#4CD964", "Terminé"),
    ]

    for color_hex, label in legend_items:
        draw_circle(c, margin_left + 4, y - 3, 4, color_hex)
        c.setFillColor(colors.black)
        c.drawString(margin_left + 15, y - 6, label)
        y -= 15

    y -= 20

    # ===============================
    # CALCUL AUTOMATIQUE DES COLONNES
    # ===============================
    headers = ["", "Nom de l'affaire", "Code", "Date", "Statut", ""]

    columns = [
        [""] * len(df),
        df["nom"].astype(str).tolist(),
        df["ref"].astype(str).tolist(),
        df["date"].dt.strftime("%d/%m/%Y").fillna("-").tolist() if not df.empty else [],
        df["statut"].astype(str).tolist(),
        [""] * len(df),
    ]

    c.setFont("Helvetica", 10)
    max_widths = []

    if len(df) > 0:
        for header, col in zip(headers, columns):
            header_w = c.stringWidth(header, "Helvetica-Bold", 10)
            cell_w = max(c.stringWidth(str(x), "Helvetica", 9) for x in col)
            max_widths.append(max(header_w, cell_w) + 12)
    else:
        max_widths = [22, 190, 60, 60, 70, 22]

    total_target = 22 + 190 + 60 + 60 + 70 + 22
    sum_widths = sum(max_widths)

    if sum_widths > total_target:
        ratio = total_target / sum_widths
        col_widths = [w * ratio for w in max_widths]
    else:
        col_widths = max_widths

    # =========================
    # TABLEAU
    # =========================
    def draw_header(ypos):
        c.setFillColor(colors.HexColor("#EFEFEF"))
        c.rect(margin_left, ypos - 18, sum(col_widths), 18, fill=1)

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 10)

        xx = margin_left + 5
        for h, w in zip(headers, col_widths):
            c.drawString(xx, ypos - 12, h)
            xx += w

        return ypos - 35

    y = draw_header(y)
    c.setFont("Helvetica", 9)

    df_sorted = df.sort_values("date")

    for _, row in df_sorted.iterrows():

        if y < 80:
            c.showPage()
            y = page_h - 60
            y = draw_header(y)
            c.setFont("Helvetica", 9)

        # Pastille urgence
        urg = get_urgence_color(row, opts)
        draw_circle(
            c,
            margin_left + col_widths[0] / 2,
            y + 4,
            4,
            urg,
        )

        # Pastille statut
        stat = COLOR_STATUT.get(row["statut"], "#F5EEDC")
        draw_circle(
            c,
            margin_left + sum(col_widths) - col_widths[-1] / 2,
            y + 5,
            4,
            stat,
        )

        # Texte
        c.setFillColor(colors.black)
        dd = "-" if pd.isna(row["date"]) else row["date"].strftime("%d/%m/%Y")

        values = ["", row["nom"], row["ref"], dd, row["statut"], ""]

        xx = margin_left + 5
        for val, w in zip(values, col_widths):
            c.drawString(xx, y, str(val))
            xx += w

        y -= 18

    # ===============================
    # SÉPARATEUR
    # ===============================
    c.setStrokeColor(colors.HexColor("#BBBBBB"))
    c.setLineWidth(1)
    c.line(margin_left, y + 12, page_w - margin_left, y + 12)
    y -= 35

    # ===============================
    # TITRE : VUE PAR SEMAINE
    # ===============================
    c.setFont("Helvetica-Bold", 14)
    c.setFillColor(colors.black)
    c.drawString(margin_left, y, "📅 Vue par semaine")
    y -= 25

    # ===============================
    # CONTENU : VUE PAR SEMAINE
    # ===============================
    df_week = df[df["date"].notna()].copy()

    if not df_week.empty:

        iso = df_week["date"].dt.isocalendar()
        df_week["week"] = iso.week
        df_week["year"] = iso.year
        df_week = df_week.sort_values("date")

        mois_fr = {
            1: "janvier", 2: "février", 3: "mars", 4: "avril",
            5: "mai", 6: "juin", 7: "juillet", 8: "août",
            9: "septembre", 10: "octobre", 11: "novembre", 12: "décembre"
        }

        for (year, week), grp in df_week.groupby(["year", "week"]):

            if y < 120:
                c.showPage()
                y = page_h - 60
                c.setFont("Helvetica-Bold", 14)
                c.setFillColor(colors.black)
                c.drawString(margin_left, y, "📅 Vue par semaine")
                y -= 25

            start = date.fromisocalendar(year, week, 1)
            end = date.fromisocalendar(year, week, 7)

            if start.month == end.month:
                peri = f"du {start.day} au {end.day} {mois_fr[start.month]} {year}"
            else:
                peri = (
                    f"du {start.day} {mois_fr[start.month]} "
                    f"au {end.day} {mois_fr[end.month]} {year}"
                )

            c.setFont("Helvetica-Bold", 12)
            c.setFillColor(colors.black)
            c.drawString(margin_left, y, f"Semaine {week:02d} — {peri}")
            y -= 18

            for _, row in grp.iterrows():

                if y < 100:
                    c.showPage()
                    y = page_h - 60

                urg = get_urgence_color(row, opts)
                draw_circle(c, margin_left + 5, y + 2, 4, urg)

                c.setFillColor(colors.black)
                c.setFont("Helvetica", 10)

                # ======= DATE FRANÇAISE LONGUE =======
                jours_fr = {
                    0: "Lundi",
                    1: "Mardi",
                    2: "Mercredi",
                    3: "Jeudi",
                    4: "Vendredi",
                    5: "Samedi",
                    6: "Dimanche"
                }

                mois_fr_long = {
                    1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril",
                    5: "Mai", 6: "Juin", 7: "Juillet", 8: "Août",
                    9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre"
                }

                d = row["date"]

                jour_txt = jours_fr[d.weekday()]
                mois_txt = mois_fr_long[d.month]

                date_longue = f"{jour_txt} {d.day:02d} {mois_txt}"

                txt = f"{date_longue} — {row['nom']} ({row['ref']})"

                max_text_width = 300 - 40
                text_width = c.stringWidth(txt, "Helvetica", 10)

                if text_width > max_text_width:
                    cutoff = len(txt)
                    while c.stringWidth(txt[:cutoff] + "…", "Helvetica", 10) > max_text_width and cutoff > 10:
                        cutoff -= 1
                    txt_to_print = txt[:cutoff] + "…"
                else:
                    txt_to_print = txt

                c.drawString(margin_left + 18, y, txt_to_print)

                stat = COLOR_STATUT.get(row["statut"], "#F5EEDC")
                draw_circle(c, margin_left + 300, y + 2, 4, stat)

                y -= 16

    else:
        c.setFont("Helvetica", 10)
        c.setFillColor(colors.black)
        c.drawString(margin_left, y, "Aucune date disponible pour la vue par semaine.")
        y -= 20

    # ===============================
    # QR-CODE (si activé)
    # ===============================
    if qr_url is not None and isinstance(qr_url, str) and qr_url.strip() != "":
        try:
            qr_img = qrcode.make(qr_url)
            qr_buf = BytesIO()
            qr_img.save(qr_buf, format="PNG")
            qr_buf.seek(0)

            qr_reader = ImageReader(qr_buf)

            qr_size = 28 * mm
            qr_x = page_w - qr_size - 25
            qr_y = 20

            c.drawImage(
                qr_reader,
                qr_x,
                qr_y,
                width=qr_size,
                height=qr_size,
                preserveAspectRatio=True,
                mask='auto'
            )

            c.setFont("Helvetica", 8)
            c.setFillColor(colors.black)
            c.drawCentredString(qr_x + qr_size / 2, qr_y - 10, "Accès mobile")

        except Exception as e:
            print("Erreur QR-code PDF:", e)

    c.showPage()
    c.save()
    buf.seek(0)
    return buf


# =======================================================
# INTERFACE STREAMLIT — CONFIG DE BASE
# =======================================================

st.title("📋 Gestion des priorités chantier")

mode = st.sidebar.selectbox("Mode", ["Lecture seule", "Administrateur"])
is_admin = False

if mode == "Administrateur":
    mdp = st.sidebar.text_input("Mot de passe", type="password")
    if mdp == ADMIN_PASSWORD:
        is_admin = True
    else:
        st.sidebar.error("Mot de passe incorrect")


# ========================================
# TEST CONFIG GITHUB (sécurisé)
# ========================================
if is_admin:
    st.subheader("🔧 Test configuration GitHub")

    try:
        cfg = get_github_cfg()
        st.write("Configuration détectée :", cfg)

        sha_test = github_get_file_sha("chantiers.json")
        st.write("SHA détecté pour chantiers.json :", sha_test)
    except Exception as e:
        st.error(f"Erreur lors du test GitHub : {e}")

df = charger_chantiers()
opts = charger_options()

# =============================
# FILTRES
# =============================

with st.expander("🔎 Filtres de recherche", expanded=False):
    col1, col2, col3 = st.columns(3)

    f_nom = col1.text_input("Rechercher un nom")
    f_ref = col2.text_input("Rechercher un code")
    f_statut = col3.multiselect(
        "Filtrer par statut",
        ["Prévu", "En cours", "En attente", "Terminé"],
    )

df_filtered = df.copy()

if f_nom:
    df_filtered = df_filtered[df_filtered["nom"].str.contains(f_nom, case=False, na=False)]
if f_ref:
    df_filtered = df_filtered[df_filtered["ref"].str.contains(f_ref, case=False, na=False)]
if f_statut:
    df_filtered = df_filtered[df_filtered["statut"].isin(f_statut)]

# =============================
# TABLEAU PRINCIPAL
# =============================

st.subheader("📌 Liste des chantiers")

if df_filtered.empty:
    st.info("Aucun chantier.")
else:
    disp = df_filtered.copy()
    disp["Urgence"] = disp.apply(lambda r: urgence_emoji(r, opts), axis=1)
    disp["État"] = disp["statut"].apply(statut_emoji)

    if not disp.empty:
        disp["date"] = disp["date"].dt.strftime("%d/%m/%Y")

    disp = disp[["Urgence", "nom", "ref", "date", "statut", "État", "commentaire"]]

    disp.rename(columns={
        "nom": "Nom de l'affaire",
        "ref": "Code affaire",
        "statut": "Statut",
        "commentaire": "Commentaire",
    }, inplace=True)

    st.dataframe(disp, use_container_width=True)

# =============================
# 📅 VUE PAR SEMAINE (TIMELINE)
# =============================

with st.expander("📅 Vue par semaine (ligne du temps)", expanded=False):

    df_week = df.copy()
    df_week = df_week[df_week["date"].notna()]

    if df_week.empty:
        st.info("Aucun chantier avec une date valide.")
    else:
        horizon = opts.get("horizon", 60)
        today = date.today()
        end_date = today + timedelta(days=horizon)

        df_week = df_week[
            (df_week["date"].dt.date >= today) &
            (df_week["date"].dt.date <= end_date)
        ]

        if df_week.empty:
            st.warning("Aucun chantier dans l’horizon sélectionné.")
        else:
            iso = df_week["date"].dt.isocalendar()
            df_week["week"] = iso.week
            df_week["year"] = iso.year

            df_week = df_week.sort_values("date")

            mois_fr = {
                1: "janvier", 2: "février", 3: "mars", 4: "avril",
                5: "mai", 6: "juin", 7: "juillet", 8: "août",
                9: "septembre", 10: "octobre", 11: "novembre", 12: "décembre"
            }

            for (year, week), g in df_week.groupby(["year", "week"]):

                start = date.fromisocalendar(int(year), int(week), 1)
                endw = date.fromisocalendar(int(year), int(week), 7)

                if start < today:
                    start = today
                if endw > end_date:
                    endw = end_date

                if start.month == endw.month:
                    peri = f"du {start.day} au {endw.day} {mois_fr[start.month]} {year}"
                else:
                    peri = (
                        f"du {start.day} {mois_fr[start.month]} "
                        f"au {endw.day} {mois_fr[endw.month]} {year}"
                    )

                st.markdown(f"### 🗓️ Semaine {int(week):02d} — {peri}")

                for _, row in g.iterrows():
                    urg = urgence_emoji(row, opts)
                    stat = statut_emoji(row["statut"])
                    dstr = row["date"].strftime("%d/%m")
                    nom = truncate_nom(row["nom"], 40)

                    c1, c2, c3, c4 = st.columns([1, 2, 6, 3])
                    c1.write(urg)
                    c2.write(dstr)
                    c3.write(f"{nom} ({row['ref']})")
                    c4.write(f"{stat} {row['statut']}")

# =======================================================
# ADMINISTRATION
# =======================================================

if is_admin:

    # IMPORT EXCEL
    with st.expander("📥 Import Excel — Sélection unitaire", expanded=False):
        file = st.file_uploader("Importer un fichier Excel", type=["xlsx"], key="excel_file")

        if file:
            imported = importer_excel(file)

            if not imported.empty:
                st.success(f"{len(imported)} chantiers trouvés (Montage)")

                names = imported["nom"].fillna("Sans nom").tolist()

                idx_choice = st.selectbox(
                    "Sélectionner un chantier",
                    range(len(imported)),
                    format_func=lambda i: names[i],
                    key="excel_select",
                )

                row_imp = imported.iloc[idx_choice]

                st.info(
                    f"**Nom :** {row_imp['nom']}\n\n"
                    f"**Code :** {row_imp['ref']}\n\n"
                    f"**Date :** "
                    f"{row_imp['date'].strftime('%d/%m/%Y') if not pd.isna(row_imp['date']) else '-'}"
                )

                if st.button("➕ Ajouter ce chantier", key="excel_add"):
                    new_row = {
                        "nom": row_imp["nom"],
                        "ref": row_imp["ref"],
                        "date": row_imp["date"],
                        "commentaire": "",
                        "statut": "Prévu",
                        "priorite": ""
                    }

                    for col in df.columns:
                        if col not in new_row:
                            new_row[col] = ""

                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                    sauvegarder_chantiers(df)
                    st.success("Chantier ajouté !")
                    st.rerun()

    # AJOUT MANUEL
    with st.expander("➕ Ajouter manuellement un chantier", expanded=False):

        n_nom = st.text_input("Nom", key="m_nom")
        n_ref = st.text_input("Code", key="m_ref")
        n_date = st.date_input("Date", key="m_date")
        n_statut = st.selectbox(
            "Statut",
            ["Prévu", "En cours", "En attente", "Terminé"],
            key="m_statut",
        )
        n_comment = st.text_area("Commentaire", key="m_comm")

        if st.button("Ajouter chantier", key="m_add"):
            new_row = {
                "nom": n_nom,
                "ref": n_ref,
                "date": pd.to_datetime(n_date),
                "commentaire": n_comment,
                "statut": n_statut,
                "priorite": ""
            }

            for col in df.columns:
                if col not in new_row:
                    new_row[col] = ""

            df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            sauvegarder_chantiers(df)

            st.success("Chantier ajouté")
            st.rerun()

    # MODIFIER / SUPPRIMER
    with st.expander("✏️ Modifier / Supprimer un chantier", expanded=False):
        if df.empty:
            st.info("Aucun chantier.")
        else:
            idx = st.selectbox(
                "Sélectionner",
                df.index,
                format_func=lambda i: f"{df.loc[i, 'nom']} - {df.loc[i, 'ref']}",
                key="edit_select",
            )

            row_ed = df.loc[idx]

            e_nom = st.text_input("Nom", row_ed["nom"], key=f"e_nom_{idx}")
            e_ref = st.text_input("Code", row_ed["ref"], key=f"e_ref_{idx}")
            e_date = st.date_input(
                "Date",
                row_ed["date"].date() if not pd.isna(row_ed["date"]) else date.today(),
                key=f"e_date_{idx}",
            )
            e_statut = st.selectbox(
                "Statut",
                ["Prévu", "En cours", "En attente", "Terminé"],
                index=["Prévu", "En cours", "En attente", "Terminé"].index(row_ed["statut"]),
                key=f"e_statut_{idx}",
            )
            e_comment = st.text_area("Commentaire", row_ed["commentaire"], key=f"e_comment_{idx}")

            c1, c2 = st.columns(2)

            if c1.button("💾 Sauvegarder", key=f"save_{idx}"):
                # Mise à jour ciblée des colonnes existantes
                df.loc[idx, "nom"] = e_nom
                df.loc[idx, "ref"] = e_ref
                df.loc[idx, "date"] = pd.to_datetime(e_date)
                df.loc[idx, "statut"] = e_statut
                df.loc[idx, "commentaire"] = e_comment

                sauvegarder_chantiers(df)
                st.success("Modifié !")
                st.rerun()

            if c2.button("🗑 Supprimer", key=f"del_{idx}"):
                df = df.drop(idx).reset_index(drop=True)
                sauvegarder_chantiers(df)
                st.success("Supprimé !")
                st.rerun()

    # OPTIONS
    with st.expander("⚙️ Options générales", expanded=False):

        o1 = st.number_input("Urgence rouge ≤ jours", value=opts["rouge"], min_value=0)
        o2 = st.number_input("Urgence orange ≤ jours", value=opts["orange"], min_value=0)
        o3 = st.number_input("Urgence jaune ≤ jours", value=opts["jaune"], min_value=0)
        o4 = st.checkbox("Afficher QR dans le PDF de liste", value=opts["show_qr"])
        o5 = st.number_input(
            "Horizon vue 'par semaine' (en jours)",
            value=opts.get("horizon", 60),
            min_value=7,
            max_value=365,
        )
        o6 = st.text_input(
            "URL à encoder dans le QR-code (PDF liste)",
            value=opts.get("qr_url", ""),
        )

        if st.button("Sauvegarder options", key="opt_save"):
            opts["rouge"] = o1
            opts["orange"] = o2
            opts["jaune"] = o3
            opts["show_qr"] = o4
            opts["horizon"] = o5
            opts["qr_url"] = o6
            sauvegarder_options(opts)
            st.success("Options enregistrées !")
            st.rerun()

# =======================================================
# EXPORTS (PDF + EXCEL)
# =======================================================

st.markdown("---")

col1, col2 = st.columns(2)

with col1:
    if st.button("📄 Export PDF (liste)", key="pdf_liste_btn"):
        qr_url_for_pdf = opts.get("qr_url") if opts.get("show_qr") and opts.get("qr_url") else None
        pdf = build_pdf_liste(df, opts, qr_url_for_pdf)
        st.download_button(
            "Télécharger PDF liste",
            data=pdf,
            file_name=f"Gestion_Projet_Priorites_{date.today()}.pdf",
            mime="application/pdf",
        )

with col2:
    if st.button("📊 Export Excel", key="xlsx_btn"):
        excel = build_excel(df)
        st.download_button(
            "Télécharger Excel",
            data=excel,
            file_name=f"Chantiers_{date.today()}.xlsx",
            mime="application/vnd.ms-excel",
        )
