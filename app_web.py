import streamlit as st
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import os
import pandas as pd
from PIL import Image

# ---------- CONFIG ----------
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive",
]

SHEET_NAME = "suivi des opérations"  # Nom du fichier Google Sheet
EXCEL_PRODUITS = "produits.xlsx"    # Fichier Excel local
LOGO_FILE = "logo.png"              # Logo

SERRES = ['B', 'C', 'D', 'E', 'F', 'G', 'H']
DELTAS = [str(i) for i in range(1, 33)]
CULTURES = ['tomate', 'pastèque', 'poivron', 'concombre', 'laitue', 'ciboulette',
            'courgette', 'herbes aromatiques']
TRAITEMENTS = ['fongicide', 'insecticide', 'acaricide', 'insecticide/acaricide',
              'raticide', 'bio-stimulant', 'désinfectant', 'engrais foliaire']
SOLUTIONS_IRRI = ['AB', 'CD', 'M', 'Urée', 'enracineur', 'désinfectant']
ECS = ['1.6', '1.8', '2', '2.5', '3', '3.5', '4']

# ---------- LOGO ----------
logo_path = os.path.join(os.path.dirname(__file__), LOGO_FILE)
if os.path.exists(logo_path):
    logo = Image.open(logo_path)
    st.image(logo, width=200)
else:
    st.warning(f"Logo introuvable : {LOGO_FILE}")

# ---------- GOOGLE SHEETS ----------
@st.cache_resource
def init_google_sheets():
    credentials_dict = json.loads(st.secrets["google"]["credentials"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, SCOPE)
    return gspread.authorize(creds)

client = init_google_sheets()

# Vérifier que le fichier existe
try:
    sheet = client.open(SHEET_NAME)
except gspread.SpreadsheetNotFound:
    st.error(f"❌ Fichier Google Sheet '{SHEET_NAME}' introuvable ou accès refusé.")
    st.stop()

# Lister tous les onglets
worksheets = sheet.worksheets()
worksheet_titles = [ws.title for ws in worksheets]

# Sidebar : choisir onglet
st.sidebar.subheader("Sélection de l'onglet")
selected_worksheet = st.sidebar.selectbox("Choisir un onglet", worksheet_titles)

# ---------- AFFICHAGE DONNÉES SHEET ----------
worksheet = sheet.worksheet(selected_worksheet)
sheet_data = worksheet.get_all_records()
df_sheet = pd.DataFrame(sheet_data)

st.subheader(f"Données de l'onglet : {selected_worksheet}")

# Sidebar : filtrer par serre si colonne existe
if 'Serre' in df_sheet.columns:
    st.sidebar.subheader("Filtrer par Serre")
    selected_serre = st.sidebar.selectbox("Choisir une serre", ['Toutes'] + SERRES)
    if selected_serre != 'Toutes':
        df_sheet = df_sheet[df_sheet['Serre'] == selected_serre]

# ---------- COLORATION ----------
def color_culture(row):
    colors = {
        'tomate': 'background-color: #ffcccc',
        'pastèque': 'background-color: #ccffcc',
        'poivron': 'background-color: #ffffcc',
        'concombre': 'background-color: #ccffff',
        'laitue': 'background-color: #e6ccff',
        'ciboulette': 'background-color: #ffd9b3',
        'courgette': 'background-color: #f2f2f2',
        'herbes aromatiques': 'background-color: #ffe6cc'
    }
    if 'Culture' in row.index:
        return [colors.get(row['Culture'], '')]*len(row)
    else:
        return ['']*len(row)

def color_traitement(row):
    colors = {
        'fongicide': 'background-color: #ffcccc',
        'insecticide': 'background-color: #ccffcc',
        'acaricide': 'background-color: #ccccff',
        'raticide': 'background-color: #ffffcc',
        'bio-stimulant': 'background-color: #ccffff',
        'désinfectant': 'background-color: #e6ccff',
        'engrais foliaire': 'background-color: #ffd9b3'
    }
    if 'Traitement' in row.index:
        return [colors.get(row['Traitement'], '')]*len(row)
    else:
        return ['']*len(row)

# Sidebar : choix style
st.sidebar.subheader("Style du tableau")
style_option = st.sidebar.radio("Colorer par :", ['Aucune', 'Culture', 'Traitement'])

if style_option == 'Culture':
    styled_df = df_sheet.style.apply(color_culture, axis=1)
elif style_option == 'Traitement':
    styled_df = df_sheet.style.apply(color_traitement, axis=1)
else:
    styled_df = df_sheet

st.dataframe(styled_df)

# ---------- EXCEL PRODUITS ----------
excel_path = os.path.join(os.path.dirname(__file__), EXCEL_PRODUITS)
if os.path.exists(excel_path):
    df_produits = pd.read_excel(excel_path)
    st.subheader("📊 Produits disponibles")
    st.dataframe(df_produits)
else:
    st.warning(f"Fichier Excel introuvable : {EXCEL_PRODUITS}")
