import streamlit as st
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime
import os
from PIL import Image
import openpyxl

# ---------- CONFIG ----------
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive",
]

SHEET_NAME = "suivi des opérations"
EXCEL_PRODUITS = "produits.xlsx"
LOGO_FILE = "logo.png"

SERRES = ['B', 'C', 'D', 'E', 'F', 'G', 'H']
DELTAS = [str(i) for i in range(1, 33)]
CULTURES = ['tomate', 'pastèque', 'poivron', 'concombre', 'laitue', 'ciboulette', 'courgette', 'herbes aromatiques']
TRAITEMENTS = ['fongicide', 'insecticide', 'acaricide', 'insecticide/acaricide', 'raticide', 'bio-stimulant',
               'désinfectant', 'engrais foliaire']
SOLUTIONS_IRRI = ['AB', 'CD', 'M', 'Urée', 'enracineur', 'désinfectant']
ECS = ['1.6', '1.8', '2', '2.5', '3', '3.5', '4']

# ---------- LOGO ----------
logo_path = os.path.join(os.path.dirname(__file__), LOGO_FILE)
if os.path.exists(logo_path):
    st.image(Image.open(logo_path), width=200)
else:
    st.warning(f"Logo introuvable : {LOGO_FILE}")

# ---------- GOOGLE SHEETS ----------
@st.cache_resource
def init_google_sheets():
    credentials_dict = json.loads(st.secrets["google"]["credentials"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, SCOPE)
    return gspread.authorize(creds)

client = init_google_sheets()

try:
    sheet = client.open(SHEET_NAME)
except gspread.SpreadsheetNotFound:
    st.error(f"❌ Fichier Google Sheet '{SHEET_NAME}' introuvable ou accès refusé.")
    st.stop()

worksheets = sheet.worksheets()
worksheet_titles = [ws.title for ws in worksheets]

# Sidebar : choix de l’onglet
st.sidebar.subheader("Sélection de l'onglet")
selected_worksheet = st.sidebar.selectbox("Choisir un onglet", worksheet_titles)

worksheet = sheet.worksheet(selected_worksheet)
df_sheet = pd.DataFrame(worksheet.get_all_records())

# ---------- FILTRAGE ----------
st.sidebar.subheader("Filtrer les données")

if 'Serre' in df_sheet.columns:
    selected_serre = st.sidebar.selectbox("Serre", ['Toutes'] + SERRES)
    if selected_serre != 'Toutes':
        df_sheet = df_sheet[df_sheet['Serre'] == selected_serre]

if 'Delta' in df_sheet.columns:
    selected_delta = st.sidebar.selectbox("Delta", ['Tous'] + DELTAS)
    if selected_delta != 'Tous':
        df_sheet = df_sheet[df_sheet['Delta'] == selected_delta]

if 'Opération' in df_sheet.columns:
    selected_op = st.sidebar.selectbox("Opération", ['Toutes'] + df_sheet['Opération'].unique().tolist())
    if selected_op != 'Toutes':
        df_sheet = df_sheet[df_sheet['Opération'] == selected_op]

# ---------- COLORATION ----------
st.sidebar.subheader("Style du tableau")
style_option = st.sidebar.radio("Colorer par :", ['Aucune', 'Culture', 'Traitement'])

def color_culture(row):
    colors = {
        'tomate': '#ffcccc',
        'pastèque': '#ccffcc',
        'poivron': '#ffffcc',
        'concombre': '#ccffff',
        'laitue': '#e6ccff',
        'ciboulette': '#ffd9b3',
        'courgette': '#f2f2f2',
        'herbes aromatiques': '#ffe6cc'
    }
    if 'Culture' in row.index:
        return [f'background-color: {colors.get(row["Culture"], "")}']*len(row)
    return ['']*len(row)

def color_traitement(row):
    colors = {
        'fongicide': '#ffcccc',
        'insecticide': '#ccffcc',
        'acaricide': '#ccccff',
        'raticide': '#ffffcc',
        'bio-stimulant': '#ccffff',
        'désinfectant': '#e6ccff',
        'engrais foliaire': '#ffd9b3'
    }
    if 'Traitement' in row.index:
        return [f'background-color: {colors.get(row["Traitement"], "")}']*len(row)
    return ['']*len(row)

if style_option == 'Culture':
    styled_df = df_sheet.style.apply(color_culture, axis=1)
elif style_option == 'Traitement':
    styled_df = df_sheet.style.apply(color_traitement, axis=1)
else:
    styled_df = df_sheet

st.subheader(f"Données de l'onglet : {selected_worksheet}")
st.dataframe(styled_df)

# ---------- AJOUT / MODIFICATION PRODUITS ----------
st.subheader("Ajouter / Modifier un produit dans Excel")

excel_path = os.path.join(os.path.dirname(__file__), EXCEL_PRODUITS)
if os.path.exists(excel_path):
    df_produits = pd.read_excel(excel_path)
else:
    df_produits = pd.DataFrame(columns=['Produit', 'Quantité', 'Prix'])  # colonnes par défaut

with st.form("produit_form"):
    produit_name = st.text_input("Nom du produit")
    quantite = st.number_input("Quantité", min_value=0, step=1)
    prix = st.number_input("Prix", min_value=0.0, step=0.01)
    submitted = st.form_submit_button("Enregistrer")
    if submitted:
        new_row = {'Produit': produit_name, 'Quantité': quantite, 'Prix': prix}
        df_produits = pd.concat([df_produits, pd.DataFrame([new_row])], ignore_index=True)
        df_produits.to_excel(excel_path, index=False)
        st.success(f"Produit '{produit_name}' enregistré !")

st.subheader("📊 Produits disponibles")
st.dataframe(df_produits)
