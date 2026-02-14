import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import openpyxl
import os
import pandas as pd
from PIL import Image
import json

# ---------- CONFIG ----------
SCOPE = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive",
]

@st.cache_resource
def init_google_sheets():
    # Utilise le JSON stocké dans les secrets de Streamlit Cloud
    credentials_dict = json.loads(st.secrets["google"]["credentials"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, SCOPE)
    return gspread.authorize(creds)

client = init_google_sheets()
SHEET_NAME = "suivi des opérations"

# Données fixes
SERRES = ['B', 'C', 'D', 'E', 'F', 'G', 'H']
DELTAS = [str(i) for i in range(1, 33)]
CULTURES = ['tomate', 'pastèque', 'poivron', 'concombre', 'laitue', 'ciboulette', 'courgette', 'herbes aromatiques']
TRAITEMENTS = ['fongicide', 'insecticide', 'acaricide', 'insecticide/acaricide', 'raticide', 'bio-stimulant',
               'désinfectant', 'engrais foliaire']
SOLUTIONS_IRRI = ['AB', 'CD', 'M', 'Urée', 'enracineur', 'désinfectant']
ECS = ['1.6', '1.8', '2', '2.5', '3', '3.5', '4']

EXCEL_PRODUITS = "produits.xlsx"


# ---------- CRÉATION AUTOMATIQUE PRODUITS.XLSX ----------
def create_produits_excel():
    if not os.path.exists(EXCEL_PRODUITS):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Produits"
        ws.append(["Designation", "Dose", "Cible"])
        wb.save(EXCEL_PRODUITS)
        return True
    return False


create_produits_excel()


# ---------- FONCTIONS ----------
@st.cache_data
def charger_produits():
    wb = openpyxl.load_workbook(EXCEL_PRODUITS)
    ws = wb.active
    produits = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row and len(row) >= 3:
            produits.append({
                'nom': str(row[0]).strip(),
                'dose': str(row[1]).strip(),
                'cible': str(row[2]).strip(),
                'details': f"{row[0]} {row[1]} {row[2]}"
            })
    return produits


def ajouter_produit(designation, dose, cible):
    wb = openpyxl.load_workbook(EXCEL_PRODUITS)
    ws = wb.active
    ws.append([designation, dose, cible])
    wb.save(EXCEL_PRODUITS)
    st.cache_data.clear()


def get_or_create_sheet(client, serre_delta):
    sh = client.open(SHEET_NAME)
    try:
        return sh.worksheet(serre_delta)
    except gspread.WorksheetNotFound:
        sheet = sh.add_worksheet(title=serre_delta, rows=1000, cols=20)
        headers = ['Date', 'Serre', 'Delta', 'Culture', 'Operation', 'Details']
        sheet.append_row(headers)
        return sheet


def get_details_produits(selected_noms, produits):
    details_produits = []
    for nom in selected_noms:
        for p in produits:
            if p['nom'] == nom:
                details_produits.append(p['details'])
                break
    return details_produits


# ---------- INTERFACE STREAMLIT ----------
st.set_page_config(
    page_title="Suivi Opérations Pépinière",
    page_icon="🌱",
    layout="wide"
)

# Sidebar logo et navigation
with st.sidebar:
    st.title("🌱 Pépinière")
    try:
        logo = Image.open(r"C:\Users\hp\PycharmProjects\operations\logo.png")
        st.image(logo, width=120)
    except:
        st.markdown("### 🌱")

    st.subheader("📦 Produits")
    produits = charger_produits()

    with st.form("ajout_produit"):
        st.markdown("**Nouveau produit:**")
        des = st.text_input("**Designation**", placeholder="ex: Amistar")
        dose = st.text_input("**Dose**", placeholder="ex: 2ml/L")
        cible = st.text_input("**Cible**", placeholder="ex: pucerons")
        col_btn, col_clear = st.columns(2)
        with col_btn:
            submitted = st.form_submit_button("➕ **Ajouter**", use_container_width=True)
        with col_clear:
            if st.form_submit_button("🗑️ **Vider liste**", use_container_width=True):
                if os.path.exists(EXCEL_PRODUITS):
                    os.remove(EXCEL_PRODUITS)
                create_produits_excel()
                st.success("✅ Liste vidée!")
                st.rerun()

    if submitted and all([des, dose, cible]):
        ajouter_produit(des, dose, cible)
        st.success(f"✅ **{des}** ({dose}) → {cible} ajouté!")
        st.rerun()
    elif submitted:
        st.error("❌ Remplissez tous les champs!")

    st.markdown("**📋 Produits disponibles:**")
    if produits:
        for i, p in enumerate(produits, 1):
            st.write(f"{i}. **{p['nom']}** ({p['dose']}) → **{p['cible']}**")
    else:
        st.warning("⚠️ Aucun produit")

# Contenu principal
st.title("📊 Suivi Opérations Pépinière")
st.markdown("**Multi-Delta | Multi-Traitement | Multi-Produits**")

# Formulaire principal
col1, col2, col3 = st.columns(3)
with col1:
    serre = st.selectbox("**Serre:**", SERRES)
with col2:
    selected_deltas = st.multiselect("**Deltas:**", DELTAS, max_selections=10)
with col3:
    culture = st.selectbox("**Culture:**", CULTURES)

operation = st.selectbox("**Opération:**", ['traitement', 'irrigation'])

# 🔥 NOUVEAU : MULTI-TRAITEMENT + Multi-Produits
selected_noms = []
selected_traitements = []
solution = ""
ec = ""

if operation == 'traitement':
    # ✅ COLONNE 1 : Multi-Traitement
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        st.markdown("**🔥 Multi-Traitement**")
        selected_traitements = st.multiselect(
            "Catégories:", TRAITEMENTS,
            max_selections=4,
            help="Sélectionnez fongicide ET insecticide ET ..."
        )
    with col_t2:
        # ✅ COLONNE 2 : Multi-Produits (libre)
        st.markdown("**📦 Multi-Produits**")
        noms_produits = [p['nom'] for p in produits]
        selected_noms = st.multiselect(
            "Produits:", noms_produits,
            max_selections=8,
            help="Tous les produits disponibles"
        )

        if selected_traitements and selected_noms:
            st.caption(f"**{len(selected_traitements)} catégories** | **{len(selected_noms)} produits**")

elif operation == 'irrigation':
    col6, col7 = st.columns(2)
    with col6:
        solution = st.selectbox("**Solution:**", SOLUTIONS_IRRI)
    with col7:
        ec = st.selectbox("**EC:**", ECS)

# Aperçu en temps réel
with st.expander("👀 **Aperçu enregistrement**", expanded=True):
    if operation == 'traitement':
        if selected_traitements:
            details = "; ".join(selected_traitements)
        else:
            details = "Aucun traitement"

        if selected_noms:
            details_produits = get_details_produits(selected_noms, produits)
            if details_produits:
                details += f" - {'; '.join(details_produits)}"
    else:
        details = f"{solution} EC{ec}"

    st.info(f"""
    **🗓️ Date:** {datetime.now().strftime("%Y-%m-%d %H:%M")}
    **🏠 Serre:** {serre or '---'}  
    **🔢 Deltas:** {', '.join(selected_deltas) if selected_deltas else 'Aucun'}  
    **🌱 Culture:** {culture or '---'}  
    **⚙️ Opération:** {operation}  
    **📝 Détails:** {details}
    """)

# Bouton ENREGISTRER
if st.button("💾 **ENREGISTRER**", type="primary", use_container_width=True):
    if not all([serre, selected_deltas, culture]):
        st.error("❌ **Serre, Deltas et Culture OBLIGATOIRES!**")
    elif operation == 'traitement' and not selected_traitements:
        st.error("❌ **Sélectionnez au moins 1 traitement!**")
    else:
        date = datetime.now().strftime("%Y-%m-%d %H:%M")

        if operation == 'traitement':
            details = "; ".join(selected_traitements)
            if selected_noms:
                details_produits = get_details_produits(selected_noms, produits)
                if details_produits:
                    details += f" - {'; '.join(details_produits)}"
        else:
            details = f"{solution} EC{ec}"

        success_count = 0
        for delta in selected_deltas:
            try:
                sheet = get_or_create_sheet(client, f"{serre}{delta}")
                row = [date, serre, delta, culture, operation, details]
                sheet.append_row(row)
                success_count += 1
            except Exception as e:
                st.error(f"❌ Delta {delta}: {e}")

        if success_count > 0:
            st.success(f"✅ **{success_count}/{len(selected_deltas)}** enregistré(s)")
            st.balloons()
            st.rerun()

# Historique
if st.checkbox("📋 **Historique**"):
    try:
        if serre and selected_deltas:
            sh = client.open(SHEET_NAME)
            feuille = sh.worksheet(f"{serre}{selected_deltas[0]}")
            data = feuille.get_all_values()
            if len(data) > 1:
                df = pd.DataFrame(data[1:], columns=data[0])
                st.dataframe(df.tail(15), use_container_width=True, height=400)
            else:
                st.info("📭 Aucun enregistrement")
    except Exception as e:
        st.error(f"❌ Google Sheets: {e}")

st.markdown("---")
st.markdown("*Suivi Pépinière 🌱 | Multi-Traitement FINAL*")
