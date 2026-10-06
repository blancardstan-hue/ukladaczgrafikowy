import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Układacz Grafików", layout="wide")

st.title("📅 Szkolny Układacz Grafików")
st.write("Witaj w aplikacji do generowania grafików! Zacznij od pobrania szablonu, wypełnij go swoimi danymi, a następnie wgraj poniżej.")

def generate_excel_template():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        
        # Arkusz 1: Lektorzy
        df_lektorzy = pd.DataFrame({
            "Lektor": [
                "Jan Kowalski", "Anna Nowak", "Marek Wiśniewski", "Katarzyna Wójcik", "Piotr Kamiński",
                "Agnieszka Lewandowska", "Michał Zieliński", "Ewa Szymańska", "Tomasz Woźniak", "Magdalena Dąbrowska"
            ],
            "Dostępność Pon": ["14:00-19:00", "", "15:00-20:00", "13:00-17:00", "", "16:00-20:00", "14:00-18:00", "", "15:00-19:00", "14:00-20:00"],
            "Dostępność Wt": ["14:00-19:00", "15:00-18:00", "", "13:00-17:00", "14:00-19:00", "", "14:00-18:00", "15:00-19:00", "", "14:00-20:00"],
            "Dostępność Śr": ["", "15:00-18:00", "15:00-20:00", "13:00-17:00", "14:00-19:00", "16:00-20:00", "", "15:00-19:00", "15:00-19:00", ""],
            "Dostępność Czw": ["14:00-19:00", "", "15:00-20:00", "", "14:00-19:00", "16:00-20:00", "14:00-18:00", "15:00-19:00", "", "14:00-20:00"],
            "Dostępność Pt": ["", "15:00-18:00", "", "13:00-17:00", "", "16:00-20:00", "14:00-18:00", "", "15:00-19:00", "14:00-18:00"],
            "
