import streamlit as st
import pandas as pd
import random
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
            "Dostępność Pon": [
                "14:00-19:00", "", "15:00-20:00", "13:00-17:00", "", 
                "16:00-20:00", "14:00-18:00", "", "15:00-19:00", "14:00-20:00"
            ],
            "Dostępność Wt": [
                "14:00-19:00", "15:00-18:00", "", "13:00-17:00", "14:00-19:00", 
                "", "14:00-18:00", "15:00-19:00", "", "14:00-20:00"
            ],
            "Dostępność Śr": [
                "", "15:00-18:00", "15:00-20:00", "13:00-17:00", "14:00-19:00", 
                "16:00-20:00", "", "15:00-19:00", "15:00-19:00", ""
            ],
            "Dostępność Czw": [
                "14:00-19:00", "", "15:00-20:00", "", "14:00-19:00", 
                "16:00-20:00", "14:00-18:00", "15:00-19:00", "", "14:00-20:00"
            ],
            "Dostępność Pt": [
                "", "15:00-18:00", "", "13:00-17:00", "", 
                "16:00-20:00", "14:00-18:00", "", "15:00-19:00", "14:00-18:00"
            ],
            "Filie": [
                "Komorów", "Michałowice", "Pruszków", "Ursus 1", "Ursus 2",
                "Nowa Wieś", "Komorów", "Pruszków", "Ursus 1", "Michałowice"
            ],
            "Poziomy": [
                "5, 6, 7, 8, Masters", "3-5 lat, 0, 1", "2, 4", "0, 1, 2, 3", "Masters",
                "5, 6, 7", "3-5 lat, 0, 1, 2, 3, 4", "7, 8", "1, 3, 5", "3-5 lat, Masters"
            ]
        })
        df_lektorzy.to_excel(writer, sheet_name="Lektorzy", index=False)
        
        # Arkusz 2: Sale
        filie = ["Komorów", "Michałowice", "Pruszków", "Ursus 1", "Ursus 2", "Nowa Wieś"]
        nazwy_sal = []
        filie_sal = []
        przeznaczenie = []
        
        for filia in filie:
            for i in range(1, 6):
                nazwy_sal.append(str(i))
                filie_sal.append(filia)
                if i == 1:
                    przeznaczenie.append("3-5 lat, 0, 1, 2, 3")
                elif i == 2:
                    przeznaczenie.append("4, 5, 6, 7, 8, Masters")
                else:
                    przeznaczenie.append("Wszystkie")

        df_sale = pd.DataFrame({
            "Nazwa Sali": nazwy_sal,
            "Filia": filie_sal,
            "Przeznaczenie": przeznaczenie
        })
        df_sale.to_excel(writer, sheet_name="Sale", index=False)
        
        # Arkusz 3: Młodsze Dzieci
        df_mlodsze = pd.DataFrame({
            "Nazwa Grupy": ["Lwy", "Tygrysy", "Żyrafy", "Misie"],
            "Po
