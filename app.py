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
        pd.DataFrame({
            "Lektor": ["Jan Kowalski", "Anna Nowak"],
            "Dostępność Pon": ["14:00-19:00", ""],
            "Dostępność Wt": ["14:00-19:00", "15:00-18:00"],
            "Dostępność Śr": ["", "15:00-18:00"],
            "Dostępność Czw": ["14:00-19:00", ""],
            "Dostępność Pt": ["", "15:00-18:00"],
            "Filie": ["Filia Centrum", "Filia Północ"],
            "Poziomy": ["1-3, Starsi", "Przedszkole, 1-3"]
        }).to_excel(writer, sheet_name="Lektorzy", index=False)
        
        # Arkusz 2: Sale
        pd.DataFrame({
            "Nazwa Sali": ["Sala Żółta", "Sala Niebieska", "Sala Czerwona"],
            "Filia": ["Filia Centrum", "Filia Centrum", "Filia Północ"],
            "Przeznaczenie": ["Przedszkole, 1-3", "Wszystkie", "Tylko Starsi"]
        }).to_excel(writer, sheet_name="Sale", index=False)
        
        # Arkusz 3: Młodsze Dzieci
        pd.DataFrame({
            "Nazwa Grupy": ["Lwy", "Tygrysy"],
            "Liczba Dzieci": [8, 10],
            "Skąd Odbiór": ["SP nr 1", "Przedszkole nr 5"],
            "Docelowa Filia": ["Filia Centrum", "Filia Centrum"],
            "Koniec Szkoły Pon": ["13:30", "14:00"],
            "Koniec Szkoły Wt": ["14:25", "14:00"],
            "Koniec Szkoły Śr": ["13:30", "14:00"],
            "Koniec Szkoły Czw": ["14:25", "14:00"],
            "Koniec Szkoły Pt": ["12:30", "14:00"],
            "Liczba Spotkań": [2, 2]
        }).to_excel(writer, sheet_name="Młodsze Dzieci", index=False)
        
        # Arkusz 4: Starsi Uczniowie
        pd.DataFrame({
            "Szkoła i Klasa": ["SP nr 1 - Klasa 5A", "SP nr 2 - Klasa 6B"],
            "Liczba Chętnych": [6, 4], 
            "Docelowa Filia": ["Filia Centrum", "Filia Północ"],
            "Koniec Lekcji Pon": ["14:25", "15:15"],
            "Koniec Lekcji Wt": ["15:15", "14:25"],
            "Koniec Lekcji Śr": ["14:25", "15:15"],
            "Koniec Lekcji Czw": ["13:30", "14:25"],
            "Koniec Lekcji Pt": ["14:25", "13:30"]
        }).to_excel(writer, sheet_name="Starsi Uczniowie", index=False)
        
        # Arkusz 5: Opiekunki
        pd.DataFrame({
            "Imię Opiekunki": ["Marta Wiśniewska"],
            "Dostępność Pon": ["12:00-16:00"],
            "Dostępność Wt": ["12:00-16:00"],
            "Dostępność Śr": ["12:00-16:00"],
            "Dostępność Czw": ["12:00-16:00"],
            "Dostępność Pt": ["12:00-16:00"]
        }).to_excel(writer, sheet_name="Opiekunki", index=False)
        
        # Arkusz 6: Trasy i Czas Dojazdów
        pd.DataFrame({
            "Punkt Początkowy": ["SP nr 1", "Przedszkole nr 5"],
            "Punkt Końcowy": ["Filia Centrum", "Filia Centrum"],
            "Czas Przejścia (min)": [15, 10]
        }).to_excel(writer, sheet_name="Trasy", index=False)
        
    return output.getvalue()

st.subheader("1. Pobierz szablon")
excel_data = generate_excel_template()
st.download_button(
    label="📥 Pobierz układacz grafików (Excel)",
    data=excel_data,
    file_name="ukladacz_grafikow.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.subheader("2. Wgraj uzupełniony plik")
uploaded_file = st.file_uploader("Wrzuć z powrotem wypełniony plik Excel", type=["xlsx"])

if uploaded_file is not None:
    st.success("Plik wgrany poprawnie! Tutaj w przyszłości pojawi się wygenerowany grafik.")
    # Tu będziemy wstawiać logikę czytania danych i układania zajęć
