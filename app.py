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
            "Filie": [
                "Filia Centrum", "Filia Północ", "Filia Centrum", "Filia Północ", "Filia Południe",
                "Filia Centrum", "Filia Południe", "Filia Północ", "Filia Centrum", "Filia Południe"
            ],
            "Poziomy": [
                "5, 6, 7, 8, Masters", "3-5 lat, 0, 1", "2, 4", "0, 1, 2, 3", "Masters",
                "5, 6, 7", "3-5 lat, 0, 1, 2, 3, 4", "7, 8", "1, 3, 5", "3-5 lat, Masters"
            ]
        })
        df_lektorzy.to_excel(writer, sheet_name="Lektorzy", index=False)
        
        # Arkusz 2: Sale
        df_sale = pd.DataFrame({
            "Nazwa Sali": ["Sala Żółta", "Sala Niebieska", "Sala Czerwona", "Sala Zielona"],
            "Filia": ["Filia Centrum", "Filia Centrum", "Filia Północ", "Filia Południe"],
            "Przeznaczenie": ["3-5 lat, 0, 1, 2, 3", "Wszystkie", "4, 5, 6, 7, 8, Masters", "Wszystkie"]
        })
        df_sale.to_excel(writer, sheet_name="Sale", index=False)
        
        # Arkusz 3: Młodsze Dzieci
        df_mlodsze = pd.DataFrame({
            "Nazwa Grupy": ["Lwy", "Tygrysy"],
            "Poziom": ["0", "1"],
            "Liczba Dzieci": [8, 10],
            "Skąd Odbiór": ["SP nr 1", "Przedszkole nr 5"],
            "Docelowa Filia": ["Filia Centrum", "Filia Centrum"],
            "Koniec Szkoły Pon": ["13:30", "14:00"],
            "Koniec Szkoły Wt": ["14:25", "14:00"],
            "Koniec Szkoły Śr": ["13:30", "14:00"],
            "Koniec Szkoły Czw": ["14:25", "14:00"],
            "Koniec Szkoły Pt": ["12:30", "14:00"],
            "Liczba Spotkań": [2, 2]
        })
        df_mlodsze.to_excel(writer, sheet_name="Młodsze Dzieci", index=False)
        
        # Arkusz 4: Starsi Uczniowie
        df_starsi = pd.DataFrame({
            "Szkoła i Klasa": ["SP nr 1 - Klasa 5A", "SP nr 2 - Klasa 6B", "SP nr 3 - Klasa 8C"],
            "Poziom": ["5", "6", "8"],
            "Liczba Chętnych": [6, 4, 8], 
            "Docelowa Filia": ["Filia Centrum", "Filia Północ", "Filia Południe"],
            "Koniec Lekcji Pon": ["14:25", "15:15", "16:05"],
            "Koniec Lekcji Wt": ["15:15", "14:25", "15:15"],
            "Koniec Lekcji Śr": ["14:25", "15:15", "16:05"],
            "Koniec Lekcji Czw": ["13:30", "14:25", "15:15"],
            "Koniec Lekcji Pt": ["14:25", "13:30", "14:25"]
        })
        df_starsi.to_excel(writer, sheet_name="Starsi Uczniowie", index=False)
        
        # Arkusz 5: Opiekunki
        df_opiekunki = pd.DataFrame({
            "Imię Opiekunki": ["Marta Wiśniewska", "Krystyna Kaczmarek", "Lucyna Mazur", "Natalia Piotrowska"],
            "Dostępność Pon": ["12:00-16:00", "13:00-17:00", "", "11:00-15:00"],
            "Dostępność Wt": ["12:00-16:00", "", "13:00-17:00", "11:00-15:00"],
            "Dostępność Śr": ["12:00-16:00", "13:00-17:00", "13:00-17:00", ""],
            "Dostępność Czw": ["12:00-16:00", "", "13:00-17:00", "11:00-15:00"],
            "Dostępność Pt": ["12:00-16:00", "13:00-17:00", "", "11:00-15:00"]
        })
        df_opiekunki.to_excel(writer, sheet_name="Opiekunki", index=False)
        
        # Arkusz 6: Trasy i Czas Dojazdów
        df_trasy = pd.DataFrame({
            "Punkt Początkowy": ["SP nr 1", "Przedszkole nr 5", "SP nr 2", "SP nr 3"],
            "Punkt Końcowy": ["Filia Centrum", "Filia Centrum", "Filia Północ", "Filia Południe"],
            "Czas Przejścia (min)": [15, 10, 20, 25]
        })
        df_trasy.to_excel(writer, sheet_name="Trasy", index=False)

        # Magia dostosowywania szerokości kolumn (wymaga openpyxl)
        for sheetname in writer.sheets:
            worksheet = writer.sheets[sheetname]
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter # Pobiera literę kolumny, np. 'A'
                for cell in col:
                    try:
                        # Sprawdzamy długość tekstu w komórce
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                # Ustawiamy szerokość na najdłuższy tekst + bufor dla estetyki
                adjusted_width = (max_length + 3)
                worksheet.column_dimensions[column].width = adjusted_width
                
    return output.getvalue()

st.subheader("1. Pobierz szablon")
excel_data = generate_excel_template()
st.download_button(
    label="📥 Pobierz szablon (Excel)",
    data=excel_data,
    file_name="ukladacz_grafikow.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.subheader("2. Wgraj uzupełniony plik")
uploaded_file = st.file_uploader("Wrzuć z powrotem wypełniony plik Excel", type=["xlsx"])

if uploaded_file is not None:
    st.success("Plik wgrany poprawnie! Wszystkie zakładki są gotowe do analizy.")
    # W przyszłości dodamy tu wczytywanie pd.read_excel(uploaded_file, sheet_name=None)
