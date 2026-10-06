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
                "Komorów", "Michałowice", "Pruszków", "Ursus 1", "Ursus 2",
                "Nowa Wieś", "Komorów", "Pruszków", "Ursus 1", "Michałowice"
            ],
            "Poziomy": [
                "5, 6, 7, 8, Masters", "3-5 lat, 0, 1", "2, 4", "0, 1, 2, 3", "Masters",
                "5, 6, 7", "3-5 lat, 0, 1, 2, 3, 4", "7, 8", "1, 3, 5", "3-5 lat, Masters"
            ]
        })
        df_lektorzy.to_excel(writer, sheet_name="Lektorzy", index=False)
        
        # Arkusz 2: Sale (Generowane dynamicznie: 6 filii x 5 sal)
        filie = ["Komorów", "Michałowice", "Pruszków", "Ursus 1", "Ursus 2", "Nowa Wieś"]
        nazwy_sal = []
        filie_sal = []
        przeznaczenie = []
        
        for filia in filie:
            for i in range(1, 6):
                nazwy_sal.append(str(i))
                filie_sal.append(filia)
                # Przykładowe zróżnicowanie przeznaczenia sal w każdej filii
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
            "Nazwa Grupy": ["Lwy", "Tygrysy"],
            "Poziom": ["0", "1"],
            "Liczba Dzieci": [8, 10],
            "Skąd Odbiór": ["SP nr 1", "Przedszkole nr 5"],
            "Docelowa Filia": ["Komorów", "Michałowice"],
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
            "Docelowa Filia": ["Pruszków", "Ursus 1", "Nowa Wieś"],
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
            "Punkt Końcowy": ["Komorów", "Michałowice", "Pruszków", "Ursus 1"],
            "Czas Przejścia (min)": [15, 10, 20, 25]
        })
        df_trasy.to_excel(writer, sheet_name="Trasy", index=False)

        # Dopasowanie szerokości kolumn
        for sheetname in writer.sheets:
            worksheet = writer.sheets[sheetname]
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                worksheet.column_dimensions[column].width = (max_length + 3)
                
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
    if uploaded_file is not None:
    st.success("Plik wgrany poprawnie! Trwa analiza zapotrzebowania...")
    
    # Wczytywanie pliku Excel
    xls = pd.ExcelFile(uploaded_file)
    df_starsi = pd.read_excel(xls, sheet_name="Starsi Uczniowie")
    
    st.header("Krok 1: Weryfikacja i formowanie grup")
    
    col1, col2 = st.columns(2)
    with col1:
        bufor_dojazdu = st.slider("Bufor czasu na dojazd ze szkoły (minuty)", min_value=15, max_value=60, value=30, step=5)
    with col2:
        czas_trwania = st.slider("Czas trwania zajęć (minuty)", min_value=45, max_value=120, value=90, step=15)
        
    st.subheader("Analiza starszych uczniów")
    
    # Logika weryfikacji liczebności grup
    grupy_odrzucone = df_starsi[df_starsi["Liczba Chętnych"] < 5]
    grupy_ostrzezenie = df_starsi[df_starsi["Liczba Chętnych"] == 5]
    grupy_zatwierdzone = df_starsi[df_starsi["Liczba Chętnych"] >= 6]
    
    if not grupy_odrzucone.empty:
        st.error(f"Odrzucono {len(grupy_odrzucone)} potencjalnych grup z powodu braku wymaganej liczby chętnych (poniżej 5 osób).")
        st.dataframe(grupy_odrzucone, use_container_width=True)
        
    if not grupy_ostrzezenie.empty:
        st.warning(f"Ostrzeżenie: {len(grupy_ostrzezenie)} grup liczy dokładnie 5 osób. Zostały dodane do grafiku, ale wymagają uwagi.")
        st.dataframe(grupy_ostrzezenie, use_container_width=True)
        
    st.success(f"Zatwierdzono {len(grupy_zatwierdzone)} pełnych grup (6-12 osób) do zaplanowania.")
    st.dataframe(grupy_zatwierdzone, use_container_width=True)
    
    # Połączenie zatwierdzonych i tych z ostrzeżeniem jako naszej bazy do układania grafiku
    aktywne_grupy_starsi = pd.concat([grupy_zatwierdzone, grupy_ostrzezenie])
