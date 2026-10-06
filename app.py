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
            "Poziom": ["0", "1", "3-5 lat", "2"],
            "Liczba Dzieci": [8, 10, 6, 12],
            "Czas trwania (min)": [60, 60, 60, 60],
            "Skąd Odbiór": ["SP nr 1", "Przedszkole nr 5", "Przedszkole nr 2", "SP nr 2"],
            "Docelowa Filia": ["Komorów", "Michałowice", "Pruszków", "Ursus 1"],
            "Koniec Szkoły Pon": ["13:30", "14:00", "12:30", "13:30"],
            "Koniec Szkoły Wt": ["14:25", "14:00", "12:30", "14:25"],
            "Koniec Szkoły Śr": ["13:30", "14:00", "12:30", "13:30"],
            "Koniec Szkoły Czw": ["14:25", "14:00", "12:30", "14:25"],
            "Koniec Szkoły Pt": ["12:30", "14:00", "12:30", "12:30"],
            "Liczba Spotkań": [2, 2, 2, 1]
        })
        df_mlodsze.to_excel(writer, sheet_name="Młodsze Dzieci", index=False)
        
        # Arkusz 4: Starsi Uczniowie - GENEROWANIE LOSOWEJ BAZY TESTOWEJ
        starsi_data = []
        godziny = ["13:30", "14:25", "15:15", "16:05", "16:55"]
        
        # Generowanie podstawówki (SP nr 1, klasy 1-8)
        for klasa_num in range(1, 9):
            poziom = str(klasa_num)
            liczba_oddzialow = random.randint(4, 6) # Generuje klasy np. od A do D, E lub F
            for litera in ['A', 'B', 'C', 'D', 'E', 'F'][:liczba_oddzialow]:
                starsi_data.append({
                    "Szkoła i Klasa": f"SP nr 1 - Klasa {klasa_num}{litera}",
                    "Poziom": poziom,
                    "Liczba Chętnych": random.randint(0, 10),
                    "Czas trwania (min)": 90,
                    "Docelowa Filia": random.choice(filie),
                    "Koniec Lekcji Pon": random.choice(godziny),
                    "Koniec Lekcji Wt": random.choice(godziny),
                    "Koniec Lekcji Śr": random.choice(godziny),
                    "Koniec Lekcji Czw": random.choice(godziny),
                    "Koniec Lekcji Pt": random.choice(godziny)
                })
                
        # Generowanie trzech liceów (LO nr 1, 2, 3, klasy 1-4)
        for lo_num in range(1, 4):
            for klasa_num in range(1, 5):
                liczba_oddzialow = random.randint(4, 6)
                for litera in ['A', 'B', 'C', 'D', 'E', 'F'][:liczba_oddzialow]:
                    starsi_data.append({
                        "Szkoła i Klasa": f"LO nr {lo_num} - Klasa {klasa_num}{litera}",
                        "Poziom": "Masters",
                        "Liczba Chętnych": random.randint(0, 10),
                        "Czas trwania (min)": 90,
                        "Docelowa Filia": random.choice(filie),
                        "Koniec Lekcji Pon": random.choice(godziny),
                        "Koniec Lekcji Wt": random.choice(godziny),
                        "Koniec Lekcji Śr": random.choice(godziny),
                        "Koniec Lekcji Czw": random.choice(godziny),
                        "Koniec Lekcji Pt": random.choice(godziny)
                    })

        df_starsi = pd.DataFrame(starsi_data)
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
            "Punkt Początkowy": ["SP nr 1", "Przedszkole nr 5", "SP nr 2", "Przedszkole nr 2", "LO nr 1", "LO nr 2", "LO nr 3"],
            "Punkt Końcowy": ["Komorów", "Michałowice", "Ursus 1", "Pruszków", "Nowa Wieś", "Ursus 2", "Komorów"],
            "Czas Przejścia (min)": [15, 10, 20, 10, 25, 15, 20]
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
    st.success("Plik wgrany poprawnie! Trwa analiza zapotrzebowania...")
    
    # Wczytywanie pliku Excel
    xls = pd.ExcelFile(uploaded_file)
    
    # Wczytywanie z zabezpieczeniem, jeśli ktoś wrzuci starszą wersję pliku bez kolumny czasu
    df_mlodsze = pd.read_excel(xls, sheet_name="Młodsze Dzieci")
    if "Czas trwania (min)" not in df_mlodsze.columns:
        df_mlodsze.insert(3, "Czas trwania (min)", 60)
        
    df_starsi = pd.read_excel(xls, sheet_name="Starsi Uczniowie")
    if "Czas trwania (min)" not in df_starsi.columns:
        df_starsi.insert(4, "Czas trwania (min)", 90)
        
    df_sale = pd.read_excel(xls, sheet_name="Sale")
    
    st.header("Krok 1: Weryfikacja i formowanie grup")
    
    # ---------------- BOKS: BUFORY CZASU NA DOJAZD ----------------
    st.subheader("Bufory czasu na dojazd ze szkoły (wg filii)")
    st.write("Ustaw indywidualny czas, jakiego potrzebują dzieci na dotarcie do konkretnej placówki z okolicznej szkoły publicznej:")
    
    filie_unikalne = sorted(df_sale["Filia"].dropna().unique())
    bufory_filii = {}
    
    # Rozkładamy suwaki dynamicznie w 3 kolumnach dla estetyki
    cols_bufory = st.columns(3)
    for idx, filia in enumerate(filie_unikalne):
        with cols_bufory[idx % 3]:
            bufory_filii[filia] = st.slider(f"Bufor - {filia} (min)", min_value=15, max_value=60, value=30, step=5, key=f"bufor_{filia}")
            
    # ---------------- BOKS: MŁODSZE DZIECI ----------------
    st.markdown("---")
    st.subheader("Młodsze Dzieci (0-3 i przedszkole) - Edycja grup")
    st.write("Kliknij dwukrotnie w wybraną komórkę, aby zmienić czas trwania zajęć lub inne parametry wyłącznie dla konkretnej grupy.")
    
    # Interaktywna tabela
    edytowane_mlodsze = st.data_editor(df_mlodsze, use_container_width=True, num_rows="dynamic", key="editor_mlodsze")

    # ---------------- BOKS: STARSI UCZNIOWIE ----------------
    st.markdown("---")
    st.subheader("Starsi Uczniowie (klasy 4-8, Licea) - Selekcja i Edycja")
    
    # Logika weryfikacji liczebności grup
    grupy_odrzucone = df_starsi[df_starsi["Liczba Chętnych"] < 5]
    grupy_ostrzezenie = df_starsi[df_starsi["Liczba Chętnych"] == 5]
    grupy_zatwierdzone = df_starsi[df_starsi["Liczba Chętnych"] >= 6]
    
    if not grupy_odrzucone.empty:
        st.error(f"Odrzucono {len(grupy_odrzucone)} klas z powodu braku wymaganej liczby chętnych (poniżej 5 osób).")
        with st.expander("Rozwiń listę odrzuconych klas (do wglądu)"):
            st.dataframe(grupy_odrzucone, use_container_width=True)
            
    # Połączenie zatwierdzonych i tych z ostrzeżeniem jako naszej bazy do układania grafiku
    aktywne_grupy_starsi = pd.concat([grupy_zatwierdzone, grupy_ostrzezenie]).reset_index(drop=True)
    
    st.success(f"Zakwalifikowano {len(aktywne_grupy_starsi)} klas do układania grafiku (w tym {len(grupy_ostrzezenie)} z ostrzeżeniami o małej liczebności).")
    
    st.write("Kliknij dwukrotnie w komórkę poniżej, aby zmienić **Czas trwania (min)** lub docelową filię dla wybranej grupy.")
    
    # Interaktywna tabela
    edytowane_starsi = st.data_editor(aktywne_grupy_starsi, use_container_width=True, num_rows="dynamic", key="editor_starsi")
