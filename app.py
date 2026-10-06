import streamlit as st
import pandas as pd
import random
from io import BytesIO

st.set_page_config(page_title="Układacz Grafików", layout="wide")

st.title("📅 Szkolny Układacz Grafików")
st.write(
    "Witaj w aplikacji! Pobierz szablon, wypełnij go, "
    "a następnie wgraj poniżej."
)

def generate_excel_template():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        
        # Arkusz 1: Lektorzy
        df_lektorzy = pd.DataFrame({
            "Lektor": [
                "Jan Kowalski", "Anna Nowak", "Marek Wiśniewski",
                "Katarzyna Wójcik", "Piotr Kamiński", 
                "Agnieszka Lewandowska", "Michał Zieliński", 
                "Ewa Szymańska", "Tomasz Woźniak", "Magdalena Dąbrowska"
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
                "Komorów", "Michałowice", "Pruszków", "Ursus 1", 
                "Ursus 2", "Nowa Wieś", "Komorów", "Pruszków", 
                "Ursus 1", "Michałowice"
            ],
            "Poziomy": [
                "5, 6, 7, 8, Masters", "3-5 lat, 0, 1", "2, 4", 
                "0, 1, 2, 3", "Masters", "5, 6, 7", 
                "3-5 lat, 0, 1, 2, 3, 4", "7, 8", "1, 3, 5", "3-5 lat, Masters"
            ]
        })
        df_lektorzy.to_excel(writer, sheet_name="Lektorzy", index=False)
        
        # Arkusz 2: Sale
        filie = [
            "Komorów", "Michałowice", "Pruszków", 
            "Ursus 1", "Ursus 2", "Nowa Wieś"
        ]
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
            "Skąd Odbiór": [
                "SP nr 1", "Przedszkole nr 5", 
                "Przedszkole nr 2", "SP nr 2"
            ],
            "Docelowa Filia": [
                "Komorów", "Michałowice", 
                "Pruszków", "Ursus 1"
            ],
            "Koniec Szkoły Pon": ["13:30", "14:00", "12:30", "13:30"],
            "Koniec Szkoły Wt": ["14:25", "14:00", "12:30", "14:25"],
            "Koniec Szkoły Śr": ["13:30", "14:00", "12:30", "13:30"],
            "Koniec Szkoły Czw": ["14:25", "14:00", "12:30", "14:25"],
            "Koniec Szkoły Pt": ["12:30", "14:00", "12:30", "12:30"],
            "Liczba Spotkań": [2, 2, 2, 1]
        })
        df_mlodsze.to_excel(writer, sheet_name="Młodsze Dzieci", index=False)
        
        # Arkusz 4: Starsi Uczniowie - BAZA TESTOWA
        starsi_data = []
        godziny = [
            "12:00", "12:30", "13:20", "13:40", "14:25", 
            "15:00", "15:15", "15:20", "16:05"
        ]
        
        wymuszone = [
            ("SP nr 1", "4", "B", 3, "13:20"),
            ("SP nr 1", "4", "C", 1, "13:40"), 
            ("SP nr 1", "4", "A", 2, "12:20"),
            ("SP nr 1", "4", "D", 4, "13:30"),
            ("SP nr 1", "5", "A", 0, "15:00"),
            ("SP nr 1", "5", "B", 1, "15:30"),
            ("SP nr 1", "5", "C", 0, "15:15"),
            ("SP nr 1", "5", "D", 1, "15:00"),
            ("SP nr 1", "8", "A", 5, "12:00"),
            ("SP nr 1", "8", "B", 9, "15:00"),
            ("SP nr 1", "8", "C", 4, "12:30"),
            ("SP nr 1", "8", "D", 2, "15:20")
        ]
        
        for sz, poz, lit, chetni, czas in wymuszone:
            starsi_data.append({
                "Szkoła i Klasa": f"{sz} - Klasa {poz}{lit}",
                "Poziom": poz,
                "Liczba Chętnych": chetni,
                "Czas trwania (min)": 90,
                "Docelowa Filia": "Komorów",
                "Koniec Lekcji Pon": czas,
                "Koniec Lekcji Wt": czas,
                "Koniec Lekcji Śr": czas,
                "Koniec Lekcji Czw": czas,
                "Koniec Lekcji Pt": czas
            })

        for klasa_num in range(1, 9):
            if klasa_num in [4, 5, 8]: 
                continue 
            liczba_oddzialow = random.randint(4, 6)
            for litera in ['A', 'B', 'C', 'D', 'E', 'F'][:liczba_oddzialow]:
                starsi_data.append({
                    "Szkoła i Klasa": f"SP nr 1 - Klasa {klasa_num}{litera}",
                    "Poziom": str(klasa_num),
                    "Liczba Chętnych": random.randint(0, 10),
                    "Czas trwania (min)": 90,
                    "Docelowa Filia": random.choice(filie),
                    "Koniec Lekcji Pon": random.choice(godziny),
                    "Koniec Lekcji Wt": random.choice(godziny),
                    "Koniec Lekcji Śr": random.choice(godziny),
                    "Koniec Lekcji Czw": random.choice(godziny),
                    "Koniec Lekcji Pt": random.choice(godziny)
                })
                
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
            "Imię Opiekunki": [
                "Marta Wiśniewska", "Krystyna Kaczmarek", 
                "Lucyna Mazur", "Natalia Piotrowska"
            ],
            "Dostępność Pon": ["12:00-16:00", "13:00-17:00", "", "11:00-15:00"],
            "Dostępność Wt": ["12:00-16:00", "", "13:00-17:00", "11:00-15:00"],
            "Dostępność Śr": ["12:00-16:00", "13:00-17:00", "13:00-17:00", ""],
            "Dostępność Czw": ["12:00-16:00", "", "13:00-17:00", "11:00-15:00"],
            "Dostępność Pt": ["12:00-16:00", "13:00-17:00", "", "11:00-15:00"]
        })
        df_opiekunki.to_excel(writer, sheet_name="Opiekunki", index=False)
        
        # Arkusz 6: Trasy i Czas Dojazdów
        df_trasy = pd.DataFrame({
            "Punkt Początkowy": [
                "SP nr 1", "Przedszkole nr 5", "SP nr 2", 
                "Przedszkole nr 2", "LO nr 1", "LO nr 2", "LO nr 3"
            ],
            "Punkt Końcowy": [
                "Komorów", "Michałowice", "Ursus 1", 
                "Pruszków", "Nowa Wieś", "Ursus 2", "Komorów"
            ],
            "Czas Przejścia (min)": [15, 10, 20, 10, 25, 15, 20]
        })
        df_trasy.to_excel(writer, sheet_name="Trasy", index=False)

        for sheetname in writer.sheets:
            worksheet = writer.sheets[sheetname]
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except: pass
                worksheet.column_dimensions[column].width = (max_length + 3)
                
    return output.getvalue()

def time_to_mins(t_str):
    if pd.isna(t_str) or not isinstance(t_str, str): return 0
    try:
        h, m = map(int, t_str.split(':'))
        return h * 60 + m
    except: 
        return 0

def mins_to_time(mins):
    if mins == 0: return ""
    return f"{int(mins // 60):02d}:{int(mins % 60):02d}"

st.subheader("1. Pobierz szablon")
excel_data = generate_excel_template()
st.download_button(
    label="📥 Pobierz układacz grafików (Excel)",
    data=excel_data,
    file_name="ukladacz_grafikow.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.subheader("2. Wgraj uzupełniony plik")
uploaded_file = st.file_uploader(
    "Wrzuć z powrotem wypełniony plik Excel", 
    type=["xlsx"]
)

if uploaded_file is not None:
    st.success("Plik wgrany poprawnie! Trwa analiza zapotrzebowania...")
    
    xls = pd.ExcelFile(uploaded_file)
    df_mlodsze = pd.read_excel(xls, sheet_name="Młodsze Dzieci")
    if "Czas trwania (min)" not in df_mlodsze.columns: 
        df_mlodsze.insert(3, "Czas trwania (min)", 60)
    
    df_starsi = pd.read_excel(xls, sheet_name="Starsi Uczniowie")
    if "Czas trwania (min)" not in df_starsi.columns: 
        df_starsi.insert(4, "Czas trwania (min)", 90)
        
    df_sale = pd.read_excel(xls, sheet_name="Sale")
    
    st.header("Krok 1: Weryfikacja i formowanie grup")
    
    st.subheader("Bufory czasu na dojazd ze szkoły")
    filie_unikalne = sorted(df_sale["Filia"].dropna().unique())
    bufory_filii = {}
    cols_bufory = st.columns(3)
    
    for idx, filia in enumerate(filie_unikalne):
        with cols_bufory[idx % 3]:
            bufory_filii[filia] = st.slider(
                f"Bufor - {filia} (min)", 
                min_value=15, max_value=60, value=30, step=5, 
                key=f"bufor_{filia}"
            )
            
    st.markdown("---")
    st.subheader("Młodsze Dzieci (0-3 i przedszkole)")
    edytowane_mlodsze = st.data_editor(
        df_mlodsze, use_container_width=True, 
        num_rows="dynamic", key="editor_mlodsze"
    )

    st.markdown("---")
    st.subheader("Starsi Uczniowie - Łączenie Grup")
    
    df_starsi_chętni = df_starsi[df_starsi["Liczba Chętnych"] > 0].copy()
    
    def extract_school(x):
        if " - " in x: return x.split(" - ")[0]
        return x

    df_starsi_chętni["Szkoła"] = df_starsi_chętni["Szkoła i Klasa"].apply(extract_school)
    
    def extract_letter(x):
        if " - " in x:
            part = x.split(" - ")[-1].replace("Klasa ", "").strip()
            return "".join([c for c in part if c.isalpha()])
        return ""
        
    df_starsi_chętni["Litera"] = df_starsi_chętni["Szkoła i Klasa"].apply(extract_letter)
    
    dni = [
        "Koniec Lekcji Pon", "Koniec Lekcji Wt", 
        "Koniec Lekcji Śr", "Koniec Lekcji Czw", "Koniec Lekcji Pt"
    ]
    
    for d in dni:
        df_starsi_chętni[d + "_mins"] = df_starsi_chętni[d].apply(time_to_mins)
        
    def calc_avg(row):
        vals = [row[d + "_mins"] for d in dni if row[d + "_mins"] > 0]
        return sum(vals)/len(vals) if vals else 0
        
    df_starsi_chętni["Avg_Time"] = df_starsi_chętni.apply(calc_avg, axis=1)

    utworzone_grupy = []
    
    for (filia, szkola, poziom), group_df in df_starsi_chętni.groupby(["Docelowa Filia", "Szkoła", "Poziom"]):
        group_df = group_df.sort_values("Avg_Time")
        
        current_paczka = []
        current_count = 0
        
        for _, row in group_df.iterrows():
            c = row["Liczba Chętnych"]
            if current_count + c > 12:
                if current_count > 0:
                    utworzone_grupy.append(current_paczka)
                current_paczka = [row]
                current_count = c
            else:
                current_paczka.append(row)
                current_count += c
                
        if current_paczka:
            utworzone_grupy.append(current_paczka)

    final_rows = []
    rejected_rows = []
    warning_rows = []

    for paczka in utworzone_grupy:
        total_chetnych = sum(r["Liczba Chętnych"] for r in paczka)
        szkola = paczka[0]["Szkoła"]
        poziom = str(paczka[0]["Poziom"])
        filia = paczka[0]["Docelowa Filia"]
        czas_trwania = max(r["Czas trwania (min)"] for r in paczka)
        
        litery = "".join(sorted([r["Litera"] for r in paczka]))
        nazwa_grupy = f"{szkola} - {poziom}{litery}"
        
        max_times = {}
        for d in dni:
            valid_mins = [r[d + "_mins"] for r in paczka if r[d + "_mins"] > 0]
            max_m = max(valid_mins) if valid_mins else 0
            max_times[d] = mins_to_time(max_m)
            
        skladowe = ", ".join([r["Szkoła i Klasa"].split(" - ")[-1] for r in paczka])
            
        row_data = {
            "Nazwa Grupy": nazwa_grupy,
            "Poziom": poziom,
            "Liczba Dzieci": total_chetnych,
            "Czas trwania (min)": czas_trwania,
            "Docelowa Filia": filia,
            "Składowe Klasy": skladowe,
            "Koniec Lekcji Pon": max_times["Koniec Lekcji Pon"],
            "Koniec Lekcji Wt": max_times["Koniec Lekcji Wt"],
            "Koniec Lekcji Śr": max_times["Koniec Lekcji Śr"],
            "Koniec Lekcji Czw": max_times["Koniec Lekcji Czw"],
            "Koniec Lekcji Pt": max_times["Koniec Lekcji Pt"]
        }
        
        if total_chetnych < 5:
            rejected_rows.append(row_data)
        elif total_chetnych == 5:
            warning_rows.append(row_data)
        else:
            final_rows.append(row_data)

    df_wynik = pd.DataFrame(final_rows)
    df_ostrzezenia = pd.DataFrame(warning_rows)
    df_odrzucone = pd.DataFrame(rejected_rows)
    
    if not df_odrzucone.empty:
        st.error(
            f"Odrzucono {len(df_odrzucone)} połączonych grup. "
            "Powód: mniej niż 5 osób."
        )
        with st.expander("Rozwiń listę odrzuconych"):
            st.dataframe(df_odrzucone, use_container_width=True)
            
    if not df_ostrzezenia.empty:
        st.warning(
            f"Ostrzeżenie: {len(df_ostrzezenia)} grup liczy dokładnie 5 osób. "
            "Wymagają uwagi."
        )
        
    df_aktywne = pd.concat([df_wynik, df_ostrzezenia]).reset_index(drop=True)
    st.success(
        f"Uformowano {len(df_aktywne)} stabilnych grup "
        "(od 5 do 12 osób)."
    )
    
    st.write("Możesz ręcznie nadpisać Czas trwania lub Filię:")
    edytowane_starsi = st.data_editor(
        df_aktywne, 
        use_container_width=True, 
        num_rows="dynamic", 
        key="editor_starsi"
    )
