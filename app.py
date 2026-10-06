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
        godziny = ["12:00", "12:30", "13:20", "13:40", "14:25", "15:00", "15:15", "15:20", "16:05"]
        
        # Przykłady wymuszone dla pewności (z Twoich promptów)
        wymuszone = [
            ("SP nr 1", "4", "B", 3, "13:20"), ("SP nr 1", "4", "C", 1, "13:40"), 
            ("SP nr 1", "4", "A", 2, "12:20"), ("SP nr 1", "4", "D", 4, "13:30"),
            ("SP nr 1", "5", "A", 0, "15:00"), ("SP nr 1", "5", "B", 1, "15:30"),
            ("SP nr 1", "5", "C", 0, "15:15"), ("SP nr 1", "5", "D", 1, "15:00"),
            ("SP nr 1", "8", "A", 5, "12:00"), ("SP nr 1", "8", "B", 9, "15:00"),
            ("SP nr 1", "8", "C", 4, "12:30"), ("SP nr 1", "8", "D", 2, "15:20")
        ]
        for sz, poz, lit, chetni, czas in wymuszone:
            starsi_data.append({
                "Szkoła i Klasa": f"{sz} - Klasa {poz}{lit}",
                "Poziom": poz, "Liczba Chętnych": chetni, "Czas trwania (min)": 90,
                "Docelowa Filia": "Komorów",
                "Koniec Lekcji Pon": czas, "Koniec Lekcji Wt": czas, "Koniec Lekcji Śr": czas,
                "Koniec Lekcji Czw": czas, "Koniec Lekcji Pt": czas
            })

        # Reszta generatora
        for klasa_num in range(1, 9):
            if klasa_num in [4, 5, 8]: continue # Ominięcie wymuszonych
            liczba_oddzialow = random.randint(4, 6)
            for litera in ['A', 'B', 'C', 'D', 'E', 'F'][:liczba_oddzialow]:
                starsi_data.append({
                    "Szkoła i Klasa": f"SP nr 1 - Klasa {klasa_num}{litera}",
                    "Poziom": str(klasa_num), "Liczba Chętnych": random.randint(0, 10), "
