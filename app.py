import streamlit as st
import pandas as pd
import random
import itertools
from io import BytesIO

st.set_page_config(page_title="Układacz Grafików", layout="wide")

st.title("📅 Szkolny Układacz Grafików")
st.write(
    "Witaj w aplikacji! Pobierz szablon, wypełnij go, "
    "a następnie wgraj poniżej, by wygenerować grafik."
)

# ==========================================
# 1. GENERATOR SZABLONU EXCEL
# ==========================================
def generate_excel_template():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        
        df_lektorzy = pd.DataFrame({
            "Lektor": [
                "Jan Kowalski", "Anna Nowak", "Marek Wiśniewski",
                "Katarzyna Wójcik", "Piotr Kamiński"
            ],
            "Dostępność Pon": [
                "14:00-19:00", "", "15:00-20:00", 
                "13:00-17:00", ""
            ],
            "Dostępność Wt": [
                "14:00-19:00", "15:00-18:00", "", 
                "13:00-17:00", "14:00-19:00"
            ],
            "Dostępność Śr": [
                "", "15:00-18:00", "15:00-20:00", 
                "13:00-17:00", "14:00-19:00"
            ],
            "Dostępność Czw": [
                "14:00-19:00", "", "15:00-20:00", 
                "", "14:00-19:00"
            ],
            "Dostępność Pt": [
                "", "15:00-18:00", "", 
                "13:00-17:00", ""
            ],
            "Filie": [
                "Komorów", "Michałowice", "Pruszków", 
                "Ursus 1", "Ursus 2"
            ],
            "Poziomy": [
                "5, 6, 7, 8, Masters", "3-5 lat, 0, 1", "2, 4", 
                "0, 1, 2, 3, 4, 5, 6, 7, 8", "Masters"
            ]
        })
        df_lektorzy.to_excel(writer, sheet_name="Lektorzy", index=False)
        
        filie = [
            "Komorów", "Michałowice", "Pruszków", 
            "Ursus 1", "Ursus 2", "Nowa Wieś"
        ]
        nazwy_sal, filie_sal, przeznaczenie = [], [], []
        
        for filia in filie:
            for i in range(1, 6):
                nazwy_sal.append(str(i))
                filie_sal.append(filia)
                if i == 1: przeznaczenie.append("3-5 lat, 0, 1, 2, 3")
                elif i == 2: przeznaczenie.append("4, 5, 6, 7, 8, Masters")
                else: przeznaczenie.append("Wszystkie")

        pd.DataFrame({
            "Nazwa Sali": nazwy_sal, 
            "Filia": filie_sal, 
            "Przeznaczenie": przeznaczenie
        }).to_excel(writer, sheet_name="Sale", index=False)
        
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
        
        starsi_data = []
        godziny = ["12:00", "13:20", "14:25", "15:15", "16:05"]
        
        for klasa_num in range(4, 9):
            for litera in ['A', 'B', 'C', 'D'][:random.randint(2, 4)]:
                starsi_data.append({
                    "Szkoła i Klasa": f"SP nr 1 - Klasa {klasa_num}{litera}",
                    "Poziom": str(klasa_num), 
                    "Liczba Chętnych": random.randint(1, 8),
                    "Czas trwania (min)": 90, 
                    "Docelowa Filia": "Komorów",
                    "Koniec Lekcji Pon": random.choice(godziny), 
                    "Koniec Lekcji Wt": random.choice(godziny),
                    "Koniec Lekcji Śr": random.choice(godziny), 
                    "Koniec Lekcji Czw": random.choice(godziny),
                    "Koniec Lekcji Pt": random.choice(godziny)
                })

        pd.DataFrame(starsi_data).to_excel(writer, sheet_name="Starsi Uczniowie", index=False)
        
        pd.DataFrame({
            "Imię Opiekunki": ["Marta Wiśniewska", "Krystyna Kaczmarek"],
            "Dostępność Pon": ["12:00-16:00", "13:00-17:00"],
            "Dostępność Wt": ["12:00-16:00", ""],
            "Dostępność Śr": ["12:00-16:00", "13:00-17:00"],
            "Dostępność Czw": ["12:00-16:00", ""],
            "Dostępność Pt": ["12:00-16:00", "13:00-17:00"]
        }).to_excel(writer, sheet_name="Opiekunki", index=False)
        
        pd.DataFrame({
            "Punkt Początkowy": [
                "SP nr 1", "Przedszkole nr 5", 
                "SP nr 2", "Przedszkole nr 2"
            ],
            "Punkt Końcowy": [
                "Komorów", "Michałowice", 
                "Ursus 1", "Pruszków"
            ],
            "Czas Przejścia (min)": [15, 10, 20, 10]
        }).to_excel(writer, sheet_name="Trasy", index=False)

        for sheetname in writer.sheets:
            ws = writer.sheets[sheetname]
            for col in ws.columns:
                max_len = max([len(str(c.value)) for c in col if c.value] + [0])
                ws.column_dimensions[col[0].column_letter].width = max_len + 3
                
    return output.getvalue()

# ==========================================
# FUNKCJE POMOCNICZE
# ==========================================
dni_short = ["Pon", "Wt", "Śr", "Czw", "Pt"]

def time_to_mins(t_str):
    if pd.isna(t_str) or not str(t_str).strip(): return -1
    try:
        h, m = map(int, str(t_str).split(':'))
        return h * 60 + m
    except: return -1

def mins_to_time(mins):
    if mins < 0: return ""
    return f"{int(mins // 60):02d}:{int(mins % 60):02d}"

def parse_availability(avail_str):
    if pd.isna(avail_str) or not str(avail_str).strip(): return []
    blocks = []
    for part in str(avail_str).split(','):
        try:
            start_s, end_s = part.split('-')
            blocks.append((time_to_mins(start_s), time_to_mins(end_s)))
        except: pass
    return blocks

def is_overlap(start1, end1, start2, end2):
    return max(start1, start2) < min(end1, end2)

def score_combo(combo):
    # Ocena układu dni (1 - Najlepszy, 3 - Najgorszy)
    if len(combo) == 1: return 1
    if len(combo) == 2:
        idx1 = dni_short.index(combo[0])
        idx2 = dni_short.index(combo[1])
        diff = abs(idx1 - idx2)
        if diff == 2 or diff == 4: return 1  # Pon-Śr, Wt-Czw, Śr-Pt, Pon-Pt
        elif diff == 3: return 2             # Pon-Czw, Wt-Pt
        elif diff == 1: return 3             # Dzień po dniu (np. Pon-Wt)
        return 10
    return 1 # Fallback dla większej ilości spotkań

def get_scored_combos(avail_dict, spotkan):
    valid_days = list(avail_dict.keys())
    if len(valid_days) < spotkan: return []
    
    combos = list(itertools.combinations(valid_days, spotkan))
    scored = [(c, score_combo(c)) for c in combos]
    scored.sort(key=lambda x: x[1]) # Najpierw najlepsze wzorce
    return [x[0] for x in scored]

def check_lektor(lek, d, st_m, en_m, grafik):
    can_work = any(b_s <= st_m and b_e >= en_m for (b_s, b_e) in lek["Avail"][d])
    if not can_work: return False
    
    # 5 min bufor dla tego samego lektora na dojscie miedzy salami
    zajety = any(
        g["Lektor"] == lek["Lektor"] and g["Dzień"] == d and 
        is_overlap(st_m, en_m, g["Start"], g["End"] + 5) 
        for g in grafik
    )
    return not zajety

def check_sala(filia, poziom, d, st_m, en_m, grafik, sale_dane):
    dostepne = [
        s["Sala"] for s in sale_dane 
        if s["Filia"] == filia and 
        (poziom in s["Poziomy"] or "Wszystkie" in s["Poziomy"])
    ]
    for s in dostepne:
        zajeta = any(
            g["Sala"] == s and g["Dzień"] == d and 
            g["Filia"] == filia and 
            is_overlap(st_m, en_m, g["Start"], g["End"]) 
            for g in grafik
        )
        if not zajeta: return s
    return None

# ==========================================
# INTERFEJS GŁÓWNY
# ==========================================
st.subheader("1. Pobierz szablon")
st.download_button(
    label="📥 Pobierz układacz grafików (Excel)",
    data=generate_excel_template(),
    file_name="ukladacz_grafikow.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.subheader("2. Wgraj uzupełniony plik")
uploaded_file = st.file_uploader("Wgraj plik Excel", type=["xlsx"])

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    df_lektorzy = pd.read_excel(xls, sheet_name="Lektorzy")
    df_sale = pd.read_excel(xls, sheet_name="Sale")
    df_trasy = pd.read_excel(xls, sheet_name="Trasy")
    
    df_mlodsze = pd.read_excel(xls, sheet_name="Młodsze Dzieci")
    if "Czas trwania (min)" not in df_mlodsze.columns: 
        df_mlodsze.insert(3, "Czas trwania (min)", 60)
    
    df_starsi = pd.read_excel(xls, sheet_name="Starsi Uczniowie")
    if "Czas trwania (min)" not in df_starsi.columns: 
        df_starsi.insert(4, "Czas trwania (min)", 90)

    st.header("Krok 1: Weryfikacja i Formowanie")
    st.subheader("Bufory czasu na dojazd ze szkoły")
    
    filie_unikalne = sorted(df_sale["Filia"].dropna().unique())
    bufory_filii = {}
    cols_bufory = st.columns(3)
    for idx, filia in enumerate(filie_unikalne):
        with cols_bufory[idx % 3]:
            bufory_filii[filia] = st.slider(
                f"{filia} (min)", 15, 60, 30, 5, key=f"b_{filia}"
            )
            
    st.markdown("---")
    st.write("**Młodsze Dzieci (0-3 i przedszkole)**")
    edyt_mlodsze = st.data_editor(
        df_mlodsze, use_container_width=True, 
        num_rows="dynamic", key="e_ml"
    )

    st.write("**Starsi Uczniowie (Automatyczne łączenie)**")
    df_s = df_starsi[df_starsi["Liczba Chętnych"] > 0].copy()
    
    def ext_sch(x): return str(x).split(" - ")[0] if " - " in str(x) else x
    def ext_lit(x): 
        if " - " not in str(x): return ""
        return "".join([c for c in str(x).split(" - ")[-1] if c.isalpha()])
        
    df_s["Szkoła"] = df_s["Szkoła i Klasa"].apply(ext_sch)
    df_s["Litera"] = df_s["Szkoła i Klasa"].apply(ext_lit)
    
    dni_long = [
        "Koniec Lekcji Pon", "Koniec Lekcji Wt", 
        "Koniec Lekcji Śr", "Koniec Lekcji Czw", "Koniec Lekcji Pt"
    ]
    for d in dni_long: df_s[d + "_mins"] = df_s[d].apply(time_to_mins)
        
    def avg_t(r):
        v = [r[d+"_mins"] for d in dni_long if r[d+"_mins"] > 0]
        return sum(v) / max(1, len(v))
    df_s["Avg_Time"] = df_s.apply(avg_t, axis=1)

    utworzone_g = []
    for _, g_df in df_s.groupby(["Docelowa Filia", "Szkoła", "Poziom"]):
        g_df = g_df.sort_values("Avg_Time")
        paczka, c_count = [], 0
        for _, row in g_df.iterrows():
            if c_count + row["Liczba Chętnych"] > 12:
                if c_count > 0: utworzone_g.append(paczka)
                paczka, c_count = [row], row["Liczba Chętnych"]
            else:
                paczka.append(row)
                c_count += row["Liczba Chętnych"]
        if paczka: utworzone_g.append(paczka)

    f_rows = []
    for p in utworzone_g:
        t_chetnych = sum(r["Liczba Chętnych"] for r in p)
        if t_chetnych < 5: continue
        
        litery = "".join(sorted([r['Litera'] for r in p]))
        r_data = {
            "Nazwa Grupy": f"{p[0]['Szkoła']} - {p[0]['Poziom']} {litery}",
            "Poziom": str(p[0]["Poziom"]),
            "Liczba Dzieci": t_chetnych,
            "Czas trwania (min)": max(r["Czas trwania (min)"] for r in p),
            "Docelowa Filia": p[0]["Docelowa Filia"],
            "Liczba Spotkań": 2 
        }
        for d in dni_long:
            v_mins = [r[d + "_mins"] for r in p if r[d + "_mins"] > 0]
            r_data[d] = mins_to_time(max(v_mins)) if v_mins else ""
        f_rows.append(r_data)

    df_aktywne = pd.DataFrame(f_rows)
    edyt_starsi = st.data_editor(
        df_aktywne, use_container_width=True, 
        num_rows="dynamic", key="e_st"
    )

    # ==========================================
    # KROK 2: GENEROWANIE GRAFIKU
    # ==========================================
    st.markdown("---")
    st.header("Krok 2: Automatyczne Układanie Grafiku")
    
    if st.button("🚀 Wygeneruj Grafik", type="primary"):
        with st.spinner("Przeszukuję miliony kombinacji..."):
            grafik = []
            nieprzypisane = []
            
            # Przygotowanie zasobów
            czas_tras = {}
            for _, r in df_trasy.iterrows():
                czas_tras[f"{r['Punkt Początkowy']}_{r['Punkt Końcowy']}"] = r["Czas Przejścia (min)"]
                
            lektorzy_dane = []
            for _, r in df_lektorzy.iterrows():
                l_poziomy = [p.strip() for p in str(r["Poziom"]).split(",")] if pd.notna(r["Poziom"]) else []
                l_filie = [f.strip() for f in str(r["Filie"]).split(",")] if pd.notna(r["Filie"]) else []
                
                avail = {}
                for idx, d_long in enumerate(dni_long):
                    k = d_long.replace("Koniec Lekcji ", "Dostępność ")
                    avail[dni_short[idx]] = parse_availability(r[k])
                
                lektorzy_dane.append({
                    "Lektor": r["Lektor"], "Poziomy": l_poziomy, 
                    "Filie": l_filie, "Avail": avail
                })
                
            sale_dane = []
            for _, r in df_sale.iterrows():
                s_poz = [p.strip() for p in str(r["Przeznaczenie"]).split(",")] if pd.notna(r["Przeznaczenie"]) else []
                sale_dane.append({
                    "Sala": str(r["Nazwa Sali"]), "Filia": r["Filia"], "Poziomy": s_poz
                })
                
            # Przygotowanie zadań (grup)
            zadania = []
            
            for _, r in edyt_mlodsze.iterrows():
                av_days = {}
                skad = str(r.get("Skąd Odbiór", ""))
                do = str(r["Docelowa Filia"])
                t_czas = czas_tras.get(f"{skad}_{do}", 15)
                for idx, d in enumerate(dni_long):
                    e_m = time_to_mins(r[d])
                    if e_m > 0: av_days[dni_short[idx]] = e_m + t_czas + 5
                
                zadania.append({
                    "Grupa": r["Nazwa Grupy"], "Poziom": str(r["Poziom"]),
                    "Filia": do, "Czas": r["Czas trwania (min)"],
                    "Spotkań": int(r.get("Liczba Spotkań", 2)), 
                    "Available_Days": av_days
                })
                
            for _, r in edyt_starsi.iterrows():
                av_days = {}
                do = str(r["Docelowa Filia"])
                bufor = bufory_filii.get(do, 30)
                for idx, d in enumerate(dni_long):
                    e_m = time_to_mins(r[d])
                    if e_m > 0: av_days[dni_short[idx]] = e_m + bufor
                
                zadania.append({
                    "Grupa": r["Nazwa Grupy"], "Poziom": str(r["Poziom"]),
                    "Filia": do, "Czas": r["Czas trwania (min)"],
                    "Spotkań": int(r.get("Liczba Spotkań", 2)), 
                    "Available_Days": av_days
                })
                
            # Trudniejsze zadania pierwsze
            zadania.sort(key=lambda x: (len(x["Available_Days"]), -x["Czas"]))
            
            for zad in zadania:
                combos = get_scored_combos(zad["Available_Days"], zad["Spotkań"])
                znaleziono = False
                
                valid_leks = [
                    l for l in lektorzy_dane 
                    if zad["Poziom"] in l["Poziomy"] and 
                    (not l["Filie"] or zad["Filia"] in l["Filie"])
                ]
                
                for combo in combos:
                    for lek in valid_leks:
                        zaplanowane_dni = []
                        for d in combo:
                            st_m = zad["Available_Days"][d]
                            st_m = st_m + (5 - st_m % 5) if st_m % 5 != 0 else st_m
                            czas = zad["Czas"]
                            
                            dzien_ok = False
                            while st_m + czas <= 1140:
                                if check_lektor(lek, d, st_m, st_m + czas, grafik):
                                    sala = check_sala(
                                        zad["Filia"], zad["Poziom"], d, 
                                        st_m, st_m + czas, grafik, sale_dane
                                    )
                                    if sala:
                                        zaplanowane_dni.append({
                                            "Grupa": zad["Grupa"], 
                                            "Poziom": zad["Poziom"], 
                                            "Filia": zad["Filia"],
                                            "Dzień": d, 
                                            "Start": st_m, 
                                            "End": st_m + czas,
                                            "Sala": sala, 
                                            "Lektor": lek["Lektor"]
                                        })
                                        dzien_ok = True
                                        break
                                st_m += 5
                            if not dzien_ok: break
                            
                        # Jeśli znaleziono układ dla WSZYSTKICH dni z JEDNYM lektorem
                        if len(zaplanowane_dni) == len(combo):
                            grafik.extend(zaplanowane_dni)
                            znaleziono = True
                            break # Wychodzimy z pętli lektora
                            
                    if znaleziono: break # Wychodzimy z pętli wzorców
                
                if not znaleziono:
                    nieprzypisane.append({
                        "Grupa": zad["Grupa"], "Filia": zad["Filia"], 
                        "Poziom": zad["Poziom"],
                        "Problem": "Brak wspólnego lektora / sali na wymagane dni",
                        "Sugestia": "Sprawdź lub wydłuż dostępność lektorów dla tej filii."
                    })

            # ==========================================
            # WIZUALIZACJA WYNIKÓW
            # ==========================================
            if nieprzypisane:
                st.error(f"🔴 Konflikty grafiku! ({len(nieprzypisane)} grup wylądowało w poczekalni).")
                st.dataframe(pd.DataFrame(nieprzypisane), use_container_width=True)
            else:
                st.balloons()
                st.success("🎉 Sukces! Przypisano wszystkie grupy zachowując formułę (np. Pon-Śr, stały lektor)!")
                
            st.subheader("Wizualizacja Grafiku (Widok Filii)")
            tabs = st.tabs(filie_unikalne)
            df_grafik = pd.DataFrame(grafik)
            
            for idx, tab in enumerate(tabs):
                filia_nazwa = filie_unikalne[idx]
                with tab:
                    if not df_grafik.empty:
                        df_f = df_grafik[df_grafik["Filia"] == filia_nazwa]
                        if not df_f.empty:
                            df_f["Godzina"] = df_f["Start"].apply(mins_to_time) + " - " + df_f["End"].apply(mins_to_time)
                            df_f["Wpis"] = df_f["Grupa"] + " (" + df_f["Lektor"] + ")"
                            
                            pivot = df_f.pivot_table(
                                index=["Dzień", "Godzina"], 
                                columns="Sala", 
                                values="Wpis", 
                                aggfunc=lambda x: " | ".join(x)
                            ).fillna("-")
                            
                            st.dataframe(pivot, use_container_width=True)
                        else:
                            st.info(f"Brak przypisanych zajęć dla filii {filia_nazwa}.")
                    else:
                        st.info("Grafik jest całkowicie pusty.")
