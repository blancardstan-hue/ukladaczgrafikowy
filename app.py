import streamlit as st
import pandas as pd
import random
import itertools
from io import BytesIO

st.set_page_config(page_title="Układacz Grafików", layout="wide")

st.title("📅 Szkolny Układacz Grafików")
st.write("Witaj w aplikacji! Pobierz szablon, wypełnij go, a następnie wgraj poniżej.")

# ==========================================
# 1. GENERATOR SZABLONU EXCEL
# ==========================================
def generate_excel_template():
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        
        # 1. Lektorzy (GENEROWANIE 70 LOSOWYCH)
        imiona = ["Anna", "Jan", "Katarzyna", "Michał", "Agnieszka", "Piotr", "Magda", "Tomasz", "Ewa", "Krzysztof"]
        nazwiska = ["Nowak", "Kowal", "Wika", "Zet", "Iks", "Igrek", "Alfa", "Beta", "Gamma", "Omega"]
        godziny_lek = ["14:00-20:00", "13:00-19:00", "15:00-20:00", "08:00-12:00,15:00-19:00", "14:30-18:30", ""]
        filie = ["Komorów", "Michałowice", "Pruszków", "Ursus 1", "Ursus 2", "Nowa Wieś"]
        poziomy_opcje = [
            "3-5 lat, 0, 1, 2, 3", 
            "4, 5, 6, 7, 8, Masters", 
            "0, 1, 2, 3, 4, 5, 6, 7, 8, Masters"
        ]
        
        lektorzy_data = []
        for i in range(1, 71):
            lektorzy_data.append({
                "Lektor": f"{random.choice(imiona)} {random.choice(nazwiska)} (L{i})",
                "Dostępność Pon": random.choice(godziny_lek),
                "Dostępność Wt": random.choice(godziny_lek),
                "Dostępność Śr": random.choice(godziny_lek),
                "Dostępność Czw": random.choice(godziny_lek),
                "Dostępność Pt": random.choice(godziny_lek),
                "Filie": ", ".join(random.sample(filie, k=random.randint(1, 3))),
                "Poziomy": random.choice(poziomy_opcje)
            })
            
        pd.DataFrame(lektorzy_data).to_excel(writer, sheet_name="Lektorzy", index=False)
        
        # 2. Sale
        sale_data = []
        for f in filie:
            for i in range(1, 6):
                prz = "3-5 lat, 0, 1, 2, 3" if i == 1 else "4, 5, 6, 7, 8, Masters" if i == 2 else "Wszystkie"
                sale_data.append({"Nazwa Sali": str(i), "Filia": f, "Przeznaczenie": prz})
        pd.DataFrame(sale_data).to_excel(writer, sheet_name="Sale", index=False)
        
        # Konfiguracja Szkół
        szkoly_sp = {
            "Komorów": ["SP Komorów"], "Michałowice": ["SP Michałowice"],
            "Pruszków": ["SP nr 1 Pruszków", "SP nr 2 Pruszków"],
            "Ursus 1": ["SP Ursus A"], "Ursus 2": ["SP Ursus B"],
            "Nowa Wieś": ["SP Nowa Wieś"]
        }
        szkoly_lo = {"Komorów": ["LO Komorów"], "Pruszków": ["LO Pruszków"], "Nowa Wieś": ["LO Nowa Wieś"]}
        godziny_pocz = ["08:00", "08:55", "09:50"]
        godziny_kon = ["12:30", "13:30", "14:25", "15:20", "16:15"]

        # 3. Młodsze Dzieci
        mlodsze_data = []
        for f, szkoly in szkoly_sp.items():
            for idx, sz in enumerate(szkoly):
                mlodsze_data.append({
                    "Nazwa Grupy": f"Zerówka {idx+1}", "Poziom": "0", "Liczba Dzieci": random.randint(6, 12),
                    "Czas trwania (min)": 60, "Skąd Odbiór": sz, "Docelowa Filia": f,
                    "Początek Szkoły Pon": "08:00", "Koniec Szkoły Pon": "12:30",
                    "Początek Szkoły Wt": "08:00", "Koniec Szkoły Wt": "12:30",
                    "Początek Szkoły Śr": "08:00", "Koniec Szkoły Śr": "12:30",
                    "Początek Szkoły Czw": "08:00", "Koniec Szkoły Czw": "13:30",
                    "Początek Szkoły Pt": "08:00", "Koniec Szkoły Pt": "12:30",
                    "Liczba Spotkań": 2
                })
        pd.DataFrame(mlodsze_data).to_excel(writer, sheet_name="Młodsze Dzieci", index=False)
        
        # 4. Starsi Uczniowie
        starsi_data = []
        for f, szkoly in szkoly_sp.items():
            for sz in szkoly:
                for kl in range(1, 9):
                    for lit in ['A', 'B', 'C', 'D'][:random.randint(2, 4)]:
                        starsi_data.append({
                            "Szkoła i Klasa": f"{sz} - Klasa {kl}{lit}", "Poziom": str(kl), 
                            "Liczba Chętnych": random.randint(0, 8), "Czas trwania (min)": 90, 
                            "Docelowa Filia": f,
                            "Początek Szkoły Pon": random.choice(godziny_pocz), "Koniec Szkoły Pon": random.choice(godziny_kon),
                            "Początek Szkoły Wt": random.choice(godziny_pocz), "Koniec Szkoły Wt": random.choice(godziny_kon),
                            "Początek Szkoły Śr": random.choice(godziny_pocz), "Koniec Szkoły Śr": random.choice(godziny_kon),
                            "Początek Szkoły Czw": random.choice(godziny_pocz), "Koniec Szkoły Czw": random.choice(godziny_kon),
                            "Początek Szkoły Pt": random.choice(godziny_pocz), "Koniec Szkoły Pt": random.choice(godziny_kon)
                        })
        for f, szkoly in szkoly_lo.items():
            for sz in szkoly:
                for kl in range(1, 5):
                    for lit in ['A', 'B', 'C', 'D'][:random.randint(2, 4)]:
                        starsi_data.append({
                            "Szkoła i Klasa": f"{sz} - Klasa {kl}{lit}", "Poziom": "Masters", 
                            "Liczba Chętnych": random.randint(0, 8), "Czas trwania (min)": 90, 
                            "Docelowa Filia": f,
                            "Początek Szkoły Pon": "08:00", "Koniec Szkoły Pon": "15:20",
                            "Początek Szkoły Wt": "08:00", "Koniec Szkoły Wt": "16:15",
                            "Początek Szkoły Śr": "08:00", "Koniec Szkoły Śr": "14:25",
                            "Początek Szkoły Czw": "08:00", "Koniec Szkoły Czw": "15:20",
                            "Początek Szkoły Pt": "08:00", "Koniec Szkoły Pt": "13:30"
                        })
        pd.DataFrame(starsi_data).to_excel(writer, sheet_name="Starsi Uczniowie", index=False)
        
        # 5. Opiekunki (Z filiami)
        pd.DataFrame({
            "Imię Opiekunki": ["Marta", "Krystyna", "Zofia", "Ewa", "Agnieszka", "Magda"],
            "Filia": ["Komorów", "Michałowice", "Pruszków", "Ursus 1", "Ursus 2", "Nowa Wieś"],
            "Dostępność Pon": ["12:00-18:00"]*6, "Dostępność Wt": ["12:00-18:00"]*6,
            "Dostępność Śr": ["12:00-18:00"]*6, "Dostępność Czw": ["12:00-18:00"]*6,
            "Dostępność Pt": ["12:00-18:00"]*6
        }).to_excel(writer, sheet_name="Opiekunki", index=False)
        
        # 6. Trasy
        trasy_data = []
        for f, szkoly in szkoly_sp.items():
            for sz in szkoly: trasy_data.append({"Początek": sz, "Koniec": f, "Czas (min)": random.choice([10, 15, 20])})
        pd.DataFrame(trasy_data).to_excel(writer, sheet_name="Trasy", index=False)

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
dni_pocz = ["Początek Szkoły Pon", "Początek Szkoły Wt", "Początek Szkoły Śr", "Początek Szkoły Czw", "Początek Szkoły Pt"]
dni_kon = ["Koniec Szkoły Pon", "Koniec Szkoły Wt", "Koniec Szkoły Śr", "Koniec Szkoły Czw", "Koniec Szkoły Pt"]

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
            st_s, en_s = part.split('-')
            blocks.append((time_to_mins(st_s), time_to_mins(en_s)))
        except: pass
    return blocks

def is_overlap(st1, en1, st2, en2):
    return max(st1, st2) < min(en1, en2)

def score_combo(combo):
    if len(combo) == 1: return 1
    if len(combo) == 2:
        diff = abs(dni_short.index(combo[0]) - dni_short.index(combo[1]))
        if diff in [2, 4]: return 1  # Pon-Śr, Wt-Czw, Śr-Pt, Pon-Pt
        elif diff == 3: return 2     # Pon-Czw, Wt-Pt
        elif diff == 1: return 3     # Dzień po dniu (bardzo nie lubiane)
        return 10
    return 1

def get_scored_combos(avail_dict, spotkan):
    valid_days = list(avail_dict.keys())
    if len(valid_days) < spotkan: return []
    combos = list(itertools.combinations(valid_days, spotkan))
    scored = [(c, score_combo(c)) for c in combos]
    scored.sort(key=lambda x: x[1]) 
    return [x[0] for x in scored]

def get_color_map(unique_vals):
    colors = [
        "#ffadad", "#ffd6a5", "#fdffb6", "#caffbf", "#9bf6ff", 
        "#a0c4ff", "#bdb2ff", "#ffc6ff", "#fbb1bd", "#e2ece9",
        "#ffcfd2", "#f1c0e8", "#cfbaf0", "#a3c4f3", "#90dbf4"
    ]
    cmap = {}
    idx = 0
    for val in unique_vals:
        if val == "-" or pd.isna(val): 
            cmap[val] = ""
        else: 
            cmap[val] = f"background-color: {colors[idx % len(colors)]}; color: #000000; font-weight: bold;"
            idx += 1
    return cmap

# ==========================================
# INTERFEJS GŁÓWNY
# ==========================================
st.subheader("1. Pobierz szablon")
st.download_button("📥 Pobierz układacz (Excel)", data=generate_excel_template(), file_name="ukladacz.xlsx")

st.subheader("2. Wgraj uzupełniony plik")
uploaded_file = st.file_uploader("Wgraj plik Excel", type=["xlsx"])

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    df_lektorzy = pd.read_excel(xls, "Lektorzy")
    df_sale = pd.read_excel(xls, "Sale")
    df_trasy = pd.read_excel(xls, "Trasy")
    df_opiekunki = pd.read_excel(xls, "Opiekunki") if "Opiekunki" in xls.sheet_names else pd.DataFrame()
    
    df_mlodsze = pd.read_excel(xls, "Młodsze Dzieci")
    if "Czas trwania (min)" not in df_mlodsze.columns: df_mlodsze.insert(3, "Czas trwania (min)", 60)
    
    df_starsi = pd.read_excel(xls, "Starsi Uczniowie")
    if "Czas trwania (min)" not in df_starsi.columns: df_starsi.insert(4, "Czas trwania (min)", 90)

    st.header("Krok 1: Formowanie")
    filie_unikalne = sorted(df_sale["Filia"].dropna().unique())
    bufory_filii = {}
    cols_bufory = st.columns(3)
    for idx, f in enumerate(filie_unikalne):
        with cols_bufory[idx % 3]: bufory_filii[f] = st.slider(f"{f} (min)", 15, 60, 30, 5, key=f"b_{f}")
            
    st.markdown("---")
    st.write("**Młodsze Dzieci**")
    edyt_mlodsze = st.data_editor(df_mlodsze, use_container_width=True, num_rows="dynamic", key="e_ml")

    st.write("**Starsi Uczniowie**")
    df_s = df_starsi[df_starsi["Liczba Chętnych"] > 0].copy()
    
    def ext_sch(x): return str(x).split(" - ")[0] if " - " in str(x) else x
    def ext_lit(x): 
        if " - " not in str(x): return ""
        part = str(x).split(" - ")[-1].replace("Klasa", "").replace(" ", "")
        return "".join([c for c in part if c.isalpha()])
        
    df_s["Szkoła"] = df_s["Szkoła i Klasa"].apply(ext_sch)
    df_s["Litera"] = df_s["Szkoła i Klasa"].apply(ext_lit)
    
    for d in dni_kon: df_s[d + "_mins"] = df_s[d].apply(time_to_mins)
    df_s["Avg_Time"] = df_s.apply(lambda r: sum([r[d+"_mins"] for d in dni_kon if r[d+"_mins"]>0]) / max(1, len([r[d+"_mins"] for d in dni_kon if r[d+"_mins"]>0])), axis=1)

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
        t_chet = sum(r["Liczba Chętnych"] for r in p)
        if t_chet < 5: continue
        
        litery = "".join(sorted([r['Litera'] for r in p]))
        r_data = {
            "Nazwa Grupy": f"{p[0]['Szkoła']} - {p[0]['Poziom']}{litery}",
            "Poziom": str(p[0]["Poziom"]), "Liczba Dzieci": t_chet,
            "Czas trwania (min)": max(r["Czas trwania (min)"] for r in p),
            "Docelowa Filia": p[0]["Docelowa Filia"], "Liczba Spotkań": 2 
        }
        for d_p, d_k in zip(dni_pocz, dni_kon):
            v_k = [r[d_k + "_mins"] for r in p if r[d_k + "_mins"] > 0]
            v_p = [time_to_mins(r[d_p]) for r in p if time_to_mins(r[d_p]) > 0]
            r_data[d_k] = mins_to_time(max(v_k)) if v_k else ""
            r_data[d_p] = mins_to_time(min(v_p)) if v_p else ""
        f_rows.append(r_data)

    df_aktywne = pd.DataFrame(f_rows)
    edyt_starsi = st.data_editor(df_aktywne, use_container_width=True, num_rows="dynamic", key="e_st")

    # ==========================================
    # KROK 2: GRAFIK
    # ==========================================
    st.markdown("---")
    if st.button("🚀 Wygeneruj Grafik", type="primary"):
        with st.spinner("Przeszukuję okna i zasoby..."):
            grafik, grafik_op, nieprzypisane = [], [], []
            
            czas_tras = {f"{r['Początek']}_{r['Koniec']}": r["Czas (min)"] for _, r in df_trasy.iterrows()}
            
            # Parsowanie Lektorów i Opiekunek
            lektorzy_dane = []
            for _, r in df_lektorzy.iterrows():
                l_poziomy = [p.strip() for p in str(r["Poziomy"]).split(",")] if pd.notna(r["Poziomy"]) else []
                l_filie = [f.strip() for f in str(r["Filie"]).split(",")] if pd.notna(r["Filie"]) else []
                avail = {dni_short[i]: parse_availability(r.get(f"Dostępność {dni_short[i]}", "")) for i in range(5)}
                lektorzy_dane.append({"Lektor": r["Lektor"], "Poziomy": l_poziomy, "Filie": l_filie, "Avail": avail})
                
            opiekunki_dane = []
            if not df_opiekunki.empty:
                for _, r in df_opiekunki.iterrows():
                    avail = {dni_short[i]: parse_availability(r.get(f"Dostępność {dni_short[i]}", "")) for i in range(5)}
                    opiekunki_dane.append({"Imię": r["Imię Opiekunki"], "Filia": r.get("Filia", ""), "Avail": avail})

            sale_dane = [{"Sala": str(r["Nazwa Sali"]), "Filia": r["Filia"], "Poziomy": [p.strip() for p in str(r["Przeznaczenie"]).split(",")]} for _, r in df_sale.iterrows()]
                
            # Budowa zadań z oknami przed szkołą i po szkole
            zadania = []
            
            def zbuduj_okna(r, typ):
                av_windows = {}
                do = str(r["Docelowa Filia"])
                bufor = bufory_filii.get(do, 30) if typ == "Starsza" else czas_tras.get(f"{r.get('Skąd Odbiór', '')}_{do}", 15) + 5
                
                for idx, (d_p, d_k) in enumerate(zip(dni_pocz, dni_kon)):
                    st_sz = time_to_mins(r.get(d_p, ""))
                    en_sz = time_to_mins(r.get(d_k, ""))
                    windows = []
                    if st_sz > 0: windows.append((480, max(480, st_sz - bufor))) # Przed szkołą
                    if en_sz > 0: windows.append((en_sz + bufor, 1050 if typ == "Młodsza" else 1230)) # Po szkole
                    if windows: av_windows[dni_short[idx]] = windows
                return av_windows

            for _, r in edyt_mlodsze.iterrows():
                zadania.append({"Grupa": r["Nazwa Grupy"], "Poziom": str(r["Poziom"]), "Filia": str(r["Docelowa Filia"]), "Czas": r["Czas trwania (min)"], "Spotkań": int(r.get("Liczba Spotkań", 2)), "Typ": "Młodsza", "Odbiór": str(r.get("Skąd Odbiór", "")), "Windows": zbuduj_okna(r, "Młodsza")})
                
            for _, r in edyt_starsi.iterrows():
                zadania.append({"Grupa": r["Nazwa Grupy"], "Poziom": str(r["Poziom"]), "Filia": str(r["Docelowa Filia"]), "Czas": r["Czas trwania (min)"], "Spotkań": int(r.get("Liczba Spotkań", 2)), "Typ": "Starsza", "Odbiór": "", "Windows": zbuduj_okna(r, "Starsza")})
                
            # Sortowanie zadań: Najtrudniejsze pierwsze
            zadania.sort(key=lambda x: (len(x["Windows"]), -x["Czas"]))
            
            for zad in zadania:
                combos = get_scored_combos(zad["Windows"], zad["Spotkań"])
                znaleziono = False
                valid_leks = [l for l in lektorzy_dane if zad["Poziom"] in l["Poziomy"] and (not l["Filie"] or zad["Filia"] in l["Filie"])]
                
                # Strategia wieku: 0=Najwcześniej, 1=Środek, 2=Najpóźniej
                waga = 0 if zad["Typ"] == "Młodsza" or zad["Poziom"] in ["0","1","2","3"] else 1 if zad["Poziom"] in ["4","5"] else 2
                
                for combo in combos:
                    for lek in valid_leks:
                        zaplanowane_dni = []
                        
                        for d in combo:
                            # Wyciągamy wszystkie sloty ze wszystkich okien tego dnia
                            sloty = []
                            for win_st, win_en in zad["Windows"][d]:
                                st_test = win_st + (5 - win_st % 5) if win_st % 5 != 0 else win_st
                                while st_test + zad["Czas"] <= win_en:
                                    sloty.append(st_test)
                                    st_test += 5
                                    
                            if waga == 0: sloty.sort()
                            elif waga == 2: sloty.sort(reverse=True)
                            
                            dzien_ok = False
                            for st_m in sloty:
                                en_m = st_m + zad["Czas"]
                                
                                # Sprawdzenie lektora
                                if any(b_s <= st_m and b_e >= en_m for (b_s, b_e) in lek["Avail"][d]) and not any(g["Lektor"] == lek["Lektor"] and g["Dzień"] == d and is_overlap(st_m, en_m, g["Start"], g["End"] + 5) for g in grafik):
                                    
                                    # Sprawdzenie sali
                                    sala = next((s["Sala"] for s in sale_dane if s["Filia"] == zad["Filia"] and (zad["Poziom"] in s["Poziomy"] or "Wszystkie" in s["Poziomy"]) and not any(g["Sala"] == s["Sala"] and g["Dzień"] == d and g["Filia"] == zad["Filia"] and is_overlap(st_m, en_m, g["Start"], g["End"]) for g in grafik)), None)
                                    
                                    if sala:
                                        # Sprawdzenie Opiekunek
                                        op_przyp, op_odp = "", ""
                                        moze_isc = True
                                        
                                        if zad["Typ"] == "Młodsza" and zad["Odbiór"]:
                                            trasa = czas_tras.get(f"{zad['Odbiór']}_{zad['Filia']}", 15)
                                            # Szukaj kogoś na start
                                            for op in opiekunki_dane:
                                                if op["Filia"] == zad["Filia"] and any(b_s <= (st_m - trasa) and b_e >= st_m for (b_s, b_e) in op["Avail"][d]) and not any(g["Opiekunka"] == op["Imię"] and g["Dzień"] == d and is_overlap(st_m-trasa, st_m, g["Start"], g["End"]) for g in grafik_op):
                                                    op_przyp = op["Imię"]; break
                                            # Szukaj kogoś na koniec
                                            for op in opiekunki_dane:
                                                if op["Filia"] == zad["Filia"] and any(b_s <= en_m and b_e >= (en_m + trasa) for (b_s, b_e) in op["Avail"][d]) and not any(g["Opiekunka"] == op["Imię"] and g["Dzień"] == d and is_overlap(en_m, en_m+trasa, g["Start"], g["End"]) for g in grafik_op):
                                                    op_odp = op["Imię"]; break
                                            
                                            if not op_przyp or not op_odp: moze_isc = False
                                        
                                        if moze_isc:
                                            if op_przyp: grafik_op.extend([{"Opiekunka": op_przyp, "Dzień": d, "Start": st_m-trasa, "End": st_m}, {"Opiekunka": op_odp, "Dzień": d, "Start": en_m, "End": en_m+trasa}])
                                            
                                            zaplanowane_dni.append({
                                                "Grupa": zad["Grupa"], "Poziom": zad["Poziom"], "Filia": zad["Filia"],
                                                "Dzień": d, "Start": st_m, "End": en_m, "Sala": sala, "Lektor": lek["Lektor"],
                                                "Op_Str": f"\n🚶 {op_przyp} ➡️ | ⬅️ {op_odp}" if op_przyp else ""
                                            })
                                            dzien_ok = True
                                            break # Sukces dla tego dnia
                            
                            if not dzien_ok: break # Przerwij combo
                            
                        if len(zaplanowane_dni) == len(combo):
                            grafik.extend(zaplanowane_dni)
                            znaleziono = True
                            break 
                    if znaleziono: break 
                
                if not znaleziono:
                    nieprzypisane.append({"Grupa": zad["Grupa"], "Filia": zad["Filia"], "Poziom": zad["Poziom"], "Problem": "Brak wspólnego zasobu na dopasowane bloki"})

            # ==========================================
            # WIZUALIZACJA
            # ==========================================
            if nieprzypisane:
                st.error(f"🔴 Konflikty grafiku! ({len(nieprzypisane)} grup wylądowało w poczekalni).")
                st.dataframe(pd.DataFrame(nieprzypisane), use_container_width=True)
            else:
                st.success("🎉 Sukces! Przypisano wszystkie grupy!")
                
            st.subheader("Wizualizacja Grafiku")
            tabs = st.tabs(filie_unikalne)
            df_grafik = pd.DataFrame(grafik)
            
            for idx, tab in enumerate(tabs):
                f_nazwa = filie_unikalne[idx]
                with tab:
                    if not df_grafik.empty:
                        df_f = df_grafik[df_grafik["Filia"] == f_nazwa].copy()
                        if not df_f.empty:
                            df_f["Godzina"] = df_f["Start"].apply(mins_to_time) + " - " + df_f["End"].apply(mins_to_time)
                            df_f["Wpis"] = df_f["Grupa"] + " (" + df_f["Lektor"] + ")" + df_f["Op_Str"]
                            df_f["Dzień"] = pd.Categorical(df_f["Dzień"], categories=dni_short, ordered=True)
                            
                            pivot = df_f.pivot_table(index=["Dzień", "Godzina"], columns="Sala", values="Wpis", aggfunc=lambda x: " | ".join(x)).fillna("-")
                            
                            # Kolorowanie
                            uniq = list(set(pivot.values.flatten()))
                            cmap = get_color_map(uniq)
                            
                            try: styled = pivot.style.map(lambda x: cmap.get(x, ""))
                            except: styled = pivot.style.applymap(lambda x: cmap.get(x, ""))
                            
                            st.dataframe(styled, use_container_width=True)
                        else: st.info(f"Brak zajęć dla {f_nazwa}.")
                    else: st.info("Grafik pusty.")
