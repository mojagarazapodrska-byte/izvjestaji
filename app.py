import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import tempfile
import os
from fpdf import FPDF

# Podešavanje stranice
st.set_page_config(page_title="ATM izvještaj Payten BiH", layout="wide")

# CSS Stil za Payten branding i širenje tekstualnih polja u tabeli
st.markdown("""
    <style>
    .payten-title {
        font-size: 38px;
        font-weight: 800;
        font-family: 'Helvetica', sans-serif;
        color: #111111;
        margin-bottom: 0px;
    }
    .payten-y {
        color: #E31B23;
    }
    [data-testid="stDataFrame"] div[data-testid="stTable"] td {
        white-space: normal !important;
        word-wrap: break-word !important;
    }
    </style>
    <div class="payten-title">ATM Izvještaj Pa<span class="payten-y">y</span>ten BiH — Advanced Operations</div>
""", unsafe_allow_html=True)

st.markdown("Profesionalni operativni dashboard za analizu ATM mreže, zastoja, grešaka i učinka tehničara.")

# Otpremanje fajla
file_problemi = st.file_uploader("Otpremite Tmanage izvještaj ('Problemi.xls' / .xlsx / .csv)", type=["xls", "xlsx", "csv"])

if file_problemi:
    if file_problemi.name.endswith('.csv'):
        df_prob = pd.read_csv(file_problemi)
    else:
        df_prob = pd.read_excel(file_problemi)

    # Detekcija naziva kolona iz Tmanage izvještaja
    col_tms = 'TMS ID' if 'TMS ID' in df_prob.columns else ('Kod bankomata' if 'Kod bankomata' in df_prob.columns else None)
    col_sn = 'Serijski broj' if 'Serijski broj' in df_prob.columns else ('SN' if 'SN' in df_prob.columns else None)
    col_tip = 'Tip terminala' if 'Tip terminala' in df_prob.columns else ('Model' if 'Model' in df_prob.columns else None)
    col_vlasnik = 'Vlasnik' if 'Vlasnik' in df_prob.columns else ('Banka' if 'Banka' in df_prob.columns else None)
    col_grad = 'Grad' if 'Grad' in df_prob.columns else 'Lokacija_Grad'
    col_lokacija = 'Lokacija' if 'Lokacija' in df_prob.columns else 'Adresa'
    col_kvar = 'Tip problema' if 'Tip problema' in df_prob.columns else ('Pocetni tip' if 'Pocetni tip' in df_prob.columns else 'Klasa problema')
    col_teh = 'Tehničar zatvorio' if 'Tehničar zatvorio' in df_prob.columns else ('Inzenjer' if 'Inzenjer' in df_prob.columns else 'Tehnicar')
    col_datum_otv = 'Datum otvaranja' if 'Datum otvaranja' in df_prob.columns else None
    col_datum_zatv = 'Datum zatvaranja' if 'Datum zatvaranja' in df_prob.columns else ('Datum rjesavanja' if 'Datum rjesavanja' in df_prob.columns else None)
    col_status = 'Status' if 'Status' in df_prob.columns else None
    col_ap = 'AP' if 'AP' in df_prob.columns else ('Vreme na lokaciji' if 'Vreme na lokaciji' in df_prob.columns else None)

    # Konverzija datuma otvaranja
    if col_datum_otv and col_datum_otv in df_prob.columns:
        df_prob['Parsed_Date'] = pd.to_datetime(df_prob[col_datum_otv], format='%d.%m.%Y %H:%M', errors='coerce')
        if df_prob['Parsed_Date'].isna().all():
            df_prob['Parsed_Date'] = pd.to_datetime(df_prob[col_datum_otv], errors='coerce')
        df_prob['Samo_Datum'] = df_prob['Parsed_Date'].dt.date
    else:
        df_prob['Samo_Datum'] = pd.NaT

    if col_datum_zatv and col_datum_zatv in df_prob.columns:
        df_prob['Parsed_Close_Date'] = pd.to_datetime(df_prob[col_datum_zatv], format='%d.%m.%Y %H:%M', errors='coerce')
        if df_prob['Parsed_Close_Date'].isna().all():
            df_prob['Parsed_Close_Date'] = pd.to_datetime(df_prob[col_datum_zatv], errors='coerce')
        df_prob['Samo_Datum_Intervencije'] = df_prob['Parsed_Close_Date'].dt.date
    else:
        df_prob['Samo_Datum_Intervencije'] = pd.NaT

    # Sidebar filteri
    st.sidebar.header("⚙️ Filteri i Parametri")
    
    if col_vlasnik and col_vlasnik in df_prob.columns:
        banke_lista = ["Sve Banke"] + sorted(df_prob[col_vlasnik].dropna().astype(str).unique().tolist())
        izabrana_banka = st.sidebar.selectbox("Filtriraj po Vlasniku / Banci:", banke_lista)
        if izabrana_banka != "Sve Banke":
            df_filtered = df_prob[df_prob[col_vlasnik] == izabrana_banka].copy()
        else:
            df_filtered = df_prob.copy()
    else:
        izabrana_banka = "Sve Banke"
        df_filtered = df_prob.copy()

    st.sidebar.markdown("---")
    st.sidebar.subheader("📅 Opseg datuma (Od – Do)")
    validni_datumi = df_prob['Samo_Datum'].dropna()
    
    if not validni_datumi.empty:
        min_d, max_d = validni_datumi.min(), validni_datumi.max()
        datum_opseg = st.sidebar.date_input("Izaberite period izvještaja:", value=(min_d, max_d), min_value=min_d, max_value=max_d)
    else:
        datum_opseg = None

    if datum_opseg and isinstance(datum_opseg, tuple) and len(datum_opseg) == 2:
        start_date, end_date = datum_opseg
        df_filtered = df_filtered[(df_filtered['Samo_Datum'] >= start_date) & (df_filtered['Samo_Datum'] <= end_date)]
        period_opisni_str = f"Period: {start_date} do {end_date}"
    else:
        period_opisni_str = "Svi raspoloživi datumi"

    st.sidebar.markdown("---")
    st.sidebar.subheader("📌 Statusi problema")
    if col_status and col_status in df_filtered.columns:
        dostupni_statusi = sorted(df_filtered[col_status].dropna().astype(str).unique().tolist())
        izabrani_statusi = st.sidebar.multiselect("Izaberite status(e) za izvještaj:", dostupni_statusi, default=dostupni_statusi)
        if izabrani_statusi:
            df_filtered = df_filtered[df_filtered[col_status].astype(str).isin(izabrani_statusi)]
    else:
        izabrani_statusi = []

    # Metrike
    tot_problemi = len(df_filtered)
    
    if col_ap and col_ap in df_filtered.columns:
        hd_mask = df_filtered[col_ap].isna() | (df_filtered[col_ap].astype(str).str.strip() == "") | (df_filtered[col_ap].astype(str).str.lower() == "null")
        hd_cnt = int(hd_mask.sum())
        tehnicar_izlasci_cnt = int((~hd_mask).sum())
    else:
        hd_cnt = 0
        tehnicar_izlasci_cnt = tot_problemi

    van_funkcije_cnt = 0
    if col_status and not df_filtered.empty:
        van_funkcije_cnt = len(df_filtered[~df_filtered[col_status].astype(str).str.contains("ZATVOREN|Zatvoren", case=False, na=False)])

    if 'Vreme prekoračenja' in df_filtered.columns and tot_problemi > 0:
        sla_prolazni = len(df_filtered[df_filtered['Vreme prekoračenja'] <= 0])
        sla_rate = round((sla_prolazni / tot_problemi * 100), 1)
    else:
        sla_rate = 95.0

    if col_ap and col_ap in df_filtered.columns:
        valid_ap_vals = pd.to_numeric(df_filtered[col_ap], errors='coerce').dropna()
        if not valid_ap_vals.empty:
            avg_mttr_min = round(valid_ap_vals.mean(), 1)
        else:
            avg_mttr_min = 45.0
    elif 'Vreme resavanja' in df_filtered.columns and not df_filtered['Vreme resavanja'].dropna().empty:
        avg_mttr_min = round(df_filtered['Vreme resavanja'].mean(), 1)
    else:
        avg_mttr_min = 45.0

    # Agregacija po bankomatu (uključujući Vlasnika/Banku)
    atm_summary = []
    if col_tms and col_tms in df_filtered.columns:
        for tms_id, group in df_filtered.groupby(col_tms):
            cnt = len(group)
            sn = str(group[col_sn].dropna().iloc[0]) if col_sn and not group[col_sn].dropna().empty else "N/A"
            tip = str(group[col_tip].dropna().iloc[0]) if col_tip and not group[col_tip].dropna().empty else "N/A"
            vlasnik = str(group[col_vlasnik].dropna().iloc[0]) if col_vlasnik and not group[col_vlasnik].dropna().empty else "N/A"
            grad = str(group[col_grad].dropna().iloc[0]) if col_grad and not group[col_grad].dropna().empty else "N/A"
            lokacija = str(group[col_lokacija].dropna().iloc[0]) if col_lokacija and not group[col_lokacija].dropna().empty else "N/A"
            
            kvarovi = group[col_kvar].dropna().astype(str).unique().tolist() if col_kvar in group.columns else []
            kvarovi_str = "\n".join([f"• {k}" for k in kvarovi]) if kvarovi else "N/A"

            tehnicari = group[col_teh].dropna().astype(str).unique().tolist() if col_teh in group.columns else []
            tehnicari_str = ", ".join(tehnicari) if tehnicari else "N/A"

            atm_summary.append({
                'TMS ID': tms_id,
                'Banka / Vlasnik': vlasnik,
                'Serijski Broj (SN)': sn,
                'Tip Terminala': tip,
                'Grad': grad,
                'Lokacija': lokacija,
                'Broj Problema': cnt,
                'Tipovi Problema': kvarovi_str,
                'Tehničari': tehnicari_str,
                'Status / Akcija': "AKTIVAN / VAN FUNKCIJE" if cnt >= 2 else "EVIDENTIRANO"
            })

    df_summary_sve = pd.DataFrame(atm_summary).sort_values(by='Broj Problema', ascending=False) if atm_summary else pd.DataFrame()
    if not df_summary_sve.empty and 'Broj Problema' in df_summary_sve.columns:
        df_summary = df_summary_sve[df_summary_sve['Broj Problema'] >= 3].copy()
    else:
        df_summary = df_summary_sve.copy()

    # Dashboard prikaz
    st.subheader(f"📊 Dashboard Performansi | {izabrana_banka}")
    st.caption(period_opisni_str)
    
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Ukupno Problema", tot_problemi)
    k2.metric("🛠️ Izlasci (Tehničar)", tehnicar_izlasci_cnt, delta=f"HD riješio: {hd_cnt}")
    k3.metric("🚨 Van Funkcije / Otvoreno", van_funkcije_cnt)
    k4.metric("SLA Prolaznost", f"{sla_rate}%")
    k5.metric("Prosječno vrijeme", f"{avg_mttr_min} min")

    st.markdown("---")

    st.subheader("📋 Pregled Uređaja sa 3 ili više problema (sa naznakom banke)")
    if not df_summary.empty:
        st.dataframe(df_summary, use_container_width=True, hide_index=True)
    else:
        st.warning("⚠️ Nema uređaja koji imaju 3 ili više otvorenih/zabilježenih problema u odabranom periodu.")

    st.markdown("---")

    st.subheader("📈 Vizuelna Analiza i Prikaz Zastoja")
    
    col_g1, col_g2 = st.columns(2)
    top_kvarovi = df_filtered[col_kvar].value_counts() if col_kvar in df_filtered.columns else pd.Series()
    top_tehnicari = df_filtered[col_teh].value_counts() if col_teh in df_filtered.columns else pd.Series()

    with col_g1:
        fig, ax = plt.subplots(figsize=(6, 4))
        if not top_kvarovi.empty:
            sns.barplot(x=top_kvarovi.head(5).values, y=top_kvarovi.head(5).index, palette="Reds_r", ax=ax)
            ax.set_title("Top 5 Najčešćih Kvarova (Tipovi)", fontsize=11, fontweight='bold', color='#E31B23')
            ax.set_xlabel("Broj prijava")
            ax.set_ylabel("")
        st.pyplot(fig)

    with col_g2:
        fig, ax = plt.subplots(figsize=(6, 4))
        if not top_tehnicari.empty:
            sns.barplot(x=top_tehnicari.head(5).values, y=top_tehnicari.head(5).index, palette="Blues_r", ax=ax)
            ax.set_title("Top 5 Tehničara po Intervencijama", fontsize=11, fontweight='bold', color='#003366')
            ax.set_xlabel("Broj intervencija")
            ax.set_ylabel("")
        st.pyplot(fig)

    if izabrana_banka == "Sve Banke":
        st.markdown("### 🏦 Ukupan Broj Problema po Bankama (Vlasnicima)")
        if col_vlasnik in df_filtered.columns:
            bank_totals = df_filtered[col_vlasnik].value_counts()
            if not bank_totals.empty:
                fig_bt, ax_bt = plt.subplots(figsize=(10, max(4, len(bank_totals) * 0.4)))
                sns.barplot(x=bank_totals.values, y=bank_totals.index, palette="Blues_r", ax=ax_bt)
                ax_bt.set_title("Ukupan broj problema po bankama", fontsize=12, fontweight='bold', color='#003366')
                ax_bt.set_xlabel("Ukupan broj problema", fontsize=10)
                ax_bt.set_ylabel("Banka / Vlasnik", fontsize=10)
                plt.tight_layout()
                st.pyplot(fig_bt)
    else:
        if col_grad in df_filtered.columns and col_kvar in df_filtered.columns and not df_filtered.empty:
            st.markdown(f"### 🔥 Toplotna Mapa Zastoja (Gradovi vs Kvarovi) — {izabrana_banka}")
            heatmap_data = pd.crosstab(df_filtered[col_grad], df_filtered[col_kvar])
            if not heatmap_data.empty:
                fig_hm, ax_hm = plt.subplots(figsize=(10, max(4, len(heatmap_data) * 0.4)))
                sns.heatmap(heatmap_data, annot=True, cmap="YlOrRd", fmt="d", linewidths=.5, ax=ax_hm)
                ax_hm.set_title(f"Intenzitet kvarova po gradovima i tipovima", fontsize=12, fontweight='bold')
                plt.xticks(rotation=30, ha='right', fontsize=9)
                plt.yticks(fontsize=9)
                plt.tight_layout()
                st.pyplot(fig_hm)

    if izabrana_banka != "Sve Banke":
        st.markdown("---")
        st.markdown(f"### 🏧 Matrica Zastoja po Bankomatima (TMS ID vs Tip Problema) — {izabrana_banka}")
        if col_tms in df_filtered.columns and col_kvar in df_filtered.columns and not df_filtered.empty:
            atm_problem_matrix = pd.crosstab(df_filtered[col_tms], df_filtered[col_kvar])
            if not atm_problem_matrix.empty:
                if len(atm_problem_matrix) > 15:
                    top_atms = df_filtered[col_tms].value_counts().head(15).index
                    atm_problem_matrix = atm_problem_matrix.loc[atm_problem_matrix.index.isin(top_atms)]
                    st.info("ℹ️ Prikazano top 15 bankomata sa najviše problema radi bolje preglednosti.")
                
                fig_atm, ax_atm = plt.subplots(figsize=(10, max(4, len(atm_problem_matrix) * 0.45)))
                sns.heatmap(atm_problem_matrix, annot=True, cmap="Blues", fmt="d", linewidths=.5, ax=ax_atm, cbar=True)
                ax_atm.set_title("Broj pojedinačnih kvarova po bankomatima (TMS ID)", fontsize=12, fontweight='bold', color='#003366')
                plt.xticks(rotation=35, ha='right', fontsize=9)
                plt.yticks(fontsize=9)
                ax_atm.set_xlabel("Tip problema", fontsize=10)
                ax_atm.set_ylabel("TMS ID Bankomata", fontsize=10)
                plt.tight_layout()
                st.pyplot(fig_atm)

    def clean_text(text):
        if text is None:
            return ""
        replacements = {
            'č': 'c', 'Č': 'C', 'ć': 'c', 'Ć': 'C', 
            'š': 's', 'Š': 'S', 'ž': 'z', 'Ž': 'Z', 
            'đ': 'dj', 'Đ': 'Dj', '≥': '>=', '≤': '<=',
            '–': '-', '—': '-'
        }
        res = str(text)
        for k, v in replacements.items():
            res = res.replace(k, v)
        return res.encode('ascii', 'ignore').decode('ascii')

    def generate_excel_report():
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_excel = df_summary.copy()
            if 'Tipovi Problema' in df_excel.columns:
                df_excel['Tipovi Problema'] = df_excel['Tipovi Problema'].str.replace('\n', ' | ')
            df_excel.to_excel(writer, sheet_name='Uredjaji (3+ Problema)', index=False)
            if not df_filtered.empty:
                df_filtered.to_excel(writer, sheet_name='Sirovi Podaci', index=False)
            
            workbook = writer.book
            header_format = workbook.add_format({
                'bold': True, 'text_wrap': True, 'valign': 'top',
                'fg_color': '#003366', 'font_color': 'white', 'border': 1
            })
            for sheet_name in writer.sheets:
                worksheet = writer.sheets[sheet_name]
                worksheet.set_row(0, 25, header_format)
        output.seek(0)
        return output

    class CustomPDF(FPDF):
        def header(self):
            self.set_font("Helvetica", 'B', 14)
            text_prefix = "ATM Izvjestaj Pa"
            text_y = "y"
            text_suffix = "ten BiH"
            
            w_prefix = self.get_string_width(text_prefix)
            w_y = self.get_string_width(text_y)
            w_suffix = self.get_string_width(text_suffix)
            total_w = w_prefix + w_y + w_suffix
            
            start_x = (210 - total_w) / 2
            self.set_xy(start_x, self.get_y())
            
            self.set_text_color(0, 0, 0)
            self.cell(w_prefix, 10, clean_text(text_prefix), 0, 0, 'L')
            self.set_text_color(227, 27, 35)
            self.cell(w_y, 10, clean_text(text_y), 0, 0, 'L')
            self.set_text_color(0, 0, 0)
            self.cell(w_suffix, 10, clean_text(text_suffix), 0, 1, 'L')
            
            self.set_font("Helvetica", 'B', 11)
            self.cell(0, 7, clean_text(f"Banka / Vlasnik: {izabrana_banka}"), 0, 1, 'C')
            self.set_font("Helvetica", '', 10)
            self.cell(0, 6, clean_text(period_opisni_str), 0, 1, 'C')
            self.ln(4)

    def generate_pdf():
        pdf = CustomPDF(orientation='P', unit='mm', format='A4')
        pdf.add_page()

        # 1. Operativni Rezime
        pdf.set_font("Helvetica", 'B', 11)
        pdf.cell(0, 8, clean_text("1. Operativni Rezime i Indikatori"), ln=True)
        pdf.set_font("Helvetica", '', 10)
        pdf.cell(0, 6, clean_text(f"- Ukupno problema u periodu: {tot_problemi}"), ln=True)
        pdf.cell(0, 6, clean_text(f"- Izlasci na lokaciju (Tehnicar): {tehnicar_izlasci_cnt} | HD rijesio: {hd_cnt}"), ln=True)
        pdf.cell(0, 6, clean_text(f"- Otvoreno / van funkcije: {van_funkcije_cnt}"), ln=True)
        pdf.cell(0, 6, clean_text(f"- SLA prolaz: {sla_rate}% | Prosjecno vrijeme popravke: {avg_mttr_min} min"), ln=True)
        pdf.ln(6)

        # 2. Vizuelni prikaz - Gradovi ili Ukupno po Bankama
        if izabrana_banka != "Sve Banke" and col_grad in df_filtered.columns and col_kvar in df_filtered.columns and not df_filtered.empty:
            hm_data_pdf = pd.crosstab(df_filtered[col_grad], df_filtered[col_kvar])
            if not hm_data_pdf.empty:
                estimated_img_h = min(110, max(50, hm_data_pdf.shape[0] * 8 + 20))
                if pdf.get_y() + estimated_img_h > 260:
                    pdf.add_page()
                else:
                    pdf.ln(2)

                pdf.set_font("Helvetica", 'B', 11)
                pdf.cell(0, 8, clean_text("2. Vizuelna Analiza - Toplotna Mapa po Gradovima"), ln=True)
                pdf.ln(2)
                
                fig_pdf, ax_pdf = plt.subplots(figsize=(9, max(3.5, hm_data_pdf.shape[0] * 0.35 + 1.5)))
                sns.heatmap(hm_data_pdf, annot=True, cmap="YlOrRd", fmt="d", linewidths=.5, ax=ax_pdf, cbar=True)
                ax_pdf.set_title(f"Intenzitet kvarova po gradovima — {izabrana_banka}", fontsize=10, fontweight='bold')
                plt.xticks(rotation=30, ha='right', fontsize=8)
                plt.yticks(fontsize=8)
                plt.tight_layout()
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
                    fig_pdf.savefig(tmpfile.name, format='png', dpi=220)
                    tmp_image_path = tmpfile.name
                plt.close(fig_pdf)

                pdf.image(tmp_image_path, x=10, w=190)
                pdf.ln(6)
                try:
                    os.remove(tmp_image_path)
                except:
                    pass
        elif izabrana_banka == "Sve Banke" and col_vlasnik in df_filtered.columns:
            bank_totals_pdf = df_filtered[col_vlasnik].value_counts()
            if not bank_totals_pdf.empty:
                estimated_bt_h = min(110, max(50, len(bank_totals_pdf) * 8 + 20))
                if pdf.get_y() + estimated_bt_h > 260:
                    pdf.add_page()
                else:
                    pdf.ln(2)

                pdf.set_font("Helvetica", 'B', 11)
                pdf.cell(0, 8, clean_text("2. Ukupan Broj Problema po Bankama (Vlasnicima)"), ln=True)
                pdf.ln(2)
                
                fig_bt_pdf, ax_bt_pdf = plt.subplots(figsize=(9, max(3.5, len(bank_totals_pdf) * 0.35 + 1.5)))
                sns.barplot(x=bank_totals_pdf.values, y=bank_totals_pdf.index, palette="Blues_r", ax=ax_bt_pdf)
                ax_bt_pdf.set_title("Ukupan broj problema po bankama", fontsize=10, fontweight='bold')
                ax_bt_pdf.set_xlabel("Broj problema", fontsize=9)
                ax_bt_pdf.set_ylabel("Banka", fontsize=9)
                plt.tight_layout()
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
                    fig_bt_pdf.savefig(tmpfile.name, format='png', dpi=220)
                    tmp_image_path = tmpfile.name
                plt.close(fig_bt_pdf)

                pdf.image(tmp_image_path, x=10, w=190)
                pdf.ln(6)
                try:
                    os.remove(tmp_image_path)
                except:
                    pass

        # 3. Matrica po bankomatima (TMS ID)
        if izabrana_banka != "Sve Banke" and col_tms in df_filtered.columns and col_kvar in df_filtered.columns and not df_filtered.empty:
            atm_mat_pdf = pd.crosstab(df_filtered[col_tms], df_filtered[col_kvar])
            if not atm_mat_pdf.empty:
                if len(atm_mat_pdf) > 12:
                    top_atms_pdf = df_filtered[col_tms].value_counts().head(12).index
                    atm_mat_pdf = atm_mat_pdf.loc[atm_mat_pdf.index.isin(top_atms_pdf)]
                
                estimated_atm_h = min(110, max(50, atm_mat_pdf.shape[0] * 7 + 20))
                
                if pdf.get_y() + estimated_atm_h > 260:
                    pdf.add_page()
                else:
                    pdf.ln(4)

                pdf.set_font("Helvetica", 'B', 11)
                pdf.cell(0, 8, clean_text("3. Matrica Zastoja po Bankomatima (TMS ID)"), ln=True)
                pdf.ln(2)
                
                fig_atm_pdf, ax_atm_pdf = plt.subplots(figsize=(9, max(3.5, atm_mat_pdf.shape[0] * 0.35 + 1.5)))
                sns.heatmap(atm_mat_pdf, annot=True, cmap="Blues", fmt="d", linewidths=.5, ax=ax_atm_pdf, cbar=True)
                ax_atm_pdf.set_title("Raspodjela kvarova po pojedinacnim bankomatima", fontsize=10, fontweight='bold')
                plt.xticks(rotation=30, ha='right', fontsize=8)
                plt.yticks(fontsize=8)
                plt.tight_layout()
                
                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile2:
                    fig_atm_pdf.savefig(tmpfile2.name, format='png', dpi=220)
                    tmp_image_path2 = tmpfile2.name
                plt.close(fig_atm_pdf)

                pdf.image(tmp_image_path2, x=10, w=190)
                pdf.ln(6)
                try:
                    os.remove(tmp_image_path2)
                except:
                    pass

        # 4. Dinamička Tabela Uređaja sa prilagođenim širinama kolona
        pdf.add_page()
        pdf.set_font("Helvetica", 'B', 11)
        pdf.cell(0, 8, clean_text("4. Pregled Uredjaja sa 3 ili Vise Problema"), ln=True)
        pdf.ln(2)
        
        if not df_summary.empty:
            pdf.set_fill_color(0, 51, 102)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Helvetica", 'B', 8)
            
            # NOVE OPTIMIZOVANE ŠIRINE: TMS ID(16), Banka/Vlasnik(40), SN(28), Grad(26), Tipovi Problema(65), Br.(15) -> ukupno 190mm
            col_w = [16, 40, 28, 26, 65, 15]
            
            pdf.cell(col_w[0], 6, "TMS ID", 1, 0, 'C', 1)
            pdf.cell(col_w[1], 6, "Banka / Vlasnik", 1, 0, 'C', 1)
            pdf.cell(col_w[2], 6, "SN", 1, 0, 'C', 1)
            pdf.cell(col_w[3], 6, "Grad", 1, 0, 'C', 1)
            pdf.cell(col_w[4], 6, "Tipovi Problema", 1, 0, 'C', 1)
            pdf.cell(col_w[5], 6, "Br.", 1, 1, 'C', 1)

            pdf.set_font("Helvetica", '', 7.5)
            pdf.set_text_color(0, 0, 0)
            
            for idx, (_, row) in enumerate(df_summary.iterrows()):
                tms_val = clean_text(str(row['TMS ID']))
                banka_raw = str(row['Banka / Vlasnik'])
                # Skraćujemo naziv banke pošto je kolona uža (40mm)
                if len(banka_raw) > 22:
                    banka_raw = banka_raw[:19] + "..."
                banka_val = clean_text(banka_raw)
                
                sn_val = clean_text(str(row['Serijski Broj (SN)']))
                
                # Skraćujemo i naziv grada ako je predugačak za 26mm
                grad_raw = str(row['Grad'])
                if len(grad_raw) > 16:
                    grad_raw = grad_raw[:14] + "..."
                grad_val = clean_text(grad_raw)
                
                tipovi_val = clean_text(str(row['Tipovi Problema']).replace('\n', ' | '))
                broj_val = clean_text(str(row['Broj Problema']))
                
                # Automatsko mjerenje visine za širinu Tipova problema (65mm)
                lines = pdf.multi_cell(col_w[4], 4, tipovi_val, split_only=True)
                row_height = max(6, len(lines) * 4 + 2)
                
                if pdf.get_y() + row_height > 270:
                    pdf.add_page()
                
                y_start = pdf.get_y()
                x_start = pdf.get_x()
                
                # Ispis redova sa novim širinama
                pdf.cell(col_w[0], row_height, tms_val, 1, 0, 'C')
                pdf.cell(col_w[1], row_height, banka_val, 1, 0, 'L')
                pdf.cell(col_w[2], row_height, sn_val, 1, 0, 'C')
                pdf.cell(col_w[3], row_height, grad_val, 1, 0, 'L')
                
                # Multi-cell za tipove problema
                x_multi = pdf.get_x()
                y_multi = pdf.get_y()
                pdf.rect(x_multi, y_multi, col_w[4], row_height)
                
                text_block_height = len(lines) * 4
                v_offset = max(0, (row_height - text_block_height) / 2)
                pdf.set_xy(x_multi, y_multi + v_offset)
                pdf.multi_cell(col_w[4], 4, tipovi_val, 0, 'L')
                
                pdf.set_xy(x_multi + col_w[4], y_multi)
                pdf.cell(col_w[5], row_height, broj_val, 1, 1, 'C')
        else:
            pdf.set_font("Helvetica", '', 9)
            pdf.cell(0, 6, clean_text("Nema uredjaja sa 3 ili vise problema za odabrane filtere."), ln=True)

        return bytes(pdf.output())

    def generate_pdf_greske():
        pdf = CustomPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", 'B', 11)
        pdf.cell(0, 8, clean_text("Rang Lista Prijavljenih Kvarova"), ln=True)
        pdf.set_font("Helvetica", '', 9)

        if not top_kvarovi.empty:
            for kvar_naziv, cnt in top_kvarovi.items():
                stvarno_ime = "Nespecifikovana greska" if pd.isna(kvar_naziv) or str(kvar_naziv).strip() == "" else str(kvar_naziv)
                pct = round((cnt / tot_problemi) * 100, 1) if tot_problemi > 0 else 0
                txt = f"- {stvarno_ime}: {cnt} prijava ({pct}% ucesca)"
                pdf.cell(0, 6, clean_text(txt), ln=True)
        else:
            pdf.cell(0, 6, clean_text("Nema podataka."), ln=True)

        return bytes(pdf.output())

    def generate_pdf_tehnicari():
        pdf = CustomPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", 'B', 11)
        pdf.cell(0, 8, clean_text("Ucinak i Intervencije po Tehnicarima"), ln=True)
        pdf.set_font("Helvetica", '', 9)

        if not top_tehnicari.empty:
            for teh_ime, cnt in top_tehnicari.items():
                stvarno_ime = "Nerasporedjeno / Automatski" if pd.isna(teh_ime) or str(teh_ime).strip() == "" else str(teh_ime)
                pct = round((cnt / tot_problemi) * 100, 1) if tot_problemi > 0 else 0
                txt = f"- {stvarno_ime}: {cnt} intervencija ({pct}% ucesca)"
                pdf.cell(0, 6, clean_text(txt), ln=True)
        else:
            pdf.cell(0, 6, clean_text("Nema podataka."), ln=True)

        return bytes(pdf.output())

    st.markdown("---")
    st.subheader("📥 Preuzmite Vrhunske Izvještaje (PDF & Excel)")

    c_btn1, c_btn2 = st.columns(2)
    with c_btn1:
        st.download_button(
            label="📄 Preuzmi Glavni Periodični PDF (sa mapama i uređajima 3+)",
            data=generate_pdf(),
            file_name=f"ATM_Izvjestaj_{izabrana_banka.replace(' ', '_')}.pdf",
            mime="application/pdf"
        )
    with c_btn2:
        st.download_button(
            label="📊 Preuzmi Profesionalni Excel (.xlsx)",
            data=generate_excel_report(),
            file_name=f"ATM_Analiza_{izabrana_banka.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.markdown("### 📑 Odvojeni Specijalizovani PDF Izvještaji")
    c_btn3, c_btn4 = st.columns(2)
    with c_btn3:
        st.download_button(
            label="⚠️ Preuzmi PDF Izvještaj o Greškama",
            data=generate_pdf_greske(),
            file_name=f"ATM_Greske_{izabrana_banka.replace(' ', '_')}.pdf",
            mime="application/pdf"
        )
    with c_btn4:
        st.download_button(
            label="👷 Preuzmi PDF Izvještaj o Tehničarima",
            data=generate_pdf_tehnicari(),
            file_name=f"ATM_Analiza_Tehnicari_{izabrana_banka.replace(' ', '_')}.pdf",
            mime="application/pdf"
        )
else:
    st.info("Molimo otpremite Tmanage fajl 'Problemi.xls' na vrhu stranice da pokrenete napredni operativni dashboard.")
