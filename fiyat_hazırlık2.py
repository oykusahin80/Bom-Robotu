import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="BOM Robotu v5.4 - Fix Excel", layout="wide")

# --- TCMB CANLI KUR ÇEKME ---
@st.cache_data(ttl=3600)
def get_tcmb_rates():
    try:
        response = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml", timeout=10)
        root = ET.fromstring(response.content)
        rates = {'USD': 1.0, 'TRY': 1.0, 'EUR': 1.0}
        for currency in root.findall('Currency'):
            code = currency.get('CurrencyCode')
            if code in ['USD', 'EUR']:
                val = currency.find('ForexSelling').text
                if val: rates[code] = float(val)
        u_rate = rates.get('USD', 32.5) 
        rates['TRY_TO_USD'] = 1 / u_rate
        rates['EUR_TO_USD'] = rates['EUR'] / u_rate
        return rates
    except:
        return {'TRY_TO_USD': 1/32.5, 'EUR_TO_USD': 1.08}

RATES = get_tcmb_rates()

# --- YARDIMCI FONKSİYONLAR ---
PN_PRIORITY = ['manufacturer part number', 'man code', 'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 'model', 'p/n', 'vendor material']
PRICE_PRIORITY = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar', 'resale', 'net']
QTY_PRIORITY = ['qty', 'adet', 'miktar', 'quantity']

def aggressive_clean(text):
    if pd.isna(text) or text == "": return ""
    return re.sub(r'[^A-Z0-9]', '', str(text).upper().strip())

def convert_any_to_usd(value):
    if pd.isna(value) or str(value).strip() == "": return None
    v_str = str(value).upper().replace(" ", "").replace("TL", "TRY")
    currency = "TRY"
    if "€" in v_str or "EUR" in v_str: currency = "EUR"
    elif "$" in v_str or "USD" in v_str: currency = "USD"
    v = re.sub(r'[^\d.,]', '', v_str)
    if ',' in v and '.' in v: v = v.replace('.', '').replace(',', '.')
    elif ',' in v: v = v.replace(',', '.')
    try:
        num = float(v)
        if currency == "EUR": return round(num * RATES['EUR_TO_USD'], 4)
        if currency == "TRY": return round(num * RATES['TRY_TO_USD'], 4)
        return round(num, 4)
    except: return None

def find_best_column(columns, priority_list):
    for kw in priority_list:
        for col in columns:
            if kw in str(col).lower(): return col
    return None

def smart_load(file):
    ext = os.path.splitext(file.name)[1].lower()
    try:
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file, header=None)
            for i, row in df.head(30).iterrows():
                if any(kw in str(row.values).lower() for kw in PN_PRIORITY):
                    return pd.read_excel(file, header=i)
            return pd.read_excel(file)
        elif ext == '.pdf':
            all_rows = []
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table: all_rows.extend(table)
            if all_rows:
                df_pdf = pd.DataFrame(all_rows)
                df_pdf.columns = df_pdf.iloc[0]
                return df_pdf.iloc[1:].reset_index(drop=True)
    except: return None

# --- ARAYÜZ ---
st.title("📊 Profesyonel BOM Analiz Sistemi v5.4")

master_file = st.file_uploader("1. Master BOM Listesi", type=['xlsx', 'xls'])
supplier_files = st.file_uploader("2. Tedarikçi Teklifleri", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True)

if master_file and supplier_files:
    df_master = smart_load(master_file)
    if df_master is not None:
        m_pn_col = find_best_column(df_master.columns, PN_PRIORITY)
        m_qty_col = find_best_column(df_master.columns, QTY_PRIORITY)
        
        if m_pn_col:
            df_master['MATCH_KEY'] = df_master[m_pn_col].apply(aggressive_clean)
            final_df = df_master.copy()
            usd_cols = []

            for s_file in supplier_files:
                df_sup = smart_load(s_file)
                if df_sup is not None:
                    s_pn = find_best_column(df_sup.columns, PN_PRIORITY)
                    s_pr = find_best_column(df_sup.columns, PRICE_PRIORITY)
                    if s_pn and s_pr:
                        s_name = os.path.splitext(s_file.name)[0][:15]
                        u_col = f"Fiyat_{s_name} ($)"
                        temp_sup = df_sup[[s_pn, s_pr]].copy()
                        temp_sup['MATCH_KEY'] = temp_sup[s_pn].apply(aggressive_clean)
                        temp_sup[u_col] = temp_sup[s_pr].apply(convert_any_to_usd)
                        temp_sup = temp_sup.dropna(subset=[u_col]).drop_duplicates('MATCH_KEY')
                        final_df = pd.merge(final_df, temp_sup[['MATCH_KEY', u_col]], on='MATCH_KEY', how='left')
                        usd_cols.append(u_col)

            if usd_cols:
                # Hesaplamalar
                final_df['En Düşük Birim ($)'] = final_df[usd_cols].min(axis=1)
                
                def get_winner(row):
                    valid = row[usd_cols].dropna()
                    if valid.empty: return "Teklif Yok"
                    return valid.idxmin().replace("Fiyat_", "").replace(" ($)", "")
                
                final_df['Kazanan Tedarikçi'] = final_df.apply(get_winner, axis=1)

                if m_qty_col:
                    final_df[m_qty_col] = pd.to_numeric(final_df[m_qty_col], errors='coerce').fillna(0)
                    final_df['Toplam Maliyet ($)'] = (final_df['En Düşük Birim ($)'] * final_df[m_qty_col]).round(2)

                # Tablo Görünümü
                display_df = final_df.drop(columns=['MATCH_KEY'])
                st.subheader("✅ Analiz Tamamlandı")
                st.dataframe(display_df, use_container_width=True)

                # --- DOĞRU EXCEL ÇIKTISI (XLSX) ---
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    display_df.to_excel(writer, index=False, sheet_name='Fiyat Analizi')
                    # Sütun genişliklerini otomatik ayarla
                    worksheet = writer.sheets['Fiyat Analizi']
                    for i, col in enumerate(display_df.columns):
                        column_len = max(display_df[col].astype(str).str.len().max(), len(col)) + 2
                        worksheet.set_column(i, i, column_len)
                
                excel_data = output.getvalue()
                st.download_button(
                    label="📩 Analiz Raporunu Excel Olarak İndir",
                    data=excel_data,
                    file_name="BOM_Analiz_Raporu.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
