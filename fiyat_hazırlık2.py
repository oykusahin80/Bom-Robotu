import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

# --- SAYFA AYARLARI ---
st.set_page_config(page_title="BOM Robotu v5.7 - Kararlı Sürüm", layout="wide")

# --- TCMB KUR SERVİSİ ---
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
        return {
            'TRY_TO_USD': 1 / u_rate,
            'EUR_TO_USD': rates['EUR'] / u_rate
        }
    except:
        return {'TRY_TO_USD': 1/32.5, 'EUR_TO_USD': 1.08}

RATES = get_tcmb_rates()

# --- GELİŞMİŞ FİYAT TEMİZLEME (0,0245 FIX) ---
def clean_and_convert_to_usd(value):
    if pd.isna(value) or str(value).strip() == "": return None
    
    text = str(value).upper().replace(" ", "").replace("TL", "TRY")
    
    # Para birimi tespiti
    unit = "TRY"
    if "€" in text or "EUR" in text: unit = "EUR"
    elif "$" in text or "USD" in text: unit = "USD"
    
    # Rakam dışı her şeyi ayıkla (sadece rakam, nokta ve virgül kalsın)
    clean_val = re.sub(r'[^0-9,.]', '', text)
    
    # TR Formatı (0,0245 veya 1.250,50) -> Standart (0.0245 veya 1250.50)
    if ',' in clean_val and '.' in clean_val:
        clean_val = clean_val.replace('.', '').replace(',', '.')
    elif ',' in clean_val:
        clean_val = clean_val.replace(',', '.')
    
    try:
        num = float(clean_val)
        if unit == "EUR": return round(num * RATES['EUR_TO_USD'], 6)
        if unit == "TRY": return round(num * RATES['TRY_TO_USD'], 6)
        return round(num, 6)
    except:
        return None

# --- SÜTUN BULUCU VE VERİ YÜKLEME ---
PN_PRIORITY = ['manufacturer part number', 'man code', 'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 'model', 'p/n']
PRICE_PRIORITY = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar', 'net']
QTY_PRIORITY = ['qty', 'adet', 'miktar', 'quantity']

def find_col(cols, priority):
    for p in priority:
        for c in cols:
            if p in str(c).lower(): return c
    return None

def smart_load(file):
    ext = os.path.splitext(file.name)[1].lower()
    try:
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file, header=None)
            for i, row in df.head(30).iterrows():
                if any(p in str(row.values).lower() for p in PN_PRIORITY):
                    return pd.read_excel(file, header=i)
            return pd.read_excel(file)
        elif ext == '.pdf':
            all_rows = []
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table: all_rows.extend(table)
            if all_rows:
                df_p = pd.DataFrame(all_rows)
                df_p.columns = df_p.iloc[0]
                return df_p.iloc[1:].reset_index(drop=True)
    except: return None
    return None

# --- ANA PROGRAM ---
st.title("🚀 Akıllı BOM Karşılaştırma (v5.7 Stable)")

m_file = st.file_uploader("1. Master BOM (Excel)", type=['xlsx', 'xls'])
s_files = st.file_uploader("2. Tedarikçi Teklifleri (Toplu)", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True)

if m_file and s_files:
    df_m = smart_load(m_file)
    if df_m is not None:
        pn_col = find_col(df_m.columns, PN_PRIORITY)
        qty_col = find_col(df_m.columns, QTY_PRIORITY)
        
        if pn_col:
            df_m['MATCH_KEY'] = df_m[pn_col].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper().strip()) if pd.notna(x) else "")
            report_df = df_m.copy()
            supplier_cols = []

            for f in s_files:
                df_s = smart_load(f)
                if df_s is not None:
                    s_pn = find_col(df_s.columns, PN_PRIORITY)
                    s_pr = find_col(df_s.columns, PRICE_PRIORITY)
                    if s_pn and s_pr:
                        s_name = os.path.splitext(f.name)[0][:15]
                        col_name = f"{s_name} ($)"
                        
                        temp_s = df_s[[s_pn, s_pr]].copy()
                        temp_s['MATCH_KEY'] = temp_s[s_pn].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper().strip()) if pd.notna(x) else "")
                        temp_s[col_name] = temp_s[s_pr].apply(clean_and_convert_to_usd)
                        
                        temp_s = temp_s.dropna(subset=[col_name]).drop_duplicates('MATCH_KEY')
                        report_df = pd.merge(report_df, temp_s[['MATCH_KEY', col_name]], on='MATCH_KEY', how='left')
                        supplier_cols.append(col_name)

            if supplier_cols:
                # En ucuz olanı bul ve çökme riskine karşı kontrol et
                report_df['En Düşük ($)'] = report_df[supplier_cols].min(axis=1)
                
                def get_winner(row):
                    valid = row[supplier_cols].dropna()
                    if valid.empty: return "Teklif Yok"
                    return valid.idxmin().replace(" ($)", "")
                
                report_df['Kazanan Tedarikçi'] = report_df.apply(get_winner, axis=1)

                if qty_col:
                    report_df[qty_col] = pd.to_numeric(report_df[qty_col], errors='coerce').fillna(0)
                    report_df['Toplam Maliyet ($)'] = (report_df['En Düşük ($)'] * report_df[qty_col]).round(4)

                final_view = report_df.drop(columns=['MATCH_KEY'])
                st.dataframe(final_view, use_container_width=True)

                # Excel İndirme İşlemi
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    final_view.to_excel(writer, index=False, sheet_name='Fiyat Analizi')
                
                st.download_button("📩 Excel Raporunu İndir", output.getvalue(), "BOM_Fiyat_Analizi.xlsx")
        else:
            st.error("Master listede PN sütunu bulunamadı.")
