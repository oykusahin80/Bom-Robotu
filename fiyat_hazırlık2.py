import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET

st.set_page_config(page_title="BOM Robotu v5.0 - Multi-Currency", layout="wide")

# --- TCMB CANLI KUR ÇEKME ---
@st.cache_data(ttl=3600) # Kurları 1 saat önbelleğe alır
def get_tcmb_rates():
    try:
        response = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml")
        root = ET.fromstring(response.content)
        rates = {'USD': 1.0, 'TRY': 1.0, 'EUR': 1.0}
        
        for currency in root.findall('Currency'):
            code = currency.get('CurrencyCode')
            if code in ['USD', 'EUR']:
                # Efektif Satış kuru baz alınmıştır
                rates[code] = float(currency.find('ForexSelling').text)
        
        # USD bazlı hesaplama için TRY'yi USD'ye bölüyoruz
        rates['TRY_TO_USD'] = 1 / rates['USD']
        rates['EUR_TO_USD'] = rates['EUR'] / rates['USD']
        return rates
    except Exception as e:
        st.error(f"Kur çekilemedi, varsayılan kurlar kullanılacak: {e}")
        return {'TRY_TO_USD': 1/32, 'EUR_TO_USD': 1.08} # Fallback

RATES = get_tcmb_rates()

# --- YARDIMCI FONKSİYONLAR ---
PN_PRIORITY = ['manufacturer part number', 'man code', 'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 'model', 'p/n', 'vendor material']
PRICE_PRIORITY = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar', 'resale', 'net']

def aggressive_clean(text):
    if pd.isna(text) or text == "": return ""
    return re.sub(r'[^A-Z0-9]', '', str(text).upper().strip())

def detect_and_convert_to_usd(value):
    if pd.isna(value) or value == "": return None, ""
    v_str = str(value).upper().replace(" ", "")
    
    # Para birimi tespiti
    currency = "TRY" # Varsayılan
    if "€" in v_str or "EUR" in v_str: currency = "EUR"
    elif "$" in v_str or "USD" in v_str: currency = "USD"
    
    # Sayıyı temizle
    v = re.sub(r'[^\d.,]', '', v_str)
    if ',' in v and '.' in v: v = v.replace('.', '').replace(',', '.')
    elif ',' in v: v = v.replace(',', '.')
    
    try:
        num = float(v)
        if currency == "EUR":
            return num * RATES['EUR_TO_USD'], " (€)"
        elif currency == "TRY":
            return num * RATES['TRY_TO_USD'], " (TL)"
        return num, " ($)"
    except:
        return None, ""

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
            for i, row in df.head(40).iterrows():
                if any(kw in " ".join(map(str, row.values)).lower() for kw in PN_PRIORITY):
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
st.title("📊 Akıllı BOM Robotu v5.0 (Multi-Currency)")
st.sidebar.info(f"🏦 TCMB Güncel Kur (USD Bazlı):\n\n1 EUR = {RATES['EUR_TO_USD']:.4f} USD\n1 TRY = {RATES['TRY_TO_USD']:.4f} USD")

master_file = st.file_uploader("1. Master BOM (Excel)", type=['xlsx', 'xls'])
supplier_files = st.file_uploader("2. Teklifler (Excel veya PDF)", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True)

if master_file and supplier_files:
    df_master = smart_load(master_file)
    if df_master is not None:
        m_pn_col = find_best_column(df_master.columns, PN_PRIORITY)
        
        if m_pn_col:
            df_master['MATCH_KEY'] = df_master[m_pn_col].apply(aggressive_clean)
            final_df = df_master.copy()
            usd_price_cols = []

            for s_file in supplier_files:
                df_sup = smart_load(s_file)
                if df_sup is not None:
                    s_pn = find_best_column(df_sup.columns, PN_PRIORITY)
                    s_pr = find_best_column(df_sup.columns, PRICE_PRIORITY)
                    
                    if s_pn and s_pr:
                        s_name = os.path.splitext(s_file.name)[0]
                        p_col_usd = f"{s_name}_USD"
                        p_col_display = f"{s_name}_Orijinal"
                        
                        # Fiyat işleme
                        temp_sup = df_sup[[s_pn, s_pr]].copy()
                        temp_sup['MATCH_KEY'] = temp_sup[s_pn].apply(aggressive_clean)
                        
                        converted_data = temp_sup[s_pr].apply(detect_and_convert_to_usd)
                        temp_sup[p_col_usd] = converted_data.apply(lambda x: x[0])
                        temp_sup[p_col_display] = temp_sup[s_pr].astype(str) + converted_data.apply(lambda x: x[1])
                        
                        temp_sup = temp_sup.drop_duplicates('MATCH_KEY')
                        final_df = pd.merge(final_df, temp_sup[['MATCH_KEY', p_col_usd, p_col_display]], on='MATCH_KEY', how='left')
                        usd_price_cols.append(p_col_usd)
                        st.write(f"✅ {s_file.name} eşleşti.")

            if usd_price_cols:
                final_df['En Düşük ($)'] = final_df[usd_price_cols].min(axis=1)
                final_df['Kazanan'] = final_df[usd_price_cols].idxmin(axis=1).fillna("Yok").str.replace("_USD", "")
                
                # Temiz görünüm için teknik sütunları gizle
                display_df = final_df.drop(columns=['MATCH_KEY'] + usd_price_cols)
                st.subheader("Karşılaştırma Sonucu")
                st.dataframe(display_df)
                
                csv = display_df.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📩 Raporu İndir", csv, "Fiyat_Analiz.csv", "text/csv")
