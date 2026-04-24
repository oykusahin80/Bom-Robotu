import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="BOM Robotu v6.4 - Final", layout="wide")

# --- 1. TCMB KUR ÇEKİMİ (SADE VE NET) ---
@st.cache_data(ttl=3600)
def get_tcmb_rates():
    try:
        response = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml", timeout=10)
        root = ET.fromstring(response.content)
        rates = {'USD': 33.0, 'EUR': 36.0}
        for curr in root.findall('Currency'):
            code = curr.get('CurrencyCode')
            if code in ['USD', 'EUR']:
                val = curr.find('ForexSelling').text
                if val: rates[code] = float(val)
        return {'EUR_USD': rates['EUR'] / rates['USD'], 'TRY_USD': 1 / rates['USD']}
    except:
        return {'EUR_USD': 1.09, 'TRY_USD': 1/33.0}

L_RATES = get_tcmb_rates()

# --- 2. ILK KODLARDAKİ ESNEK TEMİZLEME MANTIĞI ---
def aggressive_clean(text):
    if pd.isna(text) or text == "": return ""
    return re.sub(r'[^A-Z0-9]', '', str(text).upper().strip())

def parse_to_usd(val, is_arrow=False):
    if pd.isna(val) or str(val).strip() == "": return None
    v = str(val).upper().replace(" ", "")
    
    # Rakam ayıklama (v4.5'teki gibi ama 6 basamak hassasiyetle)
    c = re.sub(r'[^0-9,.]', '', v)
    if ',' in c and '.' in c: c = c.replace('.', '').replace(',', '.')
    elif ',' in c: c = c.replace(',', '.')
    
    try:
        n = float(c)
        # Kur dönüşümü
        if "EUR" in v or "€" in v or is_arrow: # Arrow ise direkt EUR/USD paritesi
            return n * L_RATES['EUR_USD']
        if "TL" in v or "TRY" in v:
            return n * L_RATES['TRY_USD']
        return n # Varsayılan USD
    except: return None

# --- 3. ILK KODLARDAKİ GENİŞ SÜTUN TANIMA ---
PN_PRIORITY = ['manufacturer part number', 'man code', 'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 'model', 'p/n', 'vendor material', 'mfr part']
PRICE_PRIORITY = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar', 'resale', 'net', 'amount']
QTY_PRIORITY = ['qty', 'adet', 'miktar', 'quantity']

def find_best_col(columns, priority_list):
    for kw in priority_list:
        for col in columns:
            if kw in str(col).lower(): return col
    return None

def smart_load(file):
    ext = os.path.splitext(file.name)[1].lower()
    try:
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file, header=None)
            # İlk 50 satırda anahtar kelime ara (v4.5 mantığı)
            for i, row in df.head(50).iterrows():
                row_str = " ".join(map(str, row.values)).lower()
                if any(kw in row_str for kw in PN_PRIORITY):
                    return pd.read_excel(file, header=i)
            return pd.read_excel(file)
        elif ext == '.pdf':
            with pdfplumber.open(file) as pdf:
                all_rows = []
                for page in pdf.pages:
                    table = page.extract_table()
                    if table: all_rows.extend(table)
            if all_rows:
                df_pdf = pd.DataFrame(all_rows)
                for i, row in df_pdf.head(20).iterrows():
                    if any(kw in " ".join(map(str, row.values)).lower() for kw in PN_PRIORITY):
                        df_pdf.columns = df_pdf.iloc[i]
                        return df_pdf.iloc[i+1:].reset_index(drop=True)
                return df_pdf
    except: return None

# --- 4. ARAYÜZ VE ANALİZ ---
st.title("📊 Akıllı BOM Robotu v6.4 (Bütüncül Sürüm)")

master_file = st.file_uploader("1. Master BOM (Excel)", type=['xlsx', 'xls'])
supplier_files = st.file_uploader("2. Teklifler", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True)

if master_file and supplier_files:
    df_master = smart_load(master_file)
    if df_master is not None:
        m_pn = find_best_col(df_master.columns, PN_PRIORITY)
        m_qty = find_best_col(df_master.columns, QTY_PRIORITY)
        
        if m_pn:
            df_master['MATCH_KEY'] = df_master[m_pn].apply(aggressive_clean)
            final_df = df_master.copy()
            price_cols = []

            for s_file in supplier_files:
                df_sup = smart_load(s_file)
                if df_sup is not None:
                    s_pn = find_best_col(df_sup.columns, PN_PRIORITY)
                    s_pr = find_best_col(df_sup.columns, PRICE_PRIORITY)
                    
                    if s_pn and s_pr:
                        s_name = os.path.splitext(s_file.name)[0][:15]
                        p_col = f"{s_name}_($)"
                        is_arrow = "ARROW" in s_file.name.upper()
                        
                        temp_sup = df_sup[[s_pn, s_pr]].copy()
                        temp_sup['MATCH_KEY'] = temp_sup[s_pn].apply(aggressive_clean)
                        temp_sup[p_col] = temp_sup[s_pr].apply(lambda x: parse_to_usd(x, is_arrow))
                        
                        temp_sup = temp_sup.dropna(subset=[p_col]).drop_duplicates('MATCH_KEY')
                        final_df = pd.merge(final_df, temp_sup[['MATCH_KEY', p_col]], on='MATCH_KEY', how='left')
                        price_cols.append(p_col)
                        st.info(f"✅ {s_file.name} eşleşti.")

            if price_cols:
                # Kazanan ve En Düşük (Satır bazlı korumalı hesaplama)
                def get_results(row):
                    valid = row[price_cols].dropna()
                    if valid.empty: return pd.Series([None, "Yok"], index=['Min', 'Win'])
                    return pd.Series([valid.min(), valid.idxmin().replace("_($)", "")], index=['Min', 'Win'])

                final_df[['En Düşük ($)', 'Kazanan']] = final_df.apply(get_results, axis=1)
                
                if m_qty:
                    final_df[m_qty] = pd.to_numeric(final_df[m_qty], errors='coerce').fillna(0)
                    final_df['Toplam ($)'] = (final_df['En Düşük ($)'] * final_df[m_qty]).round(4)

                st.dataframe(final_df.drop(columns=['MATCH_KEY']))
                
                # Gerçek Excel Çıktısı
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    final_df.drop(columns=['MATCH_KEY']).to_excel(writer, index=False)
                st.download_button("📩 Excel Raporunu İndir", out.getvalue(), "BOM_Analiz.xlsx")
