import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="BOM Robotu v6.0 - Arrow Fix", layout="wide")

# --- 1. TCMB KUR SERVİSİ (Geliştirilmiş) ---
@st.cache_data(ttl=3600)
def get_live_rates():
    try:
        response = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml", timeout=15)
        root = ET.fromstring(response.content)
        rates = {'USD': 33.0, 'EUR': 36.0} # Fallback
        for curr in root.findall('Currency'):
            code = curr.get('CurrencyCode')
            if code in ['USD', 'EUR']:
                val = curr.find('ForexSelling').text
                if val: rates[code] = float(val)
        
        return {
            'EUR_USD': rates['EUR'] / rates['USD'],
            'TRY_USD': 1 / rates['USD'],
            'USD': rates['USD'],
            'EUR': rates['EUR']
        }
    except:
        return {'EUR_USD': 1.08, 'TRY_USD': 1/33.0, 'USD': 33.0, 'EUR': 36.0}

L_RATES = get_live_rates()

# --- 2. SÜTUN TANIMLARI ---
PN_PRIO = ['manufacturer part number', 'mfr part number', 'üretici parça kodu', 'part number', 'pn', 'p/n', 'mfr p/n', 'model']
PR_PRIO = ['unit price', 'birim fiyat', 'price', 'fiyat', 'net price', 'resale']
CUR_PRIO = ['currency', 'döviz', 'birim', 'curr'] # Para birimi sütunu için
QTY_PRIO = ['qty', 'adet', 'miktar', 'quantity']

# --- 3. HASSAS FİYAT İŞLEME ---
def convert_to_usd(price_val, currency_val=None):
    if pd.isna(price_val) or str(price_val).strip() == "": return None
    
    # 1. Adım: Fiyatı temizle (0,0245 -> 0.0245)
    p_str = str(price_val).upper().replace(" ", "")
    p_clean = re.sub(r'[^0-9,.]', '', p_str)
    if ',' in p_clean and '.' in p_clean: p_clean = p_clean.replace('.', '').replace(',', '.')
    elif ',' in p_clean: p_clean = p_clean.replace(',', '.')
    
    try:
        price_num = float(p_clean)
    except:
        return None

    # 2. Adım: Para birimini tespit et
    # Önce hücre içine bak, yoksa yandaki currency sütununa bak
    unit = "USD" # Global distribütörlerde varsayılan USD olmalı
    combined_text = (p_str + str(currency_val or "")).upper()
    
    if "€" in combined_text or "EUR" in combined_text:
        return price_num * L_RATES['EUR_USD']
    elif "TL" in combined_text or "TRY" in combined_text:
        return price_num * L_RATES['TRY_USD']
    elif "$" in combined_text or "USD" in combined_text:
        return price_num
    
    # Hiçbir şey bulunamadıysa ve sayı çok küçükse (0.0245 gibi), muhtemelen EUR veya USD'dir (TL olamaz)
    if price_num < 2: # 2 TL'nin altındaki ürünler genellikle dövizdir
        return price_num # USD varsay
        
    return price_num * L_RATES['TRY_USD'] # Yüksek rakamlar TL varsayılır

def find_col(cols, prio):
    for p in prio:
        for c in cols:
            if p in str(c).lower(): return c
    return None

def smart_read(file):
    ext = os.path.splitext(file.name)[1].lower()
    try:
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file, header=None)
            for i, row in df.head(40).iterrows():
                if any(p in str(row.values).lower() for p in PN_PRIO):
                    return pd.read_excel(file, header=i)
            return pd.read_excel(file)
        elif ext == '.pdf':
            with pdfplumber.open(file) as pdf:
                rows = []
                for pg in pdf.pages:
                    tbl = pg.extract_table()
                    if tbl: rows.extend(tbl)
                if rows:
                    df_p = pd.DataFrame(rows)
                    df_p.columns = df_p.iloc[0]
                    return df_p.iloc[1:].reset_index(drop=True)
    except: return None

# --- 4. ARAYÜZ ---
st.title("📊 Akıllı BOM Robotu v6.0")
st.sidebar.markdown(f"""
### 🏦 Uygulanan Kurlar
- **1 USD:** {L_RATES['USD']:.2f} TL
- **1 EUR:** {L_RATES['EUR']:.2f} TL
- **EUR/USD:** {L_RATES['EUR_USD']:.4f}
""")

m_file = st.file_uploader("1. Master Liste", type=['xlsx', 'xls'])
s_files = st.file_uploader("2. Teklifler", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True)

if m_file and s_files:
    m_df = smart_read(m_file)
    if m_df is not None:
        m_pn = find_col(m_df.columns, PN_PRIO)
        m_qty = find_col(m_df.columns, QTY_PRIO)
        
        if m_pn:
            m_df['M_KEY'] = m_df[m_pn].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper()) if pd.notna(x) else "")
            res_df = m_df.copy()
            s_cols = []

            for f in s_files:
                s_df = smart_read(f)
                if s_df is not None:
                    s_pn = find_col(s_df.columns, PN_PRIO)
                    s_pr = find_col(s_df.columns, PR_PRIO)
                    s_cur = find_col(s_df.columns, CUR_PRIO) # Para birimi sütunu
                    
                    if s_pn and s_pr:
                        s_key = os.path.splitext(f.name)[0][:15]
                        c_name = f"{s_key} ($)"
                        
                        temp = s_df.copy()
                        temp['M_KEY'] = temp[s_pn].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper()) if pd.notna(x) else "")
                        
                        # Para birimi sütununu da işleme dahil et
                        temp[c_name] = temp.apply(lambda r: convert_to_usd(r[s_pr], r[s_cur] if s_cur else None), axis=1)
                        
                        temp = temp.dropna(subset=[c_name]).drop_duplicates('M_KEY')
                        res_df = pd.merge(res_df, temp[['M_KEY', c_name]], on='M_KEY', how='left')
                        s_cols.append(c_name)
                        st.success(f"✅ {f.name} başarıyla eşleşti.")

            if s_cols:
                res_df['En Düşük ($)'] = res_df[s_cols].min(axis=1)
                res_df['Kazanan'] = res_df[s_cols].idxmin(axis=1).str.replace(" ($)", "", regex=False)
                
                if m_qty:
                    res_df[m_qty] = pd.to_numeric(res_df[m_qty], errors='coerce').fillna(0)
                    res_df['Toplam ($)'] = (res_df['En Düşük ($)'] * res_df[m_qty]).round(4)

                st.dataframe(res_df.drop(columns=['M_KEY']), use_container_width=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    res_df.drop(columns=['M_KEY']).to_excel(writer, index=False, sheet_name='Analiz')
                st.download_button("📩 Excel Raporu", output.getvalue(), "Analiz.xlsx")
