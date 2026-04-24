import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="BOM Robotu v6.2 - Kesin Çözüm", layout="wide")

# --- 1. TCMB KUR SERVİSİ ---
@st.cache_data(ttl=3600)
def get_live_rates():
    try:
        response = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml", timeout=10)
        root = ET.fromstring(response.content)
        rates = {'USD': 32.5, 'EUR': 35.5}
        for curr in root.findall('Currency'):
            code = curr.get('CurrencyCode')
            if code in ['USD', 'EUR']:
                val = curr.find('ForexSelling').text
                if val: rates[code] = float(val)
        return {
            'EUR_USD': round(rates['EUR'] / rates['USD'], 6),
            'TRY_USD': round(1 / rates['USD'], 6),
            'USD': rates['USD'],
            'EUR': rates['EUR']
        }
    except:
        return {'EUR_USD': 1.08, 'TRY_USD': 1/32.5, 'USD': 32.5, 'EUR': 35.5}

L_RATES = get_live_rates()

# --- 2. HASSAS FİYAT İŞLEME (0.0245 Fix) ---
def parse_price_usd(p_val, file_name=""):
    if pd.isna(p_val) or str(p_val).strip() == "": return None
    
    s = str(p_val).upper().replace(" ", "").replace("TL", "TRY")
    # Sayıyı temizle (virgül/nokta standardı)
    c = re.sub(r'[^0-9,.]', '', s)
    if ',' in c and '.' in c: c = c.replace('.', '').replace(',', '.')
    elif ',' in c: c = c.replace(',', '.')
    
    try:
        n = float(c)
        # Arrow veya global bir distribütör ise ve sayı küçükse EUR/USD varsay
        # Eğer hücre içinde EUR/€ varsa veya dosya adında ARROW geçiyorsa
        is_euro = "EUR" in s or "€" in s or "ARROW" in str(file_name).upper()
        
        if is_euro:
            return round(n * L_RATES['EUR_USD'], 6)
        if "TRY" in s or "TL" in s:
            return round(n * L_RATES['TRY_USD'], 6)
        
        # Birim yoksa ama sayı çok küçükse (global tedarikçi mantığı)
        if n < 2 and "USD" not in s and "$" not in s:
            return round(n * L_RATES['EUR_USD'], 6)
            
        return round(n, 6)
    except: return None

# --- 3. SÜTUN BULMA VE OKUMA ---
def find_column(cols, keywords):
    for k in keywords:
        for c in cols:
            if k in str(c).lower(): return c
    return None

def smart_read(file):
    ext = os.path.splitext(file.name)[1].lower()
    try:
        if ext in ['.xlsx', '.xls']:
            df_raw = pd.read_excel(file, header=None)
            for i, row in df_raw.head(60).iterrows():
                row_str = " ".join(map(str, row.values)).lower()
                if any(k in row_str for k in ['part', 'kod', 'pn', 'p/n', 'mfr']):
                    return pd.read_excel(file, header=i)
            return pd.read_excel(file)
        elif ext == '.pdf':
            with pdfplumber.open(file) as pdf:
                all_r = []
                for p in pdf.pages:
                    t = p.extract_table()
                    if t: all_r.extend(t)
                if all_r:
                    df = pd.DataFrame(all_r)
                    df.columns = df.iloc[0]
                    return df.iloc[1:].reset_index(drop=True)
    except: return None

# --- 4. ANA ARAYÜZ ---
st.title("🚀 BOM Robotu v6.2 - Kesin Çözüm")
st.sidebar.info(f"Parite: 1 EUR = {L_RATES['EUR_USD']} USD")

m_file = st.file_uploader("1. Master Liste", type=['xlsx', 'xls'])
s_files = st.file_uploader("2. Teklifler", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True)

if m_file and s_files:
    m_df = smart_read(m_file)
    if m_df is not None:
        m_pn = find_column(m_df.columns, ['part number', 'pn', 'p/n', 'üretici kodu', 'kod'])
        m_qty = find_column(m_df.columns, ['qty', 'adet', 'miktar', 'quantity'])
        
        if m_pn:
            m_df['M_KEY'] = m_df[m_pn].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper()) if pd.notna(x) else "")
            res_df = m_df.copy()
            s_cols = []

            for f in s_files:
                s_df = smart_read(f)
                if s_df is not None:
                    s_pn = find_column(s_df.columns, ['part number', 'pn', 'p/n', 'mfr part', 'kod'])
                    s_pr = find_column(s_df.columns, ['unit price', 'fiyat', 'price', 'tutar'])
                    
                    if s_pn and s_pr:
                        s_key = f"{os.path.splitext(f.name)[0][:10]}_($)"
                        temp = s_df.copy()
                        temp['M_KEY'] = temp[s_pn].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper()) if pd.notna(x) else "")
                        temp[s_key] = temp[s_pr].apply(lambda x: parse_price_usd(x, f.name))
                        
                        temp = temp.dropna(subset=[s_key]).drop_duplicates('M_KEY')
                        res_df = pd.merge(res_df, temp[['M_KEY', s_key]], on='M_KEY', how='left')
                        s_cols.append(s_key)
                        st.success(f"✔️ {f.name} yüklendi.")

            if s_cols:
                # --- ÇÖKMEYİ ÖNLEYEN SATIR BAZLI HESAPLAMA ---
                def calculate_row(row):
                    valid_prices = row[s_cols].dropna()
                    if valid_prices.empty:
                        return pd.Series([None, "Teklif Yok"], index=['En Düşük ($)', 'Kazanan'])
                    
                    min_val = valid_prices.min()
                    winner = valid_prices.idxmin().replace("_($)", "")
                    return pd.Series([min_val, winner], index=['En Düşük ($)', 'Kazanan'])

                res_df[['En Düşük ($)', 'Kazanan']] = res_df.apply(calculate_row, axis=1)
                
                if m_qty:
                    res_df[m_qty] = pd.to_numeric(res_df[m_qty], errors='coerce').fillna(0)
                    res_df['Toplam Maliyet ($)'] = (res_df['En Düşük ($)'] * res_df[m_qty]).round(4)

                st.subheader("🏁 Karşılaştırma Sonucu")
                st.dataframe(res_df.drop(columns=['M_KEY']), use_container_width=True)

                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                    res_df.drop(columns=['M_KEY']).to_excel(wr, index=False)
                st.download_button("📩 Excel İndir", out.getvalue(), "BOM_Raporu.xlsx")
