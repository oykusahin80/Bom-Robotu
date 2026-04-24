import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="BOM Robotu v5.9 - Arrow Fix", layout="wide")

# --- 1. KUR SERVİSİ ---
@st.cache_data(ttl=3600)
def get_live_rates():
    try:
        response = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml", timeout=10)
        root = ET.fromstring(response.content)
        rates = {'USD': 1.0, 'TRY': 1.0, 'EUR': 1.0}
        for currency in root.findall('Currency'):
            code = currency.get('CurrencyCode')
            if code in ['USD', 'EUR']:
                val = currency.find('ForexSelling').text
                if val: rates[code] = float(val)
        usd_val = rates['USD']
        eur_val = rates['EUR']
        return {'EUR_TO_USD': eur_val / usd_val, 'TRY_TO_USD': 1 / usd_val}
    except:
        return {'EUR_TO_USD': 1.08, 'TRY_TO_USD': 1/32.5}

L_RATES = get_live_rates()

# --- 2. SÜTUN TANIMA LİSTELERİ (Genişletildi) ---
# Arrow ve diğer global distribütörlerin kullandığı tüm varyasyonlar eklendi
PN_PRIO = [
    'manufacturer part number', 'mfr part number', 'mfr part #', 'man code', 
    'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 
    'model', 'p/n', 'mfr p/n', 'customer part number', 'supplier part number'
]
PR_PRIO = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar', 'net', 'resale', 'amount']
QTY_PRIO = ['qty', 'adet', 'miktar', 'quantity', 'requested qty']

# --- 3. HASSAS FİYAT VE VERİ İŞLEME ---
def parse_price(val):
    if pd.isna(val) or str(val).strip() == "": return None
    v_s = str(val).upper().replace(" ", "").replace("TL", "TRY")
    unit = "TRY"
    if "€" in v_s or "EUR" in v_s: unit = "EUR"
    elif "$" in v_s or "USD" in v_s: unit = "USD"
    
    # Sayıyı temizle: 0,0245 gibi değerleri korumak için virgül/nokta ayarı
    c_v = re.sub(r'[^0-9,.]', '', v_s)
    if ',' in c_v and '.' in c_v: c_v = c_v.replace('.', '').replace(',', '.')
    elif ',' in c_v: c_v = c_v.replace(',', '.')
    
    try:
        n = float(c_v)
        if unit == "EUR": return n * L_RATES['EUR_TO_USD']
        if unit == "TRY": return n * L_RATES['TRY_TO_USD']
        return n
    except: return None

def find_column(cols, prio):
    for p in prio:
        for c in cols:
            if p == str(c).lower().strip() or p in str(c).lower():
                return c
    return None

def smart_read(file):
    ext = os.path.splitext(file.name)[1].lower()
    try:
        if ext in ['.xlsx', '.xls']:
            # Önce ham oku
            df = pd.read_excel(file, header=None)
            # Başlık satırı ara (ilk 50 satır)
            for i, row in df.head(50).iterrows():
                row_str = " ".join(map(str, row.values)).lower()
                if any(p in row_str for p in PN_PRIO):
                    return pd.read_excel(file, header=i)
            # Bulamazsa ilk satırı başlık say
            return pd.read_excel(file)
        elif ext == '.pdf':
            with pdfplumber.open(file) as pdf:
                all_data = []
                for page in pdf.pages:
                    table = page.extract_table()
                    if table: all_data.extend(table)
                if all_data:
                    df_p = pd.DataFrame(all_data)
                    # Başlık satırı ara
                    for i, row in df_p.head(20).iterrows():
                        if any(p in " ".join(map(str, row.values)).lower() for p in PN_PRIO):
                            df_p.columns = df_p.iloc[i]
                            return df_p.iloc[i+1:].reset_index(drop=True)
                    return df_p
    except Exception as e:
        st.warning(f"{file.name} okunurken teknik hata: {e}")
    return None

# --- 4. ARAYÜZ VE ANALİZ ---
st.title("📊 Akıllı BOM Robotu v5.9")

m_file = st.file_uploader("1. Master BOM", type=['xlsx', 'xls'])
s_files = st.file_uploader("2. Teklifler", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True)

if m_file and s_files:
    m_df = smart_read(m_file)
    if m_df is not None:
        m_pn = find_column(m_df.columns, PN_PRIO)
        m_qty = find_column(m_df.columns, QTY_PRIO)
        
        if m_pn:
            m_df['M_KEY'] = m_df[m_pn].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper()) if pd.notna(x) else "")
            res_df = m_df.copy()
            s_cols = []

            for f in s_files:
                s_df = smart_read(f)
                if s_df is not None:
                    s_pn = find_column(s_df.columns, PN_PRIO)
                    s_pr = find_column(s_df.columns, PR_PRIO)
                    
                    if s_pn and s_pr:
                        s_key = os.path.splitext(f.name)[0][:15]
                        c_name = f"{s_key} ($)"
                        
                        temp = s_df[[s_pn, s_pr]].copy()
                        temp['M_KEY'] = temp[s_pn].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper()) if pd.notna(x) else "")
                        temp[c_name] = temp[s_pr].apply(parse_price)
                        
                        temp = temp.dropna(subset=[c_name]).drop_duplicates('M_KEY')
                        res_df = pd.merge(res_df, temp[['M_KEY', c_name]], on='M_KEY', how='left')
                        s_cols.append(c_name)
                        st.info(f"✅ {f.name} okundu.")
                    else:
                        st.error(f"❌ {f.name} içinde PN veya Fiyat sütunu bulunamadı!")

            if s_cols:
                res_df['En Düşük ($)'] = res_df[s_cols].min(axis=1)
                res_df['Kazanan'] = res_df.apply(lambda r: r[s_cols].idxmin().replace(" ($)", "") if r[s_cols].notna().any() else "Yok", axis=1)
                
                if m_qty:
                    res_df[m_qty] = pd.to_numeric(res_df[m_qty], errors='coerce').fillna(0)
                    res_df['Toplam ($)'] = (res_df['En Düşük ($)'] * res_df[m_qty]).round(4)

                st.dataframe(res_df.drop(columns=['M_KEY']), use_container_width=True)
                
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    res_df.drop(columns=['M_KEY']).to_excel(writer, index=False, sheet_name='Analiz')
                st.download_button("📩 Excel İndir", output.getvalue(), "BOM_Analiz.xlsx")
        else:
            st.error("Master listede parça numarası sütunu bulunamadı.")
