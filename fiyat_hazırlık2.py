import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="BOM Robotu v6.1 - Debug Modu", layout="wide")

# --- 1. TCMB KUR SERVİSİ ---
@st.cache_data(ttl=3600)
def get_live_rates():
    try:
        response = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml", timeout=10)
        root = ET.fromstring(response.content)
        rates = {'USD': 32.5, 'EUR': 35.5} # Varsayılan kurlar
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
        return {'EUR_USD': 1.08, 'TRY_USD': 1/32.5, 'USD': 32.5, 'EUR': 35.5}

L_RATES = get_live_rates()

# --- 2. HASSAS FİYAT İŞLEME ---
def parse_price_usd(p_val, cur_val=None):
    if pd.isna(p_val) or str(p_val).strip() == "": return None
    
    # Metni temizle
    s = str(p_val).upper().replace(" ", "")
    # Para birimi tespiti (Hücre içi + Yan sütun)
    ctx = (s + str(cur_val or "")).upper()
    
    # Rakamları ayıkla
    c = re.sub(r'[^0-9,.]', '', s)
    if ',' in c and '.' in c: c = c.replace('.', '').replace(',', '.')
    elif ',' in c: c = c.replace(',', '.')
    
    try:
        n = float(c)
        if "EUR" in ctx or "€" in ctx: return n * L_RATES['EUR_USD']
        if "TL" in ctx or "TRY" in ctx: return n * L_RATES['TRY_USD']
        # Arrow gibi distribütörlerde 1'den küçükse ve birim yoksa USD/EUR kabul et
        if n < 5 and ("USD" not in ctx and "$" not in ctx):
             # Eğer Arrow'dan geliyorsa ve birim yoksa EUR/USD olma ihtimali yüksek (TL olamaz)
             return n * L_RATES['EUR_USD'] if "ARROW" in ctx else n
        return n if ("$" in ctx or "USD" in ctx) else n * L_RATES['TRY_USD']
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
            # Başlık satırını ara (ilk 60 satır)
            for i, row in df_raw.head(60).iterrows():
                row_str = " ".join(map(str, row.values)).lower()
                if "part" in row_str or "kod" in row_str or "pn" in row_str:
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
    except Exception as e:
        st.error(f"Okuma Hatası ({file.name}): {e}")
    return None

# --- 4. ANA ARAYÜZ ---
st.title("📊 Profesyonel BOM Analizi (v6.1 - Debug Modu)")

m_file = st.file_uploader("1. Master BOM Listesi", type=['xlsx', 'xls'])
s_files = st.file_uploader("2. Teklif Dosyaları", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True)

if m_file and s_files:
    m_df = smart_read(m_file)
    if m_df is not None:
        # Önemli sütunları bul
        m_pn = find_column(m_df.columns, ['part number', 'parça kodu', 'pn', 'p/n', 'üretici kodu'])
        m_qty = find_column(m_df.columns, ['qty', 'adet', 'miktar', 'quantity'])
        
        if m_pn:
            m_df['M_KEY'] = m_df[m_pn].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper()) if pd.notna(x) else "")
            res_df = m_df.copy()
            s_cols = []

            for f in s_files:
                s_df = smart_read(f)
                if s_df is not None:
                    # Tedarikçi dosyasında PN, Fiyat ve Döviz sütunlarını bul
                    s_pn = find_column(s_df.columns, ['part number', 'pn', 'p/n', 'mfr part', 'kod'])
                    s_pr = find_column(s_df.columns, ['unit price', 'birim fiyat', 'price', 'fiyat', 'net price'])
                    s_cu = find_column(s_df.columns, ['currency', 'döviz', 'birim', 'curr'])
                    
                    if s_pn and s_pr:
                        s_key = f"{os.path.splitext(f.name)[0][:10]}_($)"
                        temp = s_df.copy()
                        temp['M_KEY'] = temp[s_pn].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper()) if pd.notna(x) else "")
                        
                        # Fiyat hesaplama (Hassas mod)
                        temp[s_key] = temp.apply(lambda r: parse_price_usd(r[s_pr], r[s_cu] if s_cu else f.name), axis=1)
                        
                        temp = temp.dropna(subset=[s_key]).drop_duplicates('M_KEY')
                        res_df = pd.merge(res_df, temp[['M_KEY', s_key]], on='M_KEY', how='left')
                        s_cols.append(s_key)
                        st.success(f"✔️ {f.name} okundu. ({len(temp)} eşleşme)")
                    else:
                        st.warning(f"⚠️ {f.name}: PN veya Fiyat sütunu tanımlanamadı!")

            if s_cols:
                # Sonuç Tablosu
                res_df['En Düşük ($)'] = res_df[s_cols].min(axis=1)
                res_df['Kazanan'] = res_df[s_cols].idxmin(axis=1).str.replace("_($)", "", regex=False) if not res_df[s_cols].isna().all().all() else "Yok"
                
                if m_qty:
                    res_df[m_qty] = pd.to_numeric(res_df[m_qty], errors='coerce').fillna(0)
                    res_df['Toplam Maliyet ($)'] = (res_df['En Düşük ($)'] * res_df[m_qty]).round(4)

                st.subheader("🏁 Analiz Sonucu")
                st.dataframe(res_df.drop(columns=['M_KEY']), use_container_width=True)

                # Excel Çıktısı
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as wr:
                    res_df.drop(columns=['M_KEY']).to_excel(wr, index=False, sheet_name='Analiz')
                st.download_button("📩 Excel Raporunu İndir", out.getvalue(), "BOM_Analiz_Raporu.xlsx")
            else:
                st.error("Hiçbir tedarikçi dosyasında eşleşen veri bulunamadı.")
        else:
            st.error("Master listede Parça Numarası (PN) sütunu bulunamadı. Lütfen sütun ismini kontrol edin.")
