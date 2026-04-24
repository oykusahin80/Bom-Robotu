import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import plotly.express as px

st.set_page_config(page_title="BOM Robotu v5.2", layout="wide")

# --- TCMB CANLI KUR ÇEKME ---
@st.cache_data(ttl=3600)
def get_tcmb_rates():
    try:
        # Zaman aşımı ekleyerek TCMB sitesinin yavaşlığına karşı önlem alıyoruz
        response = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml", timeout=10)
        root = ET.fromstring(response.content)
        rates = {'USD': 1.0, 'TRY': 1.0, 'EUR': 1.0}
        for currency in root.findall('Currency'):
            code = currency.get('CurrencyCode')
            if code in ['USD', 'EUR']:
                val = currency.find('ForexSelling').text
                if val: rates[code] = float(val)
        
        # USD bazlı dönüşüm katsayıları
        u_rate = rates.get('USD', 32.5) 
        rates['TRY_TO_USD'] = 1 / u_rate
        rates['EUR_TO_USD'] = rates['EUR'] / u_rate
        return rates
    except Exception as e:
        st.warning(f"⚠️ Kur çekilemedi, sabit kurlar kullanılıyor: {e}")
        return {'TRY_TO_USD': 1/32.5, 'EUR_TO_USD': 1.08}

RATES = get_tcmb_rates()

# --- TEMİZLEME VE DÖNÜŞTÜRME ---
PN_PRIORITY = ['manufacturer part number', 'man code', 'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 'model', 'p/n', 'vendor material']
PRICE_PRIORITY = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar', 'resale', 'net']
QTY_PRIORITY = ['qty', 'adet', 'miktar', 'quantity']

def aggressive_clean(text):
    if pd.isna(text) or text == "": return ""
    return re.sub(r'[^A-Z0-9]', '', str(text).upper().strip())

def detect_and_convert_to_usd(value):
    if pd.isna(value) or str(value).strip() == "": return None, ""
    v_str = str(value).upper().replace(" ", "").replace("TL", "TRY")
    
    currency = "TRY"
    if "€" in v_str or "EUR" in v_str: currency = "EUR"
    elif "$" in v_str or "USD" in v_str: currency = "USD"
    
    v = re.sub(r'[^\d.,]', '', v_str)
    if ',' in v and '.' in v: v = v.replace('.', '').replace(',', '.')
    elif ',' in v: v = v.replace(',', '.')
    
    try:
        num = float(v)
        if currency == "EUR": return round(num * RATES['EUR_TO_USD'], 4), " (€)"
        if currency == "TRY": return round(num * RATES['TRY_TO_USD'], 4), " (TL)"
        return round(num, 4), " ($)"
    except: return None, ""

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
st.title("🚀 Akıllı BOM Karşılaştırma v5.2")
st.sidebar.info(f"🏦 TCMB Kurları:\n- 1 EUR: {RATES['EUR_TO_USD']:.4f} USD\n- 1 TRY: {RATES['TRY_TO_USD']:.4f} USD")

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
                        s_name = os.path.splitext(s_file.name)[0][:15] # İsmi kısalt
                        u_col = f"{s_name}_USD"
                        o_col = f"{s_name}_Orijinal"
                        
                        temp_sup = df_sup[[s_pn, s_pr]].copy()
                        temp_sup['MATCH_KEY'] = temp_sup[s_pn].apply(aggressive_clean)
                        
                        conv = temp_sup[s_pr].apply(detect_and_convert_to_usd)
                        temp_sup[u_col] = conv.apply(lambda x: x[0])
                        temp_sup[o_col] = temp_sup[s_pr].astype(str) + conv.apply(lambda x: x[1])
                        
                        temp_sup = temp_sup.dropna(subset=[u_col]).drop_duplicates('MATCH_KEY')
                        final_df = pd.merge(final_df, temp_sup[['MATCH_KEY', u_col, o_col]], on='MATCH_KEY', how='left')
                        usd_cols.append(u_col)

            # --- HATA KORUMALI HESAPLAMA ---
            if usd_cols:
                # Satır bazlı en düşük fiyatı bulurken boş satırları (NaN) yoksay
                final_df['En Düşük ($)'] = final_df[usd_cols].min(axis=1)
                
                # Kazananı belirle (Hata veren yer burasıydı, idxmin güvenli hale getirildi)
                def get_winner(row):
                    valid_prices = row[usd_cols].dropna()
                    if valid_prices.empty: return "Teklif Yok"
                    return valid_prices.idxmin().replace("_USD", "")
                
                final_df['Kazanan'] = final_df.apply(get_winner, axis=1)

                if m_qty_col:
                    final_df[m_qty_col] = pd.to_numeric(final_df[m_qty_col], errors='coerce').fillna(0)
                    final_df['Toplam ($)'] = (final_df['En Düşük ($)'] * final_df[m_qty_col]).round(2)

                # --- ÖZET VE TABLO ---
                st.subheader("📊 Analiz Özeti")
                wins = final_df[final_df['Kazanan'] != "Teklif Yok"]['Kazanan'].value_counts()
                if not wins.empty:
                    st.plotly_chart(px.pie(values=wins.values, names=wins.index, title="Tedarikçi Payları", hole=0.4))
                
                st.dataframe(final_df.drop(columns=['MATCH_KEY'] + usd_cols), use_container_width=True)
                st.download_button("📩 Raporu İndir", final_df.to_csv(index=False).encode('utf-8-sig'), "BOM_Analiz.csv")
        else:
            st.error("Master Listede PN Sütunu Bulunamadı!")
