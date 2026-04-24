import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="BOM Robotu v5.6 - Final Fix", layout="wide")

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

def convert_any_to_usd(value):
    if pd.isna(value) or str(value).strip() == "": return None
    
    # Metni temizle ve para birimini yakala
    v_str = str(value).upper().replace(" ", "").replace("TL", "TRY")
    currency = "TRY"
    if "€" in v_str or "EUR" in v_str: currency = "EUR"
    elif "$" in v_str or "USD" in v_str: currency = "USD"
    
    # Sadece rakam, virgül ve noktayı tut
    v = re.sub(r'[^0-9,.]', '', v_str)
    
    # --- ONDALIK AYIRAÇ DÜZELTME (KRİTİK) ---
    if ',' in v and '.' in v:
        # Örn: 1.250,50 -> Nokta silinir, virgül noktaya döner
        v = v.replace('.', '').replace(',', '.')
    elif ',' in v:
        # Örn: 0,0245 -> Virgül noktaya döner
        v = v.replace(',', '.')
    
    try:
        num = float(v)
        # Hassas çarpım (6 hane)
        if currency == "EUR": return round(num * RATES['EUR_TO_USD'], 6)
        if currency == "TRY": return round(num * RATES['TRY_TO_USD'], 6)
        return round(num, 6)
    except:
        return None

# --- DİĞER YARDIMCI FONKSİYONLAR ---
PN_PRIORITY = ['manufacturer part number', 'man code', 'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 'model', 'p/n', 'vendor material']
PRICE_PRIORITY = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar', 'resale', 'net']
QTY_PRIORITY = ['qty', 'adet', 'miktar', 'quantity']

def aggressive_clean(text):
    if pd.isna(text) or text == "": return ""
    return re.sub(r'[^A-Z0-9]', '', str(text).upper().strip())

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

# --- ARAYÜZ VE ANALİZ ---
st.title("📊 Profesyonel BOM Robotu v5.6")

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
                # --- ÇÖKMEYİ ENGELLEYEN HESAPLAMA MANTIĞI ---
                final_df['En Düşük Birim ($)'] = final_df[usd_cols].min(axis=1)

                def get_winner_safe(row):
                    valid_prices = row[usd_cols].dropna()
                    if valid_prices.empty:
                        return "Teklif Yok"
                    # En küçük değerin sütun ismini bul ve temizle
                    winner_col = valid_prices.idxmin()
                    return str(winner_col).replace("Fiyat_", "").replace(" ($)", "")

                final_df['Kazanan'] = final_df.apply(get_winner_safe, axis=1)

                if m_qty_col:
                    final_df[m_qty_col] = pd.to_numeric(final_df[m_qty_col], errors='coerce').fillna(0)
                    final_df['Toplam ($)'] = (final_df['En Düşük Birim ($)'] * final_df[m_qty_col]).round(4)

                # Tablo ve İndirme
                display_df = final_df.drop(columns=['MATCH_KEY'])
                st.dataframe(display_df, use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    display_df.to_excel(writer, index=False, sheet_name='Analiz')
                
                st.download_button("📩 Excel Raporunu İndir", output.getvalue(), "Analiz_Raporu.xlsx")
        else:
            st.error("Master listede parça numarası sütunu bulunamadı.")
