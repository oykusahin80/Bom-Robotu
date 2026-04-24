import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="BOM Robotu v5.8 - Hassas Kur Fix", layout="wide")

# --- 1. TCMB KUR SERVİSİ (Hatasız USD ve EUR çekimi) ---
@st.cache_data(ttl=3600)
def get_live_rates():
    try:
        # TCMB'den güncel kurları çek
        response = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml", timeout=10)
        root = ET.fromstring(response.content)
        
        rates = {}
        for currency in root.findall('Currency'):
            code = currency.get('CurrencyCode')
            if code in ['USD', 'EUR']:
                # ForexSelling (Döviz Satış) fiyatını alıyoruz
                rates[code] = float(currency.find('ForexSelling').text)
        
        # Dönüşüm katsayılarını hesapla (Hedefimiz her zaman USD)
        # Arrow 0,0245 EUR verdiğinde: (0,0245 * EUR_Kuru) / USD_Kuru = USD Karşılığı
        usd_val = rates['USD']
        eur_val = rates['EUR']
        
        return {
            'EUR_TO_USD': eur_val / usd_val,
            'TRY_TO_USD': 1 / usd_val,
            'RAW_USD': usd_val,
            'RAW_EUR': eur_val
        }
    except Exception as e:
        st.error(f"Kur çekme hatası: {e}. Sabit kurlar uygulanıyor.")
        return {'EUR_TO_USD': 1.08, 'TRY_TO_USD': 1/32.5}

# Kurları bir kez çek ve değişkene ata
LIVE_RATES = get_live_rates()

# --- 2. HASSAS FİYAT TEMİZLEME FONKSİYONU ---
def parse_price_to_usd(raw_value):
    if pd.isna(raw_value) or str(raw_value).strip() == "":
        return None
    
    val_str = str(raw_value).upper().replace(" ", "")
    
    # Birim Tespiti
    currency_unit = "TRY" # Varsayılan
    if "€" in val_str or "EUR" in val_str: currency_unit = "EUR"
    elif "$" in val_str or "USD" in val_str: currency_unit = "USD"
    
    # Sadece rakam, nokta ve virgülü bırak
    # Örn: "0,0245 €" -> "0,0245"
    clean_str = re.sub(r'[^0-9,.]', '', val_str)
    
    # Format Düzeltme: Arrow formatı (0,0245) -> Python formatı (0.0245)
    # Eğer hem nokta hem virgül varsa (1.250,50), noktayı sil virgülü noktaya çevir.
    if ',' in clean_str and '.' in clean_str:
        clean_str = clean_str.replace('.', '').replace(',', '.')
    elif ',' in clean_str:
        clean_str = clean_str.replace(',', '.')
    
    try:
        numeric_val = float(clean_str)
        # Hassas Kur Çarpımı
        if currency_unit == "EUR":
            return numeric_val * LIVE_RATES['EUR_TO_USD']
        elif currency_unit == "TRY":
            return numeric_val * LIVE_RATES['TRY_TO_USD']
        else: # Zaten USD
            return numeric_val
    except:
        return None

# --- 3. DİĞER YARDIMCI FONKSİYONLAR ---
PN_PRIO = ['manufacturer part number', 'man code', 'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 'model', 'p/n']
PR_PRIO = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar', 'net']
QTY_PRIO = ['qty', 'adet', 'miktar', 'quantity']

def find_column(available_cols, priority_list):
    for p in priority_list:
        for c in available_cols:
            if p in str(c).lower(): return c
    return None

def smart_read(file):
    ext = os.path.splitext(file.name)[1].lower()
    if ext in ['.xlsx', '.xls']:
        df = pd.read_excel(file, header=None)
        for i, row in df.head(30).iterrows():
            if any(p in str(row.values).lower() for p in PN_PRIO):
                return pd.read_excel(file, header=i)
        return pd.read_excel(file)
    elif ext == '.pdf':
        with pdfplumber.open(file) as pdf:
            rows = []
            for page in pdf.pages:
                table = page.extract_table()
                if table: rows.extend(table)
            if rows:
                pdf_df = pd.DataFrame(rows)
                pdf_df.columns = pdf_df.iloc[0]
                return pdf_df.iloc[1:].reset_index(drop=True)
    return None

# --- 4. ANA ARAYÜZ VE ANALİZ ---
st.title("📊 Hassas BOM Karşılaştırma v5.8")
st.sidebar.markdown(f"**Güncel Kurlar (TCMB)**\n\n1 EUR = **{LIVE_RATES['EUR_TO_USD']:.6f} $**\n1 TL = **{LIVE_RATES['TRY_TO_USD']:.6f} $**")

master_file = st.file_uploader("1. Master Liste", type=['xlsx', 'xls'])
teklif_files = st.file_uploader("2. Teklifler", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True)

if master_file and teklif_files:
    m_df = smart_read(master_file)
    if m_df is not None:
        pn_col = find_column(m_df.columns, PN_PRIO)
        qty_col = find_column(m_df.columns, QTY_PRIO)
        
        if pn_col:
            # Eşleştirme anahtarını oluştur
            m_df['M_KEY'] = m_df[pn_col].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper()) if pd.notna(x) else "")
            final_table = m_df.copy()
            active_suppliers = []

            for t_file in teklif_files:
                t_df = smart_read(t_file)
                if t_df is not None:
                    t_pn = find_column(t_df.columns, PN_PRIO)
                    t_pr = find_column(t_df.columns, PR_PRIO)
                    
                    if t_pn and t_pr:
                        s_name = os.path.splitext(t_file.name)[0][:15]
                        s_col = f"{s_name} ($)"
                        
                        # Teklif verisini işle
                        sub_df = t_df[[t_pn, t_pr]].copy()
                        sub_df['M_KEY'] = sub_df[t_pn].apply(lambda x: re.sub(r'[^A-Z0-9]', '', str(x).upper()) if pd.notna(x) else "")
                        sub_df[s_col] = sub_df[t_pr].apply(parse_price_to_usd)
                        
                        # Temizle ve birleştir
                        sub_df = sub_df.dropna(subset=[s_col]).drop_duplicates('M_KEY')
                        final_table = pd.merge(final_table, sub_df[['M_KEY', s_col]], on='M_KEY', how='left')
                        active_suppliers.append(s_col)

            if active_suppliers:
                # Sonuç Hesaplamaları
                final_table['En Düşük ($)'] = final_table[active_suppliers].min(axis=1)
                
                def identify_winner(row):
                    prices = row[active_suppliers].dropna()
                    if prices.empty: return "Teklif Yok"
                    return prices.idxmin().replace(" ($)", "")
                
                final_table['Kazanan'] = final_table.apply(identify_winner, axis=1)

                if qty_col:
                    final_table[qty_col] = pd.to_numeric(final_table[qty_col], errors='coerce').fillna(0)
                    final_table['Toplam ($)'] = (final_table['En Düşük ($)'] * final_table[qty_col]).round(4)

                # Tablo ve Çıktı
                view = final_table.drop(columns=['M_KEY'])
                st.dataframe(view, use_container_width=True)

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    view.to_excel(writer, index=False, sheet_name='Analiz')
                st.download_button("📩 Excel Raporunu İndir", output.getvalue(), "Fiyat_Analiz.xlsx")
