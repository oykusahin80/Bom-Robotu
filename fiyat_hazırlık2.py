import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="BOM Robotu v7.7 - Final Fix", layout="wide")

# --- 1. TCMB KUR SERVİSİ ---
@st.cache_data(ttl=3600)
def get_live_rates():
    try:
        response = requests.get("https://www.tcmb.gov.tr/kurlar/today.xml", timeout=10)
        root = ET.fromstring(response.content)
        rates = {'USD': 33.0, 'EUR': 36.0}
        for curr in root.findall('Currency'):
            code = curr.get('CurrencyCode')
            if code in ['USD', 'EUR']:
                val = curr.find('ForexSelling').text
                if val: rates[code] = float(val)
        return {
            'EUR_USD': rates['EUR'] / rates['USD'],
            'TRY_USD': 1 / rates['USD'],
            'USD': rates['USD'], 'EUR': rates['EUR']
        }
    except:
        return {'EUR_USD': 1.09, 'TRY_USD': 1/33.0, 'USD': 33.0, 'EUR': 36.0}

L_RATES = get_live_rates()

# --- SIDEBAR ---
st.sidebar.title("🏦 Güncel Kurlar (TCMB)")
st.sidebar.write(f"**USD / TL:** {L_RATES['USD']:.4f}")
st.sidebar.write(f"**EUR / TL:** {L_RATES['EUR']:.4f}")
st.sidebar.write(f"**EUR / USD:** {L_RATES['EUR_USD']:.4f}")
st.sidebar.divider()

# --- 2. YARDIMCI FONKSİYONLAR ---
def aggressive_clean(text):
    if pd.isna(text) or str(text).strip() == "": return ""
    return re.sub(r'[^A-Z0-9]', '', str(text).upper().strip())

def parse_to_usd(val, is_arrow=False):
    if pd.isna(val) or str(val).strip() == "": return None
    v = str(val).upper().replace(" ", "")
    c = re.sub(r'[^0-9,.]', '', v)
    if ',' in c and '.' in c: c = c.replace('.', '').replace(',', '.')
    elif ',' in c: c = c.replace(',', '.')
    try:
        n = float(c)
        if is_arrow or "EUR" in v or "€" in v: return round(n * L_RATES['EUR_USD'], 4)
        if "TL" in v or "TRY" in v: return round(n * L_RATES['TRY_USD'], 4)
        return round(n, 4)
    except: return None

def clean_stock(val):
    if pd.isna(val) or str(val).strip() == "": return 999999
    s = str(val).upper()
    if any(x in s for x in ['YOK', 'OUT', 'NO', 'ZERO']): return 0
    c = re.sub(r'[^0-9]', '', s)
    try: return int(c) if c else 999999
    except: return 999999

# --- 3. SÜTUN TANIMA ---
PN_PRIORITY = ['manufacturer part number', 'man code', 'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 'model', 'p/n']
PRICE_PRIORITY = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar']
QTY_PRIORITY = ['qty', 'adet', 'miktar', 'quantity']
STOCK_PRIORITY = ['stock', 'stok', 'qty available', 'on hand', 'mevcut']

def find_best_col(columns, priority_list):
    for kw in priority_list:
        for col in columns:
            if kw == str(col).lower().strip() or kw in str(col).lower(): return col
    return None

def smart_load(file):
    if file is None: return None
    file.seek(0)
    ext = os.path.splitext(file.name)[1].lower()
    try:
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file, header=None)
            for i, row in df.head(50).iterrows():
                if any(kw in " ".join(map(str, row.values)).lower() for kw in PN_PRIORITY):
                    file.seek(0)
                    return pd.read_excel(file, header=i)
            return pd.read_excel(file)
        elif ext == '.pdf':
            with pdfplumber.open(file) as pdf:
                all_rows = []
                for page in pdf.pages:
                    table = page.extract_table()
                    if table: all_rows.extend(table)
            if all_rows:
                df_p = pd.DataFrame(all_rows)
                for i, row in df_p.head(20).iterrows():
                    if any(kw in " ".join(map(str, row.values)).lower() for kw in PN_PRIORITY):
                        df_p.columns = df_p.iloc[i]
                        return df_p.iloc[i+1:].reset_index(drop=True)
                return df_p
    except: return None

# --- 4. ANA AKIŞ ---
st.title("📊 Profesyonel BOM Robotu v7.7")

m_file = st.file_uploader("1. Master BOM", type=['xlsx', 'xls'], key="m_up")
s_files = st.file_uploader("2. Teklifler", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True, key="s_up")

if m_file and s_files:
    df_m = smart_load(m_file)
    if df_m is not None:
        m_pn = find_best_col(df_m.columns, PN_PRIORITY)
        m_qty = find_best_col(df_m.columns, QTY_PRIORITY)
        
        if m_pn:
            # KRİTİK: Index sıfırlama (Hatanın çözümü)
            df_m = df_m.dropna(subset=[m_pn]).reset_index(drop=True)
            df_m['MATCH_KEY'] = df_m[m_pn].apply(aggressive_clean)
            final_df = df_m.copy()
            
            price_cols = []

            for f in s_files:
                df_s = smart_load(f)
                if df_s is not None:
                    s_pn = find_best_col(df_s.columns, PN_PRIORITY)
                    s_pr = find_best_col(df_s.columns, PRICE_PRIORITY)
                    s_st = find_best_col(df_s.columns, STOCK_PRIORITY)
                    
                    if s_pn and s_pr:
                        s_name = os.path.splitext(f.name)[0][:10]
                        p_col = f"{s_name}_($)"
                        st_col = f"{s_name}_Stok"
                        
                        temp = df_s.copy()
                        temp['MATCH_KEY'] = temp[s_pn].apply(aggressive_clean)
                        temp['P_USD'] = temp[s_pr].apply(lambda x: parse_to_usd(x, "ARROW" in f.name.upper()))
                        temp['S_VAL'] = temp[s_st].apply(clean_stock) if s_st else 999999
                        
                        merged = pd.merge(final_df[['MATCH_KEY']], temp[['MATCH_KEY', 'P_USD', 'S_VAL']], on='MATCH_KEY', how='left')
                        
                        final_df[p_col] = merged.apply(lambda r: f"{r['P_USD']:.4f} [S: {int(r['S_VAL'])}]" if pd.notna(r['P_USD']) else None, axis=1)
                        final_df[f"{p_col}_raw"] = merged['P_USD']
                        final_df[st_col] = merged['S_VAL']
                        
                        price_cols.append(p_col)
                        st.success(f"✔️ {f.name} okundu.")

            if price_cols:
                def get_best_offer(row):
                    valid = []
                    target_qty = pd.to_numeric(row[m_qty], errors='coerce') if m_qty else 0
                    for pc in price_cols:
                        raw_p = row[f"{pc}_raw"]
                        st_v = row[pc.replace('_($)', '_Stok')]
                        if pd.notna(raw_p) and st_v > 0:
                            valid.append((raw_p, pc.replace("_($)", ""), st_v))
                    if not valid: return pd.Series([None, "Yok"], index=['Min', 'Win'])
                    suff = [v for v in valid if v[2] >= target_qty]
                    best = min(suff if suff else valid, key=lambda x: x[0])
                    return pd.Series([best[0], best[1]], index=['Min', 'Win'])

                final_df[['En Düşük ($)', 'Kazanan']] = final_df.apply(get_best_offer, axis=1)

                # Yan Panel Özeti
                st.sidebar.title("📊 Analiz Özeti")
                st.sidebar.info(f"**Toplam Kalem:** {len(final_df)}")
                st.sidebar.success(f"**Bulunan:** {final_df['En Düşük ($)'].notna().sum()}")

                if m_qty:
                    final_df[m_qty] = pd.to_numeric(final_df[m_qty], errors='coerce').fillna(0)
                    final_df['Toplam Maliyet ($)'] = (final_df['En Düşük ($)'] * final_df[m_qty]).round(4)

                # --- RENKLENDİRME (Hatasız Versiyon) ---
                def style_logic(col_data):
                    col_name = col_data.name
                    if col_name in price_cols:
                        st_col = col_name.replace("_($)", "_Stok")
                        target_qty = pd.to_numeric(final_df[m_qty], errors='coerce') if m_qty else 0
                        
                        styles = []
                        for i in range(len(col_data)):
                            curr_stok = final_df.loc[i, st_col]
                            curr_price = final_df.loc[i, f"{col_name}_raw"]
                            # Eğer fiyat varsa VE stok ihtiyacın altındaysa kırmızı yap
                            if pd.notna(curr_price) and curr_stok < target_qty.iloc[i]:
                                styles.append('background-color: #ff4b4b; color: white')
                            else:
                                styles.append('')
                        return styles
                    return [''] * len(col_data)

                st.subheader("🏁 Karşılaştırma Sonuçları")
                display_cols = [c for c in final_df.columns if "_raw" not in str(c) and "MATCH_KEY" not in str(c)]
                
                # Tabloyu göster
                st.dataframe(final_df[display_cols].style.apply(style_logic), use_container_width=True)
                
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    final_df[display_cols].to_excel(writer, index=False)
                st.download_button("📩 Excel Raporu", out.getvalue(), "BOM_Analiz.xlsx")
