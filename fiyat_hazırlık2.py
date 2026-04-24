import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="BOM Robotu v7.4 - Görsel Uyarı", layout="wide")

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
            'USD_TRY': rates['USD'], 'EUR_TRY': rates['EUR'],
            'EUR_USD': rates['EUR'] / rates['USD'], 'TRY_USD': 1 / rates['USD']
        }
    except:
        return {'USD_TRY': 33.0, 'EUR_TRY': 36.0, 'EUR_USD': 1.09, 'TRY_USD': 1/33.0}

L_RATES = get_live_rates()

# --- SIDEBAR ---
st.sidebar.title("🏦 Güncel Kurlar (TCMB)")
st.sidebar.write(f"**USD / TL:** {L_RATES['USD_TRY']:.4f}")
st.sidebar.write(f"**EUR / TL:** {L_RATES['EUR_TRY']:.4f}")
st.sidebar.write(f"**EUR / USD:** {L_RATES['EUR_USD']:.4f}")
st.sidebar.divider()

# --- 2. TEMİZLEME VE HESAPLAMA ---
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
    if any(x in s for x in ['YOK', 'OUT', 'NO', 'ZERO', '0']): return 0
    c = re.sub(r'[^0-9]', '', s)
    try: return int(c) if c else 999999
    except: return 999999

# --- 3. ESNEK SÜTUN TANIMA ---
PN_PRIORITY = ['manufacturer part number', 'man code', 'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 'model', 'p/n']
PRICE_PRIORITY = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar', 'resale']
QTY_PRIORITY = ['qty', 'adet', 'miktar', 'quantity']
NO_PRIORITY = ['no', 'sıra no', 'item no']
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
st.title("📊 BOM Robotu v7.4 (Görsel Stok Boyama)")

master_file = st.file_uploader("1. Master BOM", type=['xlsx', 'xls'], key="m_up")
supplier_files = st.file_uploader("2. Teklifler", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True, key="s_up")

if master_file and supplier_files:
    df_master = smart_load(master_file)
    if df_master is not None:
        m_pn = find_best_col(df_master.columns, PN_PRIORITY)
        m_qty = find_best_col(df_master.columns, QTY_PRIORITY)
        m_no = find_best_col(df_master.columns, NO_PRIORITY)
        
        if m_pn:
            df_master = df_master.dropna(subset=[m_pn])
            df_master['MATCH_KEY'] = df_master[m_pn].apply(aggressive_clean)
            
            final_df = df_master.copy()
            price_cols = []
            stock_info = {} # Renklendirme için stok durumlarını tutar

            for s_file in supplier_files:
                df_sup = smart_load(s_file)
                if df_sup is not None:
                    s_pn = find_best_col(df_sup.columns, PN_PRIORITY)
                    s_pr = find_best_col(df_sup.columns, PRICE_PRIORITY)
                    s_st = find_best_col(df_sup.columns, STOCK_PRIORITY)
                    
                    if s_pn and s_pr:
                        s_name = os.path.splitext(s_file.name)[0][:15]
                        p_col = f"{s_name}_($)"
                        
                        temp_sup = df_sup.copy()
                        temp_sup['MATCH_KEY'] = temp_sup[s_pn].apply(aggressive_clean)
                        temp_sup['PR_USD'] = temp_sup[s_pr].apply(lambda x: parse_to_usd(x, "ARROW" in s_file.name.upper()))
                        temp_sup['ST_VAL'] = temp_sup[s_st].apply(clean_stock) if s_st else 999999
                        
                        # Kayıtları birleştir
                        merged = pd.merge(final_df[['MATCH_KEY']], temp_sup[['MATCH_KEY', 'PR_USD', 'ST_VAL']], on='MATCH_KEY', how='left')
                        final_df[p_col] = merged['PR_USD']
                        
                        # Stok 0 olanları işaretlemek için sakla
                        stock_info[p_col] = merged['ST_VAL']
                        price_cols.append(p_col)
                        st.success(f"✔️ {s_file.name} eklendi.")

            if price_cols:
                def get_row_results(row):
                    # Sadece stok değeri > 0 olan fiyatları filtrele
                    valid_prices = []
                    for col in price_cols:
                        price = row[col]
                        # stock_info'daki aynı satırın stok değerine bak
                        idx = row.name
                        if pd.notna(price) and stock_info[col][idx] > 0:
                            valid_prices.append((price, col))
                    
                    if not valid_prices: return pd.Series([None, "Yok"], index=['Min', 'Win'])
                    
                    min_price, win_col = min(valid_prices, key=lambda x: x[0])
                    return pd.Series([min_price, win_col.replace("_($)", "")], index=['Min', 'Win'])

                final_df[['En Düşük ($)', 'Kazanan']] = final_df.apply(get_row_results, axis=1)
                
                # --- RENKLENDİRME FONKSİYONU ---
                def style_stock(df_styled):
                    for col in price_cols:
                        # Eğer bu kolondaki ilgili satırın stoğu 0 ise kırmızı yap
                        df_styled = df_styled.apply(
                            lambda x: [
                                'background-color: #ff4b4b; color: white' if stock_info[col][i] == 0 else '' 
                                for i in range(len(x))
                            ] if x.name == col else [''] * len(x),
                            axis=0
                        )
                    return df_styled

                # Analiz Özeti (Sidebar)
                total_items = len(final_df)
                found_items = final_df['En Düşük ($)'].notna().sum()
                st.sidebar.title("📊 Analiz Özeti")
                st.sidebar.info(f"**Toplam:** {total_items} Kalem")
                st.sidebar.success(f"**Stoklu Teklif:** {found_items}")
                
                if m_qty:
                    final_df[m_qty] = pd.to_numeric(final_df[m_qty], errors='coerce').fillna(0)
                    final_df['Toplam Maliyet ($)'] = (final_df['En Düşük ($)'] * final_df[m_qty]).round(4)

                st.subheader("🏁 Karşılaştırma Sonuçları")
                st.caption("Not: Kırmızı hücreler stokta olmayan teklifleri temsil eder ve hesaplamaya dahil edilmez.")
                
                # Tabloyu boyayarak göster
                styled_df = final_df.drop(columns=['MATCH_KEY']).style.pipe(style_stock).format(precision=4, na_rep="-")
                st.dataframe(styled_df, use_container_width=True)
                
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    final_df.drop(columns=['MATCH_KEY']).to_excel(writer, index=False)
                st.download_button("📩 Raporu İndir", out.getvalue(), "BOM_Analiz_StokBoyamali.xlsx")
