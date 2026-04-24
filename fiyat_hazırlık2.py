import streamlit as st
import pandas as pd
import pdfplumber
import os
import re
import requests
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="BOM Robotu v7.3 - Stok Kontrollü", layout="wide")

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

# --- SIDEBAR (KUR VE ANALİZ BİLGİLERİ) ---
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
    """Stok verisini sayıya çevirir, yoksa 999999 (Sınırsız) varsayar"""
    if pd.isna(val) or str(val).strip() == "": return 999999
    s = str(val).upper()
    if any(x in s for x in ['YOK', 'OUT', 'NO', 'ZERO']): return 0
    c = re.sub(r'[^0-9]', '', s)
    try: return int(c) if c else 999999
    except: return 999999

# --- 3. ESNEK SÜTUN TANIMA ---
PN_PRIORITY = ['manufacturer part number', 'man code', 'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 'model', 'p/n', 'mfr part']
PRICE_PRIORITY = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar', 'resale', 'net']
QTY_PRIORITY = ['qty', 'adet', 'miktar', 'quantity']
NO_PRIORITY = ['no', 'sıra no', 'item no', 'id']
STOCK_PRIORITY = ['stock', 'stok', 'qty available', 'on hand', 'mevcut', 'availability']

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
                row_str = " ".join(map(str, row.values)).lower()
                if any(kw in row_str for kw in PN_PRIORITY):
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
st.title("📊 Profesyonel BOM Robotu v7.3 (Stok Kontrollü)")

master_file = st.file_uploader("1. Master Listeyi Seçin", type=['xlsx', 'xls'], key="m_up")
supplier_files = st.file_uploader("2. Teklif Dosyalarını Seçin (Toplu)", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True, key="s_up")

if master_file and supplier_files:
    df_master = smart_load(master_file)
    if df_master is not None:
        m_pn = find_best_col(df_master.columns, PN_PRIORITY)
        m_qty = find_best_col(df_master.columns, QTY_PRIORITY)
        m_no = find_best_col(df_master.columns, NO_PRIORITY)
        
        if m_pn:
            if m_no: df_master = df_master.dropna(subset=[m_no, m_pn], how='all')
            else: df_master = df_master.dropna(subset=[m_pn])
            df_master = df_master[df_master[m_pn].apply(lambda x: str(x).strip() != "" and pd.notna(x))]
            df_master['MATCH_KEY'] = df_master[m_pn].apply(aggressive_clean)
            
            final_df = df_master.copy()
            price_cols = []

            for s_file in supplier_files:
                df_sup = smart_load(s_file)
                if df_sup is not None:
                    s_pn = find_best_col(df_sup.columns, PN_PRIORITY)
                    s_pr = find_best_col(df_sup.columns, PRICE_PRIORITY)
                    s_st = find_best_col(df_sup.columns, STOCK_PRIORITY) # Stok sütunu arama
                    
                    if s_pn and s_pr:
                        s_name = os.path.splitext(s_file.name)[0][:15]
                        p_col = f"{s_name}_($)"
                        is_arrow = "ARROW" in s_file.name.upper()
                        
                        temp_sup = df_sup.copy()
                        temp_sup['MATCH_KEY'] = temp_sup[s_pn].apply(aggressive_clean)
                        
                        # Fiyat ve Stok işleme
                        temp_sup['RAW_PRICE'] = temp_sup[s_pr].apply(lambda x: parse_to_usd(x, is_arrow))
                        if s_st:
                            temp_sup['STOCK_VAL'] = temp_sup[s_st].apply(clean_stock)
                        else:
                            temp_sup['STOCK_VAL'] = 999999 # Bilgi yoksa var say
                            
                        # KRİTİK: Stok 0 ise fiyatı iptal et
                        temp_sup[p_col] = temp_sup.apply(lambda r: r['RAW_PRICE'] if r['STOCK_VAL'] > 0 else None, axis=1)
                        
                        temp_sup = temp_sup.dropna(subset=[p_col]).drop_duplicates('MATCH_KEY')
                        final_df = pd.merge(final_df, temp_sup[['MATCH_KEY', p_col]], on='MATCH_KEY', how='left')
                        price_cols.append(p_col)
                        st.success(f"✔️ {s_file.name} işlendi.")

            if price_cols:
                def get_row_results(row):
                    valid = row[price_cols].dropna()
                    if valid.empty: return pd.Series([None, "Yok"], index=['Min', 'Win'])
                    return pd.Series([round(valid.min(), 4), valid.idxmin().replace("_($)", "")], index=['Min', 'Win'])

                final_df[['En Düşük ($)', 'Kazanan']] = final_df.apply(get_row_results, axis=1)
                
                # SIDEBAR ÖZETİ
                total_items = len(final_df)
                found_items = final_df['En Düşük ($)'].notna().sum()
                success_rate = (found_items / total_items) * 100 if total_items > 0 else 0
                st.sidebar.title("📊 Analiz Özeti")
                st.sidebar.info(f"**Toplam:** {total_items} Kalem")
                st.sidebar.success(f"**Stokta Bulunan:** {found_items} Kalem")
                st.sidebar.warning(f"**Başarı:** %{success_rate:.1f}")
                
                win_counts = final_df[final_df['Kazanan'] != "Yok"]['Kazanan'].value_counts()
                if not win_counts.empty:
                    st.sidebar.divider()
                    st.sidebar.write("**Tedarikçi Dağılımı (Stoklu):**")
                    for winner, count in win_counts.items():
                        st.sidebar.write(f"🔹 {winner}: **{count}**")

                if m_qty:
                    final_df[m_qty] = pd.to_numeric(final_df[m_qty], errors='coerce').fillna(0)
                    final_df['Toplam Maliyet ($)'] = (final_df['En Düşük ($)'] * final_df[m_qty]).round(4)

                st.subheader("🏁 Karşılaştırma Sonuçları (Sadece Stoklu Teklifler)")
                st.dataframe(final_df.drop(columns=['MATCH_KEY']).style.format(precision=4, na_rep="-"), use_container_width=True)
                
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    final_df.drop(columns=['MATCH_KEY']).to_excel(writer, index=False)
                st.download_button("📩 Raporu İndir", out.getvalue(), "BOM_Stoklu_Analiz.xlsx")
