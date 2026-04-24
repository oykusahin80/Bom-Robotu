import streamlit as st
import pandas as pd
import pdfplumber
import os
import re

st.set_page_config(page_title="BOM Robotu v4.5", layout="wide")

# --- ÖNCELİK LİSTELERİ ---
PN_PRIORITY = ['manufacturer part number', 'man code', 'üretici parça kodu', 'parça numarası', 'part number', 'pn', 'kod', 'model', 'p/n', 'vendor material']
PRICE_PRIORITY = ['unit price', 'birim fiyat', 'fiyat', 'price', 'tutar', 'resale', 'net', 'amount', 'total']

def aggressive_clean(text):
    if pd.isna(text) or text == "": return ""
    # Harf ve rakam dışı her şeyi sil (QSH-030 -> QSH030)
    return re.sub(r'[^A-Z0-9]', '', str(text).upper().strip())

def clean_price(value):
    if pd.isna(value) or value == "": return None
    v = str(value).lower().replace(" ", "")
    v = re.sub(r'[^\d.,]', '', v)
    if ',' in v and '.' in v: v = v.replace('.', '').replace(',', '.')
    elif ',' in v: v = v.replace(',', '.')
    try: return float(v)
    except: return None

def find_best_column(columns, priority_list):
    for kw in priority_list:
        for col in columns:
            if kw in str(col).lower():
                return col
    return None

def smart_load(file):
    ext = os.path.splitext(file.name)[1].lower()
    try:
        if ext in ['.xlsx', '.xls']:
            df = pd.read_excel(file, header=None)
            # Başlık satırını bul (İlk 40 satır)
            for i, row in df.head(40).iterrows():
                if any(kw in " ".join(map(str, row.values)).lower() for kw in PN_PRIORITY):
                    file.seek(0)
                    return pd.read_excel(file, header=i)
            return pd.read_excel(file) # Bulamazsa direkt oku

        elif ext == '.pdf':
            all_rows = []
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    table = page.extract_table()
                    if table: all_rows.extend(table)
            if all_rows:
                df_pdf = pd.DataFrame(all_rows)
                # Başlık bulma
                header_idx = 0
                for i, row in df_pdf.head(15).iterrows():
                    if any(kw in " ".join(map(str, row.values)).lower() for kw in PN_PRIORITY):
                        header_idx = i
                        break
                df_pdf.columns = df_pdf.iloc[header_idx]
                return df_pdf.iloc[header_idx+1:].reset_index(drop=True)
    except Exception as e:
        st.error(f"⚠️ {file.name} yüklenirken hata: {e}")
    return None

# --- ARAYÜZ ---
st.title("📊 Akıllı Fiyat Karşılaştırma (v4.5)")

master_file = st.file_uploader("1. Master BOM (Excel)", type=['xlsx', 'xls'])
supplier_files = st.file_uploader("2. Teklifler (Excel veya PDF)", type=['xlsx', 'xls', 'pdf'], accept_multiple_files=True)

if master_file and supplier_files:
    df_master = smart_load(master_file)
    
    if df_master is not None:
        m_pn_col = find_best_column(df_master.columns, PN_PRIORITY)
        
        if m_pn_col:
            st.success(f"✅ Master Liste Okundu. Anahtar: {m_pn_col}")
            df_master['MATCH_KEY'] = df_master[m_pn_col].apply(aggressive_clean)
            final_df = df_master.copy()
            price_cols = []

            for s_file in supplier_files:
                df_sup = smart_load(s_file)
                if df_sup is not None:
                    s_pn = find_best_column(df_sup.columns, PN_PRIORITY)
                    s_pr = find_best_column(df_sup.columns, PRICE_PRIORITY)
                    
                    if s_pn and s_pr:
                        s_name = os.path.splitext(s_file.name)[0]
                        p_col = f"Fiyat_{s_name}"
                        
                        temp_sup = df_sup[[s_pn, s_pr]].copy()
                        temp_sup['MATCH_KEY'] = temp_sup[s_pn].apply(aggressive_clean)
                        temp_sup[p_col] = temp_sup[s_pr].apply(clean_price)
                        
                        temp_sup = temp_sup[temp_sup['MATCH_KEY'] != ""].drop_duplicates('MATCH_KEY')
                        final_df = pd.merge(final_df, temp_sup[['MATCH_KEY', p_col]], on='MATCH_KEY', how='left')
                        price_cols.append(p_col)
                        st.info(f"📁 {s_file.name}: {final_df[p_col].notna().sum()} eşleşme.")
                    else:
                        st.warning(f"❌ {s_file.name} içinde PN veya Fiyat sütunu bulunamadı.")

            if price_cols:
                final_df['En Düşük'] = final_df[price_cols].min(axis=1)
                mask = final_df[price_cols].notna().any(axis=1)
                final_df['Kazanan'] = "Teklif Yok"
                if mask.any():
                    final_df.loc[mask, 'Kazanan'] = final_df.loc[mask, price_cols].idxmin(axis=1).str.replace("Fiyat_", "")

                st.dataframe(final_df.drop(columns=['MATCH_KEY']))
                st.download_button("📩 Excel İndir", final_df.to_csv(index=False).encode('utf-8-sig'), "Rapor.csv")
        else:
            st.error("Master BOM'da Parça Numarası sütunu bulunamadı.")