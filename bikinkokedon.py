import streamlit as st
import pandas as pd
import numpy as np
from dbfread import DBF
import os
import io
import zipfile
import tempfile
import re
import gc # Untuk garbage collection

# Library PDF
try:
    import pdfplumber
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

# --- PENGATURAN HALAMAN ---
st.set_page_config(page_title="Aplikasi Olah Data & Ekstrak PDF", layout="wide")

# --- FUNGSI MEMBERSIHKAN DAYA & IDPEL ---
def clean_daya(val):
    if pd.isna(val): return 0.0
    s = str(val).strip()
    try:
        num = float(s)
        if 0 < num < 300: return num * 1000.0
        return num
    except:
        s_clean = s.replace('.', '').replace(',', '')
        try: return float(s_clean)
        except: return 0.0

def clean_idpel(val):
    if pd.isna(val): return ""
    s = str(val).strip()
    if s.endswith('.0'): s = s[:-2]
    digits = re.sub(r'\D', '', s)
    if len(digits) >= 12:
        return digits[:12]
    return digits

# --- FUNGSI UN-MASKING (NAMA/ALAMAT BINTANG) ---
def fix_masked_info(df_baru, df_master, df_sup_pb):
    if df_baru.empty: return df_baru
    df_baru = df_baru.reset_index(drop=True)
    nama_map, alamat_map = {}, {}

    if not df_sup_pb.empty:
        sup_clean = df_sup_pb.drop_duplicates(subset=['IDPEL'], keep='first')
        sup_clean = sup_clean[sup_clean['NAMA'].astype(str).str.strip().ne('') & ~sup_clean['NAMA'].astype(str).str.contains(r'\*', na=False)]
        nama_map.update(sup_clean.set_index('IDPEL')['NAMA'].to_dict())
        alamat_map.update(sup_clean.set_index('IDPEL')['ALAMAT'].to_dict())

    if not df_master.empty:
        master_clean = df_master.drop_duplicates(subset=['IDPEL'], keep='first')
        master_clean = master_clean[master_clean['NAMA'].astype(str).str.strip().ne('') & ~master_clean['NAMA'].astype(str).str.contains(r'\*', na=False)]
        nama_map.update(master_clean.set_index('IDPEL')['NAMA'].to_dict())
        alamat_map.update(master_clean.set_index('IDPEL')['ALAMAT'].to_dict())

    if not nama_map and not alamat_map: return df_baru

    is_masked_nama = df_baru['NAMA'].astype(str).str.contains(r'\*', na=False) | df_baru['NAMA'].isna() | (df_baru['NAMA'].astype(str).str.strip() == '')
    is_masked_alamat = df_baru['ALAMAT'].astype(str).str.contains(r'\*', na=False) | df_baru['ALAMAT'].isna() | (df_baru['ALAMAT'].astype(str).str.strip() == '')

    df_baru.loc[is_masked_nama, 'NAMA'] = df_baru.loc[is_masked_nama, 'IDPEL'].map(nama_map).fillna(df_baru.loc[is_masked_nama, 'NAMA'])
    df_baru.loc[is_masked_alamat, 'ALAMAT'] = df_baru.loc[is_masked_alamat, 'IDPEL'].map(alamat_map).fillna(df_baru.loc[is_masked_alamat, 'ALAMAT'])
    return df_baru

# --- LOGIKA EKSTRAKSI TABEL PDF TIPE 1 (HYBRID) ---
def get_pdf_tables_tipe1(page):
    table_settings = {
        "vertical_strategy": "lines",
        "horizontal_strategy": "text",
    }
    tables = page.extract_tables(table_settings)
    if not tables or len(tables) == 0 or all(len(t) < 2 for t in tables):
        tables = page.extract_tables()
    return tables

def process_hybrid_table(table):
    if not table or len(table) < 2: return None

    merged_rows = []
    for row in table:
        cleaned_row = [str(cell).replace('\n', ' ').strip() if cell and str(cell) != 'None' else "" for cell in row]
        if not any(cleaned_row): continue
        
        if not merged_rows or cleaned_row[0] != "":
            merged_rows.append(cleaned_row)
        else:
            for i in range(len(cleaned_row)):
                if cleaned_row[i] != "":
                    if i < len(merged_rows[-1]):
                        merged_rows[-1][i] = (merged_rows[-1][i] + " " + cleaned_row[i]).strip()
                    else:
                        merged_rows[-1].append(cleaned_row[i])

    if not merged_rows: return None

    header_index = 0
    header_keywords = ['ID PEL', 'IDPEL', 'NOPEL', 'NAMA', 'ALAMAT', 'AGENDA', 'REGISTER', 'URUT']
    for idx, row in enumerate(merged_rows[:10]):
        row_str = " ".join([str(c).upper() for c in row])
        if any(kw in row_str for kw in header_keywords):
            header_index = idx
            break

    headers = []
    for i, h in enumerate(merged_rows[header_index]):
        val = str(h).strip() if str(h).strip() != "" else f"KOLOM_{i+1}"
        headers.append(val)
        
    df = pd.DataFrame(merged_rows[header_index+1:], columns=headers)
    return df

# --- LOGIKA EKSTRAKSI TABEL PDF TIPE 2 (DEFAULT & ALL COLUMNS) ---
def process_standard_table(table):
    if not table or len(table) < 2: return None
    
    cleaned_table = []
    for row in table:
        cleaned_row = [str(cell).replace('\n', ' ').strip() if cell and str(cell) != 'None' else "" for cell in row]
        if any(cleaned_row): cleaned_table.append(cleaned_row)

    if len(cleaned_table) < 2: return None

    header_index = 0
    header_keywords = ['ID PEL', 'IDPEL', 'NOPEL', 'NAMA', 'ALAMAT', 'NO', 'URUT', 'TANGGAL']
    for idx, row in enumerate(cleaned_table[:5]):
        row_str = " ".join([str(c).upper() for c in row])
        if any(kw in row_str for kw in header_keywords):
            header_index = idx
            break

    headers = []
    for i, h in enumerate(cleaned_table[header_index]):
        val = str(h).strip() if str(h).strip() != "" else f"KOLOM_{i+1}"
        headers.append(val)
        
    # --- ANTI-ERROR REINDEXING: Memastikan tidak ada nama kolom yang duplikat ---
    seen = set()
    unique_headers = []
    for h in headers:
        new_h = h
        counter = 1
        while new_h in seen:
            new_h = f"{h}_{counter}" 
            counter += 1
        seen.add(new_h)
        unique_headers.append(new_h)
        
    df = pd.DataFrame(cleaned_table[header_index+1:], columns=unique_headers)
    return df

# --- FUNGSI NORMALISASI (HANYA DIGUNAKAN DI TAB 3) ---
def find_target_column(col_name):
    col_clean = re.sub(r'[^A-Z0-9]', '', str(col_name).upper())
    
    if 'NOPEL' in col_clean or 'IDPEL' in col_clean or 'IDPELANGGAN' in col_clean or col_clean == 'ID' or col_clean == 'IDPEL': return 'IDPEL'
    if 'NAMAPEMOHON' in col_clean or 'NAMA' in col_clean or 'PELANGGAN' in col_clean: return 'NAMA'
    if 'ALAMATPEMOHON' in col_clean or 'ALAMAT' in col_clean or 'LOKASI' in col_clean: return 'ALAMAT'
    if 'LAMA' in col_clean and ('TARIF' in col_clean or 'DAYA' in col_clean or 'TARIP' in col_clean): return 'TARIF_LAMA'
    if 'BARU' in col_clean and ('TARIF' in col_clean or 'DAYA' in col_clean or 'TARIP' in col_clean): return 'TARIF'
    if 'TARIF' in col_clean or 'TARIP' in col_clean or 'GOL' in col_clean: return 'TARIF'
    if 'DAYA' in col_clean or 'VA' in col_clean or 'KAPAS' in col_clean: return 'DAYA'
    if 'GARDU' in col_clean or col_clean == 'GD': return 'GARDU'
    if 'TIANG' in col_clean or col_clean == 'TG': return 'TIANG'
    if 'KOKED' in col_clean or 'KDDK' in col_clean or 'KEDUDUKAN' in col_clean: return 'KOKED'
    return col_name

def separate_idpel_and_nama(df):
    if 'IDPEL' in df.columns:
        for idx in df.index:
            val = str(df.at[idx, 'IDPEL']).strip()
            match = re.match(r'^(\d{11,12})\s+(.+)$', val)
            if match:
                df.at[idx, 'IDPEL'] = match.group(1)
                nama_val = str(df.at[idx, 'NAMA']).strip() if 'NAMA' in df.columns else ''
                if not nama_val or nama_val == 'None' or nama_val.lower() == 'nan':
                    df.at[idx, 'NAMA'] = match.group(2)
    return df

def normalize_pdf_dataframe(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=['IDPEL', 'NAMA', 'ALAMAT', 'TARIF', 'DAYA', 'GARDU', 'TIANG', 'KOKED'])

    new_cols = [find_target_column(c) for c in df.columns]
    df.columns = new_cols
    df = df.loc[:, ~df.columns.duplicated(keep='first')]

    if 'TARIF' in df.columns:
        for idx in df.index:
            t_val = str(df.at[idx, 'TARIF']).strip()
            if '/' in t_val:
                parts = t_val.split('/')
                df.at[idx, 'TARIF'] = parts[0].strip()
                if len(parts) > 1 and (not 'DAYA' in df.columns or str(df.at[idx, 'DAYA']).strip() in ['', '0', '0.0']):
                    df.at[idx, 'DAYA'] = parts[1].strip()

    if 'IDPEL' in df.columns:
        id_idx = df.columns.get_loc('IDPEL')
        if 'NAMA' not in df.columns and id_idx + 1 < len(df.columns):
            col_name = df.columns[id_idx + 1]
            if col_name not in ['ALAMAT', 'TARIF', 'DAYA', 'GARDU', 'TIANG', 'KOKED']: df = df.rename(columns={col_name: 'NAMA'})
        
        if 'NAMA' in df.columns and 'ALAMAT' not in df.columns:
            nama_idx = df.columns.get_loc('NAMA')
            if nama_idx + 1 < len(df.columns):
                col_name = df.columns[nama_idx + 1]
                if col_name not in ['TARIF', 'DAYA', 'GARDU', 'TIANG', 'KOKED']: df = df.rename(columns={col_name: 'ALAMAT'})

    df = separate_idpel_and_nama(df)

    target_cols = ['IDPEL', 'NAMA', 'ALAMAT', 'TARIF', 'DAYA', 'GARDU', 'TIANG', 'KOKED']
    for col in target_cols:
        if col not in df.columns: df[col] = ""

    df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
    df['DAYA'] = df['DAYA'].apply(clean_daya)
    for col in ['NAMA', 'ALAMAT']: df[col] = df[col].astype(str).replace('nan', '').replace('None', '').str.strip()
    df = df[df['IDPEL'].astype(str).str.strip() != ''].copy()

    return df[target_cols].reset_index(drop=True)

# --- FUNGSI PEMBACA FILE UMUM TAHAP 1 & 2 ---
def baca_ekstrak_tabel(uploaded_file):
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    fname = uploaded_file.name
    
    if ext == '.xlsx':
        df = pd.read_excel(uploaded_file)
    else:
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            if ext == '.dbf':
                df = pd.DataFrame(iter(DBF(tmp_path, char_decode_errors='ignore')))
            else:
                st.error(f"Format tidak didukung di Tahap 1/2: {ext}")
                return pd.DataFrame(), fname
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    df.columns = df.columns.astype(str).str.strip().str.upper()
    df = df.rename(columns={'KDDK': 'KOKED', 'TARIF': 'TARIP', 'GOL TARIF': 'TARIP', 'NAMAPNJ': 'ALAMAT', 'ID PEL': 'IDPEL', 'ID_PELANGGAN': 'IDPEL', 'NOPEL': 'IDPEL'})
    return df, fname

def proses_list_file(files, tipe_data):
    list_df = []
    for f in files:
        df, fname = baca_ekstrak_tabel(f)
        if df.empty: continue
        
        if tipe_data in ['baru', 'petugas']:
            for col in ['ALAMAT', 'NAMA', 'TARIP', 'KOKED']:
                if col not in df.columns: df[col] = ''
            if 'DAYA' not in df.columns: df['DAYA'] = 0
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di file {fname}")
                
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['DAYA'] = df['DAYA'].apply(clean_daya)
            df['KOKED'] = df['KOKED'].astype(str).str.strip()
            list_df.append(df[['IDPEL', 'KOKED', 'TARIP', 'DAYA', 'NAMA', 'ALAMAT']])
            
        elif tipe_data == 'lama':
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di file: {fname}")
            if 'KOKED' not in df.columns: df['KOKED'] = ''
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['KOKED'] = df['KOKED'].astype(str).str.strip()
            list_df.append(df[['IDPEL', 'KOKED']])
            
        elif tipe_data == 'master':
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di file: {fname}")
            for col in ['NAMA', 'ALAMAT', 'KOKED', 'TARIP']:
                if col not in df.columns: df[col] = ''
            if 'DAYA' not in df.columns: df['DAYA'] = 0
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['DAYA'] = df['DAYA'].apply(clean_daya)
            list_df.append(df[['IDPEL', 'NAMA', 'ALAMAT', 'KOKED', 'TARIP', 'DAYA']])
            
        elif tipe_data == 'sup_pb':
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di file: {fname}")
            if 'NAMA' not in df.columns: df['NAMA'] = ''
            if 'ALAMAT' not in df.columns: df['ALAMAT'] = ''
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            list_df.append(df[['IDPEL', 'NAMA', 'ALAMAT']])

    if not list_df: return pd.DataFrame()
    return pd.concat(list_df, ignore_index=True).drop_duplicates(subset=['IDPEL'], keep='first').reset_index(drop=True)

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ==========================================
# ANTARMUKA STREAMLIT
# ==========================================
st.title("⚡ Aplikasi Olah Data & Ekstrak PDF / ICONPRN")

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs([
    "1️⃣ 1: Bikin Data Untuk Koked", 
    "2️⃣ 2: Hasil Koked",
    "3️⃣ 3: Eksport Pdf PB",
    "4️⃣ 4: Eksport PDF",
    "5️⃣ 5: eksport ICONPRN",
    "6️⃣ 6: Olah Data",
    "7️⃣ 7: Info & Lokasi",
    "8️⃣ 8: Update Data (Versi 2)",
    "9️⃣ 9: Isi Petugas",
    "🔟 10: Split Data"
])

# ==========================================
# TAB 1: PERSIAPAN DATA
# ==========================================
with tab1:
    st.header("Tahap 1: Memecah Data Untuk Petugas Lapangan")
    col1, col2 = st.columns(2)
    with col1:
        files_baru = st.file_uploader("[Tahap 1] Data Server Bulan INI", accept_multiple_files=True, key="t1_baru")
        files_lama = st.file_uploader("[Tahap 1] Data Bulan LALU", accept_multiple_files=True, key="t1_lama")
    with col2:
        files_master = st.file_uploader("[Tahap 1] Data Master (Ganti ***)", accept_multiple_files=True, key="t1_master")
        files_suppb = st.file_uploader("[Tahap 1] Data PB Baru", accept_multiple_files=True, key="t1_suppb")

    if st.button("Proses Tahap 1", type="primary"):
        if not files_baru or not files_lama: st.error("Silakan unggah Data Bulan Ini dan Data Bulan Lalu.")
        else:
            with st.spinner("Menyiapkan data untuk petugas..."):
                try:
                    df_baru = proses_list_file(files_baru, 'baru')
                    df_lama = proses_list_file(files_lama, 'lama')
                    df_master = proses_list_file(files_master, 'master') if files_master else pd.DataFrame()
                    df_sup_pb = proses_list_file(files_suppb, 'sup_pb') if files_suppb else pd.DataFrame()

                    df_baru = fix_masked_info(df_baru, df_master, df_sup_pb)
                    df_baru = df_baru[(df_baru['DAYA'] <= 33000) & (df_baru['IDPEL'] != '524050450911')].copy()
                    
                    list_idpel_lama = set(df_lama['IDPEL'].tolist())
                    df_pb = df_baru[~df_baru['IDPEL'].isin(list_idpel_lama)].copy()
                    df_tetap = df_baru[df_baru['IDPEL'].isin(list_idpel_lama)].copy()

                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        if not df_pb.empty: zip_file.writestr("PB_SEMUA_PETUGAS.xlsx", to_excel_bytes(df_pb[['IDPEL', 'KOKED', 'NAMA', 'ALAMAT', 'TARIP', 'DAYA']]))
                        mapping = {
                            'C01': 'JCA', 'C02': 'TAA', 'C03': 'JCB', 'C04': 'JCC', 'C05': 'NCD',
                            'C06': 'KBA', 'C07': 'TAB', 'C08': 'JCE', 'C09': 'TAC', 'C10': 'MBB',
                            'C11': 'KAD', 'C12': 'TAE', 'C13': 'KCF', 'C14': 'JCG', 'C15': 'TAF',
                            'C16': 'KAG', 'C17': 'BBC', 'C18': 'TAH', 'C19': 'BCK', 'C20': 'NCH',
                            'C21': 'JBE', 'C22': 'MBF', 'C23': 'MBG', 'C24': 'KBH', 'C25': 'KBI',
                            'C26': ['KAI', 'TAI'], 'C27': 'MCI', 'C28': ['KBJ', 'NBJ'], 
                            'C29': ['JCJ', 'KCJ'], 'C30': 'TAJ', 'C31': 'MBK', 'C32': 'KBL',
                            'C33': 'TAK', 'C34': 'KAL', 'C35': 'JCL', 'C36': 'MBD'
                        }
                        def add_to_zip(df_source, prefix=""):
                            if df_source.empty: return
                            for nama_file, kriteria in mapping.items():
                                temp_df = df_source[df_source['KOKED'].str[3:6].isin(kriteria)] if isinstance(kriteria, list) else df_source[df_source['KOKED'].str[3:6] == kriteria]
                                if not temp_df.empty: zip_file.writestr(f"{prefix}{nama_file}.xlsx", to_excel_bytes(temp_df.drop_duplicates(subset=['IDPEL'])[['IDPEL', 'KOKED', 'NAMA', 'ALAMAT', 'TARIP', 'DAYA']]))
                        add_to_zip(df_pb, "PB_")
                        add_to_zip(df_tetap, "")
                    st.success("✅ File untuk petugas berhasil dibuat!")
                    st.download_button("📥 Download Distribusi (.zip)", data=zip_buffer.getvalue(), file_name="Hasil_PerPetugas.zip", mime="application/zip")
                except Exception as e:
                    st.error(f"❌ Error Tahap 1: {str(e)}")

# ==========================================
# TAB 2: REKAP & PERUBAHAN KOKED
# ==========================================
with tab2:
    st.header("Tahap 2: Gabung File Petugas & Mutasi KOKED")
    col3, col4 = st.columns(2)
    with col3: files_petugas = st.file_uploader("[Tahap 2] Data Hasil Kerja Petugas", accept_multiple_files=True, key="t2_petugas")
    with col4: files_lama_pembanding = st.file_uploader("[Tahap 2] Data Bulan Lalu", accept_multiple_files=True, key="t2_lama")

    if st.button("Proses Tahap 2", type="primary"):
        if not files_petugas or not files_lama_pembanding: st.error("Silakan unggah Data Hasil Kerja Petugas dan Data Bulan Lalu.")
        else:
            with st.spinner("Membandingkan KOKED..."):
                try:
                    df_petugas = proses_list_file(files_petugas, 'petugas')
                    df_lama_pem = proses_list_file(files_lama_pembanding, 'lama')

                    df_compare = df_petugas[['IDPEL', 'KOKED']].merge(df_lama_pem[['IDPEL', 'KOKED']], on='IDPEL', suffixes=('_BARU', '_LAMA'))
                    df_changed = df_compare[(df_compare['KOKED_BARU'] != df_compare['KOKED_LAMA']) & (df_compare['KOKED_LAMA'] != '')]
                    
                    df_petugas['NO_URUT'] = pd.to_numeric(df_petugas['KOKED'].str[7:10], errors='coerce').fillna(0).astype(int)
                    df_petugas.sort_values(by=['KOKED', 'NO_URUT'], inplace=True)

                    zip_buffer2 = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer2, "w", zipfile.ZIP_DEFLATED) as zip_file2:
                        if not df_changed.empty:
                            zip_file2.writestr("PERUBAHAN_KOKED.txt", "\n".join((df_changed['IDPEL'].astype(str) + "|" + df_changed['KOKED_BARU'].astype(str)).tolist()))
                        zip_file2.writestr("DATA_GABUNGAN_FINAL.xlsx", to_excel_bytes(df_petugas[['IDPEL', 'KOKED', 'NAMA', 'ALAMAT', 'TARIP', 'DAYA']]))

                    st.success(f"✅ Selesai! {len(df_changed)} data mutasi KOKED.")
                    st.download_button("📥 Download Hasil Akhir (.zip)", data=zip_buffer2.getvalue(), file_name="Hasil_Koked.zip", mime="application/zip", type="primary")
                except Exception as e:
                    st.error(f"❌ Error Tahap 2: {str(e)}")

# ==========================================
# TAB 3: IMPORT & EKSTRAK PDF (TIPE 1)
# ==========================================
with tab3:
    st.header("Tahap 3: Import & Ekstrak PDF (Tipe 1)")
    st.markdown("Digunakan untuk file yang terpotong barisnya. **Data di-filter ke 8 Kolom Standar.**")
    pdf_files_t1 = st.file_uploader("Upload File PDF Tipe 1", type=['pdf'], accept_multiple_files=True, key="t3_pdf")

    if st.button("Proses & Ekstrak PDF (Tipe 1)", type="primary"):
        if not pdf_files_t1: st.error("Silakan unggah setidaknya satu file PDF.")
        elif not PDF_SUPPORT: st.error("Library `pdfplumber` belum di-install.")
        else:
            all_extracted_dfs = []
            with st.spinner("Mengekstraksi data Tipe 1..."):
                for uploaded_pdf in pdf_files_t1:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp: 
                        tmp.write(uploaded_pdf.getvalue())
                        tmp_path = tmp.name
                    try:
                        with pdfplumber.open(tmp_path) as pdf:
                            for page in pdf.pages:
                                tables = get_pdf_tables_tipe1(page)
                                for table in tables:
                                    df_hybrid = process_hybrid_table(table)
                                    if df_hybrid is not None and not df_hybrid.empty: 
                                        all_extracted_dfs.append(df_hybrid)
                                page.flush_cache()
                    finally:
                        if os.path.exists(tmp_path): os.remove(tmp_path)

            if all_extracted_dfs:
                final_pdf_df = normalize_pdf_dataframe(pd.concat(all_extracted_dfs, ignore_index=True))
                st.success(f"✅ Berhasil mengekstraksi {len(final_pdf_df)} baris data!")
                st.dataframe(final_pdf_df.head(50), use_container_width=True)
                st.download_button("📥 Download Excel Tipe 1 (.xlsx)", data=to_excel_bytes(final_pdf_df), file_name="Hasil_Ekstrak_PDF_Tipe1.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("⚠️ Tidak ada data tabel yang terdeteksi.")

# ==========================================
# TAB 4: IMPORT & EKSTRAK PDF (TIPE 2)
# ==========================================
with tab4:
    st.header("Tahap 4: Import & Ekstrak PDF (Tipe 2 - Semua Kolom)")
    st.markdown("Digunakan untuk format PDF standar. **Mengekspor semua kolom utuh apa adanya.**")
    pdf_files_t2 = st.file_uploader("Upload File PDF Tipe 2", type=['pdf'], accept_multiple_files=True, key="t4_pdf")

    if st.button("Proses & Ekstrak Semua Data (Tipe 2)", type="primary"):
        if not pdf_files_t2: 
            st.error("Silakan unggah setidaknya satu file PDF.")
        elif not PDF_SUPPORT: 
            st.error("Library `pdfplumber` belum di-install.")
        else:
            all_extracted_dfs = []
            error_pages = [] 
            
            with st.spinner("🚀 Sedang mengekstraksi PDF besar..."):
                for uploaded_pdf in pdf_files_t2:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp: 
                        tmp.write(uploaded_pdf.getvalue())
                        tmp_path = tmp.name
                        
                    try:
                        with pdfplumber.open(tmp_path) as pdf:
                            total_pages = len(pdf.pages)
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            
                            for i, page in enumerate(pdf.pages):
                                try:
                                    tables = page.extract_tables() 
                                    for table in tables:
                                        df_std = process_standard_table(table)
                                        if df_std is not None and not df_std.empty: 
                                            all_extracted_dfs.append(df_std)
                                except Exception as e:
                                    error_pages.append(f"Halaman {i+1}: {str(e)}")
                                finally:
                                    page.flush_cache() 
                                
                                if (i + 1) % 5 == 0 or (i + 1) == total_pages:
                                    progress_bar.progress((i + 1) / total_pages)
                                    status_text.text(f"Memproses halaman {i+1} dari {total_pages}...")
                                    gc.collect()
                                
                    except Exception as e:
                        st.error(f"Gagal membuka file PDF: {str(e)}")
                    finally:
                        if os.path.exists(tmp_path): 
                            os.remove(tmp_path)

            if all_extracted_dfs:
                try:
                    final_pdf_df = pd.concat(all_extracted_dfs, ignore_index=True)
                    
                    if len(final_pdf_df.columns) > 0:
                        first_col = final_pdf_df.columns[0]
                        final_pdf_df = final_pdf_df[final_pdf_df[first_col] != first_col]
                    
                    final_pdf_df = final_pdf_df.reset_index(drop=True)

                    st.success(f"✅ Berhasil mengekstraksi {len(final_pdf_df)} baris dengan {len(final_pdf_df.columns)} kolom utuh!")
                    
                    if error_pages:
                        st.warning(f"⚠️ Ada {len(error_pages)} halaman yang dilewati karena error format.")
                        with st.expander("Lihat Detail Error"):
                            for err in error_pages: st.write(err)

                    st.dataframe(final_pdf_df.head(50), use_container_width=True)
                    st.download_button("📥 Download Semua Data Tipe 2 (.xlsx)", data=to_excel_bytes(final_pdf_df), file_name="Hasil_Semua_Kolom_Tipe2.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                except Exception as e:
                    st.error(f"Gagal saat menggabungkan data: {str(e)}")
            else:
                st.warning("⚠️ Tidak ada data tabel yang terdeteksi.")

# ==========================================
# TAB 5: PARSING BANYAK FILE ICONPRN KE EXCEL
# ==========================================
with tab5:
    st.header("Tahap 5: Rekap Data ICONPRN")
    st.markdown("Ekstraksi file teks `.iconprn` menjadi file Excel tagihan gabungan secara otomatis.")
    
    iconprn_files = st.file_uploader("Upload File .iconprn", accept_multiple_files=True, key="t5_iconprn")
    
    if st.button("Proses File ICONPRN", type="primary"):
        if not iconprn_files:
            st.error("Silakan unggah setidaknya satu file .iconprn.")
        else:
            with st.spinner("Mengekstrak data dari file .iconprn..."):
                semua_hasil = []
                
                def tentukan_petugas(idpel, tarif_daya, kddk):
                    daya = 0
                    match_daya = re.search(r'/(\d+)', str(tarif_daya))
                    if match_daya: daya = int(match_daya.group(1))
                    if daya > 33000: return "PLN"
                    
                    idpel_khusus = {"524051069054": "c28", "524051263717": "c36", "524051265123": "c36", "524051104194": "c04", "524051000615": "c08", "524050867033": "c08"}
                    if str(idpel) in idpel_khusus: return idpel_khusus[str(idpel)]
                        
                    if len(str(kddk)) >= 6:
                        kode_mid = str(kddk)[3:6].upper()
                        mapping_kddk = {
                            "JCA": "c01", "TAA": "c02", "JCB": "c03", "JCC": "c04", "NCD": "c05",
                            "KBA": "c06", "TAB": "c07", "JCE": "c08", "TAC": "c09", "MBB": "c10",
                            "KAD": "c11", "TAE": "c12", "KCF": "c13", "JCG": "c14", "TAF": "c15",
                            "KAG": "c16", "BBC": "c17", "TAH": "c18", "BCK": "c19", "NCH": "c20",
                            "JBE": "c21", "MBF": "c22", "MBG": "c23", "KBH": "c24", "KBI": "c25",
                            "KAI": "c26", "MCI": "c27", "KBJ": "c28", "JCJ": "c29", "TAJ": "c30",
                            "MBK": "c31", "KBL": "c32", "TAK": "c33", "KAL": "c34", "JCL": "c35",
                            "MBD": "c36", "KCJ": "c29", "NBJ": "c28", "TAI": "c26", "BBD": "c19",
                            "KCG": "c29", "MBM": "c23"
                        }
                        return mapping_kddk.get(kode_mid, "BARU")
                    return "BARU"

                for file in iconprn_files:
                    try:
                        content = file.getvalue().decode('utf-8', errors='ignore')
                        blok_pelanggan = re.split(r'PEMBERITAHUAN PELAKSANAAN PEMUTUSAN', content)
                        
                        for blok in blok_pelanggan:
                            if "ID. Pelanggan" not in blok: continue
                            
                            data = {
                                'IDPEL': "", 'Nomor TUL': "", 'Nama': "", 'KDDK': "", 'Gardu/Tiang': "", 
                                'Loket': "", 'Alamat': "", 'Nomor Meter': "", 'Tarif/Daya': "", 'Kelompok': "", 
                                'Bulan Rekening': "", 'Bulan Keterlambatan': "", 'Jumlah Rekening': 0, 
                                'Jumlah Denda': 0, 'Jumlah Tunggakan': 0, 'petugas': ""
                            }
                            
                            idpel = re.search(r'ID\. Pelanggan\s*:\s*[^0-9]*(\d{11,13})', blok)
                            if idpel: data['IDPEL'] = str(idpel.group(1).strip())

                            tul = re.search(r'NO\. TUL\s*:\s*([A-Z0-9/\-]+)', blok)
                            if tul: data['Nomor TUL'] = tul.group(1).strip()

                            nama = re.search(r'Nama\s*:\s*(.+)', blok)
                            if nama: data['Nama'] = nama.group(1).strip()

                            kddk = re.search(r'Kode Kedudukan\s*:\s*([A-Z0-9]+)', blok)
                            if kddk: data['KDDK'] = kddk.group(1).strip()

                            gardu = re.search(r'Gardu\s*/?\s*Tiang\s*:\s*(.*?)(?=\s{2,}|\s+Loket\s*:|\n|\r|$)', blok, re.IGNORECASE)
                            if gardu: data['Gardu/Tiang'] = gardu.group(1).strip()

                            loket = re.search(r'Loket\s*:\s*(.*?)(?=\s{2,}|\s+Tarip|\s+Tarif|\s+Alamat|\s+Kelompok|\n|\r|$)', blok, re.IGNORECASE)
                            if loket: 
                                val_loket = loket.group(1).strip()
                                if "Tarip" in val_loket or "Kelompok" in val_loket or "Daya" in val_loket: val_loket = ""
                                data['Loket'] = val_loket

                            alamat = re.search(r'Alamat\s*:\s*(.+)', blok)
                            if alamat: data['Alamat'] = alamat.group(1).strip()

                            meter = re.search(r'Nomor Meter\s*:\s*(\d+)', blok)
                            if meter: data['Nomor Meter'] = str(meter.group(1).strip())

                            tarif_match = re.search(r'Tarip / Daya\s*:\s*(.+?)\s+Kelompok\s*:\s*([^\n]+)', blok)
                            if tarif_match:
                                data['Tarif/Daya'] = tarif_match.group(1).strip()
                                data['Kelompok'] = tarif_match.group(2).strip()
                            else:
                                tarif = re.search(r'Tarip / Daya\s*:\s*([A-Z0-9/ ]+)', blok)
                                if tarif: data['Tarif/Daya'] = tarif.group(1).strip()
                                data['Kelompok'] = "1"

                            rek = re.search(r'Rekening\s*:\s*(.+?)\s*Rp\.\s*:\s*[^0-9]*([\d,]+)', blok)
                            if rek:
                                data['Bulan Rekening'] = rek.group(1).strip()
                                val_rek = rek.group(2).replace(',', '').replace('.', '').strip()
                                data['Jumlah Rekening'] = int(val_rek) if val_rek.isdigit() else 0

                            denda = re.search(r'Jumlah Biaya Keterlambatan s\.d bulan\s*:\s*(.+?)\s*Rp\.\s*:\s*[^0-9]*([\d,]+)', blok)
                            if denda:
                                data['Bulan Keterlambatan'] = denda.group(1).strip()
                                val_denda = denda.group(2).replace(',', '').replace('.', '').strip()
                                data['Jumlah Denda'] = int(val_denda) if val_denda.isdigit() else 0

                            tunggakan = re.search(r'Jumlah Tunggakan.*?Rp\.\s*:\s*[^0-9]*([\d,]+)', blok)
                            if tunggakan:
                                val_tung = tunggakan.group(1).replace(',', '').replace('.', '').strip()
                                data['Jumlah Tunggakan'] = int(val_tung) if val_tung.isdigit() else 0

                            data['petugas'] = tentukan_petugas(data['IDPEL'], data['Tarif/Daya'], data['KDDK'])
                            semua_hasil.append(data)

                    except Exception as e:
                        st.error(f"Gagal membaca file {file.name}: {str(e)}")

                if semua_hasil:
                    df_hasil = pd.DataFrame(semua_hasil)
                    st.success(f"✅ Berhasil mengekstrak {len(df_hasil)} data tagihan!")
                    st.dataframe(df_hasil.head(50), use_container_width=True)
                    
                    st.download_button(
                        "📥 Download Rekap Tagihan ICONPRN (.xlsx)", 
                        data=to_excel_bytes(df_hasil), 
                        file_name="rekap_tagihan_gabungan.xlsx", 
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                        type="primary"
                    )
                else:
                    st.warning("⚠️ Tidak ada data tagihan valid yang ditemukan.")

# ==========================================
# TAB 6: APLIKASI UPDATE MASTER DATA
# ==========================================
with tab6:
    st.header("Tahap 6: Update Master Data Pelanggan")
    st.markdown("Meng-update data master lama dengan data baru berdasarkan IDPEL. Data bersimbol bintang (`*`) tidak akan menimpa data master.")
    
    col_t6_1, col_t6_2 = st.columns(2)
    with col_t6_1:
        file_lama_m = st.file_uploader("Upload File Excel Data LAMA", type=['xlsx', 'xls'], key="t6_lama")
    with col_t6_2:
        file_baru_m = st.file_uploader("Upload File Excel Data BARU", type=['xlsx', 'xls'], key="t6_baru")

    if st.button("Proses Update Master Data", type="primary"):
        if not file_lama_m or not file_baru_m:
            st.error("Silakan unggah kedua file Excel (Data Lama & Data Baru).")
        else:
            with st.spinner("Memproses sinkronisasi master data..."):
                try:
                    df_lama = pd.read_excel(file_lama_m, dtype=str)
                    df_baru = pd.read_excel(file_baru_m, dtype=str)

                    kolom_asli_lama = df_lama.columns.tolist()

                    df_lama.columns = df_lama.columns.astype(str).str.strip().str.lower()
                    df_baru.columns = df_baru.columns.astype(str).str.strip().str.lower()

                    kolom_yang_sama = df_baru.columns.intersection(df_lama.columns)
                    df_baru = df_baru[kolom_yang_sama]

                    if 'idpel' not in df_lama.columns or 'idpel' not in df_baru.columns:
                        st.error("❌ Kolom 'IDPEL' tidak ditemukan di salah satu file!")
                    else:
                        df_lama['idpel'] = df_lama['idpel'].str.replace('.0', '', regex=False).str.strip()
                        df_baru['idpel'] = df_baru['idpel'].str.replace('.0', '', regex=False).str.strip()

                        df_lama.set_index('idpel', inplace=True)
                        df_baru.set_index('idpel', inplace=True)

                        mask_exist = df_baru.index.isin(df_lama.index)
                        df_baru_exist = df_baru[mask_exist].copy()
                        df_baru_new = df_baru[~mask_exist].copy()

                        if not df_baru_exist.empty:
                            def bersihkan_sel(val):
                                if pd.isna(val): return np.nan
                                s = str(val).strip()
                                if '*' in s or s.lower() == 'nan' or s == '':
                                    return np.nan
                                return val

                            df_update_clean = df_baru_exist.apply(lambda col: col.map(bersihkan_sel))
                            df_lama.update(df_update_clean)

                        if not df_baru_new.empty:
                            df_lama = pd.concat([df_lama, df_baru_new])

                        df_lama.reset_index(inplace=True)

                        mapping_kolom = dict(zip([c.lower() for c in kolom_asli_lama], kolom_asli_lama))
                        df_lama.rename(columns=mapping_kolom, inplace=True)

                        st.success(f"✅ Berhasil memperbarui data! Total data sekarang: {len(df_lama)} baris.")
                        st.dataframe(df_lama.head(50), use_container_width=True)

                        st.download_button(
                            "📥 Download Master Data Updated (.xlsx)", 
                            data=to_excel_bytes(df_lama), 
                            file_name="Master_Data_Updated.xlsx", 
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                            type="primary"
                        )

                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan: {str(e)}")
                    
# ==========================================
# TAB 7: INFO DATA & LOKASI (GOOGLE SHEETS)
# ==========================================
with tab7:
    st.header("Tahap 7: Info Data & Lokasi Pelanggan")
    
    url_g_sheet = "https://docs.google.com/spreadsheets/d/1Po-6B5KvYY0uGBnqeVSocifwe6FBnDvecpf2a_JzvN0/edit?usp=sharing"
    
    @st.cache_data(ttl=600) # Cache disimpan 10 menit
    def ambil_data_terbaru():
        import re
        url_csv = re.sub(r'/edit.*', '/export?format=csv', url_g_sheet)
        df = pd.read_csv(url_csv, dtype=str)
        df = df.fillna('')
        return df
    
    # --- TOMBOL REFRESH DATA BARU ---
    kolom_kiri, kolom_kanan = st.columns([7, 3])
    with kolom_kanan:
        if st.button("🔄 Perbarui Data dari Sheet", use_container_width=True):
            st.cache_data.clear() # Perintah membersihkan memori lama
            st.rerun() # Memuat ulang aplikasi
    # --------------------------------
    
    try:
        with st.spinner("Mengambil database terbaru..."):
            tabel_utama = ambil_data_terbaru()
            
        with st.form(key="form_pencarian_reset"):
            kategori = st.selectbox("🎯 Pilih Dasar Pencarian:", ["IDPEL", "NAMA", "NOMOR METER", "SEMUA KOLOM"])
            kata_kunci = st.text_input("🔍 Masukkan Kata Kunci:")
            tombol_cari = st.form_submit_button("🔍 Cari Data Sekarang", type="primary", use_container_width=True)
        
        if tombol_cari and kata_kunci:
            kunci_bersih = str(kata_kunci).strip().lower()
            kolom_semua = list(tabel_utama.columns)
            
            if kategori == "IDPEL":
                kolom_pencarian = [k for k in kolom_semua if "ID" in str(k).upper()]
            elif kategori == "NAMA":
                kolom_pencarian = [k for k in kolom_semua if "NAMA" in str(k).upper()]
            elif kategori == "NOMOR METER":
                kolom_pencarian = [k for k in kolom_semua if "METER" in str(k).upper() or "NO" in str(k).upper()]
            else:
                kolom_pencarian = kolom_semua
            
            if not kolom_pencarian:
                kolom_pencarian = kolom_semua

            hasil_pencarian = []
            for index, baris in tabel_utama.iterrows():
                cocok = False
                for k in kolom_pencarian:
                    isi_sel = str(baris[k]).lower()
                    if kunci_bersih in isi_sel:
                        cocok = True
                        break
                if cocok:
                    hasil_pencarian.append(baris)
            
            if len(hasil_pencarian) > 0:
                st.success(f"**Ditemukan {len(hasil_pencarian)} data:**")
                
                for row in hasil_pencarian:
                    kolom_id = next((c for c in kolom_semua if 'ID' in str(c).upper()), kolom_semua[0])
                    kolom_nama = next((c for c in kolom_semua if 'NAMA' in str(c).upper()), None)
                    
                    teks_judul = f"👤 {row[kolom_id]}"
                    if kolom_nama and str(row[kolom_nama]).strip() != '':
                        teks_judul += f" - {row[kolom_nama]}"
                        
                    with st.expander(teks_judul):
                        for col in kolom_semua:
                            st.write(f"**{col}:** {row[col]}")
                        
                        lat_val, lon_val = None, None
                        for col in kolom_semua:
                            val_str = str(row[col]).strip().replace(',', '.')
                            try:
                                num_val = float(val_str)
                                if -11.0 <= num_val <= 6.0 and lat_val is None:
                                    lat_val = str(num_val)
                                elif 95.0 <= num_val <= 141.0 and lon_val is None:
                                    lon_val = str(num_val)
                            except ValueError:
                                pass
                        
                        st.markdown("---")
                        if lat_val and lon_val:
                            url_map = f"https://www.google.com/maps/search/?api=1&query={lat_val},{lon_val}"
                            st.link_button("📍 Buka Lokasi di Google Maps", url_map, type="primary", use_container_width=True)
                        else:
                            st.warning("⚠️ Data koordinat lokasi tidak ditemukan atau formatnya bukan angka.")
            else:
                st.error("❌ Data tidak ditemukan. Cek kembali kata kuncinya.")
                
    except Exception as e:
        st.error(f"Gagal memuat database. Error: {str(e)}")

# ==========================================
# TAB 8: UPDATE MASTER DATA (VERSI 2)
# ==========================================
import numpy as np

with tab8:
    st.header("Tahap 8: Update Master Data Pelanggan (Versi 2)")
    st.write("Mengisi kolom yang KOSONG di Data Lama dengan data dari Data Baru berdasarkan IDPEL. TIDAK MENAMBAH KOLOM BARU dan data yang sudah terisi TIDAK akan ditimpa.")
    st.markdown("---")

    col_l, col_b = st.columns(2)
    with col_l:
        file_lama = st.file_uploader("Upload File Excel Data LAMA", type=['xlsx', 'xls'], key="t8_lama")
    with col_b:
        file_baru = st.file_uploader("Upload File Excel Data BARU", type=['xlsx', 'xls'], key="t8_baru")

    if st.button("Proses Update Master Data (Tab 8)", type="primary"):
        if not file_lama or not file_baru:
            st.warning("⚠️ Harap upload KEDUA file Excel (Data Lama & Data Baru)!")
        else:
            with st.spinner("Sedang memproses update data..."):
                try:
                    df_lama = pd.read_excel(file_lama)
                    df_baru = pd.read_excel(file_baru)

                    # Fungsi untuk memberi penomoran pada nama kolom yang kembar (TIDAK MENGHAPUS KOLOM)
                    def buat_kolom_unik(daftar_kolom):
                        dilihat = {}
                        kolom_baru = []
                        for col in daftar_kolom:
                            if col in dilihat:
                                dilihat[col] += 1
                                kolom_baru.append(f"{col}_{dilihat[col]}")
                            else:
                                dilihat[col] = 0
                                kolom_baru.append(col)
                        return kolom_baru

                    # 1. Standarisasi nama kolom ke huruf kecil
                    cols_lama = df_lama.columns.astype(str).str.strip().str.lower()
                    cols_baru = df_baru.columns.astype(str).str.strip().str.lower()

                    # 2. Terapkan fungsi kolom unik
                    df_lama.columns = buat_kolom_unik(cols_lama)
                    df_baru.columns = buat_kolom_unik(cols_baru)
                    
                    # SIMPAN URUTAN & NAMA KOLOM DATA LAMA (Untuk membuang kolom ekstra dari data baru nanti)
                    kolom_format_lama = df_lama.columns.tolist()

                    # Cek keberadaan kolom IDPEL
                    if 'idpel' not in df_lama.columns or 'idpel' not in df_baru.columns:
                        st.error("❌ Keduanya file harus memiliki kolom 'IDPEL'!")
                    else:
                        # Ubah sel yang isinya hanya spasi/kosong menjadi NaN agar bisa diupdate
                        df_lama = df_lama.replace(r'^\s*$', np.nan, regex=True)

                        # Format kolom IDPEL agar konsisten (string tanpa angka desimal)
                        df_lama['idpel'] = df_lama['idpel'].astype(str).str.replace('.0', '', regex=False).str.strip()
                        df_baru['idpel'] = df_baru['idpel'].astype(str).str.replace('.0', '', regex=False).str.strip()

                        # Set IDPEL sebagai index untuk pencocokan
                        df_lama.set_index('idpel', inplace=True)
                        df_baru.set_index('idpel', inplace=True)

                        # df.update dengan overwrite=False HANYA mengisi cell yang kosong (NaN) di Data Lama
                        # dan MENGABAIKAN kolom-kolom baru yang tidak ada di Data Lama
                        df_lama.update(df_baru, overwrite=False)

                        # Kembalikan IDPEL menjadi kolom biasa
                        df_lama.reset_index(inplace=True)

                        # Pastikan format dan urutan kolom persis seperti Data LAMA di awal
                        df_result = df_lama[kolom_format_lama]

                        st.success("✅ Master Data berhasil diupdate (Hanya mengisi sel kosong di kolom yang sudah ada)!")
                        st.write("Preview Hasil Update:")
                        st.dataframe(df_result.head(15), use_container_width=True)

                        # Siapkan file Excel untuk didownload
                        output_t8 = io.BytesIO()
                        with pd.ExcelWriter(output_t8, engine='openpyxl') as writer:
                            df_result.to_excel(writer, index=False, sheet_name='Master_Updated')
                        
                        st.download_button(
                            label="⬇️ Download Hasil Update Master Data",
                            data=output_t8.getvalue(),
                            file_name="Master_Data_Updated.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="t8_download"
                        )

                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan saat memproses data: {e}")

# ==========================================
# TAB 9: ISI PETUGAS (BERDASARKAN DAYA, IDPEL, & KOKED/KDDK)
# ==========================================
with tab9:
    st.header("Tahap 9: Penentuan Petugas Otomatis")
    st.write("Mengisi kolom Petugas secara otomatis berdasarkan logika Daya > 33000 (PLN), IDPEL khusus, dan 3 karakter tengah KOKED/KDDK.")
    st.markdown("---")

    file_t9 = st.file_uploader("Upload File Excel (Memiliki kolom IDPEL, KOKED/KDDK, DAYA)", type=['xlsx', 'xls'], key="t9_file")

    if st.button("Proses & Isi Petugas", type="primary"):
        if file_t9 is None:
            st.warning("⚠️ Harap upload file Excel terlebih dahulu!")
        else:
            with st.spinner("Sedang memproses penentuan petugas..."):
                try:
                    df = pd.read_excel(file_t9)
                    
                    # Buat dataframe kerja dengan nama kolom huruf kecil semua untuk kemudahan olah data
                    df_work = df.copy()
                    df_work.columns = df_work.columns.astype(str).str.strip().str.lower()

                    # Jika di Excel bernama 'kddk', otomatis disesuaikan menjadi 'koked'
                    if 'kddk' in df_work.columns and 'koked' not in df_work.columns:
                        df_work['koked'] = df_work['kddk']

                    # Cek keberadaan kolom wajib
                    if not {'idpel', 'koked', 'daya'}.issubset(set(df_work.columns)):
                        st.error("❌ Error: File Excel harus memiliki kolom bernama 'IDPEL', 'KOKED' (atau 'KDDK'), dan 'DAYA'.")
                    else:
                        def tentukan_petugas_tab9(row):
                            # 1. Konversi Daya ke Angka
                            try:
                                daya_val = str(row['daya']).replace(',', '.').strip()
                                daya = float(daya_val)
                            except:
                                daya = 0.0

                            # 2. Ambil nilai IDPEL & KOKED
                            idpel = str(row['idpel']).replace('.0', '').strip() if pd.notna(row['idpel']) else ""
                            koked = str(row['koked']).strip() if pd.notna(row['koked']) else ""

                            # --- SYARAT 1: Daya > 33000 -> PLN ---
                            if daya > 33000:
                                return "PLN"
                            
                            # --- SYARAT 2: IDPEL Khusus ---
                            idpel_khusus = {
                                "524051069054": "c28", "524051263717": "c36", "524051265123": "c36",
                                "524051104194": "c04", "524051000615": "c08", "524050867033": "c08"
                            }
                            if idpel in idpel_khusus:
                                return idpel_khusus[idpel]
                            
                            # --- SYARAT 3: MID 3 karakter KOKED (karakter ke-4 s/d 6) ---
                            if len(koked) >= 6:
                                kode_mid = koked[3:6].upper()
                                mapping_kddk = {
                                    "JCA": "c01", "TAA": "c02", "JCB": "c03", "JCC": "c04", "NCD": "c05",
                                    "KBA": "c06", "TAB": "c07", "JCE": "c08", "TAC": "c09", "MBB": "c10",
                                    "KAD": "c11", "TAE": "c12", "KCF": "c13", "JCG": "c14", "TAF": "c15",
                                    "KAG": "c16", "BBC": "c17", "TAH": "c18", "BCK": "c19", "NCH": "c20",
                                    "JBE": "c21", "MBF": "c22", "MBG": "c23", "KBH": "c24", "KBI": "c25",
                                    "KAI": "c26", "MCI": "c27", "KBJ": "c28", "JCJ": "c29", "TAJ": "c30",
                                    "MBK": "c31", "KBL": "c32", "TAK": "c33", "KAL": "c34", "JCL": "c35",
                                    "MBD": "c36", "KCJ": "c29", "NBJ": "c28", "TAI": "c26", "BBD": "c19",
                                    "KCG": "c29", "MBM": "c23"
                                }
                                if kode_mid in mapping_kddk:
                                    return mapping_kddk[kode_mid]
                            
                            return "BARU"

                        # Jalankan fungsi penentuan petugas
                        hasil_petugas = df_work.apply(tentukan_petugas_tab9, axis=1)

                        # Cari nama kolom 'petugas' asli di file Excel (jika sudah ada)
                        col_petugas_asli = next((c for c in df.columns if str(c).strip().lower() == 'petugas'), None)
                        
                        if col_petugas_asli:
                            df[col_petugas_asli] = hasil_petugas
                        else:
                            df['petugas'] = hasil_petugas

                        st.success("✅ Kolom Petugas berhasil diisi!")
                        st.write("Preview Hasil Data:")
                        st.dataframe(df.head(15), use_container_width=True)

                        # Download File
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='Data_Petugas')
                        hasil_excel = output.getvalue()

                        st.download_button(
                            label="⬇️ Download Hasil Excel",
                            data=hasil_excel,
                            file_name="Data_Isi_Petugas.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="t9_download"
                        )
                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan saat memproses data: {e}")

# ==========================================
# TAB 10: SPLIT DATA MENJADI MULTIPLE WORKSHEET / FILES
# ==========================================
import re
import zipfile

with tab10:
    st.header("Tahap 10: Split Data ke Beberapa Worksheet atau File")
    st.write("Membagi satu tabel data menjadi beberapa worksheet (sheet) atau file Excel terpisah berdasarkan nilai pada kolom tertentu.")
    st.markdown("---")

    file_t10 = st.file_uploader("Upload File Excel yang ingin di-split", type=['xlsx', 'xls'], key="t10_file")

    if file_t10 is not None:
        try:
            df_t10 = pd.read_excel(file_t10)
            st.write("Preview Data Asli:")
            st.dataframe(df_t10.head(), use_container_width=True)

            st.markdown("### Pengaturan Split Data")
            
            # 1. Pilih kolom sebagai acuan
            kolom_pilihan = st.selectbox("Split berdasarkan kolom (Specific column):", df_t10.columns.tolist())

            # 2. Opsi Prefix & Suffix untuk penamaan
            col1, col2 = st.columns(2)
            with col1:
                prefix = st.text_input("Prefix (opsional):", help="Tambahan teks di depan nama sheet/file")
            with col2:
                suffix = st.text_input("Suffix (opsional):", help="Misal: Nik Padan")

            # 3. Pilihan Output (Sheet vs File ZIP)
            mode_split = st.radio(
                "Pilih Mode Output:", 
                ["Multiple Sheets (1 File Excel)", "Multiple Files (Download sebagai ZIP)"],
                help="Pilih apakah ingin hasil split berada dalam 1 file beda sheet, atau file yang benar-benar terpisah."
            )

            if st.button("Proses Split Data", type="primary"):
                with st.spinner("Sedang membagi data..."):
                    
                    # Mengambil nilai unik dari kolom yang dipilih (abaikan yang kosong/NaN)
                    nilai_unik = df_t10[kolom_pilihan].dropna().unique()

                    if mode_split == "Multiple Sheets (1 File Excel)":
                        output_t10 = io.BytesIO()
                        with pd.ExcelWriter(output_t10, engine='openpyxl') as writer:
                            for nilai in nilai_unik:
                                df_filtered = df_t10[df_t10[kolom_pilihan] == nilai]

                                # Pembuatan nama sheet
                                nilai_str = str(nilai).strip()
                                nama_custom = f"{prefix} {nilai_str} {suffix}".strip()
                                
                                # Bersihkan karakter terlarang untuk Excel
                                nama_bersih = re.sub(r'[\\/*?:\[\]]', '', nama_custom)[:31]
                                if not nama_bersih: nama_bersih = "Data"

                                df_filtered.to_excel(writer, index=False, sheet_name=nama_bersih)

                        st.success(f"✅ Data berhasil dipisah menjadi {len(nilai_unik)} sheet dalam 1 file Excel!")
                        st.download_button(
                            label="⬇️ Download Excel (Multiple Sheets)",
                            data=output_t10.getvalue(),
                            file_name=f"Data_Split_{kolom_pilihan}_Sheets.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="t10_download_sheets"
                        )

                    else:
                        # Mode: Multiple Files via ZIP
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for nilai in nilai_unik:
                                df_filtered = df_t10[df_t10[kolom_pilihan] == nilai]
                                
                                # Pembuatan nama file
                                nilai_str = str(nilai).strip()
                                nama_custom = f"{prefix} {nilai_str} {suffix}".strip()
                                nama_bersih = re.sub(r'[\\/*?:\[\]<>|"]', '', nama_custom)
                                if not nama_bersih: nama_bersih = "Data"
                                nama_file = f"{nama_bersih}.xlsx"

                                # Bikin file excel di memory untuk file spesifik ini
                                excel_buffer = io.BytesIO()
                                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                    df_filtered.to_excel(writer, index=False, sheet_name="Data")
                                
                                # Simpan file excel ke dalam ZIP
                                zip_file.writestr(nama_file, excel_buffer.getvalue())

                        st.success(f"✅ Data berhasil dipisah menjadi {len(nilai_unik)} file Excel terpisah!")
                        st.download_button(
                            label="⬇️ Download ZIP (Multiple Files)",
                            data=zip_buffer.getvalue(),
                            file_name=f"Data_Split_{kolom_pilihan}_Files.zip",
                            mime="application/zip",
                            key="t10_download_zip"
                        )

        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat memproses data: {e}")
