import streamlit as st
import pandas as pd
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
st.set_page_config(page_title="Aplikasi Olah Koked & Ekstrak PDF", layout="wide")

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
st.title("⚡ Aplikasi Olah Koked & Ekstrak PDF")

tab1, tab2, tab3, tab4 = st.tabs([
    "1️⃣ Tahap 1: Persiapan Data", 
    "2️⃣ Tahap 2: Mutasi KOKED",
    "3️⃣ Tahap 3: PDF (Tipe 1 - Filter Kolom)",
    "4️⃣ Tahap 4: PDF (Tipe 2 - Semua Kolom)"
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
                                page.flush_cache() # Hapus cache memori
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
# TAB 4: IMPORT & EKSTRAK PDF (TIPE 2 - SEMUA KOLOM + ANTI CRASH)
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
            
            with st.spinner("🚀 Sedang mengekstraksi PDF besar... (Mohon jangan tutup/pindah halaman ini)"):
                for uploaded_pdf in pdf_files_t2:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp: 
                        tmp.write(uploaded_pdf.getvalue())
                        tmp_path = tmp.name
                        
                    try:
                        with pdfplumber.open(tmp_path) as pdf:
                            total_pages = len(pdf.pages)
                            
                            # Menggunakan st.empty() agar tampilan tidak freeze
                            progress_bar = st.progress(0)
                            status_text = st.empty()
                            
                           for i, page in enumerate(pdf.pages):
                                try:
                                    # --- FITUR ANTI WATERMARK (VERSI LEBIH KETAT) ---
                                    def filter_watermark(obj):
                                        if obj.get("object_type") == "char":
                                            # 1. Filter Ukuran: Teks tabel kecil (8-10pt). Watermark jauh lebih besar.
                                            # Buang semua teks yang ukurannya lebih dari 12pt.
                                            if obj.get("size", 0) > 12:
                                                return False
                                                
                                            # 2. Filter Transparansi (Alpha): Buang teks yang agak tembus pandang/pudar.
                                            if obj.get("non_stroking_alpha", 1) < 1:
                                                return False
                                                
                                            # 3. Filter Kemiringan (Berjaga-jaga)
                                            if not obj.get("upright", True):
                                                return False
                                        return True
                                    
                                    # Terapkan filter ke halaman
                                    clean_page = page.filter(filter_watermark)
                                    tables = clean_page.extract_tables() 
                                    # ------------------------------------------------
                                    
                                    for table in tables:
                                        df_std = process_standard_table(table)
                                        if df_std is not None and not df_std.empty: 
                                            all_extracted_dfs.append(df_std)
                                except Exception as e:
                                    error_pages.append(f"Halaman {i+1}: {str(e)}")
                                finally:
                                    # Membersihkan cache pdfplumber setiap halaman agar RAM tidak jebol
                                    page.flush_cache() 
                                
                                # Mengurangi beban update UI Browser Streamlit
                                if (i + 1) % 5 == 0 or (i + 1) == total_pages:
                                    progress_bar.progress((i + 1) / total_pages)
                                    status_text.text(f"Memproses halaman {i+1} dari {total_pages}...")
                                    gc.collect() # Panggil Garbage Collector
                                
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
