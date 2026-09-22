import streamlit as st
import pandas as pd
from dbfread import DBF
import os
import io
import zipfile
import tempfile
import re

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
    """Pembersihan IDPEL standar 12 digit, hapus huruf/karakter di belakangnya."""
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

# --- LOGIKA EKSTRAKSI TABEL PDF HYBRID ---
def get_pdf_tables(page):
    """Mengambil tabel dengan berbagai toleransi batas kolom PDF."""
    tables = page.extract_tables()
    if tables and len(tables) > 0 and any(len(t) > 1 for t in tables if t):
        return tables
    
    try:
        tables = page.extract_tables({
            "vertical_strategy": "text",
            "horizontal_strategy": "text",
            "intersection_y_tolerance": 5,
        })
        if tables and len(tables) > 0 and any(len(t) > 1 for t in tables if t):
            return tables
    except Exception:
        pass

    return []

def is_new_record(row):
    """Mengecek apakah baris memiliki 12-digit IDPEL (Data Baru)."""
    row_str = " ".join([str(c) for c in row if c])
    return bool(re.search(r'\b\d{11,12}\b', row_str))

def process_hybrid_table(table):
    if not table or len(table) < 1:
        return None

    cleaned_table = []
    for row in table:
        cleaned_row = [re.sub(r'\s+', ' ', str(cell)).strip() if cell and str(cell) != 'None' else "" for cell in row]
        if any(cleaned_row):
            cleaned_table.append(cleaned_row)

    if not cleaned_table:
        return None

    # 1. Cari Baris Header Utama
    header_keywords = ['ID PEL', 'IDPEL', 'NAMA', 'ALAMAT', 'TARIF', 'DAYA', 'AGENDA', 'REGISTER']
    header_index = -1

    for idx, row in enumerate(cleaned_table[:10]):
        row_str = " ".join([c.upper() for c in row])
        matches = sum(1 for kw in header_keywords if kw in row_str)
        if matches >= 2 or ('ID' in row_str and 'NAMA' in row_str):
            header_index = idx
            break

    if header_index == -1:
        header_index = 0

    header_row_1 = cleaned_table[header_index]
    data_start_index = header_index + 1

    # 2. Cek apakah ada Sub-Header di baris berikutnya (misal: LAMA / BARU)
    if header_index + 1 < len(cleaned_table):
        next_row = cleaned_table[header_index + 1]
        next_row_str = " ".join([c.upper() for c in next_row])
        if not is_new_record(next_row) and any(k in next_row_str for k in ['LAMA', 'BARU', 'VA']):
            combined_headers = []
            for h1, h2 in zip(header_row_1, next_row):
                combined = f"{h1} {h2}".strip()
                combined_headers.append(combined)
            header_row_1 = combined_headers
            data_start_index = header_index + 2

    headers = [h.upper() if h != "" else f"KOLOM_{i+1}" for i, h in enumerate(header_row_1)]

    # 3. Proses Baris Data & Gabungkan Lanjutan Baris (Multi-line Address)
    data_rows = cleaned_table[data_start_index:]
    valid_rows = []

    for row in data_rows:
        row_str = " ".join([c.upper() for c in row if c])
        
        # Abaikan header berulang / footer
        if ('ID PEL' in row_str or 'IDPEL' in row_str) and ('NAMA' in row_str or 'ALAMAT' in row_str):
            continue
        if 'TOTAL' in row_str or 'HALAMAN' in row_str:
            continue

        if is_new_record(row):
            valid_rows.append(list(row))
        else:
            # Jika tidak ada IDPEL baru, gabungkan teks ke baris pelanggan sebelumnya
            if valid_rows:
                for i in range(min(len(row), len(valid_rows[-1]))):
                    val = str(row[i]).strip()
                    if val and val.upper() not in ['LAMA', 'BARU', '0']:
                        if valid_rows[-1][i]:
                            valid_rows[-1][i] = (valid_rows[-1][i] + " " + val).strip()
                        else:
                            valid_rows[-1][i] = val

    if not valid_rows:
        return pd.DataFrame(columns=headers)

    # Samakan panjang kolom
    num_cols = len(headers)
    padded_rows = []
    for r in valid_rows:
        if len(r) < num_cols:
            r = r + [""] * (num_cols - len(r))
        elif len(r) > num_cols:
            r = r[:num_cols]
        padded_rows.append(r)

    return pd.DataFrame(padded_rows, columns=headers)

def find_target_column(col_name):
    """Pencocokan nama kolom otomatis."""
    col_clean = re.sub(r'[^A-Z0-9\s]', '', str(col_name).upper()).strip()
    
    if any(k in col_clean for k in ['IDPEL', 'ID PEL', 'ID_PEL', 'NO PELANGGAN', 'NOPEL', 'IDPELANGGAN']):
        return 'IDPEL'
    if any(k in col_clean for k in ['NAMA PELANGGAN', 'NAMA_PELANGGAN', 'NAMAPNJ', 'NAMA PEL']) or col_clean == 'NAMA':
        return 'NAMA'
    if any(k in col_clean for k in ['ALAMAT PELANGGAN', 'ALAMAT_PELANGGAN', 'ALAMAT', 'ALM']):
        return 'ALAMAT'
    if any(k in col_clean for k in ['TARIF', 'TARIP', 'GOLTAR', 'GOL TARIF']):
        return 'TARIF'
    if any(k in col_clean for k in ['DAYA', 'KAPASITAS']) or 'VA' in col_clean:
        return 'DAYA'
    if 'GARDU' in col_clean or col_clean == 'GD':
        return 'GARDU'
    if 'TIANG' in col_clean or col_clean == 'TG':
        return 'TIANG'
    if any(k in col_clean for k in ['KOKED', 'KDDK', 'KODE KEDUDUKAN']):
        return 'KOKED'
    return col_name

def separate_idpel_and_nama(df):
    """Jika IDPEL dan NAMA tergabung dalam satu sel, pisahkan secara otomatis."""
    if 'IDPEL' in df.columns:
        for idx in df.index:
            val = str(df.at[idx, 'IDPEL']).strip()
            match = re.match(r'^(\d{11,12})\s+(.+)$', val)
            if match:
                df.at[idx, 'IDPEL'] = match.group(1)
                nama_val = str(df.at[idx, 'NAMA']).strip() if 'NAMA' in df.columns else ''
                if not nama_val or nama_val == 'None':
                    df.at[idx, 'NAMA'] = match.group(2)
    return df

def normalize_pdf_dataframe(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=['IDPEL', 'NAMA', 'ALAMAT', 'TARIF', 'DAYA', 'GARDU', 'TIANG', 'KOKED'])

    new_cols = [find_target_column(c) for c in df.columns]
    df.columns = new_cols

    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    df = separate_idpel_and_nama(df)

    target_cols = ['IDPEL', 'NAMA', 'ALAMAT', 'TARIF', 'DAYA', 'GARDU', 'TIANG', 'KOKED']
    for col in target_cols:
        if col not in df.columns:
            df[col] = ""

    df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
    df['DAYA'] = df['DAYA'].apply(clean_daya)

    # Hapus baris kosong / baris header tersisa
    df = df[df['IDPEL'].astype(str).str.strip() != ''].copy()

    return df[target_cols].reset_index(drop=True)

# --- FUNGSI PEMBACA FILE UMUM ---
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
            elif ext == '.pdf':
                if not PDF_SUPPORT:
                    st.error("Library pdfplumber belum terinstall.")
                    return pd.DataFrame(), fname
                all_pages_data = []
                with pdfplumber.open(tmp_path) as pdf:
                    for page in pdf.pages:
                        tables = get_pdf_tables(page)
                        for table in tables:
                            df_tbl = process_hybrid_table(table)
                            if df_tbl is not None and not df_tbl.empty:
                                all_pages_data.append(df_tbl)
                if all_pages_data:
                    df = pd.concat(all_pages_data, ignore_index=True)
                else:
                    df = pd.DataFrame()
            else:
                st.error(f"Format tidak didukung: {ext}")
                return pd.DataFrame(), fname
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    df.columns = df.columns.astype(str).str.strip().str.upper()
    df = df.rename(columns={
        'KDDK': 'KOKED', 
        'TARIF': 'TARIP', 
        'GOL TARIF': 'TARIP',
        'NAMAPNJ': 'ALAMAT',
        'ID PEL': 'IDPEL',
        'ID_PELANGGAN': 'IDPEL'
    })
    return df, fname

def proses_list_file(files, tipe_data):
    list_df = []
    for f in files:
        df, fname = baca_ekstrak_tabel(f)
        if df.empty: continue
        
        if tipe_data in ['baru', 'petugas']:
            if 'ALAMAT' not in df.columns: df['ALAMAT'] = ''
            if 'NAMA' not in df.columns: df['NAMA'] = ''
            if 'TARIP' not in df.columns: df['TARIP'] = ''
            if 'DAYA' not in df.columns: df['DAYA'] = 0
            if 'KOKED' not in df.columns: df['KOKED'] = ''
            
            if 'IDPEL' not in df.columns: 
                raise ValueError(f"Kolom wajib 'IDPEL' hilang di file {fname}")
                
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

tab1, tab2, tab3 = st.tabs([
    "1️⃣ Tahap 1: Persiapan Data Petugas", 
    "2️⃣ Tahap 2: Rekap Pekerjaan & Mutasi KOKED",
    "3️⃣ Tahap 3: Import & Ekstrak PDF ke Excel"
])

# ==========================================
# TAB 1: PERSIAPAN DATA
# ==========================================
with tab1:
    st.header("Tahap 1: Memecah Data Untuk Petugas Lapangan")
    st.info("Upload data server, bagi menjadi per wilayah (C01, C02, dll), dan pisahkan Pelanggan Baru (PB).")
    
    col1, col2 = st.columns(2)
    with col1:
        files_baru = st.file_uploader("[Tahap 1] Data Server Bulan INI", accept_multiple_files=True, key="t1_baru")
        files_lama = st.file_uploader("[Tahap 1] Data Bulan LALU (Pembanding PB)", accept_multiple_files=True, key="t1_lama")
    with col2:
        files_master = st.file_uploader("[Tahap 1] Data Master (Ganti ***)", accept_multiple_files=True, key="t1_master")
        files_suppb = st.file_uploader("[Tahap 1] Data PB Baru (Pelengkap)", accept_multiple_files=True, key="t1_suppb")

    if st.button("Proses Tahap 1", type="primary"):
        if not files_baru or not files_lama:
            st.error("Silakan unggah Data Bulan Ini dan Data Bulan Lalu.")
        else:
            with st.spinner("Menyiapkan data untuk petugas..."):
                try:
                    df_baru = proses_list_file(files_baru, 'baru')
                    df_lama = proses_list_file(files_lama, 'lama')
                    df_master = proses_list_file(files_master, 'master') if files_master else pd.DataFrame()
                    df_sup_pb = proses_list_file(files_suppb, 'sup_pb') if files_suppb else pd.DataFrame()

                    df_baru = fix_masked_info(df_baru, df_master, df_sup_pb)

                    idpel_block = '524050450911'
                    df_baru = df_baru[(df_baru['DAYA'] <= 33000) & (df_baru['IDPEL'] != idpel_block)].copy()
                    
                    list_idpel_lama = set(df_lama['IDPEL'].tolist())
                    df_pb = df_baru[~df_baru['IDPEL'].isin(list_idpel_lama)].copy()
                    df_tetap = df_baru[df_baru['IDPEL'].isin(list_idpel_lama)].copy()

                    kolom_export = ['IDPEL', 'KOKED', 'NAMA', 'ALAMAT', 'TARIP', 'DAYA']
                    zip_buffer = io.BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        if not df_pb.empty:
                            zip_file.writestr("PB_SEMUA_PETUGAS.xlsx", to_excel_bytes(df_pb[kolom_export]))

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
                                if isinstance(kriteria, list): temp_df = df_source[df_source['KOKED'].str[3:6].isin(kriteria)]
                                else: temp_df = df_source[df_source['KOKED'].str[3:6] == kriteria]
                                
                                if not temp_df.empty:
                                    temp_df = temp_df.drop_duplicates(subset=['IDPEL'])
                                    zip_file.writestr(f"{prefix}{nama_file}.xlsx", to_excel_bytes(temp_df[kolom_export]))

                        add_to_zip(df_pb, "PB_")
                        add_to_zip(df_tetap, "")

                    st.success("✅ File untuk petugas berhasil dibuat!")
                    st.download_button("📥 Download Distribusi Petugas (.zip)", data=zip_buffer.getvalue(), file_name="Hasil_PerPetugas.zip", mime="application/zip")
                except Exception as e:
                    st.error(f"❌ Error Tahap 1: {str(e)}")

# ==========================================
# TAB 2: REKAP & PERUBAHAN KOKED
# ==========================================
with tab2:
    st.header("Tahap 2: Gabung File Petugas & Ekstrak Perubahan KOKED")
    st.info("Upload kembali file-file dari petugas yang KOKED-nya sudah diedit/diurutkan di lapangan, lalu bandingkan dengan bulan lalu.")
    
    col3, col4 = st.columns(2)
    with col3:
        files_petugas = st.file_uploader("[Tahap 2] Data Hasil Kerja Petugas (Bisa >1 File)", accept_multiple_files=True, key="t2_petugas")
    with col4:
        files_lama_pembanding = st.file_uploader("[Tahap 2] Data Bulan Lalu (Untuk Cek Mutasi KOKED)", accept_multiple_files=True, key="t2_lama")

    if st.button("Proses Tahap 2", type="primary"):
        if not files_petugas or not files_lama_pembanding:
            st.error("Silakan unggah Data Hasil Kerja Petugas dan Data Bulan Lalu.")
        else:
            with st.spinner("Membandingkan KOKED dan menyusun hasil akhir..."):
                try:
                    df_petugas = proses_list_file(files_petugas, 'petugas')
                    df_lama_pem = proses_list_file(files_lama_pembanding, 'lama')

                    df_compare = df_petugas[['IDPEL', 'KOKED']].merge(df_lama_pem[['IDPEL', 'KOKED']], on='IDPEL', suffixes=('_BARU', '_LAMA'))
                    df_changed = df_compare[(df_compare['KOKED_BARU'] != df_compare['KOKED_LAMA']) & (df_compare['KOKED_LAMA'] != '')]
                    
                    df_petugas['NO_URUT'] = df_petugas['KOKED'].str[7:10]
                    df_petugas['NO_URUT'] = pd.to_numeric(df_petugas['NO_URUT'], errors='coerce').fillna(0).astype(int)
                    df_petugas.sort_values(by=['KOKED', 'NO_URUT'], inplace=True)
                    kolom_export = ['IDPEL', 'KOKED', 'NAMA', 'ALAMAT', 'TARIP', 'DAYA']

                    zip_buffer2 = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer2, "w", zipfile.ZIP_DEFLATED) as zip_file2:
                        if not df_changed.empty:
                            txt_content = (df_changed['IDPEL'].astype(str) + "|" + df_changed['KOKED_BARU'].astype(str)).tolist()
                            txt_string = "\n".join(txt_content)
                            zip_file2.writestr("PERUBAHAN_KOKED.txt", txt_string)

                        zip_file2.writestr("DATA_GABUNGAN_FINAL.xlsx", to_excel_bytes(df_petugas[kolom_export]))

                    st.success(f"✅ Tahap 2 selesai! Terdeteksi {len(df_changed)} data yang KOKED-nya berubah.")
                    st.download_button(
                        label="📥 Download Hasil Akhir (.zip)", 
                        data=zip_buffer2.getvalue(), 
                        file_name="Hasil_Koked.zip", 
                        mime="application/zip",
                        type="primary"
                    )

                except Exception as e:
                    st.error(f"❌ Error Tahap 2: {str(e)}")

# ==========================================
# TAB 3: IMPORT & EKSTRAK PDF KE EXCEL
# ==========================================
with tab3:
    st.header("Tahap 3: Import & Ekstrak PDF ke File Excel")
    st.info("Unggah satu atau beberapa file PDF. Sistem akan mengekstraksi data secara presisi, merapikan baris/kalimat terpotong, menyesuaikan header secara otomatis, dan menyimpannya menjadi satu file Excel.")

    pdf_files = st.file_uploader("Upload File PDF (Bisa Pilih Banyak File)", type=['pdf'], accept_multiple_files=True, key="t3_pdf")

    if st.button("Proses & Ekstrak PDF", type="primary"):
        if not pdf_files:
            st.error("Silakan unggah setidaknya satu file PDF.")
        elif not PDF_SUPPORT:
            st.error("Library `pdfplumber` belum di-install. Install via terminal: `pip install pdfplumber`")
        else:
            all_extracted_dfs = []
            
            with st.spinner("Mengekstraksi data dari PDF..."):
                for uploaded_pdf in pdf_files:
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
                        tmp.write(uploaded_pdf.getvalue())
                        tmp_path = tmp.name

                    try:
                        with pdfplumber.open(tmp_path) as pdf:
                            for page in pdf.pages:
                                tables = get_pdf_tables(page)
                                for table in tables:
                                    df_hybrid = process_hybrid_table(table)
                                    if df_hybrid is not None and not df_hybrid.empty:
                                        all_extracted_dfs.append(df_hybrid)
                    except Exception as e:
                        st.error(f"Gagal memproses file {uploaded_pdf.name}: {e}")
                    finally:
                        if os.path.exists(tmp_path):
                            os.remove(tmp_path)

            if all_extracted_dfs:
                combined_raw_df = pd.concat(all_extracted_dfs, ignore_index=True)
                final_pdf_df = normalize_pdf_dataframe(combined_raw_df)

                st.success(f"✅ Berhasil mengekstraksi {len(final_pdf_df)} baris data dari {len(pdf_files)} file PDF!")
                st.subheader("Preview Data Hasil Ekstraksi:")
                st.dataframe(final_pdf_df.head(50), use_container_width=True)

                excel_data = to_excel_bytes(final_pdf_df)
                st.download_button(
                    label="📥 Download Data Ekstraksi PDF (.xlsx)",
                    data=excel_data,
                    file_name="Hasil_Ekstrak_PDF.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
            else:
                st.warning("⚠️ Tidak ada data atau tabel yang berhasil dideteksi dari file PDF yang diunggah.")
