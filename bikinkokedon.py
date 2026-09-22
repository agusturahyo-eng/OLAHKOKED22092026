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
def process_hybrid_table(table):
    if not table or len(table) < 2:
        return None

    merged_rows = []
    for row in table:
        cleaned_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
        if not any(cleaned_row):
            continue
        
        if not merged_rows or cleaned_row[0] != "":
            merged_rows.append(cleaned_row)
        else:
            for i in range(len(cleaned_row)):
                if cleaned_row[i] != "":
                    merged_rows[-1][i] = (merged_rows[-1][i] + " " + cleaned_row[i]).strip()

    if not merged_rows:
        return None

    headers = []
    for i, h in enumerate(merged_rows[0]):
        val = h if h != "" else f"KOLOM_{i+1}"
        headers.append(val.upper())
        
    df = pd.DataFrame(merged_rows[1:], columns=headers)
    return df

def normalize_pdf_dataframe(df):
    """Mencocokkan variasi header secara otomatis dan menyusun kolom target."""
    target_cols = ['IDPEL', 'NAMA', 'ALAMAT', 'TARIF', 'DAYA', 'GARDU', 'TIANG', 'KOKED']
    
    alias_map = {
        'IDPEL': ['IDPEL', 'ID PEL', 'ID PELANGGAN', 'ID_PELANGGAN', 'NO PELANGGAN', 'NOPEL', 'ID'],
        'NAMA': ['NAMA', 'NAMA PELANGGAN', 'NAMA_PELANGGAN', 'NAMAPNJ', 'PELANGGAN'],
        'ALAMAT': ['ALAMAT', 'ALAMAT PELANGGAN', 'ALAMAT_PELANGGAN', 'ALM'],
        'TARIF': ['TARIF', 'TARIP', 'GOL TARIF', 'GOLTAR', 'GOL TARIP'],
        'DAYA': ['DAYA', 'DAYA (VA)', 'KAPASITAS', 'VA'],
        'GARDU': ['GARDU', 'NO GARDU', 'KODE GARDU', 'GD'],
        'TIANG': ['TIANG', 'NO TIANG', 'KODE TIANG', 'TG'],
        'KOKED': ['KOKED', 'KDDK', 'KODE KEDUDUKAN']
    }

    current_cols = [str(c).strip().upper() for c in df.columns]
    df.columns = current_cols

    rename_dict = {}
    for target, aliases in alias_map.items():
        for col in current_cols:
            if col in aliases or any(alias in col for alias in aliases):
                rename_dict[col] = target
                break

    df = df.rename(columns=rename_dict)

    for col in target_cols:
        if col not in df.columns:
            df[col] = ""

    df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
    df['DAYA'] = df['DAYA'].apply(clean_daya)

    # Antisipasi jika rename menghasilkan duplikat sebelum filter target_cols
    df = df.loc[:, ~df.columns.duplicated()]
    
    # Ambil kolom yang ada di target_cols
    available_cols = [col for col in target_cols if col in df.columns]
    return df[available_cols]

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
                        tables = page.extract_tables()
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
        
        # Penanganan kolom wajib secara rinci
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
    with col4:Tentu, saya bisa membantu. Apa yang ingin Anda gabungkan? 

Silakan unggah file yang dimaksud (misalnya PDF, dokumen Word, atau gambar) atau tempelkan teksnya di sini, dan beri tahu saya format akhir yang Anda inginkan.
