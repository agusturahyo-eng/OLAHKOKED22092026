import streamlit as st
import pandas as pd
from dbfread import DBF
import os
import io
import zipfile
import tempfile
import re

# --- PENGATURAN HALAMAN ---
st.set_page_config(page_title="Aplikasi Olah Koked", layout="wide")

# --- FUNGSI MENGONVERSI & MEMBERSIHKAN DATA ---
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

def clean_idpel(x):
    if pd.isna(x): return ""
    s = str(x).split('.')[0]
    return re.sub(r'\D', '', s)

# --- FUNGSI UN-MASKING (NAMA/ALAMAT BINTANG & KOKED KOSONG) ---
def fix_masked_info(df_baru, df_master, df_sup_pb):
    if df_baru.empty: return df_baru
    df_baru = df_baru.reset_index(drop=True)
    nama_map, alamat_map, koked_map = {}, {}, {}

    # 1. Dari Data PB Baru (Fokus ambil Nama & Alamat. Koked kalau ada)
    if not df_sup_pb.empty:
        sup_clean = df_sup_pb.drop_duplicates(subset=['IDPEL'], keep='first')
        
        valid_nama = sup_clean[sup_clean['NAMA'].astype(str).str.strip().ne('') & ~sup_clean['NAMA'].astype(str).str.contains(r'\*', na=False)]
        nama_map.update(valid_nama.set_index('IDPEL')['NAMA'].to_dict())
        
        valid_alamat = sup_clean[sup_clean['ALAMAT'].astype(str).str.strip().ne('') & ~sup_clean['ALAMAT'].astype(str).str.contains(r'\*', na=False)]
        alamat_map.update(valid_alamat.set_index('IDPEL')['ALAMAT'].to_dict())
        
        if 'KOKED' in sup_clean.columns:
            valid_koked = sup_clean[sup_clean['KOKED'].astype(str).str.strip().ne('')]
            koked_map.update(valid_koked.set_index('IDPEL')['KOKED'].to_dict())

    # 2. Dari Data Master
    if not df_master.empty:
        master_clean = df_master.drop_duplicates(subset=['IDPEL'], keep='first')
        
        valid_nama = master_clean[master_clean['NAMA'].astype(str).str.strip().ne('') & ~master_clean['NAMA'].astype(str).str.contains(r'\*', na=False)]
        nama_map.update(valid_nama.set_index('IDPEL')['NAMA'].to_dict())
        
        valid_alamat = master_clean[master_clean['ALAMAT'].astype(str).str.strip().ne('') & ~master_clean['ALAMAT'].astype(str).str.contains(r'\*', na=False)]
        alamat_map.update(valid_alamat.set_index('IDPEL')['ALAMAT'].to_dict())
        
        if 'KOKED' in master_clean.columns:
            valid_koked = master_clean[master_clean['KOKED'].astype(str).str.strip().ne('')]
            koked_map.update(valid_koked.set_index('IDPEL')['KOKED'].to_dict())

    # Cek yang perlu ditambal di Data Baru
    is_masked_nama = df_baru['NAMA'].astype(str).str.contains(r'\*', na=False) | df_baru['NAMA'].isna() | (df_baru['NAMA'].astype(str).str.strip() == '')
    is_masked_alamat = df_baru['ALAMAT'].astype(str).str.contains(r'\*', na=False) | df_baru['ALAMAT'].isna() | (df_baru['ALAMAT'].astype(str).str.strip() == '')
    is_empty_koked = df_baru['KOKED'].isna() | (df_baru['KOKED'].astype(str).str.strip() == '')

    # Apply mapping
    if nama_map:
        df_baru.loc[is_masked_nama, 'NAMA'] = df_baru.loc[is_masked_nama, 'IDPEL'].map(nama_map).fillna(df_baru.loc[is_masked_nama, 'NAMA'])
    if alamat_map:
        df_baru.loc[is_masked_alamat, 'ALAMAT'] = df_baru.loc[is_masked_alamat, 'IDPEL'].map(alamat_map).fillna(df_baru.loc[is_masked_alamat, 'ALAMAT'])
    if koked_map:
        df_baru.loc[is_empty_koked, 'KOKED'] = df_baru.loc[is_empty_koked, 'IDPEL'].map(koked_map).fillna(df_baru.loc[is_empty_koked, 'KOKED'])

    return df_baru

# --- FUNGSI PEMBACA FILE (STREAMLIT ADAPTER + HYBRID PDF) ---
def baca_ekstrak_tabel(uploaded_file):
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    fname = uploaded_file.name
    
    if ext in ['.xlsx', '.xls', '.csv']:
        if ext == '.csv':
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
    else:
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            if ext == '.dbf':
                # Menggunakan dbfread agar tidak error
                df = pd.DataFrame(iter(DBF(tmp_path, char_decode_errors='ignore')))
            elif ext == '.pdf':
                try:
                    import pdfplumber
                    all_dfs = []
                    table_settings = {
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "text",
                    }
                    with pdfplumber.open(tmp_path) as pdf:
                        for page in pdf.pages:
                            table = page.extract_table(table_settings)
                            if not table or len(table) < 2: continue
                            
                            merged_rows = []
                            for row in table:
                                cleaned_row = [str(cell).replace('\n', ' ').strip() if cell else "" for cell in row]
                                if not any(cleaned_row): continue
                                
                                if not merged_rows or cleaned_row[0] != "":
                                    merged_rows.append(cleaned_row)
                                else:
                                    for i in range(len(cleaned_row)):
                                        if cleaned_row[i] != "" and i < len(merged_rows[-1]):
                                            merged_rows[-1][i] = (str(merged_rows[-1][i]) + " " + cleaned_row[i]).strip()

                            headers = [str(h).replace('\n', ' ').strip().upper() if h else f"KOLOM_{i}" for i, h in enumerate(merged_rows[0])]
                            df_page = pd.DataFrame(merged_rows[1:], columns=headers)
                            if not df_page.empty:
                                all_dfs.append(df_page)
                    
                    df = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()
                except ImportError:
                    st.error("Library PDF (pdfplumber) belum terinstall. Pastikan ada di requirements.txt!")
                    return pd.DataFrame(), fname
            else:
                st.error(f"Format tidak didukung: {ext}")
                return pd.DataFrame(), fname
        finally:
            os.remove(tmp_path)

    # Standardisasi Nama Kolom secara Fleksibel
    if not df.empty:
        df.columns = df.columns.astype(str).str.replace('\n', ' ').str.strip().str.upper()
        rename_dict = {}
        
        for col in df.columns:
            col_clean = col.upper().strip()
            
            if any(kw in col_clean for kw in ['ID PEL', 'IDPEL', 'NOPEL', 'NO PEL']) and 'TETANGGA' not in col_clean:
                rename_dict[col] = 'IDPEL'
            elif any(kw in col_clean for kw in ['NAMA PEMOHON', 'NAMA PELANGGAN']) or col_clean == 'NAMA':
                rename_dict[col] = 'NAMA'
            elif any(kw in col_clean for kw in ['ALAMAT PEMOHON', 'ALAMAT PELANGGAN', 'NAMAPNJ']) or col_clean == 'ALAMAT':
                rename_dict[col] = 'ALAMAT'
            elif col_clean in ['TARIF', 'TARIP', 'TARIF BARU']:
                rename_dict[col] = 'TARIP'
            elif col_clean in ['DAYA', 'DAYA BARU']:
                rename_dict[col] = 'DAYA'
            elif col_clean in ['KDDK', 'KOKED']:
                rename_dict[col] = 'KOKED'
                
        df = df.rename(columns=rename_dict)
        
        # MENCEGAH ERROR KOLOM DUPLIKAT (Fix untuk Ambiguous Series Error)
        df = df.loc[:, ~df.columns.duplicated(keep='first')]
        
    return df, fname

def proses_list_file(files, tipe_data):
    list_df = []
    for f in files:
        df, fname = baca_ekstrak_tabel(f)
        if df.empty: continue
        
        if tipe_data == 'baru' or tipe_data == 'petugas':
            if 'KOKED' not in df.columns: df['KOKED'] = ''
            if 'ALAMAT' not in df.columns: df['ALAMAT'] = ''
            if 'TARIP' not in df.columns: df['TARIP'] = ''
            if 'DAYA' not in df.columns: df['DAYA'] = 0
            
            for col in ['IDPEL', 'NAMA']:
                if col not in df.columns: raise ValueError(f"Kolom wajib '{col}' hilang di file {fname}")
                
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
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di file PB: {fname}")
            if 'NAMA' not in df.columns: df['NAMA'] = ''
            if 'ALAMAT' not in df.columns: df['ALAMAT'] = ''
            
            # Jika PB Baru tidak punya kolom KOKED, tidak akan error (hanya dibuat kosong)
            if 'KOKED' not in df.columns: df['KOKED'] = ''
            
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['KOKED'] = df['KOKED'].astype(str).str.strip()
            # Hanya ambil data yang diperlukan
            list_df.append(df[['IDPEL', 'NAMA', 'ALAMAT', 'KOKED']])

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
st.title("⚡ Aplikasi Olah Koked")

tab1, tab2 = st.tabs(["1️⃣ Tahap 1: Persiapan Data Petugas", "2️⃣ Tahap 2: Rekap Pekerjaan & Mutasi KOKED"])

# ==========================================
# TAB 1: PERSIAPAN DATA (SEBELUM KE LAPANGAN)
# ==========================================
with tab1:
    st.header("Tahap 1: Memecah Data Untuk Petugas Lapangan")
    st.info("Upload data server, bagi menjadi per wilayah, dan pisahkan Pelanggan Baru (PB).")
    
    col1, col2 = st.columns(2)
    with col1:
        files_baru = st.file_uploader("[Tahap 1] Data Server Bulan INI", accept_multiple_files=True, key="t1_baru")
        files_lama = st.file_uploader("[Tahap 1] Data Bulan LALU (Pembanding PB)", accept_multiple_files=True, key="t1_lama")
    with col2:
        files_master = st.file_uploader("[Tahap 1] Data Master (Ganti ***)", accept_multiple_files=True, key="t1_master")
        files_suppb = st.file_uploader("[Tahap 1] Data PB Baru (PDF/Excel - Pelengkap)", accept_multiple_files=True, key="t1_suppb")

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

                    # Unmasking Data (Nama/Alamat/Koked)
                    df_baru = fix_masked_info(df_baru, df_master, df_sup_pb)

                    idpel_block = '524050450911'
                    df_baru = df_baru[(df_baru['DAYA'] <= 33000) & (df_baru['IDPEL'] != idpel_block)].copy()
                    
                    list_idpel_lama = set(df_lama['IDPEL'].tolist())
                    df_pb = df_baru[~df_baru['IDPEL'].isin(list_idpel_lama)].copy()
                    df_tetap = df_baru[df_baru['IDPEL'].isin(list_idpel_lama)].copy()

                    kolom_export = ['IDPEL', 'KOKED', 'NAMA', 'ALAMAT', 'TARIP', 'DAYA']
                    zip_buffer = io.BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
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
                    st.download_button("📥 Download Distribusi Petugas (.zip)", data=zip_buffer.getvalue(), file_name="Distribusi_Petugas.zip", mime="application/zip")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

# ==========================================
# TAB 2: REKAP & PERUBAHAN KOKED (SETELAH LAPANGAN)
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
                    with zipfile.ZipFile(zip_buffer2, "a", zipfile.ZIP_DEFLATED, False) as zip_file2:
                        
                        if not df_changed.empty:
                            txt_content = (df_changed['IDPEL'].astype(str) + "|" + df_changed['KOKED_BARU'].astype(str)).tolist()
                            txt_string = "\n".join(txt_content)
                            zip_file2.writestr("PERUBAHAN_KOKED.txt", txt_string)

                        zip_file2.writestr("DATA_GABUNGAN_FINAL.xlsx", to_excel_bytes(df_petugas[kolom_export]))

                    st.success(f"✅ Tahap 2 selesai! Terdeteksi {len(df_changed)} data yang KOKED-nya berubah.")
                    st.download_button(
                        label="📥 Download Hasil Akhir (.zip)", 
                        data=zip_buffer2.getvalue(), 
                        file_name="Hasil_Akhir_Mutasi.zip", 
                        mime="application/zip",
                        type="primary"
                    )

                except Exception as e:
                    st.error(f"❌ Error Tahap 2: {str(e)}")
