import streamlit as st
import pandas as pd
from dbfread import DBF
import os
import io
import zipfile
import tempfile

# --- PENGATURAN HALAMAN ---
st.set_page_config(page_title="Distribusi Data Multi-File", layout="wide")

# --- FUNGSI MENGONVERSI & MEMBERSIHKAN DAYA ---
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
    return s

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

    if not nama_map and not alamat_map:
        return df_baru

    is_masked_nama = df_baru['NAMA'].astype(str).str.contains(r'\*', na=False) | df_baru['NAMA'].isna() | (df_baru['NAMA'].astype(str).str.strip() == '')
    is_masked_alamat = df_baru['ALAMAT'].astype(str).str.contains(r'\*', na=False) | df_baru['ALAMAT'].isna() | (df_baru['ALAMAT'].astype(str).str.strip() == '')

    df_baru.loc[is_masked_nama, 'NAMA'] = df_baru.loc[is_masked_nama, 'IDPEL'].map(nama_map).fillna(df_baru.loc[is_masked_nama, 'NAMA'])
    df_baru.loc[is_masked_alamat, 'ALAMAT'] = df_baru.loc[is_masked_alamat, 'IDPEL'].map(alamat_map).fillna(df_baru.loc[is_masked_alamat, 'ALAMAT'])
    return df_baru

# --- FUNGSI PEMBACA FILE (STREAMLIT ADAPTER) ---
def baca_ekstrak_tabel(uploaded_file):
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    fname = uploaded_file.name
    
    if ext == '.xlsx':
        df = pd.read_excel(uploaded_file)
    else:
        # File DBF/PDF butuh file fisik, kita buat temporary file
        with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        try:
            if ext == '.dbf':
                df = pd.DataFrame(iter(DBF(tmp_path, char_decode_errors='ignore')))
            elif ext == '.pdf':
                try:
                    import pdfplumber
                    all_data, headers = [], None
                    with pdfplumber.open(tmp_path) as pdf:
                        for page in pdf.pages:
                            for table in page.extract_tables():
                                for row in table:
                                    if not any(row): continue
                                    if headers is None: headers = [str(cell).strip().upper() if cell else '' for cell in row]
                                    else:
                                        row_upper = [str(cell).strip().upper() if cell else '' for cell in row]
                                        if 'IDPEL' in row_upper or 'KOKED' in row_upper or 'KDDK' in row_upper: continue
                                        all_data.append(row)
                    df = pd.DataFrame([r[:len(headers)] for r in all_data], columns=headers)
                except ImportError:
                    st.error("Library PDF (pdfplumber) belum terinstall.")
                    return pd.DataFrame(), fname
            else:
                st.error(f"Format tidak didukung: {ext}")
                return pd.DataFrame(), fname
        finally:
            os.remove(tmp_path)

    df.columns = df.columns.str.strip().str.upper()
    df = df.rename(columns={'KDDK': 'KOKED', 'TARIF': 'TARIP', 'NAMAPNJ': 'ALAMAT'})
    return df, fname

# --- BACA MASING-MASING KATEGORI ---
def proses_list_file(files, tipe_data):
    list_df = []
    for f in files:
        df, fname = baca_ekstrak_tabel(f)
        if df.empty: continue
        
        if tipe_data == 'baru':
            if 'ALAMAT' not in df.columns: df['ALAMAT'] = ''
            for col in ['IDPEL', 'KOKED', 'TARIP', 'DAYA', 'NAMA']:
                if col not in df.columns: raise ValueError(f"Kolom wajib '{col}' hilang di file Baru: {fname}")
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['DAYA'] = df['DAYA'].apply(clean_daya)
            df['KOKED'] = df['KOKED'].astype(str).str.strip()
            list_df.append(df[['IDPEL', 'KOKED', 'TARIP', 'DAYA', 'NAMA', 'ALAMAT']])
            
        elif tipe_data == 'lama':
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di file Lama: {fname}")
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            list_df.append(df[['IDPEL']])
            
        elif tipe_data == 'master':
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di file Master: {fname}")
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
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            list_df.append(df[['IDPEL', 'NAMA', 'ALAMAT']])

    if not list_df: return pd.DataFrame()
    return pd.concat(list_df, ignore_index=True).drop_duplicates(subset=['IDPEL'], keep='first').reset_index(drop=True)

# --- HELPER: EXCEL KE BYTES ---
def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

# ==========================================
# ANTARMUKA STREAMLIT
# ==========================================
st.title("⚡ Aplikasi Pembanding & Distribusi Multi-File")
st.markdown("Unggah file data di bawah ini, proses, lalu unduh hasilnya dalam satu file `.zip`.")

col1, col2 = st.columns(2)
with col1:
    files_baru = st.file_uploader("1. Data Bulan INI (Sumber Utama)", accept_multiple_files=True)
    files_lama = st.file_uploader("2. Data Bulan LALU (Pembanding PB)", accept_multiple_files=True)
with col2:
    files_master = st.file_uploader("3. Data Master (Opsional - Ganti ***)", accept_multiple_files=True)
    files_suppb = st.file_uploader("4. Data PB Baru (Opsional - Pelengkap)", accept_multiple_files=True)

if st.button("Mulai Proses Data", type="primary"):
    if not files_baru or not files_lama:
        st.error("Silakan unggah Data Bulan Ini dan Data Bulan Lalu terlebih dahulu.")
    else:
        with st.spinner("Memproses data, merapikan kolom, dan menyusun file Excel..."):
            try:
                df_baru = proses_list_file(files_baru, 'baru')
                df_lama = proses_list_file(files_lama, 'lama')
                df_master = proses_list_file(files_master, 'master') if files_master else pd.DataFrame()
                df_sup_pb = proses_list_file(files_suppb, 'sup_pb') if files_suppb else pd.DataFrame()

                # Perbaiki Nama ***
                df_baru = fix_masked_info(df_baru, df_master, df_sup_pb)

                # Filter Daya & Blokir
                idpel_block = '524050450911'
                df_baru = df_baru[(df_baru['DAYA'] <= 33000) & (df_baru['IDPEL'] != idpel_block)].copy()

                # Pisahkan Lama & PB
                list_idpel_lama = set(df_lama['IDPEL'].tolist())
                df_pb = df_baru[~df_baru['IDPEL'].isin(list_idpel_lama)].copy()
                df_tetap = df_baru[df_baru['IDPEL'].isin(list_idpel_lama)].copy()

                # Urutkan berdasarkan KOKED
                for df in [df_pb, df_tetap]:
                    if not df.empty:
                        df['NO_URUT'] = df['KOKED'].str[7:10]
                        df['NO_URUT'] = pd.to_numeric(df['NO_URUT'], errors='coerce').fillna(0).astype(int)
                        df.sort_values(by=['KOKED', 'NO_URUT'], inplace=True)

                kolom_export = ['IDPEL', 'KOKED', 'NAMA', 'ALAMAT', 'TARIP', 'DAYA']
                
                # --- SIAPKAN ZIP FILE DI MEMORI ---
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    
                    # 1. PB SEMUA PETUGAS
                    if not df_pb.empty:
                        df_pb_clean = df_pb.drop_duplicates(subset=['IDPEL'])
                        zip_file.writestr("PB_SEMUA_PETUGAS.xlsx", to_excel_bytes(df_pb_clean[kolom_export]))

                    # 2. MASTER TERUPDATE
                    df_master_base = df_master[kolom_export].copy() if not df_master.empty else pd.DataFrame(columns=kolom_export)
                    df_master_combined = pd.concat([df_master_base, df_pb[kolom_export]], ignore_index=True)
                    df_master_combined = df_master_combined.drop_duplicates(subset=['IDPEL']).reset_index(drop=True)
                    
                    # Timpa nilai master yang lama dengan data baru yang masuk (jika relevan)
                    df_baru_map = df_baru.set_index('IDPEL')
                    in_baru = df_master_combined['IDPEL'].isin(df_baru_map.index)
                    for col in ['KOKED', 'TARIP', 'DAYA']:
                        df_master_combined.loc[in_baru, col] = df_master_combined.loc[in_baru, 'IDPEL'].map(df_baru_map[col].to_dict())
                    
                    zip_file.writestr("MASTER_UPDATED.xlsx", to_excel_bytes(df_master_combined[kolom_export]))

                    # 3. PECAH PER WILAYAH / PETUGAS
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

                st.success("✅ Pemrosesan selesai! Silakan unduh hasilnya.")
                st.download_button(
                    label="📥 Download Hasil (.zip)",
                    data=zip_buffer.getvalue(),
                    file_name="Hasil_Distribusi.zip",
                    mime="application/zip",
                    type="primary"
                )

            except Exception as e:
                st.error(f"❌ Terjadi kesalahan: {str(e)}")