import streamlit as st
import pandas as pd
import io
import re
import os
import tempfile
from dbfread import DBF

st.set_page_config(page_title="Aplikasi Pengolahan Data PB", layout="wide")

# --- FUNGSI PENGOLAHAN STRING ---
def clean_idpel(x):
    if pd.isna(x): return ""
    # Hapus .0 di akhir jika terbaca sebagai float, lalu hapus karakter non-digit
    s = str(x).split('.')[0]
    return re.sub(r'\D', '', s)

# --- FUNGSI MEMBACA FILE ---
def read_file(file_obj, filename):
    ext = filename.split('.')[-1].lower()
    try:
        if ext == 'csv':
            df = pd.read_csv(file_obj, dtype=str)
        elif ext in ['xls', 'xlsx']:
            df = pd.read_excel(file_obj, dtype=str)
        elif ext == 'dbf':
            # Menggunakan dbfread sebagai pengganti simpledbf
            with tempfile.NamedTemporaryFile(delete=False, suffix='.dbf') as tmp:
                tmp.write(file_obj.getvalue())
                tmp_path = tmp.name
            try:
                df = pd.DataFrame(iter(DBF(tmp_path, char_decode_errors='ignore')))
            finally:
                os.remove(tmp_path)
        else:
            st.error(f"Format file {ext} tidak didukung ({filename})")
            return pd.DataFrame()
            
        if not df.empty:
            df.columns = df.columns.str.strip().str.upper()
            # Mencegah error "The truth value of a Series is ambiguous" akibat duplikat kolom
            df = df.loc[:, ~df.columns.duplicated(keep='first')]
            
        return df
    except Exception as e:
        st.error(f"Gagal membaca {filename}: {e}")
        return pd.DataFrame()

# --- FUNGSI STANDARISASI KOLOM BERDASARKAN TIPE ---
def proses_list_file(file_list, tipe_data='master'):
    list_df = []
    for file in file_list:
        df = read_file(file, file.name)
        if df.empty: continue
        
        fname = file.name

        if tipe_data == 'baru' or tipe_data == 'petugas':
            # Opsional di Data Baru & Petugas
            if 'KOKED' not in df.columns: df['KOKED'] = ''
            if 'ALAMAT' not in df.columns: df['ALAMAT'] = ''
            if 'TARIP' not in df.columns: df['TARIP'] = ''
            if 'DAYA' not in df.columns: df['DAYA'] = 0
            
            for col in ['IDPEL', 'NAMA']:
                if col not in df.columns:
                    raise ValueError(f"Kolom wajib '{col}' hilang di file {fname}")
            
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['KOKED'] = df['KOKED'].astype(str).str.strip()
            list_df.append(df)
            
        elif tipe_data == 'master':
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di file Master: {fname}")
            if 'NAMA' not in df.columns: df['NAMA'] = ''
            if 'ALAMAT' not in df.columns: df['ALAMAT'] = ''
            if 'KOKED' not in df.columns: df['KOKED'] = ''
            
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['KOKED'] = df['KOKED'].astype(str).str.strip()
            list_df.append(df[['IDPEL', 'NAMA', 'ALAMAT', 'KOKED']])
            
        elif tipe_data == 'sup_pb':
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di file PB Baru: {fname}")
            if 'NAMA' not in df.columns: df['NAMA'] = ''
            if 'ALAMAT' not in df.columns: df['ALAMAT'] = ''
            
            # KOKED di PB Baru kalau ada dimasukkan, kalau tidak ada tidak error
            if 'KOKED' not in df.columns: df['KOKED'] = ''
            
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['KOKED'] = df['KOKED'].astype(str).str.strip()
            # Hanya ambil IDPEL, NAMA, ALAMAT, dan KOKED (jika kebetulan ada)
            list_df.append(df[['IDPEL', 'NAMA', 'ALAMAT', 'KOKED']])

    if not list_df: return pd.DataFrame()
    return pd.concat(list_df, ignore_index=True)

# --- FUNGSI UN-MASKING (NAMA/ALAMAT BINTANG & KOKED KOSONG) ---
def fix_masked_info(df_baru, df_master, df_sup_pb):
    """
    Menambal nama/alamat yang mengandung '*' di df_baru dengan data asli dari df_sup_pb atau df_master.
    KOKED utamanya dari Data Baru, tapi jika di Data Baru kosong, kita ambilkan dari PB Baru/Master jika ada.
    """
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

# --- FUNGSI MENGAMBIL KODE PETUGAS DARI KOKED ---
def get_petugas_code(koked):
    k = str(koked).strip().upper()
    if len(k) >= 5 and k[:3] == '524':
        return k[3:5]
    return "SEMUA_PETUGAS"

# --- MAPPING NAMA PETUGAS UNTUK TAB 2 (Data Petugas) ---
def map_nama_petugas(df_main, df_petugas):
    if 'KODE' not in df_petugas.columns or 'NAMA_PETUGAS' not in df_petugas.columns:
        df_main['NAMA_PETUGAS'] = 'TIDAK DIKETAHUI'
        return df_main
        
    df_petugas['KODE'] = df_petugas['KODE'].astype(str).str.strip().str.upper()
    petugas_dict = df_petugas.drop_duplicates(subset=['KODE']).set_index('KODE')['NAMA_PETUGAS'].to_dict()
    
    # Ambil 2 digit petugas dari KOKED
    df_main['KODE_PETUGAS'] = df_main['KOKED'].astype(str).apply(lambda x: x[3:5] if (len(x) >= 5 and x[:3] == '524') else "")
    df_main['NAMA_PETUGAS'] = df_main['KODE_PETUGAS'].map(petugas_dict).fillna('TIDAK DIKETAHUI')
    return df_main

# --- UI STREAMLIT ---
st.title("Aplikasi Pengolahan Data PB (KOKED dari Data Baru)")

# TAB NAVIGASI
tab1, tab2 = st.tabs(["1. Pisah Data PB per Petugas", "2. Filter Data Petugas (Status PB)"])

# ==========================================
# TAB 1: Pisah Data PB per Petugas
# ==========================================
with tab1:
    st.header("1. Pisah Data PB (Bulan INI) per Petugas")
    st.info("KOKED dan target Data PB diambil dari 'Data Server Bulan INI'. File Master dan Data PB Baru hanya digunakan untuk membuka NAMA dan ALAMAT yang di-bintang (*).")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        file_baru = st.file_uploader("Upload Data Server Bulan INI (DBF/Excel/CSV)", accept_multiple_files=True, key='baru')
    with col2:
        file_master = st.file_uploader("Upload Data Master Lama (DBF/Excel/CSV)", accept_multiple_files=True, key='master')
    with col3:
        file_sup_pb = st.file_uploader("Upload Data PB Baru (Sudah CSV/Excel)", accept_multiple_files=True, key='sup_pb')

    if st.button("Proses & Pisahkan Data (Tab 1)"):
        if not file_baru:
            st.warning("Data Server Bulan INI wajib diupload!")
        else:
            with st.spinner("Memproses data..."):
                try:
                    df_baru = proses_list_file(file_baru, tipe_data='baru')
                    df_master = proses_list_file(file_master, tipe_data='master') if file_master else pd.DataFrame()
                    df_sup_pb = proses_list_file(file_sup_pb, tipe_data='sup_pb') if file_sup_pb else pd.DataFrame()

                    # Un-masking Nama, Alamat, (dan Koked jika di Data Baru kosong)
                    df_baru = fix_masked_info(df_baru, df_master, df_sup_pb)

                    # Buat kolom PETUGAS_MAP dari KOKED Data Baru
                    df_baru['PETUGAS_MAP'] = df_baru['KOKED'].apply(get_petugas_code)
                    
                    # Pisahkan DataFrame berdasarkan kode petugas
                    grouped = df_baru.groupby('PETUGAS_MAP')
                    
                    st.success("Berhasil diproses! Silakan download hasilnya di bawah:")
                    
                    # Tampilkan tombol download
                    cols = st.columns(4)
                    idx = 0
                    for kode, group_df in grouped:
                        group_df_clean = group_df.drop(columns=['PETUGAS_MAP'])
                        
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            group_df_clean.to_excel(writer, index=False, sheet_name='Data')
                        excel_data = output.getvalue()
                        
                        nama_file = f"PB_SEMUA_PETUGAS.xlsx" if kode == "SEMUA_PETUGAS" else f"PB_{kode}.xlsx"
                        
                        cols[idx % 4].download_button(
                            label=f"Download {nama_file} ({len(group_df_clean)} baris)",
                            data=excel_data,
                            file_name=nama_file,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"dl_tab1_{kode}"
                        )
                        idx += 1
                        
                except Exception as e:
                    st.error(f"Terjadi kesalahan: {e}")

# ==========================================
# TAB 2: Filter Data Petugas
# ==========================================
with tab2:
    st.header("2. Filter Data Laporan Petugas (PB Belum/Sudah Proses)")
    st.info("Bandingkan Data Laporan Petugas (Bulan INI) dengan Data Master Lama (Bulan LALU) untuk mencari PB Baru.")
    
    colA, colB, colC = st.columns(3)
    with colA:
        file_petugas = st.file_uploader("Upload Laporan Petugas Bulan INI (Excel/CSV)", accept_multiple_files=True, key='petugas_tab2')
    with colB:
        file_master_2 = st.file_uploader("Upload Data Master Bulan LALU (Excel/DBF/CSV)", accept_multiple_files=True, key='master_tab2')
    with colC:
        file_mapping = st.file_uploader("Upload Mapping Petugas (Excel) - Opsional", accept_multiple_files=False, key='mapping_tab2', help="Harus berisi kolom KODE dan NAMA_PETUGAS")

    if st.button("Proses Data Petugas (Tab 2)"):
        if not file_petugas or not file_master_2:
            st.warning("Laporan Petugas Bulan INI dan Data Master Bulan LALU wajib diupload!")
        else:
            with st.spinner("Memproses data..."):
                try:
                    df_petugas = proses_list_file(file_petugas, tipe_data='petugas')
                    df_master_2 = proses_list_file(file_master_2, tipe_data='master')
                    
                    df_petugas['IDPEL'] = df_petugas['IDPEL'].astype(str)
                    df_master_2['IDPEL'] = df_master_2['IDPEL'].astype(str)
                    
                    master_idpel = set(df_master_2['IDPEL'].unique())
                    
                    # Status PB Baru / Lama
                    df_petugas['STATUS_PB'] = df_petugas['IDPEL'].apply(lambda x: 'PB LAMA' if x in master_idpel else 'PB BARU')
                    
                    # Mapping Nama Petugas jika ada
                    if file_mapping is not None:
                        df_map = read_file(file_mapping, file_mapping.name)
                        df_map.columns = df_map.columns.str.strip().str.upper()
                        df_petugas = map_nama_petugas(df_petugas, df_map)
                    
                    st.success("Data berhasil diproses!")
                    
                    # Preview Data
                    st.subheader("Preview Data Hasil Filter:")
                    st.dataframe(df_petugas.head(20))
                    
                    # Hitung Rekap
                    total_pb = len(df_petugas)
                    pb_lama = len(df_petugas[df_petugas['STATUS_PB'] == 'PB LAMA'])
                    pb_baru = len(df_petugas[df_petugas['STATUS_PB'] == 'PB BARU'])
                    
                    st.write(f"**Total Data:** {total_pb} | **PB Lama:** {pb_lama} | **PB Baru:** {pb_baru}")
                    
                    # Buat file Excel dengan 3 Sheet
                    output2 = io.BytesIO()
                    with pd.ExcelWriter(output2, engine='xlsxwriter') as writer:
                        df_petugas.to_excel(writer, index=False, sheet_name='Semua Data')
                        df_petugas[df_petugas['STATUS_PB'] == 'PB BARU'].to_excel(writer, index=False, sheet_name='Hanya PB Baru')
                        df_petugas[df_petugas['STATUS_PB'] == 'PB LAMA'].to_excel(writer, index=False, sheet_name='Hanya PB Lama')
                    
                    excel_data2 = output2.getvalue()
                    
                    st.download_button(
                        label="Download Hasil Evaluasi Petugas (.xlsx)",
                        data=excel_data2,
                        file_name="Evaluasi_Laporan_Petugas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_tab2"
                    )
                    
                except Exception as e:
                    st.error(f"Terjadi kesalahan: {e}")
