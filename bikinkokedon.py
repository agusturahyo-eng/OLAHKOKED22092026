import streamlit as st
import pandas as pd
import numpy as np
from dbfread import DBF
import os
import sys
import io
import csv
import shutil
import zipfile
import tempfile
import bisect
import re
import gc
import openpyxl

# --- PENGATURAN HALAMAN (WAJIB PALING ATAS DI STREAMLIT) ---
st.set_page_config(page_title="Aplikasi Olah Data & Cetak TUL", layout="wide")

# 1. Inisialisasi status login
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

# 2. Jika BELUM login, tampilkan form
if not st.session_state.authenticated:
    st.title("Aplikasi Terkunci 🔒")
    password = st.text_input("Masukkan Password:", type="password")
    
    if st.button("Masuk"):
        if password == "Tlg@1234":  # Ganti dengan password Anda
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Password salah!")
            
    # HENTIKAN aplikasi di sini jika belum login. 
    st.stop()

# Library PDF
try:
    import pdfplumber
    PDF_SUPPORT = True
except ImportError:
    PDF_SUPPORT = False

try:
    import fitz  # PyMuPDF
    PYMUPDF_SUPPORT = True
except ImportError:
    PYMUPDF_SUPPORT = False

# --- KONFIGURASI & FUNGSI GLOBAL SUPABASE ---
SUPABASE_URL = "https://wnzedfyiublmzmjkzsrq.supabase.co" 
SUPABASE_KEY = "sb_publishable_g1WlECVLkTdNBMwL2QxWlA_gOu6l4VR"
SUPABASE_TABLE = "dataplg3"

@st.cache_resource
def init_koneksi():
    from supabase import create_client
    return create_client(SUPABASE_URL, SUPABASE_KEY)

@st.cache_data(ttl=1800, max_entries=2, show_spinner=False)
def fetch_master_supabase(columns="*"):
    """Mengambil seluruh data master dari Supabase secara bertahap."""
    client = init_koneksi()
    all_rows = []
    batch_size = 1000
    start = 0
    
    while True:
        respon = client.table(SUPABASE_TABLE).select(columns).range(start, start + batch_size - 1).execute()
        data = respon.data
        if not data:
            break
        all_rows.extend(data)
        if len(data) < batch_size:
            break
        start += batch_size
        
    return pd.DataFrame(all_rows)

def tampilkan_info_header_master():
    st.info(
        "📋 **Keterangan Header Kolom Master (`dataplg3` di Supabase):**\n\n"
        "`IDPEL` | `KDDK` *(otomatis dipetakan ke KOKED)* | `NAMA` | `ALAMAT` | `TARIF` *(otomatis dipetakan ke TARIP)* | "
        "`DAYA` | `NOMOR GARDU` | `NOTIANG` | `MEREKKWH` | `NOMORKWH` | `KOORDINAT X` | `KOORDINAT Y` | `NOIDENTITAS` | `NO HP`"
    )

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
    return s

def fix_masked_info(df_baru, df_master, df_sup_pb):
    if df_baru.empty: return df_baru
    df_baru = df_baru.reset_index(drop=True)
    nama_map, alamat_map = {}, {}

    if not df_sup_pb.empty:
        sup_clean = df_sup_pb.drop_duplicates(subset=['IDPEL'], keep='first').copy()
        sup_clean = sup_clean[sup_clean['NAMA'].astype(str).str.strip().ne('') & ~sup_clean['NAMA'].astype(str).str.contains(r'\*', na=False)]
        nama_map.update(sup_clean.set_index('IDPEL')['NAMA'].to_dict())
        alamat_map.update(sup_clean.set_index('IDPEL')['ALAMAT'].to_dict())

    if not df_master.empty:
        master_clean = df_master.drop_duplicates(subset=['IDPEL'], keep='first').copy()
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
    df = df.rename(columns={'KDDK': 'KOKED', 'TARIF': 'TARIP', 'NAMAPNJ': 'ALAMAT'})
    df = df.loc[:, ~df.columns.duplicated(keep='first')]
    return df, fname

def siapkan_df_master_standar(df, tipe_data='master'):
    if df.empty: return pd.DataFrame()
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip().str.upper()
    df = df.rename(columns={'KDDK': 'KOKED', 'TARIF': 'TARIP', 'NAMAPNJ': 'ALAMAT'})
    df = df.loc[:, ~df.columns.duplicated(keep='first')]

    if 'IDPEL' not in df.columns: raise ValueError("Kolom 'IDPEL' tidak ditemukan di Data Master Supabase.")

    if tipe_data == 'master':
        for col in ['NAMA', 'ALAMAT', 'KOKED', 'TARIP']:
            if col not in df.columns: df[col] = ''
        if 'DAYA' not in df.columns: df['DAYA'] = 0
        df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
        df['DAYA'] = df['DAYA'].apply(clean_daya)
        df['KOKED'] = df['KOKED'].astype(str).str.strip()
        return df[['IDPEL', 'NAMA', 'ALAMAT', 'KOKED', 'TARIP', 'DAYA']].drop_duplicates(subset=['IDPEL'], keep='first').reset_index(drop=True)
    elif tipe_data == 'lama':
        if 'KOKED' not in df.columns: df['KOKED'] = ''
        df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
        df['KOKED'] = df['KOKED'].astype(str).str.strip()
        return df[['IDPEL', 'KOKED']].drop_duplicates(subset=['IDPEL'], keep='first').reset_index(drop=True)
    return df

def proses_list_file(files, tipe_data):
    list_df = []
    for f in files:
        df, fname = baca_ekstrak_tabel(f)
        if df.empty: continue
        
        if tipe_data == 'baru':
            if 'ALAMAT' not in df.columns: df['ALAMAT'] = ''
            for col in ['IDPEL', 'KOKED', 'TARIP', 'DAYA', 'NAMA']:
                if col not in df.columns: raise ValueError(f"Kolom wajib '{col}' tidak ditemukan di Data Baru: {fname}")
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['DAYA'] = df['DAYA'].apply(clean_daya)
            df['KOKED'] = df['KOKED'].astype(str).str.strip()
            list_df.append(df[['IDPEL', 'KOKED', 'TARIP', 'DAYA', 'NAMA', 'ALAMAT']])
            
        elif tipe_data == 'lama':
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di Data Lama: {fname}")
            if 'KOKED' not in df.columns: df['KOKED'] = ''
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['KOKED'] = df['KOKED'].astype(str).str.strip()
            list_df.append(df[['IDPEL', 'KOKED']])
            
        elif tipe_data == 'master':
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di Master: {fname}")
            for col in ['NAMA', 'ALAMAT', 'KOKED', 'TARIP']:
                if col not in df.columns: df[col] = ''
            if 'DAYA' not in df.columns: df['DAYA'] = 0
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['DAYA'] = df['DAYA'].apply(clean_daya)
            df['KOKED'] = df['KOKED'].astype(str).str.strip()
            list_df.append(df[['IDPEL', 'NAMA', 'ALAMAT', 'KOKED', 'TARIP', 'DAYA']])
            
        elif tipe_data == 'sup_pb':
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di Data PB Baru: {fname}")
            if 'NAMA' not in df.columns: df['NAMA'] = ''
            if 'ALAMAT' not in df.columns: df['ALAMAT'] = ''
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            list_df.append(df[['IDPEL', 'NAMA', 'ALAMAT']])

        elif tipe_data == 'petugas': 
            for col in ['ALAMAT', 'NAMA', 'TARIP', 'KOKED']:
                if col not in df.columns: df[col] = ''
            if 'DAYA' not in df.columns: df['DAYA'] = 0
            if 'IDPEL' not in df.columns: raise ValueError(f"Kolom 'IDPEL' hilang di file {fname}")
            df['IDPEL'] = df['IDPEL'].apply(clean_idpel)
            df['DAYA'] = df['DAYA'].apply(clean_daya)
            df['KOKED'] = df['KOKED'].astype(str).str.strip()
            list_df.append(df[['IDPEL', 'KOKED', 'TARIP', 'DAYA', 'NAMA', 'ALAMAT']])

    if not list_df: return pd.DataFrame()
    return pd.concat(list_df, ignore_index=True).drop_duplicates(subset=['IDPEL'], keep='first').reset_index(drop=True)

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()


# --- FUNGSI TAB 3 & 11 (PDFPLUMBER HYBRID) ---
def filter_watermark_obj(obj):
    if obj.get("object_type") == "char":
        if not obj.get("upright", True): return False
        if obj.get("size", 0) > 14: return False
    return True

def bersihkan_halaman_pdf(page):
    try: return page.filter(filter_watermark_obj)
    except: return page

def get_pdf_tables_tipe1(page):
    clean_p = bersihkan_halaman_pdf(page)
    table_settings = {"vertical_strategy": "lines", "horizontal_strategy": "text"}
    tables = clean_p.extract_tables(table_settings)
    if not tables or len(tables) == 0 or all(len(t) < 2 for t in tables):
        tables = clean_p.extract_tables()
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

    headers = [str(h).strip() if str(h).strip() != "" else f"KOLOM_{i+1}" for i, h in enumerate(merged_rows[header_index])]
    return pd.DataFrame(merged_rows[header_index+1:], columns=headers)

# --- FUNGSI NORMALISASI TAB 3 ---
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
    if df is None or df.empty: return pd.DataFrame(columns=['IDPEL', 'NAMA', 'ALAMAT', 'TARIF', 'DAYA', 'GARDU', 'TIANG', 'KOKED'])
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


# ================= FUNGSI UNTUK TAB 12 (CETAK TUL) =================
def tentukan_petugas_tul(idpel, tarif_daya, kddk):
    daya = 0
    match_daya = re.search(r'/(\d+)', str(tarif_daya))
    if match_daya: daya = int(match_daya.group(1))
    if daya > 33000: return "PLN"
    
    idpel_khusus = {
        "524051069054": "c28", "524051263717": "c36", "524051265123": "c36", 
        "524051104194": "c04", "524051000615": "c08", "524050867033": "c08"
    }
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

def baca_data_dari_iconprn_tul(teks_mentah):
    hasil = []
    blok_halaman = re.split(r'PEMBERITAHUAN PELAKSANAAN PEMUTUSAN', teks_mentah)
    for blok in blok_halaman:
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

        # AMAN: Nomor meter mendukung alphanumeric
        meter = re.search(r'Nomor Meter\s*:\s*([A-Z0-9]+)', blok, re.IGNORECASE)
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

        data['petugas'] = tentukan_petugas_tul(data['IDPEL'], data['Tarif/Daya'], data['KDDK'])
        hasil.append(data)
    return hasil

STANDAR_KOLOM_TUL = {
    "IDPEL": ["idpel", "id pelanggan", "id_pelanggan", "no pelanggan"],
    "Nomor TUL": ["nomor tul", "no tul", "no. tul", "nomor_tul"],
    "Nama": ["nama", "nama pelanggan", "nama_pelanggan"],
    "KDDK": ["kddk", "kode kedudukan", "kedudukan"],
    "Gardu/Tiang": ["gardu/tiang", "gardu", "nama gardu/tiang", "gardutiang", "tiang"],
    "Loket": ["loket", "kode loket"],
    "Alamat": ["alamat", "alamat pelanggan"],
    "Nomor Meter": ["nomor meter", "no meter", "nomor_meter", "nomormeter"],
    "Tarif/Daya": ["tarif/daya", "tarip / daya", "tarifdaya", "tarif", "daya"],
    "Kelompok": ["kelompok", "klp"],
    "Bulan Rekening": ["bulan rekening", "bulan_rekening", "rekening", "blth"],
    "Bulan Keterlambatan": ["bulan keterlambatan", "bulan_keterlambatan", "keterlambatan"],
    "Jumlah Rekening": ["jumlah rekening", "jumlah_rekening", "jumlah_rekening_", "rp rekening", "tagihan"],
    "Jumlah Denda": ["jumlah denda", "jumlah_denda", "jumlah_denda_", "denda", "bk", "biaya keterlambatan"],
    "Jumlah Tunggakan": ["jumlah tunggakan", "jumlah_tunggakan", "jumlah_tunggakan_", "total", "total tunggakan"],
    "petugas": ["petugas", "kode petugas", "nama petugas", "cater"]
}

def deteksi_kolom_otomatis(df_cols):
    mapping = {}
    cols_lower = {c.strip().lower(): c for c in df_cols}
    for target, kandidat_list in STANDAR_KOLOM_TUL.items():
        found = None
        for k in kandidat_list:
            if k in cols_lower:
                found = cols_lower[k]
                break
        if not found:
            for c_low, c_orig in cols_lower.items():
                if any(k in c_low for k in kandidat_list):
                    found = c_orig
                    break
        mapping[target] = found
    return mapping

def get_printer_init_code(printer_choice):
    ESC = "\x1b"
    if "LQ-2190" in printer_choice: return f"{ESC}g{ESC}0"
    elif "17 CPI" in printer_choice: return f"{ESC}P\x0f{ESC}0"
    else: return f"{ESC}M\x12{ESC}0"

def format_rupiah(val):
    try:
        if pd.isna(val) or str(val).strip() == "": return "0"
        bersih = str(val).replace(".", "").replace(",", "").strip()
        return f"{int(float(bersih)):,}".replace(",", ".")
    except Exception:
        return str(val)

def shift_line(text, offset_x):
    if not text: return ""
    if offset_x > 0: return (" " * offset_x) + text
    elif offset_x < 0:
        leading = len(text) - len(text.lstrip(" "))
        return text[min(abs(offset_x), leading):]
    return text

def ambil_nilai(row, col_map, key, default=""):
    if col_map is None: return str(row.get(key, default)).strip()
    nama_kolom = col_map.get(key)
    if nama_kolom and nama_kolom in row and pd.notna(row[nama_kolom]):
        return str(row[nama_kolom]).strip()
    return default

def buat_isian_blangko(row, col_map, init_code, kota, tgl, jab, manager, is_cetak_kota, geser_tgl, offset_x=0, offset_y=0, page_lines=44):
    ESC = "\x1b"
    BOLD_ON = f"{ESC}E"  
    BOLD_OFF = f"{ESC}F" 

    no_tul = ambil_nilai(row, col_map, "Nomor TUL")
    nama = ambil_nilai(row, col_map, "Nama")[:33]
    idpel = ambil_nilai(row, col_map, "IDPEL")
    kddk = ambil_nilai(row, col_map, "KDDK")
    alamat = ambil_nilai(row, col_map, "Alamat")[:50]
    no_meter = ambil_nilai(row, col_map, "Nomor Meter")
    gardu = ambil_nilai(row, col_map, "Gardu/Tiang")
    loket = ambil_nilai(row, col_map, "Loket")
    tarif = ambil_nilai(row, col_map, "Tarif/Daya")
    kelompok = ambil_nilai(row, col_map, "Kelompok", "0")
    bln_rek = ambil_nilai(row, col_map, "Bulan Rekening")
    bln_lambat = ambil_nilai(row, col_map, "Bulan Keterlambatan")

    rp_rek = format_rupiah(ambil_nilai(row, col_map, "Jumlah Rekening", "0")).rjust(15)
    rp_denda = format_rupiah(ambil_nilai(row, col_map, "Jumlah Denda", "0")).rjust(15)
    rp_total = format_rupiah(ambil_nilai(row, col_map, "Jumlah Tunggakan", "0")).rjust(15)

    raw_lines = [""] * page_lines

    def set_line(idx, content):
        target = idx + offset_y
        if 0 <= target < page_lines:
            raw_lines[target] = shift_line(content, offset_x)

    set_line(1,  f"{'':<73}{no_tul}")
    set_line(7,  f"{'':<21}{nama}")
    set_line(8,  f"{'':<21}{idpel:<58}{kddk}")         
    set_line(9,  f"{'':<21}{alamat}")
    set_line(10, f"{'':<21}{no_meter}")
    set_line(11, f"{'':<21}{gardu:<53}{loket}")        
    set_line(12, f"{'':<21}{tarif:<53}{kelompok}")     
    
    set_line(14, f"{'':<12}{bln_rek:<62}{rp_rek}")
    set_line(15, f"{'':<39}{bln_lambat:<35}{rp_denda}")
    set_line(17, f"{'':<74}{rp_total}")
    
    str_kota = f"{kota}, " if is_cetak_kota else ""
    str_jabatan = jab if is_cetak_kota else ""
    
    posisi_tgl = max(0, 64 + geser_tgl)
    set_line(30, f"{'':<{posisi_tgl}}{str_kota}{tgl}")
    set_line(31, f"{'':<69}{str_jabatan}")
    set_line(36, f"{'':<64}{BOLD_ON}{manager}{BOLD_OFF}")

    raw_lines[0] = init_code + raw_lines[0]
    return "\r\n".join(raw_lines) + "\r\n"

def buat_blangko_dan_isi(row, col_map, init_code, kota, tgl, jab, manager, page_lines=44):
    ESC = "\x1b"
    BOLD_ON = f"{ESC}E"
    BOLD_OFF = f"{ESC}F"

    no_tul = ambil_nilai(row, col_map, "Nomor TUL")
    nama = ambil_nilai(row, col_map, "Nama")[:33]
    idpel = ambil_nilai(row, col_map, "IDPEL")
    kddk = ambil_nilai(row, col_map, "KDDK")
    alamat = ambil_nilai(row, col_map, "Alamat")[:50]
    no_meter = ambil_nilai(row, col_map, "Nomor Meter")
    gardu = ambil_nilai(row, col_map, "Gardu/Tiang")
    loket = ambil_nilai(row, col_map, "Loket")
    tarif = ambil_nilai(row, col_map, "Tarif/Daya")
    kelompok = ambil_nilai(row, col_map, "Kelompok", "0")
    bln_rek = ambil_nilai(row, col_map, "Bulan Rekening")
    bln_lambat = ambil_nilai(row, col_map, "Bulan Keterlambatan")

    rp_rek = format_rupiah(ambil_nilai(row, col_map, "Jumlah Rekening", "0")).rjust(15)
    rp_denda = format_rupiah(ambil_nilai(row, col_map, "Jumlah Denda", "0")).rjust(15)
    rp_total = format_rupiah(ambil_nilai(row, col_map, "Jumlah Tunggakan", "0")).rjust(15)

    lines = [
        f"{init_code}PT. PLN (PERSERO) UID JAWA TENGAH DAN DIY",
        f"UP3 KLATEN                                                    NO. TUL :  {no_tul}",
        "ULP TULUNG",
        "",
        "             PEMBERITAHUAN PELAKSANAAN PEMUTUSAN SEMENTARA SAMBUNGAN TENAGA LISTRIK             ",
        "             ======================================================================             ",
        "Kepada Yth. ",
        f"Nama              :  {nama}",
        f"ID. Pelanggan     :  {idpel:<42}Kode Kedudukan : {kddk}",
        f"Alamat            :  {alamat}",
        f"Nomor Meter       :  {no_meter}",
        f"Nama Gardu/Tiang  :  {gardu:<44}Loket    : {loket}",
        f"Tarip / Daya      :  {tarif:<42}Kelompok : {kelompok}",
        "",
        f"Rekening :  {bln_rek:<55}Rp. : {rp_rek}",
        f"Jumlah Biaya Keterlambatan s.d bulan : {bln_lambat:<28}Rp. : {rp_denda}",
        "                                                                       ---------------",
        f"Jumlah Tunggakan (belum termasuk biaya Administrasi)               Rp. : {rp_total}",
        "",
        "   Dengan ini diberitahukan dengan hormat bahwa pada hari ini aliran listrik di rumah/alamat seperti tersebut diatas   ",
        "terpaksa diputus untuk sementara karena rekening listrik belum dilunasi pada waktu yang telah ditetapkan.",
        "Penyambungan kembali akan dilakukan pada setiap hari jam kerja apabila rekening serta biaya keterlambatan dilunasi",
        "di tempat penerimaan pembayaran rekening listrik, kantor pos, atau bank yang ditunjuk oleh PLN.                       ",
        "   Apabila dalam jangka waktu 60 hari terhitung sejak dilakukan pemutusan sementara tunggakan belum dilunasi,         ",
        "maka instalasi milik PLN akan dibongkar, dan penyambungan kembali dapat dilaksanakan setelah Saudara menyelesaikan    ",
        "Biaya Penyambungan yang diperlakukan sebagai sambungan baru serta tetap diwajibkan membayar tagihan listrik           ",
        "yang belum dilunasi beserta dendanya.                     ",
        "",
        "UNTUK MENGHINDARI RESIKO, MOHON TIDAK TITIP PEMBAYARAN REKENING KEPADA PETUGAS",
        "",
        f"                                                                {kota}, {tgl}",
        f"                                                                     {jab}",
        "",
        "|------------------------------------------------------|",
        "|       PADA WAKTU MELAKUKAN PEMBAYARAN DIMOHON        |",
        "|         MENUNJUKKAN SURAT PEMBERITAHUAN INI          |",
        f"|------------------------------------------------------|        {BOLD_ON}{manager}{BOLD_OFF}",
        "                              TGL  STAND PUTUS  PELANGGAN        ",
        "",
        "A5 TUL VI-01/PETUGAS PEMUTUS...... ... ... ... ...........                   ",
        "ABAIKAN PEMBERITAHUAN INI JIKA SUDAH MEMBAYAR TAGIHAN"
    ]
    while len(lines) < page_lines:
        lines.append("")
    return "\r\n".join(lines[:page_lines]) + "\r\n"

def buat_blangko_kosong(init_code, jml_lembar=50, page_lines=44):
    lines = [
        f"{init_code}PT. PLN (PERSERO) UID JAWA TENGAH DAN DIY",
        "UP3 KLATEN                                                    NO. TUL :  ",
        "ULP TULUNG", "",
        "             PEMBERITAHUAN PELAKSANAAN PEMUTUSAN SEMENTARA SAMBUNGAN TENAGA LISTRIK             ",
        "             ======================================================================             ",
        "Kepada Yth. ",
        "Nama              :                                                  ",
        "ID. Pelanggan     :  \t\t                              Kode Kedudukan :     ",
        "Alamat            :                                                  ",
        "Nomor Meter       :                                                                ",
        "Nama Gardu/Tiang  :  \t                                      Loket    :              ",
        "Tarip / Daya      :  \t                                      Kelompok :               ", "",
        "Rekening :        \t\t\t\t                   Rp. :  ",
        "Jumlah Biaya Keterlambatan s.d bulan :                             Rp. :  ",
        "                                                                        ---------------",
        "Jumlah Tunggakan (belum termasuk biaya Administrasi)               Rp. :  ", "",
        "   Dengan ini diberitahukan dengan hormat bahwa pada hari ini aliran listrik di rumah/alamat seperti tersebut diatas   ",
        "terpaksa diputus untuk sementara karena rekening listrik belum dilunasi pada waktu yang telah ditetapkan.",
        "Penyambungan kembali akan dilakukan pada setiap hari jam kerja apabila rekening serta biaya keterlambatan dilunasi",
        "di tempat penerimaan pembayaran rekening listrik, kantor pos, atau bank yang ditunjuk oleh PLN.                       ",
        "   Apabila dalam jangka waktu 60 hari terhitung sejak dilakukan pemutusan sementara tunggakan belum dilunasi,         ",
        "maka instalasi milik PLN akan dibongkar, dan penyambungan kembali dapat dilaksanakan setelah Saudara menyelesaikan    ",
        "Biaya Penyambungan yang diperlakukan sebagai sambungan baru serta tetap diwajibkan membayar tagihan listrik           ",
        "yang belum dilunasi beserta dendanya.                     ", "",
        "UNTUK MENGHINDARI RESIKO, MOHON TIDAK TITIP PEMBAYARAN REKENING KEPADA PETUGAS", "",
        "                                                                           ",
        "                                                                               ", "",
        "|------------------------------------------------------|",
        "|       PADA WAKTU MELAKUKAN PEMBAYARAN DIMOHON        |",
        "|         MENUNJUKKAN SURAT PEMBERITAHUAN INI          |",
        "|------------------------------------------------------|",
        "                              TGL  STAND PUTUS  PELANGGAN        ", "",
        "A5 TUL VI-01/PETUGAS PEMUTUS...... ... ... ... ...........                   ",
        "ABAIKAN PEMBERITAHUAN INI JIKA SUDAH MEMBAYAR TAGIHAN"
    ]
    while len(lines) < page_lines:
        lines.append("")
    return ("\r\n".join(lines[:page_lines]) + "\r\n") * jml_lembar

# ==========================================
# ANTARMUKA STREAMLIT
# ==========================================
st.title("⚡ Aplikasi Olah Data & Cetak TUL")

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab11, tab12 = st.tabs([
    "1️⃣ 1: Bikin Data Untuk Koked", 
    "2️⃣ 2: Hasil Koked",
    "3️⃣ 3: Eksport Pdf PB",
    "4️⃣ 4: Eksport PDF",
    "5️⃣ 5: eksport ICONPRN",
    "6️⃣ 6: Olah Data",
    "7️⃣ 7: Info & Lokasi",
    "8️⃣ 8: Update Data (Versi 2)",
    "9️⃣ 9: Isi Petugas",
    "🔟 10: Split Data",
    "1️⃣1️⃣ 11: Ekstrak Laporan Asli",
    "🖨️ 12: Cetak TUL VI-01"
])

# ==========================================
# TAB 1: PERSIAPAN DATA
# ==========================================
with tab1:
    st.header("Tahap 1: Memecah Data Untuk Petugas Lapangan")
    st.markdown("Digunakan untuk membuat data perpetugas buat koked, data pelanggan baru, dan data master")
    tampilkan_info_header_master()
    col1, col2 = st.columns(2)
    with col1:
        files_baru = st.file_uploader("[Tahap 1] Data Server Bulan INI", accept_multiple_files=True, key="t1_baru")
        files_lama = st.file_uploader("[Tahap 1] Data Bulan LALU", accept_multiple_files=True, key="t1_lama")
    with col2:
        sumber_master_t1 = st.radio(
            "Sumber Data Master (Ganti ***):",
            ["Upload File (.xlsx / .dbf)", "Ambil dari Supabase (dataplg3)"],
            horizontal=True,
            key="sumber_master_t1"
        )
        if sumber_master_t1 == "Upload File (.xlsx / .dbf)":
            files_master = st.file_uploader("[Tahap 1] Data Master (Ganti ***)", accept_multiple_files=True, key="t1_master")
        else:
            files_master = None
            st.success("✅ Data Master akan diambil otomatis dari Supabase (`IDPEL, KDDK, NAMA, ALAMAT, TARIF, DAYA`).")
        files_suppb = st.file_uploader("[Tahap 1] Data PB Baru", accept_multiple_files=True, key="t1_suppb")

    if st.button("Proses Tahap 1", type="primary"):
        if not files_baru or not files_lama: st.error("Silakan unggah Data Bulan Ini dan Data Bulan Lalu.")
        else:
            with st.spinner("Menyiapkan data untuk petugas..."):
                try:
                    df_baru = proses_list_file(files_baru, 'baru')
                    df_lama = proses_list_file(files_lama, 'lama')
                    if sumber_master_t1 == "Ambil dari Supabase (dataplg3)":
                        df_sup_raw = fetch_master_supabase("IDPEL,KDDK,NAMA,ALAMAT,TARIF,DAYA")
                        df_master = siapkan_df_master_standar(df_sup_raw, 'master')
                    else:
                        df_master = proses_list_file(files_master, 'master') if files_master else pd.DataFrame()
                    df_sup_pb = proses_list_file(files_suppb, 'sup_pb') if files_suppb else pd.DataFrame()

                    if df_baru.empty: raise ValueError("Data Bulan Ini kosong atau gagal dibaca.")
                    if df_lama.empty: raise ValueError("Data Bulan Lalu kosong atau gagal dibaca.")

                    df_baru = fix_masked_info(df_baru, df_master, df_sup_pb)
                    idpel_block = '524050450911'
                    df_baru = df_baru[(df_baru['DAYA'] <= 33000) & (df_baru['IDPEL'] != idpel_block)].copy().reset_index(drop=True)

                    list_idpel_lama = set(df_lama['IDPEL'].tolist())
                    df_pb = df_baru[~df_baru['IDPEL'].isin(list_idpel_lama)].copy().reset_index(drop=True)
                    df_tetap = df_baru[df_baru['IDPEL'].isin(list_idpel_lama)].copy().reset_index(drop=True)

                    for df in [df_pb, df_tetap]:
                        if not df.empty:
                            df['NO_URUT'] = df['KOKED'].astype(str).str[7:10]
                            df['NO_URUT'] = pd.to_numeric(df['NO_URUT'], errors='coerce').fillna(0).astype(int)
                            df.sort_values(by=['KOKED', 'NO_URUT'], inplace=True)

                    kolom_export = ['IDPEL', 'KOKED', 'NAMA', 'ALAMAT', 'TARIP', 'DAYA']
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        if not df_pb.empty:
                            df_pb_clean = df_pb.drop_duplicates(subset=['IDPEL'], keep='first')
                            pb_bytes = to_excel_bytes(df_pb_clean[kolom_export])
                            zip_file.writestr("PB_SEMUA_PETUGAS.xlsx", pb_bytes)

                        df_master_base = df_master[kolom_export].copy() if not df_master.empty else pd.DataFrame(columns=kolom_export)
                        df_pb_master_format = df_pb[kolom_export].copy()
                        df_master_combined = pd.concat([df_master_base, df_pb_master_format], ignore_index=True)
                        df_master_combined = df_master_combined.drop_duplicates(subset=['IDPEL'], keep='first').reset_index(drop=True)

                        df_baru_map = df_baru.set_index('IDPEL')
                        in_baru = df_master_combined['IDPEL'].isin(df_baru_map.index)
                        
                        df_master_combined.loc[in_baru, 'KOKED'] = df_master_combined.loc[in_baru, 'IDPEL'].map(df_baru_map['KOKED'].to_dict())
                        df_master_combined.loc[in_baru, 'TARIP'] = df_master_combined.loc[in_baru, 'IDPEL'].map(df_baru_map['TARIP'].to_dict())
                        df_master_combined.loc[in_baru, 'DAYA'] = df_master_combined.loc[in_baru, 'IDPEL'].map(df_baru_map['DAYA'].to_dict())

                        nama_baru_map = df_baru_map['NAMA'].to_dict()
                        alamat_baru_map = df_baru_map['ALAMAT'].to_dict()

                        updated_nama = df_master_combined['IDPEL'].map(nama_baru_map)
                        valid_nama = in_baru & updated_nama.notna() & (updated_nama.astype(str).str.strip() != '') & (~updated_nama.astype(str).str.contains(r'\*', na=False))
                        df_master_combined.loc[valid_nama, 'NAMA'] = updated_nama[valid_nama]

                        updated_alamat = df_master_combined['IDPEL'].map(alamat_baru_map)
                        valid_alamat = in_baru & updated_alamat.notna() & (updated_alamat.astype(str).str.strip() != '') & (~updated_alamat.astype(str).str.contains(r'\*', na=False))
                        df_master_combined.loc[valid_alamat, 'ALAMAT'] = updated_alamat[valid_alamat]

                        zip_file.writestr("MASTER_UPDATED.xlsx", to_excel_bytes(df_master_combined[kolom_export]))

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
                        def export_excel(df_source, prefix=""):
                            if df_source.empty: return
                            for nama_file, kriteria in mapping.items():
                                if isinstance(kriteria, list): temp_df = df_source[df_source['KOKED'].astype(str).str[3:6].isin(kriteria)]
                                else: temp_df = df_source[df_source['KOKED'].astype(str).str[3:6] == kriteria]
                                if not temp_df.empty:
                                    temp_df = temp_df.drop_duplicates(subset=['IDPEL'], keep='first')
                                    zip_file.writestr(f"{prefix}{nama_file}.xlsx", to_excel_bytes(temp_df[kolom_export]))
                        export_excel(df_pb, "PB_")
                        export_excel(df_tetap, "")
                        
                    st.success("✅ File untuk petugas, master, dan pelanggan baru berhasil dibuat!")
                    st.download_button("📥 Download Distribusi (.zip)", data=zip_buffer.getvalue(), file_name="Hasil_PerPetugas.zip", mime="application/zip")
                except Exception as e:
                    st.error(f"❌ Error Tahap 1: {str(e)}")

# ==========================================
# TAB 2: REKAP & PERUBAHAN KOKED
# ==========================================
with tab2:
    st.header("Tahap 2: Gabung File Petugas & Mutasi KOKED")
    st.markdown("Digunakan untuk membuat hasil koked untuk di upload di PLN")
    tampilkan_info_header_master()
    col3, col4 = st.columns(2)
    with col3: files_petugas = st.file_uploader("[Tahap 2] Data Hasil Kerja Petugas", accept_multiple_files=True, key="t2_petugas")
    with col4:
        sumber_master_t2 = st.radio(
            "Sumber Data Master / Pembanding:",
            ["Upload File (.xlsx / .dbf)", "Ambil dari Supabase (dataplg3)"],
            horizontal=True,
            key="sumber_master_t2"
        )
        if sumber_master_t2 == "Upload File (.xlsx / .dbf)":
            files_lama_pembanding = st.file_uploader("[Tahap 2] Data Master / Bulan Sekarang", accept_multiple_files=True, key="t2_lama")
        else:
            files_lama_pembanding = None
            st.success("✅ Data Master Pembanding (`IDPEL, KDDK`) akan diambil otomatis dari Supabase.")

    if st.button("Proses Tahap 2", type="primary"):
        if not files_petugas or (sumber_master_t2 == "Upload File (.xlsx / .dbf)" and not files_lama_pembanding): 
            st.error("Silakan unggah Data Hasil Kerja Petugas dan pilih/unggah Data Master Pembanding.")
        else:
            with st.spinner("Membandingkan KOKED..."):
                try:
                    df_petugas = proses_list_file(files_petugas, 'petugas')
                    if sumber_master_t2 == "Ambil dari Supabase (dataplg3)":
                        df_sup_t2 = fetch_master_supabase("IDPEL,KDDK")
                        df_lama_pem = siapkan_df_master_standar(df_sup_t2, 'lama')
                    else:
                        df_lama_pem = proses_list_file(files_lama_pembanding, 'lama')

                    df_petugas.columns = df_petugas.columns.astype(str).str.strip().str.upper()
                    df_lama_pem.columns = df_lama_pem.columns.astype(str).str.strip().str.upper()

                    df_petugas.rename(columns={'KDDK': 'KOKED', 'TARIP': 'TARIF'}, inplace=True)
                    df_lama_pem.rename(columns={'KDDK': 'KOKED', 'TARIP': 'TARIF'}, inplace=True)

                    df_compare = df_petugas[['IDPEL', 'KOKED']].merge(df_lama_pem[['IDPEL', 'KOKED']], on='IDPEL', suffixes=('_BARU', '_LAMA'))
                    df_changed = df_compare[(df_compare['KOKED_BARU'] != df_compare['KOKED_LAMA']) & (df_compare['KOKED_LAMA'] != '')]
                    df_petugas['NO_URUT'] = pd.to_numeric(df_petugas['KOKED'].str[7:10], errors='coerce').fillna(0).astype(int)
                    df_petugas.sort_values(by=['KOKED', 'NO_URUT'], inplace=True)

                    for col in ['IDPEL', 'KOKED', 'NAMA', 'ALAMAT', 'TARIF', 'DAYA']:
                        if col not in df_petugas.columns: df_petugas[col] = ''

                    zip_buffer2 = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer2, "w", zipfile.ZIP_DEFLATED) as zip_file2:
                        if not df_changed.empty: zip_file2.writestr("PERUBAHAN_KOKED.txt", "\n".join((df_changed['IDPEL'].astype(str) + "|" + df_changed['KOKED_BARU'].astype(str)).tolist()))
                        zip_file2.writestr("DATA_GABUNGAN_FINAL.xlsx", to_excel_bytes(df_petugas[['IDPEL', 'KOKED', 'NAMA', 'ALAMAT', 'TARIF', 'DAYA']]))

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
            with st.spinner("Mengekstraksi data Tipe 1 (Tanpa Watermark)..."):
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
# TAB 4: IMPORT & EKSTRAK PDF (TIPE 2 - TURBO & RAPIH SESUAI ASLI)
# ==========================================
with tab4:
    st.header("Tahap 4: Ekstrak PDF Tipe 2 (Turbo & Format Utuh)")
    st.markdown("Menggunakan mesin **Turbo (PyMuPDF)** yang sangat cepat, namun telah dioptimalkan agar membaca **tabel secara utuh sesuai aslinya (16+ Kolom)** tanpa memotong atau menggabungkan data.")
    pdf_files_t4 = st.file_uploader("Upload File PDF Laporan", type=['pdf'], accept_multiple_files=True, key="t4_pdf")

    if st.button("🚀 Proses & Ekstrak Data (Tab 4)", type="primary"):
        if not pdf_files_t4: st.error("Silakan unggah setidaknya satu file PDF.")
        elif not PYMUPDF_SUPPORT: st.error("⚠️ Library `pymupdf` belum terpasang!")
        else:
            all_extracted_rows = []
            master_headers = None
            
            with st.spinner("🚀 Mengekstraksi PDF dengan Mode Turbo Cepat & Rapih..."):
                progress_bar = st.progress(0)
                status_text = st.empty()
                total_files = len(pdf_files_t4)
                total_baris = 0

                for f_idx, uploaded_pdf in enumerate(pdf_files_t4):
                    fname = uploaded_pdf.name
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp: 
                        tmp.write(uploaded_pdf.getvalue())
                        tmp_pdf_path = tmp.name
                    
                    try:
                        doc = fitz.open(tmp_pdf_path)
                        total_hal = len(doc)
                        for i, page in enumerate(doc):
                            tabs = page.find_tables()
                            if tabs.tables:
                                for tab in tabs.tables:
                                    raw_data = tab.extract()
                                    if not raw_data or len(raw_data) < 2: continue
                                        
                                    cleaned_table = []
                                    for row in raw_data:
                                        cleaned_row = [str(c).replace('\n', ' ').strip() if c is not None else "" for c in row]
                                        if any(cleaned_row): cleaned_table.append(cleaned_row)
                                            
                                    if not cleaned_table: continue
                                        
                                    if master_headers is None:
                                        header_keywords = ['NO. AGENDA', 'NAMA PEMOHON', 'NOPEL', 'NO. SIP']
                                        for r_idx, row in enumerate(cleaned_table[:10]):
                                            row_str = " ".join([c.upper() for c in row])
                                            if any(kw in row_str for kw in header_keywords):
                                                master_headers = row
                                                cleaned_table = cleaned_table[r_idx + 1:]
                                                break
                                        if master_headers is None:
                                            master_headers = [f"Kolom_{x+1}" for x in range(len(cleaned_table[0]))]

                                    for row in cleaned_table:
                                        row_str = " ".join([c.upper() for c in row])
                                        if any(kw in row_str for kw in ['NO. AGENDA', 'NAMA PEMOHON', 'NOPEL', 'HALAMAN']): continue
                                            
                                        if len(row) < len(master_headers): row.extend([""] * (len(master_headers) - len(row)))
                                        elif len(row) > len(master_headers): row = row[:len(master_headers)]
                                            
                                        all_extracted_rows.append(row)
                                        total_baris += 1

                            fitz.TOOLS.store_shrink(100)
                            current_progress = (f_idx + ((i + 1) / max(1, total_hal))) / total_files
                            progress_bar.progress(min(1.0, current_progress))
                            status_text.text(f"⚡ File {f_idx+1}/{total_files} ({fname}): Halaman {i+1}/{total_hal} (Terkumpul: {total_baris} baris)...")
                        doc.close()
                    except Exception as e:
                        st.error(f"Gagal memproses file {fname}: {str(e)}")
                    finally:
                        if os.path.exists(tmp_pdf_path): os.remove(tmp_pdf_path)
                        gc.collect()

            if all_extracted_rows and master_headers:
                status_text.empty()
                progress_bar.progress(1.0)
                unique_headers = []
                dilihat = {}
                for idx, col in enumerate(master_headers):
                    col_bersih = str(col).replace('\n', ' ').strip()
                    if not col_bersih: col_bersih = f"Kolom_Kosong_{idx+1}"
                    if col_bersih in dilihat:
                        dilihat[col_bersih] += 1
                        unique_headers.append(f"{col_bersih}_{dilihat[col_bersih]}")
                    else:
                        dilihat[col_bersih] = 0
                        unique_headers.append(col_bersih)
                
                df_final = pd.DataFrame(all_extracted_rows, columns=unique_headers)
                df_final = df_final.replace(r'^\s*$', np.nan, regex=True).dropna(how='all')
                df_final = df_final.fillna("")
                
                st.success(f"🎉 Selesai Cepat! Berhasil mengekstraksi {len(df_final)} baris data format utuh!")
                st.dataframe(df_final.head(50), use_container_width=True)
                st.download_button("📥 Download Ekstrak PDF Tipe 2 (.xlsx)", data=to_excel_bytes(df_final), file_name="Hasil_Ekstrak_Tipe2_Turbo_Rapih.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
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
                            
                            # MEMAKAI REGEX YANG BENAR (A-Z0-9)
                            meter = re.search(r'Nomor Meter\s*:\s*([A-Z0-9]+)', blok, re.IGNORECASE)
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
                    st.download_button("📥 Download Rekap Tagihan ICONPRN (.xlsx)", data=to_excel_bytes(df_hasil), file_name="rekap_tagihan_gabungan.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                else:
                    st.warning("⚠️ Tidak ada data tagihan valid yang ditemukan.")

# ==========================================
# TAB 6: APLIKASI UPDATE / LENGKAPI DATA DARI MASTER
# ==========================================
with tab6:
    st.header("Tahap 6: Update / Lengkapi Data dari Master Data")
    st.markdown("Meng-update / melengkapi **Data LAMA (file yang ingin dicarikan datanya)** menggunakan **Data BARU (Master Data)** berdasarkan IDPEL. Data bersimbol bintang (`*`) tidak akan menimpa data.")
    tampilkan_info_header_master()
    col_t6_1, col_t6_2 = st.columns(2)
    with col_t6_1: file_lama_m = st.file_uploader("Upload File Excel Data LAMA (Data yang ingin dilengkapi)", type=['xlsx', 'xls'], key="t6_lama")
    with col_t6_2:
        sumber_baru_t6 = st.radio("Sumber Data BARU (Master Data Referensi):", ["Upload File Excel", "Ambil dari Supabase (dataplg3)"], horizontal=True, key="sumber_baru_t6")
        if sumber_baru_t6 == "Upload File Excel":
            file_baru_m = st.file_uploader("Upload File Excel Data BARU (Master Data)", type=['xlsx', 'xls'], key="t6_baru")
            tambah_baris_baru_t6 = st.checkbox("Tambahkan juga IDPEL baru dari file kedua jika belum ada di file pertama", value=True, key="t6_add_new")
        else:
            file_baru_m = None
            tambah_baris_baru_t6 = False
            st.success("✅ Data BARU (Master Data) akan diambil otomatis dari tabel `dataplg3` Supabase.")

    if st.button("Proses Update Data (Tab 6)", type="primary"):
        if not file_lama_m or (sumber_baru_t6 == "Upload File Excel" and not file_baru_m):
            st.error("Silakan unggah File Data LAMA dan pilih/unggah Sumber Data BARU (Master).")
        else:
            with st.spinner("Memproses sinkronisasi dari master data..."):
                try:
                    df_lama = pd.read_excel(file_lama_m, dtype=str)
                    if sumber_baru_t6 == "Ambil dari Supabase (dataplg3)": df_baru = fetch_master_supabase("*").astype(str)
                    else: df_baru = pd.read_excel(file_baru_m, dtype=str)

                    kolom_asli_lama = df_lama.columns.tolist()
                    df_lama.columns = df_lama.columns.astype(str).str.strip().str.lower()
                    df_baru.columns = df_baru.columns.astype(str).str.strip().str.lower()

                    if 'koked' in df_lama.columns and 'kddk' in df_baru.columns and 'koked' not in df_baru.columns: df_baru.rename(columns={'kddk': 'koked'}, inplace=True)
                    elif 'kddk' in df_lama.columns and 'koked' in df_baru.columns and 'kddk' not in df_baru.columns: df_baru.rename(columns={'koked': 'kddk'}, inplace=True)
                    if 'tarip' in df_lama.columns and 'tarif' in df_baru.columns and 'tarip' not in df_baru.columns: df_baru.rename(columns={'tarif': 'tarip'}, inplace=True)
                    elif 'tarif' in df_lama.columns and 'tarip' in df_baru.columns and 'tarif' not in df_baru.columns: df_baru.rename(columns={'tarip': 'tarif'}, inplace=True)

                    kolom_yang_sama = df_baru.columns.intersection(df_lama.columns)
                    df_baru = df_baru[kolom_yang_sama]

                    if 'idpel' not in df_lama.columns or 'idpel' not in df_baru.columns: st.error("❌ Kolom 'IDPEL' tidak ditemukan di salah satu sumber data!")
                    else:
                        df_lama['idpel'] = df_lama['idpel'].str.replace('.0', '', regex=False).str.strip()
                        df_baru['idpel'] = df_baru['idpel'].str.replace('.0', '', regex=False).str.strip()

                        df_lama = df_lama.drop_duplicates(subset=['idpel'], keep='first')
                        df_baru = df_baru.drop_duplicates(subset=['idpel'], keep='first')

                        df_lama.set_index('idpel', inplace=True)
                        df_baru.set_index('idpel', inplace=True)

                        mask_exist = df_baru.index.isin(df_lama.index)
                        df_baru_exist = df_baru[mask_exist].copy()
                        df_baru_new = df_baru[~mask_exist].copy()

                        if not df_baru_exist.empty:
                            def bersihkan_sel(val):
                                if pd.isna(val): return np.nan
                                s = str(val).strip()
                                if '*' in s or s.lower() in ['nan', 'none', '<na>'] or s == '': return np.nan
                                return val
                            df_update_clean = df_baru_exist.apply(lambda col: col.map(bersihkan_sel))
                            df_lama.update(df_update_clean)

                        if tambah_baris_baru_t6 and not df_baru_new.empty: df_lama = pd.concat([df_lama, df_baru_new])

                        df_lama.reset_index(inplace=True)
                        mapping_kolom = dict(zip([c.lower() for c in kolom_asli_lama], kolom_asli_lama))
                        df_lama.rename(columns=mapping_kolom, inplace=True)

                        st.success(f"✅ Berhasil memperbarui data! Total data sekarang: {len(df_lama)} baris.")
                        st.dataframe(df_lama.head(50), use_container_width=True)
                        st.download_button("📥 Download Data Updated (.xlsx)", data=to_excel_bytes(df_lama), file_name="Data_Updated_Tab6.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan: {str(e)}")
                    
# ==========================================
# TAB 7: INFO DATA & LOKASI (SUPABASE DATABASE)
# ==========================================
with tab7:
    st.header("Tahap 7: Info Data & Lokasi Pelanggan")
    st.markdown("Digunakan untuk mencari data pelanggan berdasar idpel, nama, nomormeter")
    tampilkan_info_header_master()
    try:
        supabase = init_koneksi()
        with st.form(key="form_pencarian_supabase"):
            kategori = st.selectbox("🎯 Pilih Dasar Pencarian:", ["IDPEL", "NAMA", "NOMOR METER", "SEMUA KOLOM"])
            kata_kunci = st.text_input("🔍 Masukkan Kata Kunci:")
            tombol_cari = st.form_submit_button("⚡ Cari Data ", type="primary", use_container_width=True)
            
        if tombol_cari and kata_kunci:
            kunci_bersih = str(kata_kunci).strip()
            with st.spinner("Mencari langsung di Database Server..."):
                if kategori == "IDPEL": respon = supabase.table("dataplg3").select("*").ilike("IDPEL", f"%{kunci_bersih}%").execute()
                elif kategori == "NAMA": respon = supabase.table("dataplg3").select("*").ilike("NAMA", f"%{kunci_bersih}%").execute()
                elif kategori == "NOMOR METER": respon = supabase.table("dataplg3").select("*").ilike("NOMORKWH", f"%{kunci_bersih}%").execute()
                else: 
                    kondisi_or = f"IDPEL.ilike.%{kunci_bersih}%,NAMA.ilike.%{kunci_bersih}%,ALAMAT.ilike.%{kunci_bersih}%,NOMORKWH.ilike.%{kunci_bersih}%"
                    respon = supabase.table("dataplg3").select("*").or_(kondisi_or).execute()
                data_hasil = respon.data 
                
            if len(data_hasil) > 0:
                st.success(f"**Berhasil! Ditemukan {len(data_hasil)} :**")
                for baris in data_hasil:
                    id_val = baris.get('IDPEL', 'Detail')
                    nama_val = baris.get('NAMA', '')
                    teks_judul = f"👤 {id_val}"
                    if str(nama_val).strip() != '': teks_judul += f" - {nama_val}"
                        
                    with st.expander(teks_judul):
                        for nama_kolom, isi_kolom in baris.items():
                            if isi_kolom is not None and str(isi_kolom).strip() != "":
                                st.write(f"**{nama_kolom}:** {isi_kolom}")
                        
                        lat_val, lon_val = None, None
                        kx = baris.get("KOORDINAT X", "")
                        ky = baris.get("KOORDINAT Y", "")
                        try:
                            if kx and ky:
                                num_x = float(str(kx).replace(',', '.'))
                                num_y = float(str(ky).replace(',', '.'))
                                if -11.0 <= num_x <= 6.0: lat_val, lon_val = str(num_x), str(num_y)
                                elif -11.0 <= num_y <= 6.0: lat_val, lon_val = str(num_y), str(num_x)
                        except (ValueError, TypeError): pass
                            
                        st.markdown("---")
                        if lat_val and lon_val:
                            url_map = f"https://www.google.com/maps/search/?api=1&query={lat_val},{lon_val}"
                            st.link_button("📍 Buka Lokasi di Google Maps", url_map, type="primary", use_container_width=True)
                        else:
                            st.warning("⚠️ Data koordinat lokasi tidak ditemukan atau format tidak sesuai.")
            else:
                st.error("❌ Data tidak ditemukan di Database. Cek kembali penulisan kata kunci.")
    except Exception as e:
        st.error(f"Gagal menghubungkan ke database: {str(e)}")
        
# ==========================================
# TAB 8: UPDATE MASTER DATA (VERSI 2)
# ==========================================
with tab8:
    st.header("Tahap 8: Lengkapi Kolom Kosong dari Master Data (Versi 2)")
    st.write("Mengisi kolom yang KOSONG di **Data LAMA (file yang ingin dilengkapi)** dengan data dari **Data BARU (Master Data)** berdasarkan IDPEL. TIDAK MENAMBAH KOLOM BARU dan data yang sudah terisi TIDAK akan ditimpa.")
    tampilkan_info_header_master()
    st.markdown("---")
    col_l, col_b = st.columns(2)
    with col_l: file_lama = st.file_uploader("Upload File Excel Data LAMA (Data yang ingin dilengkapi)", type=['xlsx', 'xls'], key="t8_lama")
    with col_b:
        sumber_baru_t8 = st.radio("Sumber Data BARU (Master Data Referensi):", ["Upload File Excel", "Ambil dari Supabase (dataplg3)"], horizontal=True, key="sumber_baru_t8")
        if sumber_baru_t8 == "Upload File Excel": file_baru = st.file_uploader("Upload File Excel Data BARU (Master Data)", type=['xlsx', 'xls'], key="t8_baru")
        else:
            file_baru = None
            st.success("✅ Data BARU (Master Data) akan diambil otomatis dari tabel `dataplg3` Supabase.")

    if st.button("Proses Lengkapi Data (Tab 8)", type="primary"):
        if not file_lama or (sumber_baru_t8 == "Upload File Excel" and not file_baru):
            st.warning("⚠️ Harap upload file Excel Data LAMA dan pilih/upload Sumber Data BARU (Master)!")
        else:
            with st.spinner("Sedang memproses pengisian data dari Master..."):
                try:
                    df_lama = pd.read_excel(file_lama, dtype=str)
                    if sumber_baru_t8 == "Ambil dari Supabase (dataplg3)": df_baru = fetch_master_supabase("*").astype(str)
                    else: df_baru = pd.read_excel(file_baru, dtype=str)

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

                    cols_lama = df_lama.columns.astype(str).str.strip().str.lower()
                    cols_baru = df_baru.columns.astype(str).str.strip().str.lower()
                    df_lama.columns = buat_kolom_unik(cols_lama)
                    df_baru.columns = buat_kolom_unik(cols_baru)
                    
                    if 'koked' in df_lama.columns and 'kddk' in df_baru.columns and 'koked' not in df_baru.columns: df_baru.rename(columns={'kddk': 'koked'}, inplace=True)
                    elif 'kddk' in df_lama.columns and 'koked' in df_baru.columns and 'kddk' not in df_baru.columns: df_baru.rename(columns={'koked': 'kddk'}, inplace=True)
                    if 'tarip' in df_lama.columns and 'tarif' in df_baru.columns and 'tarip' not in df_baru.columns: df_baru.rename(columns={'tarif': 'tarip'}, inplace=True)
                    elif 'tarif' in df_lama.columns and 'tarip' in df_baru.columns and 'tarif' not in df_baru.columns: df_baru.rename(columns={'tarip': 'tarif'}, inplace=True)

                    kolom_format_lama = df_lama.columns.tolist()

                    if 'idpel' not in df_lama.columns or 'idpel' not in df_baru.columns:
                        st.error("❌ Kedua sumber data harus memiliki kolom 'IDPEL'!")
                    else:
                        df_lama = df_lama.replace(r'^\s*$', np.nan, regex=True).replace(['nan', 'NaN', '<NA>', 'None'], np.nan)
                        df_baru = df_baru.replace(r'^\s*$', np.nan, regex=True).replace(['nan', 'NaN', '<NA>', 'None'], np.nan)

                        for col in df_lama.columns:
                            df_lama[col] = df_lama[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                            df_lama[col] = df_lama[col].replace(['nan', 'None', '<NA>'], np.nan)
                        for col in df_baru.columns:
                            df_baru[col] = df_baru[col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                            df_baru[col] = df_baru[col].replace(['nan', 'None', '<NA>'], np.nan)

                        df_lama = df_lama.drop_duplicates(subset=['idpel'], keep='first')
                        df_baru = df_baru.drop_duplicates(subset=['idpel'], keep='first')
                        df_lama.set_index('idpel', inplace=True)
                        df_baru.set_index('idpel', inplace=True)

                        df_lama.update(df_baru, overwrite=False)
                        df_lama.reset_index(inplace=True)
                        df_result = df_lama[kolom_format_lama]

                        df_export = df_result.copy()
                        for col in df_export.columns:
                            df_export[col] = df_export[col].fillna("").astype(str)
                            df_export[col] = df_export[col].replace({'nan': '', 'None': '', '<NA>': ''})

                        st.success("✅ Data berhasil dilengkapi dari Master Data!")
                        st.write("Preview Hasil Update:")
                        st.dataframe(df_result.head(15), use_container_width=True)

                        output_t8 = io.BytesIO()
                        with pd.ExcelWriter(output_t8, engine='openpyxl') as writer:
                            df_export.to_excel(writer, index=False, sheet_name='Master_Updated')
                            ws = writer.sheets['Master_Updated']
                            for col in ws.columns:
                                max_len = 0
                                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                                for cell in col:
                                    if cell.value is not None and str(cell.value).strip() != "":
                                        cell.value = str(cell.value)
                                        cell.data_type = 's'
                                        cell.number_format = '@'
                                    val_str = str(cell.value) if cell.value is not None else ""
                                    max_len = max(max_len, len(val_str))
                                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                        st.download_button(label="⬇️ Download Hasil Lengkapi Data (.xlsx)", data=output_t8.getvalue(), file_name="Data_Dilengkapi_Master.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="t8_download")
                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan saat memproses data: {e}")

# ==========================================
# TAB 9: ISI PETUGAS
# ==========================================
with tab9:
    st.header("Tahap 9: Penentuan Petugas Otomatis")
    st.write("Mengisi kolom Petugas secara otomatis berdasarkan logika Daya > 33000 (PLN), IDPEL khusus, dan 3 karakter tengah KOKED/KDDK.")
    st.markdown("---")
    file_t9 = st.file_uploader("Upload File Excel (Memiliki kolom IDPEL, KOKED/KDDK, DAYA)", type=['xlsx', 'xls'], key="t9_file")

    if st.button("Proses & Isi Petugas", type="primary"):
        if file_t9 is None: st.warning("⚠️ Harap upload file Excel terlebih dahulu!")
        else:
            with st.spinner("Sedang memproses penentuan petugas..."):
                try:
                    df = pd.read_excel(file_t9, dtype=str)
                    df_work = df.copy()
                    df_work.columns = df_work.columns.astype(str).str.strip().str.lower()

                    if 'kddk' in df_work.columns and 'koked' not in df_work.columns: df_work['koked'] = df_work['kddk']
                    if not {'idpel', 'koked', 'daya'}.issubset(set(df_work.columns)):
                        st.error("❌ Error: File Excel harus memiliki kolom bernama 'IDPEL', 'KOKED' (atau 'KDDK'), dan 'DAYA'.")
                    else:
                        def tentukan_petugas_tab9(row):
                            try:
                                daya_val = str(row['daya']).replace(',', '.').strip()
                                daya = float(daya_val)
                            except: daya = 0.0

                            idpel = str(row['idpel']).replace('.0', '').strip() if pd.notna(row['idpel']) and str(row['idpel']).lower() != 'nan' else ""
                            koked = str(row['koked']).strip() if pd.notna(row['koked']) and str(row['koked']).lower() != 'nan' else ""

                            if daya > 33000: return "PLN"
                            
                            idpel_khusus = {
                                "524051069054": "c28", "524051263717": "c36", "524051265123": "c36",
                                "524051104194": "c04", "524051000615": "c08", "524050867033": "c08"
                            }
                            if idpel in idpel_khusus: return idpel_khusus[idpel]
                            
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
                                if kode_mid in mapping_kddk: return mapping_kddk[kode_mid]
                            return "BARU"

                        hasil_petugas = df_work.apply(tentukan_petugas_tab9, axis=1)
                        col_petugas_asli = next((c for c in df.columns if str(c).strip().lower() == 'petugas'), None)
                        
                        if col_petugas_asli: df[col_petugas_asli] = hasil_petugas
                        else: df['petugas'] = hasil_petugas

                        for col in df.columns:
                            df[col] = df[col].fillna("").astype(str)
                            df[col] = df[col].replace({'nan': '', 'None': '', '<NA>': ''})
                            if 'idpel' in col.lower() or 'nik' in col.lower():
                                df[col] = df[col].str.replace(r'\.0$', '', regex=True)

                        st.success("✅ Kolom Petugas berhasil diisi!")
                        st.write("Preview Hasil Data:")
                        st.dataframe(df.head(15), use_container_width=True)

                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='Data_Petugas')
                            ws = writer.sheets['Data_Petugas']
                            for col in ws.columns:
                                max_len = 0
                                col_letter = openpyxl.utils.get_column_letter(col[0].column)
                                for cell in col:
                                    if cell.value is not None and str(cell.value).strip() != "":
                                        cell.value = str(cell.value)
                                        cell.data_type = 's'
                                        cell.number_format = '@'
                                    val_str = str(cell.value) if cell.value is not None else ""
                                    max_len = max(max_len, len(val_str))
                                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                        st.download_button(label="⬇️ Download Hasil Excel", data=output.getvalue(), file_name="Data_Isi_Petugas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="t9_download")
                except Exception as e:
                    st.error(f"❌ Terjadi kesalahan saat memproses data: {e}")

# ==========================================
# TAB 10: SPLIT DATA
# ==========================================
with tab10:
    st.header("Tahap 10: Split Data ke Beberapa Worksheet atau File")
    st.write("Membagi satu tabel data menjadi beberapa worksheet (sheet) atau file Excel terpisah berdasarkan nilai pada kolom tertentu.")
    st.markdown("---")
    file_t10 = st.file_uploader("Upload File Excel yang ingin di-split", type=['xlsx', 'xls'], key="t10_file")

    if file_t10 is not None:
        try:
            df_t10 = pd.read_excel(file_t10, dtype=str)
            for col in df_t10.columns:
                df_t10[col] = df_t10[col].fillna("").astype(str)
                df_t10[col] = df_t10[col].replace({'nan': '', 'None': '', '<NA>': ''})
                if 'idpel' in str(col).lower() or 'nik' in str(col).lower():
                    df_t10[col] = df_t10[col].str.replace(r'\.0$', '', regex=True)

            st.write("Preview Data Asli:")
            st.dataframe(df_t10.head(), use_container_width=True)

            st.markdown("### Pengaturan Split Data")
            kolom_pilihan = st.selectbox("Split berdasarkan kolom (Specific column):", df_t10.columns.tolist())

            col1, col2 = st.columns(2)
            with col1: prefix = st.text_input("Prefix (opsional):", help="Tambahan teks di depan nama sheet/file")
            with col2: suffix = st.text_input("Suffix (opsional):", help="Misal: Nik Padan")

            mode_split = st.radio("Pilih Mode Output:", ["Multiple Sheets (1 File Excel)", "Multiple Files (Download sebagai ZIP)"], help="Pilih apakah ingin hasil split berada dalam 1 file beda sheet, atau file yang benar-benar terpisah.")

            if st.button("Proses Split Data", type="primary"):
                with st.spinner("Sedang membagi data..."):
                    df_t10_filtered = df_t10[df_t10[kolom_pilihan].str.strip() != ""]
                    nilai_unik = df_t10_filtered[kolom_pilihan].unique()

                    if mode_split == "Multiple Sheets (1 File Excel)":
                        output_t10 = io.BytesIO()
                        with pd.ExcelWriter(output_t10, engine='openpyxl') as writer:
                            for nilai in nilai_unik:
                                df_filtered = df_t10[df_t10[kolom_pilihan] == nilai]
                                nilai_str = str(nilai).strip()
                                nama_custom = f"{prefix} {nilai_str} {suffix}".strip()
                                nama_bersih = re.sub(r'[\\/*?:\[\]]', '', nama_custom)[:31]
                                if not nama_bersih: nama_bersih = "Data"

                                df_filtered.to_excel(writer, index=False, sheet_name=nama_bersih)
                                ws = writer.sheets[nama_bersih]
                                for col in ws.columns:
                                    max_len = 0
                                    col_letter = openpyxl.utils.get_column_letter(col[0].column)
                                    for cell in col:
                                        if cell.value is not None and str(cell.value).strip() != "":
                                            cell.value = str(cell.value)
                                            cell.data_type = 's'
                                            cell.number_format = '@'
                                        val_str = str(cell.value) if cell.value is not None else ""
                                        max_len = max(max_len, len(val_str))
                                    ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

                        st.success(f"✅ Data berhasil dipisah menjadi {len(nilai_unik)} sheet dalam 1 file Excel!")
                        st.download_button(label="⬇️ Download Excel (Multiple Sheets)", data=output_t10.getvalue(), file_name=f"Data_Split_{kolom_pilihan}_Sheets.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="t10_download_sheets")

                    else:
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                            for nilai in nilai_unik:
                                df_filtered = df_t10[df_t10[kolom_pilihan] == nilai]
                                nilai_str = str(nilai).strip()
                                nama_custom = f"{prefix} {nilai_str} {suffix}".strip()
                                nama_bersih = re.sub(r'[\\/*?:\[\]<>|"]', '', nama_custom)
                                if not nama_bersih: nama_bersih = "Data"
                                nama_file = f"{nama_bersih}.xlsx"

                                excel_buffer = io.BytesIO()
                                with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                                    df_filtered.to_excel(writer, index=False, sheet_name="Data")
                                    ws = writer.sheets["Data"]
                                    for col in ws.columns:
                                        max_len = 0
                                        col_letter = openpyxl.utils.get_column_letter(col[0].column)
                                        for cell in col:
                                            if cell.value is not None and str(cell.value).strip() != "":
                                                cell.value = str(cell.value)
                                                cell.data_type = 's'
                                                cell.number_format = '@'
                                            val_str = str(cell.value) if cell.value is not None else ""
                                            max_len = max(max_len, len(val_str))
                                        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
                                zip_file.writestr(nama_file, excel_buffer.getvalue())

                        st.success(f"✅ Data berhasil dipisah menjadi {len(nilai_unik)} file Excel terpisah!")
                        st.download_button(label="⬇️ Download ZIP (Multiple Files)", data=zip_buffer.getvalue(), file_name=f"Data_Split_{kolom_pilihan}_Files.zip", mime="application/zip", key="t10_download_zip")
        except Exception as e:
            st.error(f"❌ Terjadi kesalahan saat memproses data: {e}")

# ==========================================
# TAB 11: EKSTRAK PDF LAPORAN FORMAT ASLI (AKURAT)
# ==========================================
with tab11:
    st.header("Tahap 11: Ekstrak PDF Laporan Sesuai Asli (Akurat)")
    st.markdown("Menggunakan mesin Hybrid (seperti Tab 3) yang sudah terbukti berhasil membaca garis laporan, tetapi fitur ini **mempertahankan seluruh kolom asli (16 kolom)** tanpa memfilter datanya.")
    pdf_files_t11 = st.file_uploader("Upload File PDF Laporan (Tab 11)", type=['pdf'], accept_multiple_files=True, key="t11_pdf")

    if st.button("Proses & Ekstrak Laporan (Akurat)", type="primary"):
        if not pdf_files_t11: st.error("Silakan unggah setidaknya satu file PDF.")
        elif not PDF_SUPPORT: st.error("Library `pdfplumber` belum terinstall. Pastikan sudah ada di requirements.txt!")
        else:
            all_extracted_dfs = []
            with st.spinner("⏳ Mengekstraksi tabel dengan mesin Hybrid..."):
                progress_bar = st.progress(0)
                status_text = st.empty()
                total_files = len(pdf_files_t11)
                for f_idx, uploaded_pdf in enumerate(pdf_files_t11):
                    fname = uploaded_pdf.name
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp: 
                        tmp.write(uploaded_pdf.getvalue())
                        tmp_path = tmp.name
                    try:
                        with pdfplumber.open(tmp_path) as pdf:
                            total_hal = len(pdf.pages)
                            for i, page in enumerate(pdf.pages):
                                tables = get_pdf_tables_tipe1(page)
                                for table in tables:
                                    df_hybrid = process_hybrid_table(table)
                                    if df_hybrid is not None and not df_hybrid.empty: 
                                        all_extracted_dfs.append(df_hybrid)
                                page.flush_cache()
                                current_progress = (f_idx + ((i + 1) / max(1, total_hal))) / total_files
                                progress_bar.progress(min(1.0, current_progress))
                                status_text.text(f"⚡ File {f_idx+1}/{total_files} ({fname}): Halaman {i+1}/{total_hal} diproses...")
                    except Exception as e:
                        st.error(f"Gagal memproses file {fname}: {str(e)}")
                    finally:
                        if os.path.exists(tmp_path): os.remove(tmp_path)
                gc.collect()

            if all_extracted_dfs:
                status_text.empty()
                progress_bar.progress(1.0)
                df_final = pd.concat(all_extracted_dfs, ignore_index=True)
                
                unique_headers = []
                dilihat = {}
                for idx, col in enumerate(df_final.columns):
                    col_bersih = str(col).replace('\n', ' ').strip()
                    if not col_bersih: col_bersih = f"Kolom_Kosong_{idx+1}"
                    if col_bersih in dilihat:
                        dilihat[col_bersih] += 1
                        unique_headers.append(f"{col_bersih}_{dilihat[col_bersih]}")
                    else:
                        dilihat[col_bersih] = 0
                        unique_headers.append(col_bersih)
                
                df_final.columns = unique_headers
                df_final = df_final.replace(r'^\s*$', np.nan, regex=True).dropna(how='all')
                df_final = df_final.fillna("")
                
                st.success(f"🎉 Selesai! Berhasil mengekstraksi {len(df_final)} baris data format tabel asli!")
                st.dataframe(df_final.head(50), use_container_width=True)
                st.download_button("📥 Download Excel Laporan (.xlsx)", data=to_excel_bytes(df_final), file_name="Hasil_Ekstrak_Laporan_Akurat.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", key="dl_tab11")
            else:
                st.warning("⚠️ Tidak ada data tabel yang terdeteksi dengan format tersebut.")

# ==========================================
# TAB 12: CETAK TUL VI-01
# ==========================================
with tab12:
    st.header("Tahap 12: Cetak TUL VI-01 (Presisi & Rata Kanan)")
    st.markdown("Digunakan untuk mencetak Pemberitahuan Pelaksanaan Pemutusan Sementara Sambungan Tenaga Listrik ke format `.iconprn` yang siap dikirim ke printer Dot Matrix.")
    
    col_t12_1, col_t12_2 = st.columns([1, 2])
    with col_t12_1:
        st.subheader("🖨️ Pengaturan Cetak")
        mode_cetak = st.radio("Pilih Tahap Pencetakan:", ["Tahap 2: Isi Blangko Saja", "Tahap 1 & 2 Sekaligus", "Tahap 1 Saja: Cetak Blangko Kosong"], index=0, key="t12_mode")
        tipe_printer = st.selectbox("Pilih Printer yang Digunakan:", ["Epson LQ-2190 (Standar Blangko 15 CPI - ESC g)", "Epson LX-310 / LX-300+II (9-Pin Condensed 17 CPI)", "Epson LX-310 / LX-300+II (9-Pin Elite 12 CPI)"], help="Jika memakai kertas blangko lama, SELALU pilih opsi LQ-2190.", key="t12_printer")
        
        with st.expander("📐 Kalibrasi Posisi Kertas (Global)"):
            geser_kanan = st.number_input("Geser Kanan/Kiri Semua Teks (Spasi):", min_value=-15, max_value=30, value=0, key="t12_kanan")
            geser_bawah = st.number_input("Geser Turun/Naik Semua Teks (Baris):", min_value=-5, max_value=10, value=0, key="t12_bawah")
            tinggi_halaman = st.number_input("Jumlah Baris per Lembar:", min_value=30, max_value=66, value=44, key="t12_tinggi")

        with st.expander("✍️ Pengaturan Tanda Tangan"):
            cetak_kota_jabatan = st.checkbox("Cetak tulisan Kota & Jabatan", value=False, key="t12_ctk_jab")
            geser_tgl = st.number_input("Geser Kiri/Kanan KHUSUS Tanggal (Spasi):", min_value=-30, max_value=30, value=8, key="t12_geser_tgl")
            kota_cetak = st.text_input("Kota", value="TULUNG", key="t12_kota")
            tgl_cetak = st.text_input("Tanggal Cetak", value="05-10-2026", key="t12_tgl")
            jabatan = st.text_input("Jabatan", value="MANAGER", key="t12_jab")
            nama_manager = st.text_input("Nama Manager", value="MARTONO AJI PRABOWO", key="t12_mgr")

    with col_t12_2:
        init_esc = get_printer_init_code(tipe_printer)
        
        if "Tahap 1 Saja" in mode_cetak:
            st.subheader("📄 Cetak Blangko Kosong (.iconprn)")
            jml = st.number_input("Jumlah Lembar Blangko:", min_value=1, max_value=1000, value=50, key="t12_jml_blangko")
            raw_blanko = buat_blangko_kosong(init_esc, jml, tinggi_halaman)
            st.download_button(label=f"🖨️ Download BLANKO_{jml}_LEMBAR.iconprn", data=raw_blanko.encode("latin1", errors="replace"), file_name=f"BLANKO_{jml}_LEMBAR.iconprn", mime="application/octet-stream", key="t12_dl_blangko")
        else:
            uploaded_files_t12 = st.file_uploader("📂 Upload File Data (.xlsx, .xls, .iconprn, atau .zip)", type=["xlsx", "xls", "iconprn", "prn", "zip"], accept_multiple_files=True, key="t12_uploads")

            if uploaded_files_t12:
                df_t12 = None
                col_map_final = None 
                semua_data_iconprn = []
                excel_uploaded = False

                with st.spinner("Memproses file masukan..."):
                    for uploaded_file in uploaded_files_t12:
                        nama_file = uploaded_file.name.lower()
                        if nama_file.endswith(('.xlsx', '.xls')): excel_uploaded = True
                        elif nama_file.endswith('.zip'):
                            with zipfile.ZipFile(uploaded_file, 'r') as z:
                                for z_name in z.namelist():
                                    if z_name.lower().endswith(('.iconprn', '.prn', '.txt')):
                                        teks_mentah = z.read(z_name).decode('latin1', errors='ignore')
                                        semua_data_iconprn.extend(baca_data_dari_iconprn_tul(teks_mentah))
                        else:
                            teks_mentah = uploaded_file.getvalue().decode('latin1', errors='ignore')
                            semua_data_iconprn.extend(baca_data_dari_iconprn_tul(teks_mentah))
                            
                if excel_uploaded:
                    file_excel = next(f for f in uploaded_files_t12 if f.name.lower().endswith(('.xlsx', '.xls')))
                    xls = pd.ExcelFile(file_excel)
                    pilih_sheet = st.selectbox("Pilih Sheet Excel:", xls.sheet_names, key="t12_sheet") if len(xls.sheet_names) > 1 else xls.sheet_names[0]
                    df_excel = pd.read_excel(file_excel, sheet_name=pilih_sheet, dtype=str)
                    
                    auto_map = deteksi_kolom_otomatis(df_excel.columns.tolist())
                    kolom_hilang = [k for k, v in auto_map.items() if v is None]

                    with st.expander("🔗 Pengaturan Pencocokan Kolom Excel", expanded=len(kolom_hilang) > 0):
                        if kolom_hilang: st.warning(f"⚠️ Ada nama kolom yang berbeda: **{', '.join(kolom_hilang)}**. Silakan pilih manual:")
                        else: st.success("✅ Semua kolom Excel otomatis dikenali!")

                        opsi_kolom = ["(Kosongkan)"] + df_excel.columns.tolist()
                        col_map_final = {}
                        cols_ui = st.columns(3)
                        for idx, (field_blangko, terdeteksi) in enumerate(auto_map.items()):
                            with cols_ui[idx % 3]:
                                idx_default = opsi_kolom.index(terdeteksi) if terdeteksi in opsi_kolom else 0
                                col_map_final[field_blangko] = st.selectbox(f"[{field_blangko}]:", opsi_kolom, index=idx_default, key=f"t12_map_{field_blangko}")
                                if col_map_final[field_blangko] == "(Kosongkan)": col_map_final[field_blangko] = None
                    df_t12 = df_excel

                else:
                    if semua_data_iconprn: df_t12 = pd.DataFrame(semua_data_iconprn)
                        
                if df_t12 is not None and len(df_t12) > 0:
                    st.success(f"✅ Berhasil memuat {len(df_t12)} tagihan pelanggan!")
                    
                    kol_petugas = col_map_final.get("petugas") if col_map_final else "petugas"
                    
                    if kol_petugas and kol_petugas in df_t12.columns:
                        df_t12[kol_petugas] = df_t12[kol_petugas].fillna("Tanpa_Petugas").astype(str)
                        daftar_ptg = sorted(df_t12[kol_petugas].unique().tolist())

                        st.subheader("👷 Pilih Petugas & Urutan Cetak")
                        ptg_terpilih = st.multiselect("Pilih Kode Petugas:", daftar_ptg, default=[daftar_ptg[0]] if daftar_ptg else [], key="t12_ptg")
                        df_saring = df_t12[df_t12[kol_petugas].isin(ptg_terpilih)].reset_index(drop=True)
                    else:
                        st.subheader("📋 Rentang Urutan Cetak")
                        ptg_terpilih = ["Semua"]
                        df_saring = df_t12.reset_index(drop=True)

                    total_data = len(df_saring)
                    dynamic_key = f"{'_'.join(ptg_terpilih)}_{total_data}"
                    
                    c2, c3 = st.columns(2)
                    with c2: urut_awal = st.number_input("Mulai Urutan ke-:", min_value=1, max_value=max(1, total_data), value=1, key=f"t12_awal_{dynamic_key}")
                    with c3: urut_akhir = st.number_input("Sampai Urutan ke-:", min_value=1, max_value=max(1, total_data), value=max(1, total_data), key=f"t12_akhir_{dynamic_key}")

                    df_cetak = df_saring.iloc[urut_awal - 1 : urut_akhir]
                    st.info(f"Siap mencetak **{len(df_cetak)} lembar** (Petugas: **{', '.join(ptg_terpilih)}**, Urutan {urut_awal} s/d {urut_akhir}).")
                    st.dataframe(df_cetak, use_container_width=True, height=220)

                    if len(df_cetak) > 0:
                        hasil_iconprn = ""
                        for _, baris in df_cetak.iterrows():
                            if "Tahap 2" in mode_cetak: hasil_iconprn += buat_isian_blangko(baris, col_map_final, init_esc, kota_cetak, tgl_cetak, jabatan, nama_manager, cetak_kota_jabatan, geser_tgl, geser_kanan, geser_bawah, tinggi_halaman)
                            else: hasil_iconprn += buat_blangko_dan_isi(baris, col_map_final, init_esc, kota_cetak, tgl_cetak, jabatan, nama_manager, tinggi_halaman)

                        nama_ptg_file = "_".join(ptg_terpilih)
                        
                        b1, b2 = st.columns(2)
                        with b1:
                            st.download_button(
                                label=f"🖨️ CETAK PETUGAS {nama_ptg_file} ({len(df_cetak)} Lbr)",
                                data=hasil_iconprn.encode("latin1", errors="replace"),
                                file_name=f"ISI_BLANGKO_{nama_ptg_file}_{urut_awal}_{urut_akhir}.iconprn",
                                mime="application/octet-stream",
                                type="primary",
                                use_container_width=True,
                                key=f"t12_dl_cetak_{dynamic_key}"
                            )
                        with b2:
                            if kol_petugas and kol_petugas in df_t12.columns:
                                buf = io.BytesIO()
                                with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                                    for kode_p, grp in df_cetak.groupby(kol_petugas):
                                        teks_p = ""
                                        for _, r in grp.iterrows():
                                            if "Tahap 2" in mode_cetak: teks_p += buat_isian_blangko(r, col_map_final, init_esc, kota_cetak, tgl_cetak, jabatan, nama_manager, cetak_kota_jabatan, geser_tgl, geser_kanan, geser_bawah, tinggi_halaman)
                                            else: teks_p += buat_blangko_dan_isi(r, col_map_final, init_esc, kota_cetak, tgl_cetak, jabatan, nama_manager, tinggi_halaman)
                                        zf.writestr(f"PETUGAS_{kode_p}_{len(grp)}lembar.iconprn", teks_p.encode("latin1", errors="replace"))
                                
                                st.download_button(
                                    label="📦 DOWNLOAD PAKET ZIP",
                                    data=buf.getvalue(),
                                    file_name=f"PAKET_CETAK_{urut_awal}_sd_{urut_akhir}.zip",
                                    mime="application/zip",
                                    use_container_width=True,
                                    key=f"t12_dl_zip_{dynamic_key}"
                                )
