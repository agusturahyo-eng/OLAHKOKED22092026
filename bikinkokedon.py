import re
import os
import glob
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================================================================
# FUNGSI UNTUK MENERJEMAHKAN RUMUS EXCEL "PETUGAS" OTOMATIS KE PYTHON
# =========================================================================
def tentukan_petugas(idpel, tarif_daya, kddk):
    daya = 0
    match_daya = re.search(r'/(\d+)', tarif_daya)
    if match_daya:
        daya = int(match_daya.group(1))
        
    if daya > 33000:
        return "PLN"
    
    idpel_khusus = {
        "524051069054": "c28", "524051263717": "c36", "524051265123": "c36",
        "524051104194": "c04", "524051000615": "c08", "524050867033": "c08"
    }
    if idpel in idpel_khusus:
        return idpel_khusus[idpel]
        
    if len(kddk) >= 6:
        kode_mid = kddk[3:6].upper()
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

def parsing_banyak_file_iconprn_ke_excel(folder_path=".", output_excel="rekap_tagihan_gabungan.xlsx"):
    search_pattern = os.path.join(folder_path, "*.iconprn")
    daftar_file = glob.glob(search_pattern)

    if not daftar_file:
        print(f"ERROR: Tidak ditemukan file berekstensi '.iconprn' di folder ini!")
        return

    print(f"Ditemukan {len(daftar_file)} file .iconprn untuk diproses.\n")
    semua_hasil = []

    for file_path in daftar_file:
        nama_file = os.path.basename(file_path)
        print(f"Memproses file: {nama_file}...")
        
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()

            blok_pelanggan = re.split(r'PEMBERITAHUAN PELAKSANAAN PEMUTUSAN', content)
            jumlah_data_file = 0

            for blok in blok_pelanggan:
                if "ID. Pelanggan" not in blok:
                    continue

                # =====================================================
                # TEMPLATE DICTIONARY: MEMAKSA URUTAN KOLOM TETAP
                # =====================================================
                data = {
                    'IDPEL': "",
                    'Nomor TUL': "",
                    'Nama': "",
                    'KDDK': "",
                    'Gardu/Tiang': "",
                    'Loket': "",
                    'Alamat': "",
                    'Nomor Meter': "",
                    'Tarif/Daya': "",
                    'Kelompok': "",
                    'Bulan Rekening': "",
                    'Bulan Keterlambatan': "",
                    'Jumlah Rekening': 0,
                    'Jumlah Denda': 0,
                    'Jumlah Tunggakan': 0,
                    'petugas': ""
                }

                idpel = re.search(r'ID\. Pelanggan\s*:\s*[^0-9]*(\d{11,13})', blok)
                if idpel: data['IDPEL'] = str(idpel.group(1).strip())

                tul = re.search(r'NO\. TUL\s*:\s*([A-Z0-9/\-]+)', blok)
                if tul: data['Nomor TUL'] = tul.group(1).strip()

                nama = re.search(r'Nama\s*:\s*(.+)', blok)
                if nama: data['Nama'] = nama.group(1).strip()

                kddk = re.search(r'Kode Kedudukan\s*:\s*([A-Z0-9]+)', blok)
                if kddk: data['KDDK'] = kddk.group(1).strip()

                # PERBAIKAN: Berhenti mencari jika menemukan >=2 spasi berjejer, atau bertemu kata "Loket"
                gardu = re.search(r'Gardu\s*/?\s*Tiang\s*:\s*(.*?)(?=\s{2,}|\s+Loket\s*:|\n|\r|$)', blok, re.IGNORECASE)
                if gardu: data['Gardu/Tiang'] = gardu.group(1).strip()

                # PERBAIKAN: Loket hanya akan membaca karakter sampai ia bertemu spasi ganda, atau teks 'Tarip', 'Alamat', dll.
                loket = re.search(r'Loket\s*:\s*(.*?)(?=\s{2,}|\s+Tarip|\s+Tarif|\s+Alamat|\s+Kelompok|\n|\r|$)', blok, re.IGNORECASE)
                if loket: 
                    val_loket = loket.group(1).strip()
                    # Filter pembersih terakhir jika masih ada kebocoran
                    if "Tarip" in val_loket or "Kelompok" in val_loket or "Daya" in val_loket:
                        val_loket = ""
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
                jumlah_data_file += 1

            print(f"  -> Berhasil mengekstrak {jumlah_data_file} data.")

        except Exception as e:
            print(f"  -> Gagal membaca file {nama_file}. Error: {e}")

    if semua_hasil:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Rekap Tagihan"

        headers = list(semua_hasil[0].keys())
        ws.append(headers)

        font_header = Font(name="Arial", size=10, bold=False) 
        font_data = Font(name="Arial", size=10)
        border_tipis = Border(
            left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
        )

        align_left = Alignment(horizontal="left", vertical="center")
        align_right = Alignment(horizontal="right", vertical="center")

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = font_header
            cell.border = border_tipis

        kolom_nominal = ["Jumlah Rekening", "Jumlah Denda", "Jumlah Tunggakan"]
        kolom_teks = ["IDPEL", "Nomor Meter", "Kelompok"]

        # Tulis Data
        for row_idx, row_data in enumerate(semua_hasil, start=2):
            for col_idx, key in enumerate(headers, start=1):
                val = row_data[key]
                cell = ws.cell(row=row_idx, column=col_idx, value=val)
                cell.font = font_data
                cell.border = border_tipis

                if key in kolom_nominal:
                    cell.number_format = '#,##0'  # Format Ribuan Excel
                    cell.alignment = align_right
                elif key in kolom_teks:
                    cell.number_format = '@'      
                    cell.alignment = align_left
                else:
                    cell.alignment = align_left

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 1, 10)

        wb.save(output_excel)

        print("\n==================================================")
        print(f"SUKSES! Total {len(semua_hasil)} data berhasil digabungkan.")
        print(f"File Excel tersimpan: '{output_excel}'")
        print("==================================================")

if __name__ == "__main__":
    parsing_banyak_file_iconprn_ke_excel()
    input("\nTekan Enter untuk menutup program...")

