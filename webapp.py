import streamlit as st
import json
import pandas as pd
import re
import os
import tempfile
from pathlib import Path
import numpy as np

# Coba import pdfplumber, jika error gunakan pymupdf
try:
    import pdfplumber
    PDF_LIBRARY = "pdfplumber"
except ImportError:
    try:
        import fitz  # pymupdf
        PDF_LIBRARY = "pymupdf"
    except ImportError:
        st.error("❌ Tidak ada library PDF yang terinstall. Pastikan pdfplumber atau pymupdf ada di requirements.txt")
        PDF_LIBRARY = None

# Set page configuration
st.set_page_config(
    page_title="PDF to Excel Converter - SLIK Report",
    page_icon="📊",
    layout="wide"
)

# Custom CSS untuk styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #2e86ab;
        margin-top: 1.5rem;
        margin-bottom: 1rem;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 15px;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

# Judul aplikasi
st.markdown('<div class="main-header">🔄 PDF to Excel Converter - SLIK Report</div>', unsafe_allow_html=True)

# Sidebar untuk upload file
st.sidebar.title("📁 Upload PDF Files")
uploaded_files = st.sidebar.file_uploader(
    "Pilih file PDF SLIK Report", 
    type=["pdf"], 
    accept_multiple_files=True
)

# Fungsi ekstraksi PDF yang kompatibel dengan kedua library
def extract_text_from_pdf(pdf_file, library):
    """Mengekstrak teks dari PDF menggunakan library yang tersedia"""
    text_data = []
    
    if library == "pdfplumber":
        with pdfplumber.open(pdf_file) as pdf:  
            for page in pdf.pages:  
                text = page.extract_text()  
                if text: 
                    text_data.append(text)
                    
    elif library == "pymupdf":
        doc = fitz.open(stream=pdf_file.read(), filetype="pdf")
        for page in doc:
            text = page.get_text()
            if text:
                text_data.append(text)
        doc.close()
        
    return text_data

def pdf_to_json(pdf_file, json_file_path):
    """Mengkonversi file PDF ke JSON"""
    if PDF_LIBRARY is None:
        st.error("Library PDF tidak tersedia")
        return []
        
    text_data = extract_text_from_pdf(pdf_file, PDF_LIBRARY)

    # SIMPAN TEKS DALAM FORMAT JSON
    with open(json_file_path, 'w', encoding='utf-8') as output_file:  
        json.dump(text_data, output_file, ensure_ascii=False, indent=4)
    
    return text_data

# Tampilkan informasi library yang digunakan
if PDF_LIBRARY:
    st.sidebar.info(f"📚 Using: {PDF_LIBRARY}")
else:
    st.error("""
    ❌ Library PDF tidak terinstall. 
    
    Tambahkan salah satu ke requirements.txt:
    - pdfplumber==0.10.3
    - atau pymupdf==1.23.8
    """)

def read_json_files(json_files):
    """Membaca semua file JSON dan menggabungkannya"""
    dataframes = []
    for json_file in json_files:
        try:
            df = pd.read_json(json_file)  
            dataframes.append(df)
        except Exception as e:
            st.error(f"Gagal membaca file {json_file}: {str(e)}")
    
    if not dataframes:
        return None
    
    combined_df = pd.concat(dataframes, ignore_index=True)
    return combined_df

def process_kredit_data(combined_data):
    """Memproses data kredit dari data gabungan"""
    # Pilih Baris == 'Informasi Debitur' atau 'Jenis Kredit/Pembiayaan'
    kredit1 = combined_data[combined_data[0].str.contains('Sistem Layanan Informasi Keuangan|Jenis Kredit/Pembiayaan', na=False)]
    
    # Split Antar Baris ('\n')
    kredit2 = pd.DataFrame(kredit1)
    lines = kredit2[0].apply(lambda x: x.split('\n')).explode().tolist()
    
    # List Simpan Data
    data_list = []
    
    # Dictionary untuk Kolom
    data_dict = {
        "BANK": None,
        "Baki Debet": None,
        "No Rekening": None,
        "Kualitas": None,
        "Jenis Kredit/Pembiayaan": None,
        "Akad Kredit/Pembiayaan": None,
        "Frekuensi Perpanjangan Kredit/": None,
        "No Akad Awal": None,
        "Tanggal Akad Awal": None,
        "No Akad Akhir": None,
        "Tanggal Akad Akhir": None,
        "Tanggal Awal Kredit": None,
        "Tanggal Mulai": None,
        "Tanggal Jatuh Tempo": None,
        "Kategori Debitur": None,
        "Jenis Penggunaan": None,
        "Sektor Ekonomi": None,
        "Kredit Program Pemerintah": None,
        "Kab/Kota Lokasi Proyek": None,
        "Valuta": None,
        "Suku Bunga/Imbalan": None,
        "Jenis Suku Bunga/Imbalan": None,
        "Keterangan": None,
        "Nama Debitur": None,
        "Nama Group": None
    }
    
    # Variabel untuk Save Nama Deb & Group
    nama_debitur = None
    nama_group = None

    for i, line in enumerate(lines):
        if line.strip() == "Nomor Laporan":  # Hanya baris yang persis "Nomor Laporan"
            if i + 2 >= 0:  # Pastikan indeks tidak negatif
                group_line = lines[i + 2].strip()
                nama_group = " ".join(group_line.split()[:3])
                data_dict["Nama Group"] = nama_group
                
        # Baris == "Penyajian informasi debitur pada Sistem Layanan Informasi"
        if "Penyajian informasi debitur pada Sistem Layanan Informasi" in line:
            # Ambil Baris Ke-4 Setelah Teks Tersebut
            if i + 3 < len(lines):
                nama_line = lines[i + 3].strip()  
                # Ambil 5 Kata Pertama
                nama_debitur = " ".join(nama_line.split()[:5])
                data_dict["Nama Debitur"] = nama_debitur

        # NAMA PELAPOR/BANK
        if " - " in line and "Rp" in line:
            if i > 0 and ("Pelapor Cabang" in lines[i-1]):
                bank_part = line.split("Rp")[0].strip()
                data_dict["BANK"] = bank_part

        # BAKI DEBET
            os_part = line.split("Rp")[1].strip()
            os_value = os_part.split(" ")[0]
            data_dict["Baki Debet"] = os_value

        # NO REKENING & KUALITAS
        elif "No Rekening" in line and "Kualitas" in line:
            nokredit_parts = line.split("No Rekening")[1].strip("Kualitas")[0].strip()
            data_dict["No Rekening"] = nokredit_parts
            kualitas_parts = line.split("Kualitas")[1].strip()
            data_dict["Kualitas"] = kualitas_parts
        
        # SIFAT KREDIT/PEMBIAYAAN & JUMLAH HARI TUNGGAKAN
        elif "Sifat Kredit/Pembiayaan" in line and "Jumlah Hari Tunggakan":
            sifatkre_parts = line.split("Sifat Kredit/Pembiayaan")[1].strip()
            data_dict["Sifat Kredit/Pembiayaan"] = sifatkre_parts

        # JENIS KREDIT/PEMBIAYAAN & NILAI PROYEK
        elif "Jenis Kredit/Pembiayaan" in line and "Nilai Proyek":
            jeniskre_parts = line.split("Jenis Kredit/Pembiayaan")[1].strip()
            data_dict["Jenis Kredit/Pembiayaan"] = jeniskre_parts

        # AKAD KREDIT/PEMBIAYAAN & PLAFON AWAL
        elif "Akad Kredit/Pembiayaan" in line and "Plafon Awal":
            akadkre_parts = line.split("Akad Kredit/Pembiayaan")[1].strip()
            data_dict["Akad Kredit/Pembiayaan"] = akadkre_parts

        # FREKUENSI PERPANJANGAN KREDIT/ & PLAFON
        elif "Frekuensi Perpanjangan Kredit/" in line and "Plafon":
            frekre_parts = line.split("Frekuensi Perpanjangan Kredit/")[1].strip()
            data_dict["Frekuensi Perpanjangan Kredit/"] = frekre_parts

        # NO AKAD AWAL & REALISASI/PENCAIRAN BULAN BERJALAN
        elif "No Akad Awal" in line and "Realisasi/Pencairan Bulan Berjalan":
            noakadawal_parts = line.split("No Akad Awal")[1].strip()
            data_dict["No Akad Awal"] = noakadawal_parts

        # TANGGAL AKAD AWAL & NILAI DALAM MATA UANG ASAL
        elif "Tanggal Akad Awal" in line and "Nilai dalam Mata Uang Asal":
            tglakadawal_parts = line.split("Tanggal Akad Awal")[1].strip()
            data_dict["Tanggal Akad Awal"] = tglakadawal_parts

        # NO AKAD AKHIR & SEBAB MACET
        elif "No Akad Akhir" in line and "Sebab Macet":
            noakadakhir_parts = line.split("No Akad Akhir")[1].strip()
            data_dict["No Akad Akhir"] = noakadakhir_parts

        # TANGGAL AKAD AKHIR & TANGGAL MACET
        elif "Tanggal Akad Akhir" in line and "Tanggal Macet":
            tglakadakhir_parts = line.split("Tanggal Akad Akhir")[1].strip()
            data_dict["Tanggal Akad Akhir"] = tglakadakhir_parts

        # TANGGAL AWAL KREDIT & TUNGGAKAN POKOK
        elif "Tanggal Awal Kredit" in line and "Tunggakan Pokok":
            tglawalkre_parts = line.split("Tanggal Awal Kredit")[1].strip()
            data_dict["Tanggal Awal Kredit"] = tglawalkre_parts

        # TANGGAL MULAI & TUNGGAKAN BUNGA
        elif "Tanggal Mulai" in line and "Tunggakan Bunga":
            tglmulai_parts = line.split("Tanggal Mulai")[1].strip()
            data_dict["Tanggal Mulai"] = tglmulai_parts

        # TANGGAL JATUH TEMPO & FREKUENSI TUNGGAKAN
        elif "Tanggal Jatuh Tempo" in line and "Frekuensi Tunggakan":
            tgljatem_parts = line.split("Tanggal Jatuh Tempo")[1].strip()
            data_dict["Tanggal Jatuh Tempo"] = tgljatem_parts

        # KATEGORI DEBITUR & DENDA
        elif "Kategori Debitur" in line and "Denda":
            katdeb_parts = line.split("Kategori Debitur")[1].strip()
            data_dict["Kategori Debitur"] = katdeb_parts

        # JENIS PENGGUNAAN & FREKUENSI RESTRUKTURISASI
        elif "Jenis Penggunaan" in line and "Frekuensi Restrukturisasi":
            jenispeng_parts = line.split("Jenis Penggunaan")[1].strip()
            data_dict["Jenis Penggunaan"] = jenispeng_parts

        # SEKTOR EKONOMI & TANGGAL RESTRUKTURISASI AKHIR
        elif "Sektor Ekonomi" in line and "Tanggal Restrukturisasi Akhir":
            sekon_parts = line.split("Sektor Ekonomi")[1].strip()
            data_dict["Sektor Ekonomi"] = sekon_parts

        # KREDIT PROGRAM PEMERINTAH & CARA RESTRUKTURISASI
        elif "Kredit Program Pemerintah" in line and "Cara Restrukturisasi":
            kreditprogram_parts = line.split("Kredit Program Pemerintah")[1].strip()
            data_dict["Kredit Program Pemerintah"] = kreditprogram_parts

        # KAB/KOTA LOKASI PROYEK & KONDISI
        elif "Kab/Kota Lokasi Proyek" in line and "Kondisi":
            lokasi_parts = line.split("Kab/Kota Lokasi Proyek")[1].strip()
            data_dict["Kab/Kota Lokasi Proyek"] = lokasi_parts

        # VALUTA & TANGGAL KONDISI
        elif "Valuta" in line and "Tanggal Kondisi":
            valuta_parts = line.split("Valuta")[1].strip()
            data_dict["Valuta"] = valuta_parts

        # SUKU BUNGA/IMBALAN & JENIS BUNGA/IMBALAN
        elif "Suku Bunga/Imbalan" in line and "Jenis Suku Bunga/Imbalan" in line:  
            parts = line.split(" ")  
            data_dict["Suku Bunga/Imbalan"] = parts[2]
            jenisbunga_parts = line.split("Jenis Suku Bunga/Imbalan")[1].strip()
            data_dict["Jenis Suku Bunga/Imbalan"] = jenisbunga_parts

        # KETERANGAN
        elif "Keterangan" in line:
            # Cek apakah baris sebelumnya mengandung Bank Beneficiary
            if i > 0 and ("Jenis Suku Bunga/Imbalan" in lines[i-1]):
                keterangan_parts = line.split("Keterangan")[1].strip()
                data_dict["Keterangan"] = keterangan_parts
                
                # MENYIMPAN DICTIONARY KEDALAM LIST
                data_list.append(data_dict.copy())

    kredit3 = pd.DataFrame(data_list)
    
    if not kredit3.empty:
        # PROSES CLEANING SESUAI DENGAN KODE IPYNB ANDA
        kredit4 = kredit3.copy()
        kredit4 = kredit4.dropna(subset=['Jenis Kredit/Pembiayaan'])
        
        # Split kolom-kolom yang perlu dipisah
        columns_to_split = [
            ('Sifat Kredit/Pembiayaan', 'Jumlah Hari Tunggakan'),
            ('Jenis Kredit/Pembiayaan', 'Nilai Proyek'),
            ('Akad Kredit/Pembiayaan', 'Plafon Awal'),
            ('Frekuensi Perpanjangan Kredit/', 'Plafon'),
            ('No Akad Awal', 'Realisasi/Pencairan Bulan Berjalan'),
            ('Tanggal Akad Awal', 'Nilai dalam Mata Uang Asal'),
            ('No Akad Akhir', 'Sebab Macet'),
            ('Tanggal Akad Akhir', 'Tanggal Macet'),
            ('Tanggal Awal Kredit', 'Tunggakan Pokok'),
            ('Tanggal Mulai', 'Tunggakan Bunga'),
            ('Tanggal Jatuh Tempo', 'Frekuensi Tunggakan'),
            ('Kategori Debitur', 'Denda'),
            ('Jenis Penggunaan', 'Frekuensi Restrukturisasi'),
            ('Sektor Ekonomi', 'Tanggal Restrukturisasi Akhir'),
            ('Kredit Program Pemerintah', 'Cara Restrukturisasi'),
            ('Kab/Kota Lokasi Proyek', 'Kondisi'),
            ('Valuta', 'Tanggal Kondisi')
        ]
        
        for col, split_col in columns_to_split:
            if col in kredit4.columns:
                kredit4[[col, split_col]] = kredit4[col].apply(
                    lambda x: pd.Series(x.split(split_col, 1)) if pd.notna(x) else pd.Series([None, None])
                )
        
        # Cleaning data numerik
        def clean_data(value):
            if pd.isna(value):
                return value
            value = str(value).replace('Rp', '')
            value = value.split(',')[0]
            value = value.replace('.', '')
            return value

        columns_to_clean = ['Baki Debet', 'Nilai Proyek', 'Plafon Awal', 'Plafon', 
                           'Tunggakan Pokok', 'Tunggakan Bunga', 'Denda', 
                           'Realisasi/Pencairan Bulan Berjalan', 'Nilai dalam Mata Uang Asal']
        
        for col in columns_to_clean:
            if col in kredit4.columns:
                kredit4[col] = kredit4[col].apply(clean_data)
        
        # Cleaning tanggal
        month_mapping = {
            'amgoursttiusas': 'Agustus',
            'nmoovretimsabsier': 'November',
            'jmuloi': 'Juli',
            'smepotretimsabsier': 'September',
            'omkotortbisears': 'Oktober',
            'mmeoi': 'Mei',
            'jmanourtaisrai': 'Januari',
            'meborrutiasaris': 'Februari',
            'amporilr': 'April',
            'mmaorerttis': 'Maret',
            'dmeosertmisabseir': 'Desember',
            'jmunoi': 'Juni'
        }

        def clean_first_part(text):
            return re.sub(r'[^\d]', '', str(text))

        def clean_third_part(text):
            return re.sub(r'[^\d]', '', str(text))

        def correct_second_part(text):
            text_lower = str(text).lower()
            for wrong, correct in month_mapping.items():
                if wrong in text_lower:
                    return correct
            return text

        def process_date_string(date_str):
            if pd.isna(date_str):
                return date_str
            parts = str(date_str).split()
            if len(parts) != 3:
                return date_str 
            
            first_part_cleaned = clean_first_part(parts[0])
            second_part_corrected = correct_second_part(parts[1])
            third_part_cleaned = clean_third_part(parts[2])
            
            return f"{first_part_cleaned} {second_part_corrected} {third_part_cleaned}"

        date_columns = ['Tanggal Akad Awal', 'Tanggal Akad Akhir', 'Tanggal Awal Kredit',
                       'Tanggal Mulai', 'Tanggal Jatuh Tempo', 'Tanggal Macet',
                       'Tanggal Restrukturisasi Akhir', 'Tanggal Kondisi']

        for col in date_columns:
            if col in kredit4.columns:
                kredit4[col] = kredit4[col].apply(process_date_string)

        # Convert bulan ke bahasa Inggris
        def convert_month_name(month_id):
            bulan_mapping = {
                'Januari': 'January',
                'Februari': 'February',
                'Maret': 'March',
                'April': 'April',
                'Mei': 'May',
                'Juni': 'June',
                'Juli': 'July',
                'Agustus': 'August',
                'September': 'September',
                'Oktober': 'October',
                'November': 'November',
                'Desember': 'December'
            }
            return bulan_mapping.get(month_id, month_id)

        def fix_date(date_str):
            try:
                if pd.isna(date_str) or not isinstance(date_str, str):
                    return date_str
                    
                parts = date_str.split()
                if len(parts) == 3:
                    day, month_id, year = parts
                    month_en = convert_month_name(month_id)
                    return f"{day} {month_en} {year}"
                return date_str
            except:
                return date_str

        for col in date_columns:
            if col in kredit4.columns:
                kredit4[col] = kredit4[col].apply(fix_date)

        # Convert to datetime
        def clean_and_convert_to_date(date_str):
            if not date_str:
                return None
            try:
                return pd.to_datetime(date_str, format='%d %B %Y', errors='coerce')
            except:
                return None

        for col in date_columns:
            if col in kredit4.columns:
                kredit4[col] = kredit4[col].apply(clean_and_convert_to_date)

        # Filter data
        if 'Keterangan' in kredit4.columns:
            kredit4 = kredit4[~kredit4['Keterangan'].str.contains('Tgl Penilaian Penilai Independen|Garansi|L/C', na=False, case=False)]

        # Cleaning persentase
        def clean_percentage(value):
            if pd.isnull(value):
                return value 
            value = str(value).replace('%', '').strip()
            try:
                return float(value) / 100
            except ValueError:
                return None

        if 'Suku Bunga/Imbalan' in kredit4.columns:
            kredit4['Suku Bunga/Imbalan'] = kredit4['Suku Bunga/Imbalan'].apply(clean_percentage)

        # Standardisasi nilai
        if 'Jenis Kredit/Pembiayaan' in kredit4.columns:
            kredit4['Jenis Kredit/Pembiayaan'] = kredit4['Jenis Kredit/Pembiayaan'].str.strip().replace(
                {
                    'Kredit atau Pembiayaan untuk': 'Kredit atau Pembiayaan untuk Pembayaran Bersama (Sindikasi)',
                    'Kartu Kredit atau Kartu Pembiayaan': 'Kartu Kredit atau Kartu Pembiayaan Syariah',
                    'Kredit atau Pembiayaan kepada Pihak': 'Kredit atau Pembiayaan kepada Pihak Ketiga Melalui Lembaga Lain Secara Channeling',
                    'Kredit atau Pembiayaan kepada Non-UMKM': 'Kredit atau Pembiayaan kepada Non-UMKM melalui Lembaga Lain Secara Executing',
                    'Kredit atau Pembiayaan kepada UMKM': 'Kredit atau Pembiayaan kepada UMKM Melalui Lembaga Lain Secara Executing',
                    'Kredit/ Pembiayaan Kepada Non-UMKM': 'Kredit atau Pembiayaan kepada Non-UMKM melalui Lembaga Lain Secara Executing',
                    'Kredit/Pembiayaan Dalam Rangka': 'Kredit atau Pembiayaan Dalam Rangka Pembiayaan Bersama (Sindikasi)',
                },
                regex=False)

        if 'Kategori Debitur' in kredit4.columns:
            kredit4['Kategori Debitur'] = kredit4['Kategori Debitur'].str.strip().replace(
                {'Bukan Debitur Usaha Mikro, Kecil, dan': 'Bukan Debitur Usaha Mikro, Kecil, dan Menengah'},
                regex=False)

        if 'Sektor Ekonomi' in kredit4.columns:
            kredit4['Sektor Ekonomi'] = kredit4['Sektor Ekonomi'].str.strip().replace({
                'Industri Rokok dan Produk Tembakau': 'Industri Rokok dan Produk Tembakau Lainnya',
                'Industri Penggilingan Beras dan Jagung': 'Industri Penggilingan Beras dan Jagung dan Industri Tepung Beras dan Jagung',
                'Industri Penggilingan Padi dan': 'Industri Penggilingan Padi dan Penyosohan Beras',
                'Perdagangan Besar Mesin-mesin, Suku': 'Perdagangan Besar Mesin-mesin, Suku Cadang dan Perlengkapannya',
                'Perdagangan Eceran Mesin-mesin': 'Perdagangan Eceran Mesin-mesin (Kecuali Mobil dan Sepeda Motor) dan Suku Cadang, termasuk Alat-alat Tranportasi',
                'Perdagangan Besar Mesin, Peralatan': 'Perdagangan Besar Mesin, Peralatan dan Perlengkapannya',
                'Perdagangan Impor Suku Cadang': 'Perdagangan Impor Suku Cadang Mesin-mesin, Suku Cadang dan Perlengkapan Lain',
                'Rumah Tangga Untuk Pemilikan Mobil': 'Rumah Tangga Untuk Pemilikian Mobil Roda Empat',
            }, regex=False)

        if 'Kredit Program Pemerintah' in kredit4.columns:
            kredit4['Kredit Program Pemerintah'] = kredit4['Kredit Program Pemerintah'].str.strip().replace({
                'Kredit yang bukan merupakan kredit/': 'Kredit yang bukan merupakan kredit/pembiayaan dalam rangka program pemerintah'
            }, regex=False)

        # Proses data bank
        try:
            # Coba baca file kode bank
            kodebank_path = "ZZ Data Kode Bank.xlsx"
            if os.path.exists(kodebank_path):
                kodebank = pd.read_excel(kodebank_path)
                
                # Function untuk memisahkan bank dan cabang
                def pisahkan_bank_cabang(row, kodebank):
                    if pd.isna(row):
                        return pd.Series([None, None])
                    for keterangan in kodebank['KETERANGAN']:
                        if keterangan in str(row):
                            cabang = str(row).replace(keterangan, '').strip()
                            return pd.Series([keterangan, cabang])
                    return pd.Series([row, '']) 

                if 'BANK' in kredit4.columns:
                    kredit4[['BANK', 'CABANG']] = kredit4['BANK'].apply(lambda x: pisahkan_bank_cabang(x, kodebank))

                # Split Kode Bank & Nama Bank
                if 'BANK' in kredit4.columns:
                    kredit4[['KODE BANK', 'NAMA BANK']] = kredit4['BANK'].str.split(' - ', n=1, expand=True)

                # Mapping nama bank
                if 'KETERANGAN' in kodebank.columns and 'NAMA BANK' in kodebank.columns:
                    mapping_dict = dict(zip(kodebank['KETERANGAN'], kodebank['NAMA BANK']))
                    kredit4['BANK'] = kredit4['BANK'].map(mapping_dict)
                    
        except Exception as e:
            st.warning(f"Tidak dapat memproses data kode bank: {str(e)}")

        # Tambahkan kategori dan rename kolom
        kredit4['Kategori'] = 'Kredit/Pembiayaan'
        
        kredit4 = kredit4.rename(columns={
            "Baki Debet": "Baki Debet/Nominal",
            "No Rekening": "No Rek/LC/Surat",
            "Jenis Kredit/Pembiayaan": "Jenis Kredit/LC/Garansi/Surat/Fasilitas",
            "Tanggal Mulai": "Tanggal Mulai/Terbit",
            "Tanggal Macet": "Tanggal Macet/Wanprestasi",
            "Jenis Penggunaan": "Tujuan/Jenis Penggunaan",
        })
        
        return kredit4
    else:
        return pd.DataFrame()

def clean_dataframe(df):
    """Membersihkan dataframe dengan menghapus spasi berlebihan"""
    if not df.empty:
        return df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
    return df

def main():
    if uploaded_files:
        st.markdown('<div class="sub-header">📊 File yang Diupload</div>', unsafe_allow_html=True)
        st.write(f"Jumlah file PDF: {len(uploaded_files)}")
        
        for i, file in enumerate(uploaded_files):
            st.write(f"{i+1}. {file.name}")
        
        # Tombol untuk memulai proses
        if st.button("🚀 Mulai Konversi ke Excel", type="primary"):
            with st.spinner("Sedang memproses file..."):
                # Step 1: Konversi PDF ke JSON
                st.markdown('<div class="sub-header">📝 Langkah 1: Konversi PDF ke JSON</div>', unsafe_allow_html=True)
                
                json_files = []
                temp_dir = tempfile.mkdtemp()
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for i, uploaded_file in enumerate(uploaded_files):
                    status_text.text(f"Memproses {uploaded_file.name}...")
                    
                    # Simpan file PDF sementara
                    pdf_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(pdf_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # Konversi ke JSON
                    json_filename = os.path.splitext(uploaded_file.name)[0] + ".json"
                    json_path = os.path.join(temp_dir, json_filename)
                    
                    pdf_to_json(pdf_path, json_path)
                    json_files.append(json_path)
                    
                    progress_bar.progress((i + 1) / len(uploaded_files))
                
                status_text.text("✅ Semua file PDF berhasil dikonversi ke JSON")
                st.markdown('<div class="success-box">✅ Konversi PDF ke JSON selesai!</div>', unsafe_allow_html=True)
                
                # Step 2: Baca file JSON
                st.markdown('<div class="sub-header">📖 Langkah 2: Membaca File JSON</div>', unsafe_allow_html=True)
                
                combined_data = read_json_files(json_files)
                
                if combined_data is not None:
                    st.markdown('<div class="success-box">✅ Pembacaan file JSON berhasil!</div>', unsafe_allow_html=True)
                    st.write(f"Total data yang digabungkan: {len(combined_data)} baris")
                    
                    # Step 3: Proses data kredit
                    st.markdown('<div class="sub-header">🔧 Langkah 3: Memproses Data Kredit</div>', unsafe_allow_html=True)
                    
                    kredit_data = process_kredit_data(combined_data)
                    
                    if not kredit_data.empty:
                        st.markdown('<div class="success-box">✅ Pemrosesan data kredit selesai!</div>', unsafe_allow_html=True)
                        
                        # Step 4: Cleaning final data
                        st.markdown('<div class="sub-header">✨ Langkah 4: Cleaning Data Final</div>', unsafe_allow_html=True)
                        
                        final_data = clean_dataframe(kredit_data)
                        
                        # Tampilkan preview data
                        st.markdown('<div class="sub-header">👁️ Preview Data</div>', unsafe_allow_html=True)
                        st.dataframe(final_data.head(10))
                        
                        st.write(f"Shape data: {final_data.shape}")
                        
                        # Step 5: Download Excel
                        st.markdown('<div class="sub-header">💾 Langkah 5: Download File Excel</div>', unsafe_allow_html=True)
                        
                        # Simpan ke Excel
                        excel_buffer = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
                        final_data.to_excel(excel_buffer.name, index=False, engine='openpyxl')
                        
                        # Baca file Excel untuk download
                        with open(excel_buffer.name, "rb") as f:
                            excel_data = f.read()
                        
                        # Tombol download
                        st.download_button(
                            label="📥 Download Excel File",
                            data=excel_data,
                            file_name="SLIK_Report_Converted.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        st.markdown('<div class="success-box">✅ Proses selesai! File Excel siap diunduh.</div>', unsafe_allow_html=True)
                        
                        # Bersihkan file temporary
                        os.unlink(excel_buffer.name)
                    else:
                        st.error("❌ Gagal memproses data kredit. Data tidak ditemukan.")
                else:
                    st.error("❌ Gagal membaca file JSON.")
                
                # Bersihkan directory temporary
                for json_file in json_files:
                    if os.path.exists(json_file):
                        os.unlink(json_file)
                if os.path.exists(temp_dir):
                    os.rmdir(temp_dir)
    else:
        st.markdown('<div class="info-box">📋 Silakan upload file PDF SLIK Report di sidebar</div>', unsafe_allow_html=True)
        
        # Informasi tentang aplikasi
        st.markdown("""
        ### ℹ️ Tentang Aplikasi
        
        Aplikasi ini mengkonversi file PDF SLIK Report menjadi format Excel dengan tahapan:
        
        1. **Read PDF** - Membaca file PDF yang diupload
        2. **Convert to JSON** - Mengkonversi konten PDF ke format JSON
        3. **Read JSON File** - Membaca dan menggabungkan file JSON
        4. **Convert to Excel** - Memproses data dan menghasilkan file Excel
        
        ### 📋 Format Output Excel
        
        File Excel yang dihasilkan akan berisi data dengan kolom-kolom berikut:
        - BANK
        - Baki Debet/Nominal
        - No Rek/LC/Surat
        - Kualitas
        - Jenis Kredit/LC/Garansi/Surat/Fasilitas
        - Dan kolom-kolom lainnya sesuai format SLIK Report
        """)

if __name__ == "__main__":

    main()

