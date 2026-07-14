import pandas as pd
import os
import time

print("--> Proses pembersihan data (cleaning)")

file_path = 'Piutang.xls'

if not os.path.exists(file_path):
    print(f"--> File '{file_path}' tidak ditemukan. Pastikan file ada.")
    exit()

def load_dataset(path):
    try:
        return pd.read_excel(path, header=3)
    except:
        return None

df = load_dataset(file_path)

if df is None:
    print("--> Gagal membaca file Piutang.xls.")
    exit()

target_indices = [2, 3, 5, 9, 11, 14, 16, 18, 20, 22]
df_clean = df.iloc[:, target_indices].copy()

new_columns = [
    'Kode Pelanggan', 
    'No. Faktur',     
    'Tgl Faktur',     
    'Jatuh Tempo',    
    'Nilai Faktur',   
    'Sisa Piutang',   
    'Umur JT',        
    'Nama Pelanggan', 
    'Nama Penjual',   
    'Nama Kontak'    
]
df_clean.columns = new_columns

df_clean['Kode Pelanggan'] = df_clean['Kode Pelanggan'].ffill()
df_clean = df_clean.dropna(subset=['No. Faktur'])

def format_clean(val):
    if pd.isna(val):
        return ""
    s = str(val)
    if s.endswith('.0'):
        return s[:-2]
    if s.endswith(',00'):
        return s[:-3]
    return s

cols_to_clean = ['Kode Pelanggan', 'Nilai Faktur', 'Sisa Piutang']
for col in cols_to_clean:
    df_clean[col] = df_clean[col].apply(format_clean)

indo_months_in = {
    'Jan': 'Jan', 'Feb': 'Feb', 'Mar': 'Mar', 'Apr': 'Apr', 'Mei': 'May', 'Jun': 'Jun',
    'Jul': 'Jul', 'Agu': 'Aug', 'Sep': 'Sep', 'Okt': 'Oct', 'Nop': 'Nov', 'Des': 'Dec',
    'Peb': 'Feb', 'Ags': 'Aug', 
    
    'jan': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'apr': 'Apr', 'mei': 'May', 'jun': 'Jun',
    'jul': 'Jul', 'agu': 'Aug', 'ags': 'Aug', 'sep': 'Sep', 'okt': 'Oct', 'nop': 'Nov', 
    'nov': 'Nov', 'des': 'Dec'
}

def clean_and_format_date(x):
    if pd.isna(x) or str(x).strip() == '':
        return None
    date_str = str(x)
    for indo, eng in indo_months_in.items():
        if indo in date_str:
            date_str = date_str.replace(indo, eng)
            break
    try:
        return pd.to_datetime(date_str, dayfirst=True, errors='coerce')
    except:
        return None

df_clean['Tgl Faktur'] = df_clean['Tgl Faktur'].apply(clean_and_format_date)
df_clean['Jatuh Tempo'] = df_clean['Jatuh Tempo'].apply(clean_and_format_date)

if 'Tgl Faktur' in df_clean.columns:
    idx_tgl = df_clean.columns.get_loc('Tgl Faktur')
    df_clean.insert(idx_tgl + 1, 'SS', '')

if 'Jatuh Tempo' in df_clean.columns:
    idx_jt = df_clean.columns.get_loc('Jatuh Tempo')
    df_clean.insert(idx_jt + 1, 'SS', '', allow_duplicates=True)

df_clean.reset_index(drop=True, inplace=True)

output_filename = "ARClean_temp.xlsx"

print("--> Proses menyimpan data ke ARClean_temp.xlsx")
try:
    df_clean.to_excel(output_filename, index=False)
    print(f"--> Data bersih siap ({len(df_clean)} baris) dan disimpan di {output_filename}.")
except Exception as e:
    print(f"--> Terjadi error saat menyimpan: {e}")

print("--> Menghapus file asli")
time.sleep(1)
try:
    if os.path.exists(file_path):
        os.remove(file_path)
        print(f"--> Berhasil menghapus: {file_path}")
except Exception as e:
    print(f"--> Gagal menghapus {file_path}: {e}")

print("--> Semua proses selesai!")
