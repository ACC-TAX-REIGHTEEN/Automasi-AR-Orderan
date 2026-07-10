import pandas as pd
import os

config_file = 'config.conf'
run_giro = False

if os.path.exists(config_file):
    with open(config_file, 'r') as f:
        lines = [line.strip() for line in f.readlines()]
    if '[GIRO]' in lines:
        idx = lines.index('[GIRO]')
        if idx + 1 < len(lines):
            line_val = lines[idx + 1].replace(' ', '').lower()
            if line_val == 'giro_stats=ya':
                run_giro = True

if not run_giro:
    print("--> Proses GIRO di-skip berdasarkan config.conf.")
    exit()

df = pd.read_excel('Giro.xls', header=None)

data_bersih = []
for i in range(len(df)):
    baris = df.iloc[i].dropna().tolist()
    if len(baris) >= 9:
        data_bersih.append(baris)

kolom = [
    'No. Pelanggan', 'Nama Pelanggan', 'Tgl Faktur', 
    'No. Faktur. (SO)', 'No. Form', 'Total Diterima', 
    'Nilai terima', 'Nama Bank', 'Tgl Cek'
]

hasil_df = pd.DataFrame(data_bersih, columns=kolom)

hasil_df['Total Diterima'] = hasil_df['Total Diterima'].astype(str).str.replace(',', '.', regex=False)
hasil_df['Total Diterima'] = pd.to_numeric(hasil_df['Total Diterima'], errors='coerce')
hasil_df['Nilai terima'] = hasil_df['Nilai terima'].astype(str).str.replace(',', '.', regex=False)
hasil_df['Nilai terima'] = pd.to_numeric(hasil_df['Nilai terima'], errors='coerce')
hasil_df.to_excel('Giro_temp.xlsx', index=False)

print("--> File Giro_temp.xlsx telah berhasil dibuat dengan format angka pada kolom Total Diterima.")
