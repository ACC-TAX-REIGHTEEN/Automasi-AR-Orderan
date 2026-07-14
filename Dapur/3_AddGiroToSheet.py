import pandas as pd
import os
from collections import defaultdict
import datetime

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

print("--> Proses pembaruan kolom tanggal JT")

giro_file = 'Giro_temp.xlsx'
target_file = 'ARClean_temp.xlsx'

if not os.path.exists(giro_file):
    print(f"--> ERROR: File '{giro_file}' tidak ditemukan!")
    exit()
if not os.path.exists(target_file):
    print(f"--> ERROR: File '{target_file}' tidak ditemukan!")
    exit()

indo_months = {
    'Jan': 1,  'Feb': 2,  'Peb': 2,  'Mar': 3,  'Apr': 4,  'Mei': 5,  'Jun': 6,
    'Jul': 7,  'Agu': 8,  'Ags': 8,  'Sep': 9,  'Okt': 10, 'Nov': 11, 'Nop': 11, 'Des': 12,
    
    'jan': 1,  'feb': 2,  'mar': 3,  'apr': 4,  'mei': 5,  'jun': 6,
    'jul': 7,  'agu': 8,  'ags': 8,  'sep': 9,  'okt': 10, 'nov': 11, 'nop': 11, 'des': 12
}

def clean_invoice_str(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        return s[:-2]
    return s

print(f"--> Membaca data referensi dari {giro_file}...")
df_giro = pd.read_excel(giro_file)

giro_groups = defaultdict(list)
for _, row in df_giro.iterrows():
    no_faktur_so = clean_invoice_str(row['No. Faktur. (SO)'])
    tgl_cek = row['Tgl Cek']
    if no_faktur_so and not pd.isna(tgl_cek):
        giro_groups[no_faktur_so].append(tgl_cek)

mapping_jt = {}
for no_faktur, tgl_list in giro_groups.items():
    parsed_dates = []
        
    for tgl in tgl_list:
        if isinstance(tgl, (pd.Timestamp, datetime.datetime)):
            d_int, m_int, y_int = tgl.day, tgl.month, tgl.year
            d_str = str(d_int).zfill(2)
            m_short = str(m_int).zfill(2)
            y_short = str(y_int)[-2:]
            parsed_dates.append((y_int, m_int, d_int, d_str, m_short, y_short))
        else:
            parts = str(tgl).strip().split()
            if len(parts) == 3:
                try:
                    d_int = int(parts[0])
                    d_str = parts[0].zfill(2)
                    m_name = parts[1]
                    m_int = indo_months.get(m_name, 1)
                    m_short = str(m_int).zfill(2)
                    y_int = int(parts[2])
                    y_short = parts[2][-2:]
                    parsed_dates.append((y_int, m_int, d_int, d_str, m_short, y_short))
                except ValueError:
                    continue
                        
    if not parsed_dates:
        continue
            
    parsed_dates.sort(key=lambda x: (x[0], x[1], x[2]))
        
    date_by_month_year = defaultdict(list)
    for item in parsed_dates:
        y_int, m_int, d_int, d_str, m_short, y_short = item
        group_key = (y_int, m_int, m_short, y_short)
        if d_str not in date_by_month_year[group_key]:
            date_by_month_year[group_key].append(d_str)
                
    group_strings = []
    for key in sorted(date_by_month_year.keys()):
        y_int, m_int, m_short, y_short = key
        days_str = ",".join(date_by_month_year[key])
        group_strings.append(f"{days_str}/{m_short}/{y_short}")
            
    mapping_jt[no_faktur] = "JT " + " & ".join(group_strings)

print(f"--> Membaca file target {target_file}...")
try:
    df_target = pd.read_excel(target_file)
except Exception as e:
    print(f"--> ERROR: Gagal membaca file target: {e}")
    exit()

print("--> Mencocokkan nomor faktur dan menyusun nilai baru...")
hasil_kolom_jt = []
for fktr in df_target['No. Faktur']:
    fktr_cleaned = clean_invoice_str(fktr)
    nilai_jt = mapping_jt.get(fktr_cleaned, "")
    hasil_kolom_jt.append(nilai_jt)

df_target['Tanggal JT'] = hasil_kolom_jt

print(f"--> Menyimpan hasil perubahan ke {target_file}...")
try:
    df_target.to_excel(target_file, index=False)
    print("--> Proses berhasil! Kolom 'Tanggal JT' berhasil ditambahkan di paling kanan dan diperbarui.")
except Exception as e:
    print(f"--> Terjadi error saat menyimpan file: {e}")
