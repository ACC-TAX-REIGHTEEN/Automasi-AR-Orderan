import configparser
import datetime
import os
import pandas as pd
import numpy as np

indo_months_in = {
    'Jan': 'Jan', 'Feb': 'Feb', 'Mar': 'Mar', 'Apr': 'Apr', 'Mei': 'May', 'Jun': 'Jun',
    'Jul': 'Jul', 'Agu': 'Aug', 'Sep': 'Sep', 'Okt': 'Oct', 'Nop': 'Nov', 'Des': 'Dec',
    'Peb': 'Feb', 'Ags': 'Aug', 
    
    'jan': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'apr': 'Apr', 'mei': 'May', 'jun': 'Jun',
    'jul': 'Jul', 'agu': 'Aug', 'ags': 'Aug', 'sep': 'Sep', 'okt': 'Oct', 'nop': 'Nov', 
    'nov': 'Nov', 'des': 'Dec'
}

def parse_tgl_faktur(val):
    if pd.isna(val) or str(val).strip().lower() in ['', 'nan', 'none']:
        return np.nan
    
    if isinstance(val, (pd.Timestamp, datetime.datetime, datetime.date)):
        return val.strftime('%Y-%m-%d %H:%M:%S')
    
    val_str = str(val).strip()
    
    parts = val_str.split()
    if len(parts) == 3:
        day, month, year = parts[0], parts[1], parts[2]
        if month in indo_months_in:
            month = indo_months_in[month]
            val_str = f"{day} {month} {year}"
    
    try:
        dt = pd.to_datetime(val_str, dayfirst=True)
        return dt.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        return val_str

def check_config_status():
    config_file = 'config.conf'
    if not os.path.exists(config_file):
        print(f"--> Warning: File '{config_file}' tidak ditemukan. Mengabaikan pengecekan config.")
        return True

    config = configparser.ConfigParser()
    try:
        config.read(config_file)
        if 'ACC' in config and 'data' in config['ACC']:
            acc_data = config['ACC']['data'].strip().upper()
            if acc_data != 'DEPO':
                print(f"--> Status ACC data adalah '{acc_data}'. Seluruh proses Python di-SKIP.")
                return False
            else:
                print("--> Status ACC data = DEPO. Memproses data...")
                return True
        else:
            print("--> Section [ACC] atau key 'data' tidak ditemukan di config.conf.")
            return True
    except Exception as e:
        print(f"--> Error membaca config.conf: {e}")
        return True

def clean_data_autofit(input_file, output_file):
    if not check_config_status():
        return

    print(f"--> Sedang memproses file: {input_file}...")
    
    try:
        df = pd.read_excel(input_file, header=None)
    except Exception as e:
        print(f"--> Error membaca file: {e}")
        return
        
    target_headers = [
        "Kode", 
        "No. Faktur", 
        "Tgl Faktur", 
        "Nilai Faktur", 
        "Sisa Piutang", 
        "Umur JT", 
        "Nama Pelanggan", 
        "Sales", 
        "Negara Pelanggan"
    ]
    
    header_map = {}
    start_row = 0
    
    for i in range(min(150, len(df))):
        for j in range(len(df.columns)):
            val = str(df.iat[i, j]).strip()
            if val in target_headers and val not in header_map:
                header_map[val] = j
                if val == "No. Faktur":
                    start_row = i
                    
    if "No. Faktur" not in header_map:
        print("--> Error: Kolom No. Faktur tidak ditemukan.")
        return

    col_faktur = header_map["No. Faktur"]
    df_data = df.iloc[start_row + 1:].copy()
    
    df_data[col_faktur] = df_data[col_faktur].astype(str).str.strip()
    kondisi_kosong_faktur = df_data[col_faktur].str.lower().isin(['nan', 'none', ''])
    df_data.loc[kondisi_kosong_faktur, col_faktur] = np.nan
    
    df_clean = df_data.dropna(subset=[col_faktur]).copy()
    
    header_labels = ["No. Faktur", "Faktur", "No.", "Total", "Halaman", "Page", "Tanggal"]
    df_clean = df_clean[~df_clean[col_faktur].astype(str).str.contains('|'.join(header_labels), case=False, na=False)]
    
    if "Kode" in header_map:
        col_kode = header_map["Kode"]
        df_clean[col_kode] = df_clean[col_kode].astype(str).str.strip()
        kondisi_kosong_kode = df_clean[col_kode].str.lower().isin(['nan', 'none', ''])
        df_clean.loc[kondisi_kosong_kode, col_kode] = np.nan
        df_clean = df_clean.dropna(subset=[col_kode]).copy()
        
    def get_col_data(header_name):
        if header_name in header_map:
            return df_clean[header_map[header_name]]
        return np.nan

    temp_df = pd.DataFrame({
        "Kode Pelanggan": get_col_data("Kode"),
        "No. Faktur": get_col_data("No. Faktur"),
        "Tgl Faktur": get_col_data("Tgl Faktur"),
        "_SS_1": np.nan,
        "Jatuh Tempo": np.nan,
        "_SS_2": np.nan,
        "Nilai Faktur": get_col_data("Nilai Faktur"),
        "Sisa Piutang": get_col_data("Sisa Piutang"),
        "Umur JT": get_col_data("Umur JT"),
        "Nama Pelanggan": get_col_data("Nama Pelanggan"),
        "Nama Penjual": get_col_data("Sales"),
        "Nama Kontak": get_col_data("Negara Pelanggan")
    })

    temp_df["Tgl Faktur"] = temp_df["Tgl Faktur"].apply(parse_tgl_faktur)

    def parse_to_float(val):
        if pd.isna(val) or str(val).strip() == "":
            return np.nan
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).strip()
        
        if '.' in s and ',' in s:
            if s.rfind(',') > s.rfind('.'):
                s = s.replace('.', '').replace(',', '.')
            else:
                s = s.replace(',', '')
        elif ',' in s:
            if len(s) - s.rfind(',') <= 3:
                s = s.replace(',', '.')
            else:
                s = s.replace(',', '')
        elif '.' in s:
            if s.count('.') > 1 or (len(s) - s.rfind('.') == 4):
                s = s.replace('.', '')
        try:
            return float(s)
        except:
            return np.nan

    cols_to_clean = ['Nilai Faktur', 'Sisa Piutang']
    for col in cols_to_clean:
        if col in temp_df.columns:
            temp_df[col] = temp_df[col].apply(parse_to_float)

    final_headers = [
        "Kode Pelanggan",
        "No. Faktur",
        "Tgl Faktur",
        "SS",
        "Jatuh Tempo",
        "SS",
        "Nilai Faktur",
        "Sisa Piutang",
        "Umur JT",
        "Nama Pelanggan",
        "Nama Penjual",
        "Nama Kontak"
    ]
    temp_df.columns = final_headers
    df_final = temp_df.copy()
    df_final.reset_index(drop=True, inplace=True)
    
    try:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Data Bersih')
            workbook = writer.book
            worksheet = writer.sheets['Data Bersih']
            
            format_angka = workbook.add_format({'num_format': '#,##0.00'})
            
            for i in range(len(df_final.columns)):
                col_name = df_final.columns[i]
                col_data = df_final.iloc[:, i]
                
                panjang_maksimal = col_data.apply(lambda x: len(str(x)) if pd.notna(x) else 0).max() if not df_final.empty else 0
                max_len = max(panjang_maksimal, len(col_name)) + 2
                
                if col_name in cols_to_clean:
                    worksheet.set_column(i, i, max_len, format_angka)
                else:
                    worksheet.set_column(i, i, max_len)
                    
        print(f"--> SUKSES! File tersimpan rapi di: {output_file}")
        
    except Exception as e:
        print(f"--> Error saat menyimpan file: {e}")

    return df_final

input_filename = 'Piutang.xls' 
output_filename = 'ARClean_temp.xlsx'

if __name__ == "__main__":
    clean_data_autofit(input_filename, output_filename)