import os
import re
import time
import configparser
from datetime import datetime
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

def load_config():
    config = configparser.ConfigParser()
    config.read('config.conf')
    return config

def standardize_code(code):
    if pd.isna(code):
        return "" 
    if isinstance(code, float) and code.is_integer():
        code = int(code)    
    s = str(code).strip().upper()
    s = re.sub(r'\s*-\s*', '-', s)
    s = re.sub(r'^(SL|YY|MKS|MGL|PW|PWT|PA|KDI)\s+(\d+)', r'\1-\2', s)
    s = re.sub(r'(\d+)([A-Z])$', r'\1 \2', s)
    return s

def format_idr(value):
    try:
        val_float = float(value)
        return f"{val_float:,.0f}".replace(",", ".")
    except (ValueError, TypeError):
        return str(value)

def parse_order_date(date_str):
    try:
        dt = pd.to_datetime(date_str, errors='coerce')
        if pd.notna(dt):
            return dt.strftime('%d/%m/%Y')
    except Exception:
        pass
    return str(date_str)

def format_excel_date(date_val):
    if pd.isna(date_val):
        return ""
    if isinstance(date_val, datetime) or hasattr(date_val, 'strftime'):
        return date_val.strftime('%d %b %Y')
    return str(date_val)

def read_excel_auto_header(file_path, sheet_name, target_column):
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    target_clean = str(target_column).strip().upper()
    
    for idx, row in df_raw.iterrows():
        row_cleaned = [str(val).strip().upper() for val in row.dropna()]
        if target_clean in row_cleaned:
            df_clean = df_raw.iloc[idx + 1:].copy()
            df_clean.columns = df_raw.iloc[idx].astype(str).str.strip()
            return df_clean.reset_index(drop=True)
            
    raise KeyError(f"Kolom target '{target_column}' tidak ditemukan di baris manapun pada file '{file_path}' (Sheet: {sheet_name})")

def run_ar_process():
    print(f"--> [{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] Memulai sinkronisasi data AR...")
    
    config = load_config()
    
    ow_sheet_name = config.get('OWCOLKEY', 'ow_excel_sheet', fallback='SOLO')
    ow_col_name = config.get('OWCOLKEY', 'ow_excel_col', fallback='Nomor Invoice')
    
    avg_sheet_name = config.get('AVG', 'avg_excel_sheet', fallback='Solo')
    avg_key_col = config.get('AVG', 'avg_excel_key', fallback='NO. PELANGGAN')
    
    ar_url = config.get('AR', 'ar_url')
    ar_sheet_name = config.get('AR', 'ar_sheet')
    ar_key_col_name = config.get('AR', 'ar_key_col')
    ar_prod_key_col_name = config.get('AR', 'ar_prod_key_col').strip()
    ar_target_col_name = config.get('AR', 'ar_target_col')
    
    flag_fraud = config.get('AR', 'ar_data_fraud', fallback='No')
    flag_codecus = config.get('AR', 'ar_data_codecus', fallback='Ya')
    flag_namecus = config.get('AR', 'ar_data_namecus', fallback='Ya')
    fallback_prod = config.get('AR', 'ar_data_prod', fallback='PCMO')
    flag_dt_order = config.get('AR', 'ar_data_dt_order', fallback='Ya')
    flag_calc = config.get('AR', 'ar_data_calc', fallback='Ya')
    
    flag_avg_age = config.get('AR', 'ar_data_avg_age', fallback='No')
    flag_avg_val = config.get('AR', 'ar_data_avg_val', fallback='No')
    flag_avg_inv = config.get('AR', 'ar_data_avg_inv', fallback='No')
    flag_avg_plaf = config.get('AR', 'ar_data_avg_plaf', fallback='Ya')
    flag_avg_pay = config.get('AR', 'ar_data_avg_pay', fallback='Ya')
    flag_avg_his = config.get('AR', 'ar_data_avg_his', fallback='Ya')
    
    flag_inv_numb = config.get('AR', 'ar_data_inv_numb', fallback='Ya')
    flag_inv_dt = config.get('AR', 'ar_data_inv_dt', fallback='No')
    flag_inv_due = config.get('AR', 'ar_data_inv_due', fallback='No')
    flag_inv_val = config.get('AR', 'ar_data_inv_val', fallback='Ya')
    flag_inv_ar = config.get('AR', 'ar_data_inv_ar', fallback='Ya')
    flag_owing = config.get('AR', 'ar_data_owing', fallback='Ya')
    flag_giro = config.get('AR', 'ar_data_giro', fallback='Ya')
    flag_age = config.get('AR', 'ar_data_age', fallback='Ya')

    if not os.path.exists('ARClean_temp.xlsx') or not os.path.exists('Owing_temp.xlsx') or not os.path.exists('Avg_temp.xlsx'):
        print("--> Kesalahan: Pastikan ARClean_temp.xlsx, Owing_temp.xlsx, dan Avg_temp.xlsx ada di folder yang sama!")
        return

    try:
        df_ar_clean = read_excel_auto_header('ARClean_temp.xlsx', sheet_name=0, target_column='Kode Pelanggan')
        df_owing = read_excel_auto_header('Owing_temp.xlsx', sheet_name=ow_sheet_name, target_column=ow_col_name)
        df_avg = read_excel_auto_header('Avg_temp.xlsx', sheet_name=avg_sheet_name, target_column=avg_key_col)
    except KeyError as ke:
        print(f"--> Terjadi kegagalan deteksi struktur kolom tabel Excel: {ke}")
        return

    df_ar_clean['Clean_Kode'] = df_ar_clean['Kode Pelanggan'].apply(standardize_code)
    df_avg['Clean_Kode'] = df_avg[avg_key_col].apply(standardize_code)
    
    owing_set = set(df_owing[ow_col_name].dropna().astype(str).str.strip())

    if flag_fraud == 'No' and 'Nama Penjual' in df_ar_clean.columns:
        df_ar_clean = df_ar_clean[~df_ar_clean['Nama Penjual'].astype(str).str.contains('FRAUD', case=False, na=False)]

    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file('credentials.json', scopes=scope)
    client = gspread.authorize(creds)
    
    ss = client.open_by_url(ar_url)
    wks = ss.worksheet(ar_sheet_name)
    sheet_id = wks.id
    
    all_rows = wks.get_all_values()
    if not all_rows:
        print("--> Sheet kosong.")
        return
        
    header = all_rows[0]
    
    try:
        key_col_idx = header.index(ar_key_col_name)
        target_col_idx = header.index(ar_target_col_name)
    except ValueError as e:
        print(f"--> Kesalahan nama kolom di Google Sheets tidak ditemukan: {e}")
        return
        
    prod_col_idx = header.index(ar_prod_key_col_name) if ar_prod_key_col_name in header else None

    requests = []
    current_date = datetime.now().date()

    for row_idx, row in enumerate(all_rows[1:], start=2):
        if key_col_idx >= len(row):
            continue
            
        if target_col_idx < len(row) and row[target_col_idx].strip() != "":
            continue

        raw_key = row[key_col_idx]
        std_key = standardize_code(raw_key)
        
        if not std_key:
            continue

        raw_order_dt = row[0] if len(row) > 0 else ""
        formatted_order_dt = parse_order_date(raw_order_dt)
        
        product_val = fallback_prod
        if prod_col_idx is not None and prod_col_idx < len(row) and row[prod_col_idx].strip():
            product_val = row[prod_col_idx].strip()

        user_ar_rows = df_ar_clean[df_ar_clean['Clean_Kode'] == std_key]
        
        total_sisa_piutang = 0
        if not user_ar_rows.empty and 'Sisa Piutang' in user_ar_rows.columns:
            total_sisa_piutang = user_ar_rows['Sisa Piutang'].sum()
            
        formatted_piutang_val = format_idr(total_sisa_piutang)

        user_avg_row = df_avg[df_avg['Clean_Kode'] == std_key]
        
        avg_age = "#N/A"
        avg_val = "#N/A"
        avg_inv = "#N/A"
        plafon = "#N/A"
        avg_pay = "#N/A"
        avg_his = "#N/A"

        if not user_avg_row.empty:
            idx_row = user_avg_row.iloc[0]
            if flag_avg_age == 'Ya': avg_age = str(idx_row.get(config.get('AVG', 'avg_excel_age'), '#N/A'))
            if flag_avg_val == 'Ya': avg_val = format_idr(idx_row.get(config.get('AVG', 'avg_excel_val'), '#N/A'))
            if flag_avg_inv == 'Ya': avg_inv = str(idx_row.get(config.get('AVG', 'avg_excel_inv'), '#N/A'))
            if flag_avg_plaf == 'Ya': plafon = format_idr(idx_row.get(config.get('AVG', 'avg_excel_plaf'), '#N/A'))
            if flag_avg_pay == 'Ya': avg_pay = format_idr(idx_row.get(config.get('AVG', 'avg_excel_pay'), '#N/A'))
            if flag_avg_his == 'Ya': 
                his_key = config.get('AVG', 'avg_excel_his', fallback=config.get('AVG', 'avg_excel_history', fallback='AVG HISTORY BAYAR (HARI)'))
                his_val = idx_row.get(his_key, '#N/A')
                avg_his = f"{his_val} HR" if his_val != '#N/A' else '#N/A'

        note_lines = []
        
        header_line_parts = []
        if flag_codecus == 'Ya': header_line_parts.append(std_key)
        if flag_namecus == 'Ya' and not user_ar_rows.empty:
            header_line_parts.append(str(user_ar_rows.iloc[0].get('Nama Pelanggan', '')))
        header_line_parts.append(product_val)
        note_lines.append("\t".join([p for p in header_line_parts if p]))
        
        if flag_dt_order == 'Ya':
            note_lines.append(formatted_order_dt)
        else:
            note_lines.append("")
            
        note_lines.append("")

        if flag_calc == 'Ya': note_lines.append(f"Piutang\t {formatted_piutang_val} ")
        if flag_avg_age == 'Ya': note_lines.append(f"Avg Umur Piutang {avg_age}")
        if flag_avg_val == 'Ya': note_lines.append(f"Avg Nilai Faktur {avg_val}")
        if flag_avg_inv == 'Ya': note_lines.append(f"Avg Jumlah Faktur {avg_inv}")
        if flag_inv_val == 'Ya': note_lines.append(f"Inv\t {len(user_ar_rows)} ")
        if flag_avg_plaf == 'Ya': note_lines.append(f"Plafon\t {plafon} ")
        if flag_avg_pay == 'Ya': note_lines.append(f"Rata-Rata Bayar\t {avg_pay} ")
        if flag_avg_his == 'Ya': note_lines.append(f"Rata-Rata History Bayar\t {avg_his.replace(' HR', '')}\tHari")

        for _, inv_row in user_ar_rows.iterrows():
            inv_part = []
            
            if flag_inv_numb == 'Ya':
                inv_part.append(str(inv_row.get('No. Faktur', '')))
            
            if flag_inv_dt == 'Ya':
                inv_part.append(format_excel_date(inv_row.get('Tgl Faktur')))
            
            if flag_inv_due == 'Ya':
                inv_part.append(format_excel_date(inv_row.get('Jatuh Tempo')))
                
            if flag_inv_ar == 'Ya':
                inv_part.append(format_idr(inv_row.get('Sisa Piutang', 0)))
                
            if flag_age == 'Ya' and pd.notna(inv_row.get('Tgl Faktur')):
                try:
                    tgl_faktur_dt = pd.to_datetime(inv_row['Tgl Faktur']).date()
                    selisih_hari = (current_date - tgl_faktur_dt).days
                    inv_part.append(f"{selisih_hari}\tHR")
                except Exception:
                    inv_part.append("-\tHR")

            line_str = "\t".join([str(x) for x in inv_part if x != ""])
            
            if flag_owing == 'Ya':
                no_faktur_str = str(inv_row.get('No. Faktur', '')).strip()
                if no_faktur_str in owing_set:
                    line_str += " (OWING)"
                    
            if flag_giro == 'Ya' and 'Tanggal JT' in inv_row.index and pd.notna(inv_row['Tanggal JT']):
                tgl_jt_giro = format_excel_date(inv_row['Tanggal JT'])
                if tgl_jt_giro.strip():
                    line_str += f" (JT {tgl_jt_giro})"
                    
            note_lines.append(line_str)

        final_note_text = "\n".join(note_lines)
        
        req_item = {
            "updateCells": {
                "rows": [{
                    "values": [{
                        "userEnteredValue": {"stringValue": formatted_piutang_val if flag_calc == 'Ya' else ""},
                        "note": final_note_text
                    }]
                }],
                "fields": "userEnteredValue,note",
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_idx - 1,
                    "endRowIndex": row_idx,
                    "startColumnIndex": target_col_idx,
                    "endColumnIndex": target_col_idx + 1
                }
            }
        }
        requests.append(req_item)

    if requests:
        body = {"requests": requests}
        ss.batch_update(body)
        print(f"--> Berhasil memperbarui {len(requests)} baris data kosong di Google Sheet!")
    else:
        print("--> Tidak ada data target kosong baru yang perlu diperbarui.")

if __name__ == "__main__":
    while True:
        try:
            config_load = load_config()
            interval_menit = int(config_load.get('AR', 'ar_time_interval', fallback=15))
        except Exception:
            interval_menit = 15
            
        try:
            run_ar_process()
        except Exception as err:
            print(f"--> Terjadi error runtime saat proses berjalan: {err}")
            
        print(f"--> Menunggu interval selama {interval_menit} menit berikutnya...\n")
        time.sleep(interval_menit * 60)