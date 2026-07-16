import configparser
import requests
import os

def download_sheet_from_gdrive(url, output_name):
    
    if '/d/' in url:
        base_id = url.split('/d/')[1].split('/')[0]
        download_url = f"https://docs.google.com/spreadsheets/d/{base_id}/export?format=xlsx"
    else:
        print(f"--> Gagal: URL untuk {output_name} bukan format Google Sheets yang valid!")
        return

    try:
        print(f"--> Sedang mengunduh {output_name} dari Google Sheets Cadangan...")
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(download_url, timeout=15, headers=headers)
        response.raise_for_status()
        with open(output_name, 'wb') as f:
            f.write(response.content)
        print(f"--> Berhasil! File disimpan sebagai {output_name}")
    except requests.exceptions.Timeout:
        print(f"--> Gagal: Koneksi terlalu lama (Timeout) saat mengunduh {output_name}")
    except Exception as e:
        print(f"--> Gagal mengunduh {output_name}: {e}")

def main():
    config_file = 'config.conf'
    if not os.path.exists(config_file):
        print("--> File config.conf tidak ditemukan!")
        return

    config = configparser.ConfigParser(allow_no_value=True)
    config.read(config_file)

    owing_url = config.get('OWING', 'url', fallback='').strip() if config.has_section('OWING') else ''
    aravg_url = config.get('ARAVG', 'url', fallback='').strip() if config.has_section('ARAVG') else ''

    fback_ow_url = config.get('FBACKOW', 'url', fallback='').strip() if config.has_section('FBACKOW') else ''
    fback_aravg_url = config.get('FBACKARAVG', 'url', fallback='').strip() if config.has_section('FBACKARAVG') else ''

    print("--- Memulai Pengecekan URL Konfigurasi (Lokal & Google Sheets Cadangan) ---")

    if not owing_url:
        print("--> [OWING] URL Utama kosong.")
        if fback_ow_url:
            download_sheet_from_gdrive(fback_ow_url, 'Owing_temp.xlsx')
        else:
            print("--> Peringatan: URL Utama [OWING] dan Cadangan [FBACKOW] KOSONG! File Owing_temp.xlsx tidak dapat diperbarui.")
    else:
        print("--> [OWING] URL Utama Terisi. Mengabaikan unduhan cadangan.")

    if not aravg_url:
        print("--> [ARAVG] URL Utama kosong.")
        if fback_aravg_url:
            download_sheet_from_gdrive(fback_aravg_url, 'Avg_temp.xlsx')
        else:
            print("--> Peringatan: URL Utama [ARAVG] dan Cadangan [FBACKARAVG] KOSONG! File Avg_temp.xlsx tidak dapat diperbarui.")
    else:
        print("--> [ARAVG] URL Utama Terisi. Mengabaikan unduhan cadangan.")

if __name__ == "__main__":
    main()