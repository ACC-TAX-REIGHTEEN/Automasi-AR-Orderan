import configparser
import requests
import os

def download_file(url, output_name):
    try:
        print(f"--> Sedang mengunduh {output_name}...")
        response = requests.get(url)
        response.raise_for_status()
        with open(output_name, 'wb') as f:
            f.write(response.content)
        print(f"--> Berhasil! File disimpan sebagai {output_name}")
    except Exception as e:
        print(f"--> Gagal mengunduh {output_name}: {e}")

def main():
    config_file = 'config.conf'
    if not os.path.exists(config_file):
        print("--> File config.conf tidak ditemukan!")
        return

    config = configparser.ConfigParser(allow_no_value=True)
    config.read(config_file)

    owing_url = ''
    aravg_url = ''

    if config.has_section('OWING') and config.has_option('OWING', 'url'):
        owing_url = config.get('OWING', 'url')
    if config.has_section('ARAVG') and config.has_option('ARAVG', 'url'):
        aravg_url = config.get('ARAVG', 'url')

    owing_str = str(owing_url or '').strip()
    aravg_str = str(aravg_url or '').strip()

    fallback_downloads = {
        'Owing_temp.xlsx': 'https://github.com/ACC-TAX-REIGHTEEN/Automasi-AR-Orderan/raw/refs/heads/main/Contoh%20Data/Owing_temp.xlsx',
        'Avg_temp.xlsx': 'https://github.com/ACC-TAX-REIGHTEEN/Automasi-AR-Orderan/raw/refs/heads/main/Contoh%20Data/Avg_temp.xlsx'
    }

    print("--- Memulai Pengecekan URL Konfigurasi ---")

    if not owing_str:
        print("--> [OWING] URL Kosong! Mengunduh data cadangan dari GitHub...")
        download_file(fallback_downloads['Owing_temp.xlsx'], 'Owing_temp.xlsx')
    else:
        print("--> [OWING] URL Terisi. Mengabaikan unduhan cadangan Owing.")

    if not aravg_str:
        print("--> [ARAVG] URL Kosong! Mengunduh data cadangan dari GitHub...")
        download_file(fallback_downloads['Avg_temp.xlsx'], 'Avg_temp.xlsx')
    else:
        print("--> [ARAVG] URL Terisi. Mengabaikan unduhan cadangan Avg.")

if __name__ == "__main__":
    main()