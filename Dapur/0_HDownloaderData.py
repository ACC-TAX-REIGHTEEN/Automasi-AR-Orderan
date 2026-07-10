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

    if owing_str and aravg_str:
        print("--> Kedua URL terisi, melewati proses unduhan")
        return

    downloads = {
        'Owing_temp.xlsx': 'https://github.com/ACC-TAX-REIGHTEEN/.xlsx',
        'Avg_temp.xlsx': 'https://github.com/ACC-TAX-REIGHTEEN/.xlsx'
    }

    for output_name, url in downloads.items():
        download_file(url, output_name)

if __name__ == "__main__":
    main()
