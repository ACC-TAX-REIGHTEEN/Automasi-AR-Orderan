import configparser
from datetime import datetime
import pandas as pd

config = configparser.ConfigParser()
config.read('config.conf')

giro_stats = config.get('GIRO', 'giro_stats', fallback='No')

if giro_stats.strip() == 'Ya':
    print("--> Menjalankan program: Status GIRO aktif.")

    indo_months_in = {
    'Jan': 'Jan', 'Feb': 'Feb', 'Mar': 'Mar', 'Apr': 'Apr', 'Mei': 'May', 'Jun': 'Jun',
    'Jul': 'Jul', 'Agu': 'Aug', 'Sep': 'Sep', 'Okt': 'Oct', 'Nop': 'Nov', 'Des': 'Dec',
    'Peb': 'Feb', 'Ags': 'Aug', 
    'jan': 'Jan', 'feb': 'Feb', 'mar': 'Mar', 'apr': 'Apr', 'mei': 'May', 'jun': 'Jun',
    'jul': 'Jul', 'agu': 'Aug', 'ags': 'Aug', 'sep': 'Sep', 'okt': 'Oct', 'nop': 'Nov', 
    'nov': 'Nov', 'des': 'Dec'
    }

    def parse_indo_date(date_str):
        if pd.isna(date_str):
            return pd.NaT
        parts = str(date_str).split()
        if len(parts) == 3:
            day, month, year = parts
            en_month = indo_months_in.get(month, month)
            return datetime.strptime(f'{day} {en_month} {year}', '%d %b %Y')
        return pd.NaT

    df = pd.read_excel('Giro_temp.xlsx')
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    df['temp_date'] = df['Tgl Cek'].apply(parse_indo_date)
    df = df[df['temp_date'] >= today]
    df = df.drop(columns=['temp_date'])

    df.to_excel('Giro_temp.xlsx', index=False)
    print(
        "--> Data berhasil diperbarui. Baris data dari hari kemarin mundur telah dihapus."
    )

else:
    print(
        "--> Program di-skip. Status 'giro_stats' di config.conf bernilai 'No' atau tidak aktif."
    )