import pandas as pd

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