import pandas as pd

df = pd.read_excel('ARClean_temp.xlsx')

mapping_bulan = {
    'Jan': 1,  'Feb': 2,  'Peb': 2,  'Mar': 3,  'Apr': 4,  'Mei': 5,  'Jun': 6,
    'Jul': 7,  'Agu': 8,  'Ags': 8,  'Sep': 9,  'Okt': 10, 'Nov': 11, 'Nop': 11, 'Des': 12,
    
    'jan': 1,  'feb': 2,  'mar': 3,  'apr': 4,  'mei': 5,  'jun': 6,
    'jul': 7,  'agu': 8,  'ags': 8,  'sep': 9,  'okt': 10, 'nov': 11, 'nop': 11, 'des': 12
}

def ubah_format(nilai):
    if pd.isna(nilai):
        return nilai
    dt = pd.to_datetime(nilai)
    return f"{dt.day} {mapping_bulan[dt.month]} {dt.year}"

df['Tgl Faktur'] = df['Tgl Faktur'].apply(ubah_format)
df['Jatuh Tempo'] = df['Jatuh Tempo'].apply(ubah_format)

df.to_excel('ARClean_temp.xlsx', index=False)

print("--> Format tanggal pada kolom Tgl Faktur dan Jatuh Tempo berhasil diubah.")
