import pandas as pd

df_ar = pd.read_excel('ARClean_temp.xlsx')
df_fb = pd.read_excel('FallbackCash_temp.xlsx')

df_ar.columns = df_ar.columns.astype(str).str.strip()
df_fb.columns = df_fb.columns.astype(str).str.strip()

col_ar_kode = [c for c in df_ar.columns if 'Kode Pelanggan' in c][0]
col_fb_no = [c for c in df_fb.columns if 'No. Pelanggan' in c][0]
col_fb_nama = [c for c in df_fb.columns if 'Nama Pelanggan' in c][0]
col_fb_kontak = [c for c in df_fb.columns if 'Nama' in c and 'kontak' in c.lower()][0]

def standarkan_kode(df, nama_kolom):
    return df[nama_kolom].fillna('').astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

df_ar['key_temp'] = standarkan_kode(df_ar, col_ar_kode)
df_fb['key_temp'] = standarkan_kode(df_fb, col_fb_no)

df_missing = df_fb[~df_fb['key_temp'].isin(df_ar['key_temp'])].copy()

if not df_missing.empty:
    df_baru = pd.DataFrame(columns=df_ar.columns)
    df_baru['Kode Pelanggan'] = df_missing[col_fb_no]
    df_baru['Nama Pelanggan'] = df_missing[col_fb_nama]
    df_baru['Nama Kontak'] = df_missing[col_fb_kontak]
    
    df_ar = df_ar.drop(columns=['key_temp'])
    df_baru = df_baru.drop(columns=['key_temp'], errors='ignore')
    
    df_hasil = pd.concat([df_ar, df_baru], ignore_index=True)
else:
    df_hasil = df_ar.drop(columns=['key_temp'])

df_hasil.to_excel('ARClean_temp.xlsx', index=False)

print("--> Data dari FallbackCash_temp.xlsx berhasil dibandingkan dan ditambahkan ke ARClean_temp.xlsx")