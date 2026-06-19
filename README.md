# 📑 Lembar Penagihan Harian

> **Otomasi laporan piutang per penagih dari ekspor Accurate, lengkap dengan cetak siap-TTD dan sinkronisasi ke Google Sheets**

Skrip Python berbasis pipeline yang membaca ekspor daftar piutang dari **Accurate** (`ExportFile.xls`), mengelompokkan tagihan per **penagih/sales** berdasarkan mapping kode pelanggan, lalu menghasilkan dua output sekaligus:
* **lembar penagihan siap cetak** per penagih dalam format Excel resmi dengan kolom tanda tangan, dan
*  **rekap otomatis** yang disisipkan langsung ke Google Sheets untuk monitoring tim.

---

## 📋 Daftar Isi

- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder](#-struktur-folder)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Pipeline (Step-by-Step)](#-alur-pipeline-step-by-step)
- [Konfigurasi `piutang.conf`](#-konfigurasi-piutangconf)
- [Setup Google Sheets API](#-setup-google-sheets-api)
- [Output](#-output)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## ✨ Fitur Utama

- **Deteksi header otomatis** — Membaca file ekspor Accurate dengan posisi header yang bisa berubah-ubah (scan hingga 150 baris pertama), tanpa perlu mapping kolom manual.
- **Parser angka cerdas** — Menangani format angka campuran (titik sebagai pemisah ribuan vs desimal, koma sebagai desimal) secara otomatis.
- **Pengelompokan per penagih** — Mengelompokkan piutang berdasarkan kode pelanggan sesuai daftar penagih di `piutang.conf`, lengkap dengan subtotal per penagih.
- **Template Excel dinamis** — Menyalin gaya, border, merge cell, lebar kolom, dan formula dari `TEMPLATE.xlsx` untuk setiap kelompok penagih, sehingga hasil cetak konsisten secara visual.
- **Auto-fit & styling cetak** — Lebar kolom dan tinggi baris menyesuaikan otomatis, termasuk baris tanda tangan yang diberi tinggi khusus.
- **Sinkronisasi langsung ke Google Sheets** — Data hasil akhir disisipkan otomatis ke spreadsheet tim menggunakan Google Service Account, tanpa copy-paste manual.
- **Auto-cleanup** — Semua file sementara (`*temp.xlsx`, `ExportFile.xls`, `Print_AR.xlsx`) dibersihkan otomatis setelah proses selesai.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl xlsxwriter xlrd gspread google-auth
```

| Library | Kegunaan |
|---|---|
| `pandas` | Baca, bersihkan, dan kelompokkan data piutang |
| `openpyxl` | Baca/tulis `.xlsx`, copy style/formula dari template |
| `xlsxwriter` | Buat file Excel sementara dengan formatting angka |
| `xlrd` | Baca file legacy `.xls` dari Accurate |
| `gspread` | Klien Python untuk Google Sheets API |
| `google-auth` | Autentikasi via Service Account (`google.oauth2.service_account`) |

### Akun Google
Dibutuhkan **Google Cloud Service Account** dengan akses ke Google Sheets API dan Google Drive API. Lihat bagian [Setup Google Sheets API](#-setup-google-sheets-api) di bawah.

---

## 📁 Struktur Folder

```
📦 Lembar-Penagihan-Harian/
│
├── 📄 Ambil AR.py                    ← File utama. Jalankan ini untuk memulai
├── 📄 ExportFile.xls                 ← [INPUT] Ekspor daftar piutang dari Accurate (letakkan di sini)
│
└── 📁 Dapur/                         ← Folder kerja internal (jangan diubah strukturnya)
    ├── 📄 __init__.py
    ├── 📄 1_CleanerAcc.py            ← Bersihkan & rapikan ExportFile.xls
    ├── 📄 2_FilterAR.py              ← Kelompokkan piutang per penagih + subtotal
    ├── 📄 3_CalculateAR.py           ← Susun ke TEMPLATE.xlsx siap cetak
    ├── 📄 4_HelperCleaningData.py    ← Ratakan data (unmerge + isi nama) untuk upload
    ├── 📄 5_InjectDataToSS.py        ← Sisipkan data ke Google Sheets
    ├── 📄 TEMPLATE.xlsx              ← Template resmi lembar penagihan (jangan dihapus)
    ├── 📄 credentials.json           ← Kredensial Google Service Account (rahasia!)
    └── 📄 piutang.conf               ← Konfigurasi mapping penagih & metadata laporan
```

---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input

1. Export daftar piutang/AR (Account Receivable) dari **Accurate** ke format `.xls`.
2. Simpan file tersebut dengan nama **`ExportFile.xls`** di folder utama proyek (sejajar dengan `Ambil AR.py`).

> File ekspor harus mengandung kolom-kolom berikut (nama persis, urutan bebas): `No. Faktur`, `Tgl Faktur`, `Kode`, `Nama Pelanggan`, `Negara Pelanggan`, `Alamat 1 Pelanggan`, `Kota Pelanggan`, `Jatuh Tempo`, `Nilai Faktur`, `Sisa Piutang`, `Umur JT`, `Telepon Pelanggan`, `Sales`, `Area`.

### Langkah 2 — Atur `piutang.conf`

Buka `Dapur/piutang.conf` dan sesuaikan daftar penagih, kode pelanggan yang ditangani masing-masing, serta metadata laporan (perusahaan, divisi, tanggal, nama penginput). Lihat detail format di bagian [Konfigurasi](#-konfigurasi-piutangconf).

### Langkah 3 — Siapkan kredensial Google Sheets

Ganti isi `Dapur/credentials.json` dengan kredensial Service Account Anda sendiri, dan isi `spreadsheet_id` serta ID worksheet di `Dapur/5_InjectDataToSS.py`. Lihat panduan lengkap di [Setup Google Sheets API](#-setup-google-sheets-api).

### Langkah 4 — Jalankan

```bash
python "Ambil AR.py"
```

atau klik dua kali file tersebut jika sudah ada asosiasi Python di sistem Anda.

### Langkah 5 — Ambil hasil

Setelah proses selesai:
- File **`Print_AR.xlsx`** (lembar penagihan siap cetak, sudah berisi data semua penagih dipisah per blok) akan disalin ke folder utama.
- Data ringkas otomatis tersisip ke Google Sheets yang dikonfigurasi.

---

## 🔄 Alur Pipeline (Step-by-Step)

Pipeline ini dijalankan secara berurutan oleh `Ambil AR.py`:

```
[Mulai]
   │
   ├─── Validasi awal
   │       Cek folder Dapur/ ada
   │       Cek semua file syarat ada (5 skrip + credentials.json + piutang.conf)
   │       Jika ada yang kurang → tampilkan pesan & berhenti
   │
   ├─── Pembersihan sisa proses sebelumnya
   │       Hapus *temp.xlsx dan ExportFile.xls lama dari folder Dapur/
   │
   ├─── Salin ExportFile.xls → Dapur/
   │       Jika tidak ditemukan di folder utama → berhenti dengan pesan error
   │
   ├─── [1] 1_CleanerAcc.py
   │       Baca ExportFile.xls (header=None, scan hingga baris ke-150)
   │       Deteksi otomatis posisi kolom berdasarkan nama header target
   │       Buang baris kosong/header berulang (mis. "Total", "Halaman")
   │       Parse angka "Nilai Faktur" dan "Sisa Piutang" (format ID/EN)
   │       Simpan: ExportFile_clean_temp.xlsx
   │
   ├─── [2] 2_FilterAR.py
   │       Baca piutang.conf → bentuk mapping {Kode Pelanggan: Nama Penagih}
   │       Filter baris yang kode pelanggannya ada di mapping
   │       Hitung kolom "Terbayar" = Nilai Faktur − Sisa Piutang
   │       Kelompokkan per Penagih → urutkan per Nama Pelanggan, Kode, No. Faktur
   │       Sisipkan baris "TOTAL [Nama Penagih]" + baris pemisah kosong antar grup
   │       Simpan: Laporan_Piutang_Penagih_temp.xlsx
   │
   ├─── [3] 3_CalculateAR.py
   │       Baca piutang.conf → ambil metadata (PERUSAHAAN, DIVISI, TANGGAL, INPUT)
   │       Baca TEMPLATE.xlsx sebagai master style/formula
   │       Untuk setiap kelompok Penagih:
   │         • Salin blok header (baris 1–4) dari template, isi nama penagih & metadata
   │         • Salin baris data (baris 5) untuk tiap faktur, isi No, Kode, Nama, dst.
   │         • Salin baris total (baris 6), isi formula SUM otomatis
   │         • Salin footer (baris 7+, termasuk kolom TTD Sales & Collector)
   │       Auto-fit lebar kolom & tinggi baris (baris TTD diberi tinggi 50)
   │       Simpan: Print_AR.xlsx
   │
   ├─── Salin *AR.xlsx → folder utama
   │       (Print_AR.xlsx disalin ke folder utama sebagai hasil cetak)
   │
   ├─── [4] 4_HelperCleaningData.py
   │       Baca Print_AR.xlsx, unmerge semua sel yang tergabung
   │       Deteksi baris "Nama [Penagih]" → isi kolom A tiap baris data dengan nama tsb
   │       Hapus baris non-data: judul laporan, header kolom, baris TOTAL, baris TTD
   │       Ratakan tinggi baris (18.75) untuk tampilan rapi di spreadsheet
   │       Simpan: Print_AR_temp.xlsx
   │
   ├─── [5] 5_InjectDataToSS.py
   │       Baca piutang.conf → ambil metadata laporan
   │       Baca Print_AR_temp.xlsx → susun setiap baris menjadi 14 kolom:
   │         Perusahaan, Penagih, Divisi, Tanggal, Input, No, Kode, Nama,
   │         Umur JT, No.Faktur, Tgl Faktur, Nilai Faktur, Terbayar, Sisa Piutang
   │       Autentikasi ke Google Sheets via credentials.json (Service Account)
   │       Buka spreadsheet & worksheet target berdasarkan ID
   │       Sisipkan semua baris baru tepat sebelum baris terakhir yang ada
   │
   └─── Pembersihan akhir
           Hapus semua *temp.xlsx, ExportFile.xls, dan Print_AR.xlsx dari Dapur/
           Selesai ✅
```

---

## ⚙️ Konfigurasi `piutang.conf`

File konfigurasi berbasis blok `[HEADER]` yang dibaca berurutan dari atas ke bawah. Ada dua jenis blok yang bisa diulang berpasangan, dan beberapa blok metadata tunggal di akhir.

```ini
[NAMA SALES]
MGL.RATNO

[KODE PELANGGAN]
MGL-3055
MGL-3222
MGL-3115

[NAMA SALES]
MGL.SPV DANI

[KODE PELANGGAN]
MGL-3438
MGL-3386

[PERUSAHAAN]
PTM

[DIVISI]
PCMO

[TANGGAL]
8/4/2026

[INPUT]
Indah
```

### Aturan penulisan

| Blok | Kegunaan | Boleh diulang? |
|---|---|---|
| `[NAMA SALES]` | Nama penagih/sales (satu baris setelahnya) | ✅ Ya, berpasangan dengan `[KODE PELANGGAN]` berikutnya |
| `[KODE PELANGGAN]` | Daftar kode pelanggan yang ditangani penagih di atasnya (satu kode per baris) | ✅ Ya |
| `[PERUSAHAAN]` | Nama perusahaan, dicetak di header lembar penagihan | ❌ Hanya sekali |
| `[DIVISI]` | Nama divisi, dicetak di header lembar penagihan | ❌ Hanya sekali |
| `[TANGGAL]` | Tanggal laporan, dicetak di header | ❌ Hanya sekali |
| `[INPUT]` | Nama penginput data, dicetak di header | ❌ Hanya sekali |

> ⚠️ **Penting:** Setiap blok `[NAMA SALES]` **harus langsung diikuti** oleh blok `[KODE PELANGGAN]` setelahnya. Urutan ini menentukan pemetaan — kode pelanggan akan dikaitkan dengan nama sales yang ditulis paling terakhir sebelum blok kode tersebut.
>
> Kode pelanggan yang tidak terdaftar di `piutang.conf` akan **otomatis diabaikan** (tidak masuk laporan).

---

## 🔑 Setup Google Sheets API

### 1. Buat Service Account

1. Buka [Google Cloud Console](https://console.cloud.google.com/) → buat project baru (atau gunakan yang sudah ada).
2. Aktifkan **Google Sheets API** dan **Google Drive API** untuk project tersebut.
3. Buka menu **IAM & Admin → Service Accounts** → buat Service Account baru.
4. Pada tab **Keys**, buat key baru dengan tipe **JSON** → file akan otomatis terunduh.

### 2. Pasang kredensial

Ganti isi file `Dapur/credentials.json` dengan isi file JSON yang baru diunduh. Format yang dibutuhkan:

```json
{
  "type": "service_account",
  "project_id": "...",
  "private_key_id": "...",
  "private_key": "-----BEGIN PRIVATE KEY-----...",
  "client_email": "nama-akun@nama-project.iam.gserviceaccount.com",
  "client_id": "...",
  "auth_uri": "https://accounts.google.com/o/oauth2/auth",
  "token_uri": "https://oauth2.googleapis.com/token",
  "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
  "client_x509_cert_url": "...",
  "universe_domain": "googleapis.com"
}
```

### 3. Beri akses ke spreadsheet target

Buka Google Sheets yang akan dijadikan tujuan data, klik **Share/Bagikan**, lalu tambahkan alamat email `client_email` dari `credentials.json` sebagai **Editor**.

### 4. Isi ID Spreadsheet dan Worksheet

Buka `Dapur/5_InjectDataToSS.py`, lalu ganti dua nilai berikut:

```python
# Sebelum
spreadsheet_id = 'YOUR SPREADSHEETS ID'
...
worksheet = sh.get_worksheet_by_id(YOUR ID SPREADSHEETS)

# Sesudah (contoh)
spreadsheet_id = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms'
...
worksheet = sh.get_worksheet_by_id(0)
```

> **Cara mendapatkan Spreadsheet ID:** lihat URL Google Sheets: `https://docs.google.com/spreadsheets/d/**[ID DI SINI]**/edit`
>
> **Cara mendapatkan Worksheet ID (gid):** lihat di URL setelah membuka tab/sheet tertentu: `...edit#gid=**[ID DI SINI]**`

---

## 📤 Output

### 1. `Print_AR.xlsx` — Lembar penagihan siap cetak

Disalin ke folder utama setelah proses selesai. Berisi satu blok terpisah per penagih, masing-masing dengan:
- Header laporan (nama perusahaan, divisi, tanggal, nama penagih, nama penginput)
- Tabel rincian faktur: No, Kode, Nama Pelanggan, Umur JT, No. Faktur, Tgl Faktur, Nilai Faktur, Terbayar, Sisa Piutang
- Baris **TOTAL TAGIHAN** dengan formula SUM otomatis
- Kolom tanda tangan **TTD Sales & Collector** di footer

> File ini **dihapus dari folder `Dapur/`** di akhir proses, namun **salinannya tetap ada di folder utama** untuk dicetak.

### 2. Baris baru di Google Sheets

Setiap baris faktur (per penagih, tanpa baris total/header) disisipkan sebagai satu baris baru dengan 14 kolom:

| Kolom | Isi |
|---|---|
| Perusahaan | Dari `piutang.conf` |
| Penagih | Nama sales/penagih |
| Divisi | Dari `piutang.conf` |
| Tanggal | Dari `piutang.conf` |
| Input | Nama penginput, dari `piutang.conf` |
| No. | Nomor urut dalam grup penagih |
| Kode | Kode pelanggan |
| Nama | Nama pelanggan |
| Umur JT | Umur jatuh tempo |
| No. Faktur | Nomor faktur |
| Tgl Faktur | Tanggal faktur |
| Nilai Faktur | Nilai faktur asli |
| Terbayar | Nilai Faktur − Sisa Piutang |
| Sisa Piutang | Sisa piutang saat ini |

Data disisipkan tepat sebelum baris terakhir yang sudah ada di worksheet (mempertahankan baris terbawah, misalnya baris total atau footer permanen).

---

## 🛠️ Troubleshooting

### ❌ `File ExportFile.xls tidak ditemukan untuk diproses`
Pastikan file ekspor Accurate ada di folder utama (sejajar dengan `Ambil AR.py`) dengan nama persis `ExportFile.xls`.

### ❌ `Error: Kolom No. Faktur tidak ditemukan` (dari `1_CleanerAcc.py`)
Skrip mencari header `"No. Faktur"` dalam 150 baris pertama. Jika ekspor Accurate Anda memiliki struktur berbeda (header di luar 150 baris, atau nama kolom berbeda), sesuaikan nilai `target_headers` di `Dapur/1_CleanerAcc.py`.

### ❌ Laporan kosong / penagih tidak muncul di hasil
Periksa apakah kode pelanggan di data Accurate **sama persis** dengan yang terdaftar di blok `[KODE PELANGGAN]` pada `piutang.conf` (termasuk huruf besar/kecil dan tanda hubung). Kode yang tidak cocok akan otomatis diabaikan oleh `2_FilterAR.py`.

### ❌ `Error: Spreadsheets tidak ditemukan. Pastikan email Service Account sudah diberi akses Editor`
Buka Google Sheets target → Share → tambahkan email `client_email` dari `credentials.json` sebagai Editor.

### ❌ `Error: Worksheet dengan ID tersebut tidak ditemukan`
Periksa kembali nilai `worksheet_id` (gid) yang diisi di `5_InjectDataToSS.py`. Pastikan worksheet/tab dengan ID tersebut benar-benar ada di spreadsheet.

### ❌ Error autentikasi Google (`DefaultCredentialsError` / `invalid_grant`)
Periksa isi `credentials.json` — pastikan file JSON valid (bukan placeholder) dan `private_key` tersalin lengkap termasuk baris `-----BEGIN PRIVATE KEY-----` dan `-----END PRIVATE KEY-----`.

### ❌ Format angka Nilai Faktur/Sisa Piutang salah (misal: 1.234.567 terbaca sebagai 1.234)
Fungsi `parse_to_float()` di `1_CleanerAcc.py` mendeteksi format angka berdasarkan posisi titik/koma terakhir. Jika ekspor Accurate Anda menggunakan format angka yang tidak umum, periksa dan sesuaikan logika parsing di fungsi tersebut.

---

## 📌 Catatan Penting

- **Jangan ubah struktur folder `Dapur/`** — semua skrip menggunakan path relatif dan saling bergantung pada urutan eksekusi.
- **File `TEMPLATE.xlsx` wajib ada** dan tidak boleh dihapus atau diubah strukturnya — bila ingin mengubah tata letak cetak, edit langsung di sini (baris 1–4 = header, baris 5 = baris data, baris 6 = baris total, baris 7+ = footer/TTD).
- **`credentials.json` bersifat rahasia** — jangan pernah commit file ini dengan isi asli ke repository publik. Gunakan `.gitignore` atau simpan sebagai GitHub Secret jika dipakai dalam CI/CD.
- **Jalankan dari folder utama** — bukan dari dalam folder `Dapur/`. Gunakan `Ambil AR.py` sebagai satu-satunya titik masuk.
- **Penyisipan ke Google Sheets bersifat tambahan (insert), bukan replace** — setiap kali dijalankan, baris baru akan terus ditambahkan. Pastikan ekspektasi tim terhadap riwayat data di spreadsheet sudah jelas.
- **Selalu cek ulang hasil cetak** sebelum dibagikan ke penagih — terutama formula TOTAL TAGIHAN dan kelengkapan data per grup.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
