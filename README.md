# 📑 Automasi AR Orderan

> **Sinkronisasi data piutang AR ke Google Sheets order management secara real-time dan berkelanjutan — dari ekspor Accurate ke cell note terstruktur siap pakai tim admin sales**

Pipeline Python delapan langkah yang membaca ekspor AR dari Accurate (`Piutang.xls`) dan data Giro (`Giro.xls`), mengunduh data referensi dari tiga Google Sheets (Owing, AVG performa pelanggan, FallbackCash), menggabungkan semua sumber, lalu menyuntikkan **nilai total piutang** dan **ringkasan lengkap per pelanggan sebagai cell note** ke kolom target di Google Sheets order tracker — berjalan dalam loop otomatis setiap menit (bawaan per-15 menit) untuk menangkap pesanan baru secara real-time.

---

## 📋 Daftar Isi

- [Gambaran Umum & Arsitektur](#-gambaran-umum--arsitektur)
- [Fitur Utama](#-fitur-utama)
- [Prasyarat](#-prasyarat)
- [Struktur Folder & File](#-struktur-folder--file)
- [Cara Penggunaan](#-cara-penggunaan)
- [Alur Kerja Pipeline](#-alur-kerja-pipeline)
- [Detail Tiap Skrip](#-detail-tiap-skrip)
- [Konfigurasi `config.conf`](#-konfigurasi-configconf)
- [Setup Google Sheets API](#-setup-google-sheets-api)
- [Format Output: Nilai Sel & Cell Note](#-format-output-nilai-sel--cell-note)
- [Mode Demo (Tanpa URL)](#-mode-demo-tanpa-url)
- [Troubleshooting](#-troubleshooting)
- [Catatan Penting](#-catatan-penting)

---

## 🗂️ Gambaran Umum & Arsitektur

Sistem ini dirancang untuk mendukung **tim sales atau admin** yang menggunakan Google Sheets sebagai tracker order harian. Setiap baris di Google Sheets mewakili satu pesanan pelanggan. Saat baris baru dibuat (kolom target masih kosong), skrip secara otomatis mengisi:

- **Nilai sel** → total Sisa Piutang pelanggan dalam format IDR
- **Cell note** → ringkasan lengkap kondisi kredit pelanggan: plafon, rata-rata bayar, riwayat, daftar faktur aktif beserta nilai asli, titip bayar, status OWING, dan tanggal giro

Semua data dikompilasi dari empat sumber:

| Sumber | Format | Isi |
|---|---|---|
| `Piutang.xls` | File lokal | Ekspor AR dari Accurate — faktur aktif per pelanggan |
| `Giro.xls` | File lokal (opsional) | Rekap cek giro — tanggal cair per nomor faktur |
| Google Sheets OWING | Diunduh otomatis | Nomor faktur berstatus OWING (terima tapi belum lunas) |
| Google Sheets AVG | Diunduh otomatis | Performa historis pelanggan (plafon, rata-rata bayar, tiering) |
| Google Sheets FBACK | Diunduh otomatis | Daftar pelanggan cash yang tidak muncul di Piutang.xls |

---

## ✨ Fitur Utama

- **Loop sinkronisasi berkelanjutan** — Skrip 6 berjalan terus-menerus dengan interval yang dapat dikonfigurasi (bawaan per-15 menit), secara otomatis mendeteksi baris baru di Google Sheets dan mengisinya.
- **Only-empty fill** — Hanya mengisi sel yang masih kosong di kolom target; baris yang sudah terisi tidak akan ditimpa.
- **Multi-kode pelanggan per sel** — Satu sel di kolom `ar_key_col` dapat berisi lebih dari satu kode pelanggan (dipisah `&`); skrip menggabungkan data semua kode tersebut.
- **Standardisasi kode pelanggan otomatis** — Menormalisasi berbagai format penulisan kode (SL001, YY 1234, MGL-5678, dll.) ke format kanonik sebelum pencocokan.
- **Cell note terstruktur dengan 20+ flag** — Setiap komponen dalam ringkasan cell note dapat diaktifkan atau dinonaktifkan secara individual via `config.conf`.
- **Filter giro kadaluarsa otomatis** — Entri Giro dengan tanggal cek yang sudah lewat (sebelum hari ini) otomatis dihapus, sehingga hanya jadwal giro mendatang yang ditampilkan di cell note.
- **Nilai faktur asli & titip bayar per faktur** — Setiap baris faktur di cell note dapat menampilkan nilai asli faktur dan jumlah pembayaran yang sudah dititipkan (`Nilai Faktur − Sisa Piutang`).
- **Fallback pelanggan cash** — Pelanggan yang tidak memiliki piutang aktif (tidak muncul di Piutang.xls) tetap dapat ditampilkan profilnya berkat data dari sheet FBACK.
- **Mode demo otomatis** — Jika URL Google Sheets Owing dan AVG belum dikonfigurasi, skrip mengunduh contoh data sampel dari repositori GitHub secara otomatis.
- **Ekslusi FRAUD** — Baris dengan kata `FRAUD` di kolom `Nama Penjual` dapat difilter berdasarkan flag `ar_data_fraud`.
- **Auto-detect header Excel** — Membaca `Owing_temp.xlsx` dan `Avg_temp.xlsx` dengan deteksi posisi header dinamis, tanpa tergantung pada nomor baris yang tetap.

---

## 🔧 Prasyarat

### Python
Python **3.8+** disarankan.

### Library yang dibutuhkan

```bash
pip install pandas openpyxl xlrd requests gspread google-auth
```

| Library | Digunakan di | Kegunaan |
|---|---|---|
| `pandas` | Semua skrip | Baca, transformasi, dan simpan data Excel |
| `openpyxl` | Skrip 1–5 | Engine penulisan `.xlsx` |
| `xlrd` | Skrip 1, 2 | Baca file legacy `.xls` dari Accurate |
| `requests` | Skrip 0, 0H | Unduh file dari Google Sheets & GitHub |
| `gspread` | Skrip 6 | Klien Google Sheets API |
| `google-auth` | Skrip 6 | Autentikasi via Service Account |
| `configparser`, `re`, `datetime`, `collections`, `os`, `time`, `glob`, `shutil`, `subprocess` | Semua | Standard library |

> **Catatan `xlrd`:** Gunakan versi yang kompatibel dengan `.xls` (format Accurate):
> ```bash
> pip install "xlrd>=1.0.0,<2.0.0"
> ```

### Akun Google
Diperlukan **Google Cloud Service Account** dengan akses ke Google Sheets API dan Google Drive API. Lihat bagian [Setup Google Sheets API](#-setup-google-sheets-api).

---

## 📁 Struktur Folder & File

```
📦 Automasi-AR-Orderan/
│
├── 📄 Jalankan Automasi.py              ← Orkestrator utama. Jalankan ini
├── 📄 Piutang.xls                       ← [INPUT] Ekspor AR dari Accurate (wajib)
├── 📄 Giro.xls                          ← [INPUT] Rekap giro/cek masuk (opsional)
├── 📄 Ekspor Data.png                   ← Panduan visual cara ekspor dari Accurate
│
├── 📁 Dapur/                            ← Folder pipeline (jangan diubah)
│   ├── 📄 __init__.py
│   ├── 📄 0_DownloaderData.py           ← Unduh Owing/AVG/FallbackCash dari Google Sheets
│   ├── 📄 0_HDownloaderData.py          ← Unduh data contoh jika URL belum dikonfigurasi
│   ├── 📄 1_CleanerAccAR.py             ← Bersihkan Piutang.xls → ARClean_temp.xlsx
│   ├── 📄 2_CleanerAccGiro.py           ← Bersihkan Giro.xls → Giro_temp.xlsx (jika aktif)
│   ├── 📄 2_HCleanerAccGiroDue.py       ← Filter giro kadaluarsa (hapus Tgl Cek < hari ini)
│   ├── 📄 3_AddGiroToSheet.py           ← Tambahkan Tanggal JT ke ARClean dari Giro
│   ├── 📄 4_PatchFallbackCash.py        ← Tambahkan pelanggan cash dari FallbackCash
│   ├── 📄 5_AdjDateFormat.py            ← Format ulang tanggal ke format Indonesia
│   ├── 📄 6_InjectDataToSS.py           ← Loop sinkronisasi → inject ke Google Sheets
│   ├── 📄 config.conf                   ← Konfigurasi utama (wajib diisi sebelum pakai)
│   └── 📄 credentials.json              ← Kredensial Google Service Account (rahasia!)
│
└── 📁 Contoh Data/                      ← Data sampel untuk mode demo
    ├── 📄 Owing_temp.xlsx
    └── 📄 Avg_temp.xlsx
```


---

## 🚀 Cara Penggunaan

### Langkah 1 — Siapkan file input

Letakkan di folder utama (sejajar dengan `Jalankan Automasi.py`):
- **`Piutang.xls`** — ekspor laporan AR dari Accurate (wajib, nama harus persis)
- **`Giro.xls`** — rekap pembayaran giro (opsional; jika tidak ada, langkah giro di-skip)

> Lihat `Ekspor Data.png` untuk panduan visual cara mengekspor dari Accurate.

### Langkah 2 — Isi `config.conf`

Buka `Dapur/config.conf` dan isi minimal tiga nilai kritis:

```ini
[OWING]
url = https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_OWING/edit

[ARAVG]
url = https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_AVG/edit

[AR]
ar_url = https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_ORDER_TRACKER/edit
ar_sheet = NamaSheetTarget
ar_key_col = KODE PELANGGAN
ar_target_col = Nominal Nota Belum Lunas
```

Lihat panduan lengkap di bagian [Konfigurasi `config.conf`](#-konfigurasi-configconf).

### Langkah 3 — Pasang kredensial Google Service Account

Ganti isi `Dapur/credentials.json` dengan file JSON Service Account Anda. Lihat [Setup Google Sheets API](#-setup-google-sheets-api).

### Langkah 4 — Berikan akses ke semua Google Sheets

Bagikan semua spreadsheet yang dikonfigurasi (OWING, ARAVG, FBACK, AR target) ke alamat `client_email` di `credentials.json` sebagai **Editor**.

### Langkah 5 — Jalankan

```bash
python "Jalankan Automasi.py"
```

### Langkah 6 — Pantau dan hentikan

Skrip 6 berjalan dalam loop tak terbatas. Tekan **`Ctrl+C`** untuk menghentikan sinkronisasi saat diperlukan.

```
--> Memulai eksekusi 0_DownloaderData.py
--> Sedang mengunduh Owing_temp.xlsx...
--> Berhasil! File disimpan sebagai Owing_temp.xlsx
...
--> Memulai eksekusi 2_CleanerAccGiro.py
--> Memulai eksekusi 2_HCleanerAccGiroDue.py
--> Menjalankan program: Status GIRO aktif.
--> Data berhasil diperbarui. Baris data dari hari kemarin mundur telah dihapus.
...
--> Memulai eksekusi 6_InjectDataToSS.py
--> Tekan Ctrl+C untuk menghentikan loop sinkronisasi.
--> [2025-04-01 09:00:00] Memulai sinkronisasi data AR...
--> Berhasil memperbarui 7 baris data kosong di Google Sheet!
--> Menunggu interval selama 15 menit berikutnya...
```

---

## 🔄 Alur Kerja Pipeline

```
[Mulai: Jalankan Automasi.py]
   │
   ├─── Validasi file wajib
   │       Cek Piutang.xls ada → gagal jika tidak ada
   │       Cek folder Dapur/ ada → gagal jika tidak ada
   │       Cek 10 file syarat di Dapur/ ada → gagal jika kurang
   │       (termasuk 2_HCleanerAccGiroDue.py)
   │
   ├─── Bersihkan Dapur/ dari file *.xls/*.xlsx sisa run lama
   │
   ├─── SALIN (bukan pindah) file input → Dapur/
   │       Piutang.xls → Dapur/Piutang.xls (wajib)
   │       Giro.xls → Dapur/Giro.xls (jika ada)
   │
   ├─── [0] 0_DownloaderData.py
   │       Unduh Owing_temp.xlsx dari URL [OWING]
   │       Unduh Avg_temp.xlsx dari URL [ARAVG]
   │       Unduh FallbackCash_temp.xlsx dari URL [FBACK]
   │
   ├─── [0H] 0_HDownloaderData.py
   │       Jika KEDUA URL [OWING] dan [ARAVG] kosong:
   │         → Unduh sampel Owing_temp.xlsx dari GitHub
   │         → Unduh sampel Avg_temp.xlsx dari GitHub
   │       Jika salah satu URL terisi → skip
   │
   ├─── [1] 1_CleanerAccAR.py
   │       Baca Piutang.xls (header baris ke-4)
   │       Seleksi 10 kolom → rename → ffill kode → drop NaN faktur
   │       Bersihkan angka (.0, ,00) → parse tanggal Indonesia
   │       Sisipkan kolom spacer SS → Simpan ARClean_temp.xlsx
   │       Hapus Piutang.xls dari Dapur/
   │
   ├─── [2] 2_CleanerAccGiro.py
   │       Cek [GIRO] giro_stats di config.conf
   │       Jika 'Ya': Bersihkan Giro.xls → Giro_temp.xlsx
   │       Jika 'Tidak': Skip seluruh step ini
   │
   ├─── [2H] 2_HCleanerAccGiroDue.py 
   │       Cek [GIRO] giro_stats di config.conf
   │       Jika 'Ya':
   │         Baca Giro_temp.xlsx
   │         Parse kolom 'Tgl Cek' → konversi ke datetime
   │         Filter: hanya simpan baris dengan Tgl Cek >= hari ini
   │         Hapus baris giro yang sudah kadaluarsa (tanggal cek terlewat)
   │         Timpa Giro_temp.xlsx dengan data yang sudah difilter
   │       Jika 'Tidak': Skip
   │
   ├─── [3] 3_AddGiroToSheet.py
   │       Cek [GIRO] giro_stats → Skip jika tidak aktif
   │       Bangun mapping {No. Faktur → "JT DD/MM/YY & ..."}
   │         (hanya dari giro yang tanggalnya masih mendatang)
   │       Tambah kolom 'Tanggal JT' ke ARClean_temp.xlsx
   │
   ├─── [4] 4_PatchFallbackCash.py
   │       Bandingkan kode di ARClean_temp vs FallbackCash_temp
   │       Kode yang ada di FallbackCash tapi tidak di AR →
   │         tambahkan sebagai baris baru (Kode + Nama + Kontak saja)
   │       Update ARClean_temp.xlsx
   │
   ├─── [5] 5_AdjDateFormat.py
   │       Ubah kolom Tgl Faktur & Jatuh Tempo dari datetime
   │       ke format teks Indonesia: "15 Jan 2025"
   │       Update ARClean_temp.xlsx
   │
   └─── [6] 6_InjectDataToSS.py  ← LOOP ∞
           while True:
             ├─ Baca config → ambil semua flag & URL
             ├─ Baca ARClean_temp + Owing_temp + Avg_temp
             ├─ Autentikasi Google Sheets via credentials.json
             ├─ Buka spreadsheet & sheet target
             ├─ Scan semua baris → cari yang kolom target KOSONG
             ├─ Per baris kosong:
             │    └─ Standarisasi kode pelanggan (support multi-kode &)
             │    └─ Lookup AR data → hitung total Sisa Piutang
             │    └─ Lookup AVG data → ambil plafon, bayar, history, tier
             │    └─ Bangun cell note terstruktur (sesuai flag aktif)
             │    └─ Per baris faktur: tampilkan nilai asli & titip bayar jika aktif
             │    └─ Tambahkan flag OWING dan JT Giro mendatang per faktur
             │    └─ Siapkan batch updateCells request
             ├─ Kirim batch update ke Google Sheets API
             └─ Tunggu interval_menit → ulangi
```

---

## 🔍 Detail Tiap Skrip

### Skrip 0 — `0_DownloaderData.py` (Downloader utama)

Mengunduh tiga file referensi dari Google Sheets ke folder `Dapur/`:

| Seksi config | Output file | Isi |
|---|---|---|
| `[OWING]` | `Owing_temp.xlsx` | Daftar nomor faktur berstatus OWING |
| `[ARAVG]` | `Avg_temp.xlsx` | Data performa historis per pelanggan |
| `[FBACK]` | `FallbackCash_temp.xlsx` | Daftar pelanggan cash tanpa piutang aktif |

URL dikonversi otomatis dari format `https://docs.google.com/spreadsheets/d/ID/edit` ke endpoint export XLSX Google Sheets.

---

### Skrip 0H — `0_HDownloaderData.py` (Downloader helper/demo)

Mengunduh file contoh dari repositori GitHub **hanya jika** kedua URL `[OWING]` dan `[ARAVG]` di `config.conf` masih kosong. Memungkinkan sistem berjalan dalam **mode demo** tanpa konfigurasi Google Sheets terlebih dahulu.

---

### Skrip 1 — `1_CleanerAccAR.py`

**Input:** `Piutang.xls` (header di baris ke-4)

Memilih 10 kolom berdasarkan indeks posisi dan mengubah namanya:

| Indeks asli | Nama baru |
|---|---|
| 2 | `Kode Pelanggan` |
| 3 | `No. Faktur` |
| 5 | `Tgl Faktur` |
| 9 | `Jatuh Tempo` |
| 11 | `Nilai Faktur` |
| 14 | `Sisa Piutang` |
| 16 | `Umur JT` |
| 18 | `Nama Pelanggan` |
| 20 | `Nama Penjual` |
| 22 | `Nama Kontak` |

Dua kolom spacer `SS` disisipkan (setelah `Tgl Faktur` dan setelah `Jatuh Tempo`). Output: `ARClean_temp.xlsx`. File `Piutang.xls` di Dapur/ dihapus setelah berhasil.

---

### Skrip 2 — `2_CleanerAccGiro.py` (kondisional)

Diaktifkan hanya jika `[GIRO] giro_stats = Ya`. Membersihkan `Giro.xls` → `Giro_temp.xlsx` dengan memilih 9 kolom standar dan membersihkan kolom angka.

---

### Skrip 2H — `2_HCleanerAccGiroDue.py`

**Input:** `Giro_temp.xlsx` (hasil Skrip 2)

Skrip baru yang berjalan tepat setelah Skrip 2, sebelum data giro diproses lebih lanjut. Fungsinya adalah **memfilter entri giro kadaluarsa** — hanya mempertahankan baris yang tanggal cairnya (`Tgl Cek`) masih hari ini atau di masa mendatang.

**Logika:**
```
Baca Giro_temp.xlsx
  ├─ Parse kolom 'Tgl Cek' → datetime
  │   (mendukung Timestamp, string format Indo: "15 Jan 2025",
  │    termasuk varian Peb, Ags, dan semua bentuk lowercase)
  ├─ Filter: simpan hanya baris dengan Tgl Cek >= hari ini (jam 00:00:00)
  └─ Timpa Giro_temp.xlsx dengan hasil filter
```

**Dampak:** Jadwal giro yang sudah terlewat tidak akan muncul di kolom `Tanggal JT` dalam cell note Google Sheets. Hanya tanggal cair yang masih berlaku yang ditampilkan, sehingga informasi di cell note selalu relevan dengan kondisi terkini.

**Skip otomatis:** Jika `[GIRO] giro_stats` bukan `Ya`, skrip ini langsung keluar tanpa memproses apapun.

---

### Skrip 3 — `3_AddGiroToSheet.py`

Membaca `Giro_temp.xlsx` (yang kini sudah difilter oleh Skrip 2H) dan membangun peta `{No. Faktur → "JT DD/MM/YY & ..."}`. Kolom `Tanggal JT` ditambahkan ke `ARClean_temp.xlsx`.

---

### Skrip 4 — `4_PatchFallbackCash.py`

Menambahkan baris baru untuk pelanggan cash yang ada di `FallbackCash_temp.xlsx` tapi tidak punya piutang aktif di `ARClean_temp.xlsx`. Kolom yang diisi: `Kode Pelanggan`, `Nama Pelanggan`, `Nama Kontak`.

---

### Skrip 5 — `5_AdjDateFormat.py`

Mengonversi `Tgl Faktur` dan `Jatuh Tempo` dari datetime ke teks Indonesia: `"15 Jan 2025"`, `"28 Feb 2025"`, dst.

---

### Skrip 6 — `6_InjectDataToSS.py` (Loop utama)

Inti sistem. Memindai Google Sheets, menemukan baris dengan kolom target kosong, dan mengisinya dengan nilai + cell note AR terstruktur.

**Standardisasi kode pelanggan:**
```
"MGL1234" → "MGL-1234"
"sl 001"  → "SL-001"
"YY1234A" → "YY-1234 A"
Awalan prefix yang dikenali: SL, YY, MKS, MGL, PW, PWT, PLU, SG, SMG, TGL, PA, KDI
```

**Multi-kode per sel:** `"MGL-001 & MGL-002"` → lookup keduanya, gabungkan hasilnya.

**Kalkulasi titip bayar (`ar_data_inv_pay`):** `Nilai Faktur − Sisa Piutang`. Hanya ditampilkan jika hasilnya > 0 (ada pembayaran sebagian yang sudah dititipkan). Format: `Ttp Byr: 1.234.567`.

**Batch update:** Semua perubahan dalam satu iterasi dikirim sekaligus via `batchUpdate` API untuk efisiensi.

---

## ⚙️ Konfigurasi `config.conf`

### `[OWING]`, `[ARAVG]`, `[FBACK]` — URL sumber data referensi

```ini
[OWING]
url = https://docs.google.com/spreadsheets/d/ID_SPREADSHEET/edit

[ARAVG]
url = https://docs.google.com/spreadsheets/d/ID_SPREADSHEET/edit

[FBACK]
url = https://docs.google.com/spreadsheets/d/ID_SPREADSHEET/edit
```

Kosongkan `url =` untuk skip pengunduhan.

---

### `[GIRO]` — Aktifkan/nonaktifkan fitur Giro

```ini
[GIRO]
giro_stats = Ya    ; Aktifkan: Ya | Nonaktifkan: No
```

Mengontrol tiga skrip sekaligus: `2_CleanerAccGiro.py`, `2_HCleanerAccGiroDue.py`, dan `3_AddGiroToSheet.py`.

---

### `[OWCOLKEY]` — Struktur file Owing

```ini
[OWCOLKEY]
ow_excel_sheet = SOLO        ; Nama sheet di Owing_temp.xlsx
ow_excel_col = Nomor Invoice  ; Nama kolom yang berisi nomor faktur
```

---

### `[AVG]` — Struktur file AVG performa pelanggan

```ini
[AVG]
avg_excel_sheet = Solo
avg_excel_key = NO. PELANGGAN
avg_excel_age = AVG UMUR PIUTANG
avg_excel_val = AVG NILAI FAKTUR
avg_excel_inv = JUMLAH INVOICE
avg_excel_plaf = PLAFON
avg_excel_pay = AVG BAYAR
avg_excel_his = AVG HISTORY BAYAR (HARI)
avg_excel_tier = TIERING
```

---

### `[AR]` — Target Google Sheets & semua flag output

```ini
[AR]
ar_url = https://docs.google.com/spreadsheets/d/ID/edit
ar_time_interval = 15
ar_sheet = Solo
ar_key_col = KODE PELANGGAN
ar_prod_key_col =
ar_target_col = Nominal Nota Belum Lunas
```

#### Flag filter & konteks

| Key | Default | Keterangan |
|---|---|---|
| `ar_data_fraud` | `No` | `No` = filter baris FRAUD; `Ya` = tampilkan semua |
| `ar_data_codecus` | `Ya` | Tampilkan kode pelanggan di header note |
| `ar_data_namecus` | `Ya` | Tampilkan nama pelanggan di header note |
| `ar_data_prod` | `PCMO` | Nama produk/divisi fallback jika kolom kosong |
| `ar_data_dt_order` | `Ya` | Tampilkan tanggal order di note |
| `ar_data_calc` | `Ya` | Isi nilai sel dengan total Sisa Piutang |

#### Flag ringkasan performa (`RINGKASAN PERFORMA PIUTANG`)

| Key | Default | Keterangan |
|---|---|---|
| `ar_data_avg_age` | `No` | Rata-rata umur piutang |
| `ar_data_avg_val` | `No` | Rata-rata nilai faktur |
| `ar_data_avg_inv` | `No` | Rata-rata jumlah invoice |
| `ar_data_avg_plaf` | `Ya` | Plafon kredit pelanggan |
| `ar_data_avg_pay` | `Ya` | Rata-rata pembayaran |
| `ar_data_avg_his` | `Ya` | Histori rata-rata hari bayar |
| `ar_data_avg_tier` | `No` | Tiering pelanggan |

#### Flag daftar faktur (`DAFTAR RINCIAN FAKTUR AKTIF`)

| Key | Default | Keterangan |
|---|---|---|
| `ar_data_inv_numb` | `Ya` | Nomor faktur |
| `ar_data_inv_dt` | `No` | Tanggal faktur |
| `ar_data_inv_due` | `No` | Tanggal jatuh tempo |
| `ar_data_inv_val` | `Ya` | Total jumlah faktur aktif (hitungan baris) |
| `ar_data_inv_orig` | `No` | Nilai nominal faktur asli (`Nilai Faktur`) |
| `ar_data_inv_ar` | `Ya` | Sisa piutang per faktur |
| `ar_data_inv_pay` | `Ya` | Titip bayar per faktur: `Ttp Byr: X` (`Nilai Faktur − Sisa Piutang`). Hanya muncul jika nilainya > 0 |
| `ar_data_owing` | `Ya` | Tandai `(OWING)` di baris faktur |
| `ar_data_giro` | `Ya` | Tampilkan tanggal giro mendatang `(JT ...)` per faktur |
| `ar_data_age` | `Ya` | Hitung umur piutang (hari dari Tgl Faktur hingga hari ini) |


---

## 🔑 Setup Google Sheets API

### 1. Buat Service Account

1. Buka [Google Cloud Console](https://console.cloud.google.com/) → buat atau pilih project.
2. Aktifkan **Google Sheets API** dan **Google Drive API**.
3. Masuk ke **IAM & Admin → Service Accounts** → buat Service Account baru.
4. Di tab **Keys** → buat key baru tipe **JSON** → file terunduh otomatis.

### 2. Pasang kredensial

Ganti isi `Dapur/credentials.json` dengan file JSON yang diunduh:

```json
{
  "type": "service_account",
  "project_id": "...",
  "private_key_id": "...",
  "private_key": "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n",
  "client_email": "nama@project.iam.gserviceaccount.com",
  "client_id": "...",
  ...
}
```

### 3. Berikan akses ke semua spreadsheet

Buka setiap Google Sheets (OWING, ARAVG, FBACK, AR target), klik **Share**, tambahkan `client_email` sebagai **Editor** dan atau agar dapat terlihat oleh semua orang.

### 4. Struktur spreadsheet target (AR order tracker)

- Baris pertama = header kolom
- Kolom `ar_key_col` → kode pelanggan per baris pesanan
- Kolom `ar_target_col` → diisi otomatis (awalnya kosong)
- Kolom A → tanggal order (digunakan jika `ar_data_dt_order = Ya`)

---

## 📤 Format Output: Nilai Sel & Cell Note

### Nilai Sel

Total `Sisa Piutang` dalam format IDR:

```
1.234.567
```

### Cell Note

```
MGL-1234    Toko Makmur Jaya    PCMO
Tanggal Order: 15/04/2025

========================================
RINGKASAN PERFORMA PIUTANG
========================================
Piutang                  :  1.234.567
Plafon Kredibilitas      :  5.000.000
Rata-Rata Bayar          :  3.500.000
Rata-Rata History Bayar  :  28 Hari
Total Faktur Aktif (Inv) :  3

========================================
DAFTAR RINCIAN FAKTUR AKTIF
========================================
100001   2.000.000   800.000   15 HR (OWING)
100002   500.000     Ttp Byr: 265.433   234.567   45 HR (JT 01/05/25)
100003   200.000     60 HR
```

Kolom per baris faktur (kiri ke kanan, sesuai flag yang aktif):

| Komponen | Flag | Contoh |
|---|---|---|
| Nomor faktur | `ar_data_inv_numb` | `100001` |
| Tanggal faktur | `ar_data_inv_dt` | `15 Jan 2025` |
| Jatuh tempo | `ar_data_inv_due` | `15 Feb 2025` |
| Nilai faktur asli | `ar_data_inv_orig` | `2.000.000` |
| Sisa piutang | `ar_data_inv_ar` | `800.000` |
| Titip bayar | `ar_data_inv_pay` | `Ttp Byr: 265.433` |
| Umur piutang | `ar_data_age` | `15 HR` |
| Status OWING | `ar_data_owing` | `(OWING)` |
| Tanggal giro | `ar_data_giro` | `(JT 01/05/25)` |


---

## 🎮 Mode Demo (Tanpa URL)

Jika belum memiliki Google Sheets yang terkonfigurasi:

1. Kosongkan `url =` di `[OWING]` dan `[ARAVG]`
2. Isi `ar_url`, `ar_sheet`, `ar_key_col`, `ar_target_col` dengan spreadsheet uji
3. Jalankan seperti biasa

`0_HDownloaderData.py` akan mengunduh contoh dari GitHub:
- `Owing_temp.xlsx` — contoh daftar OWING
- `Avg_temp.xlsx` — contoh data performa pelanggan

> File contoh tersedia di folder `Contoh Data/` untuk referensi struktur kolom.

---

## 🛠️ Troubleshooting

### ❌ `File Piutang.xls tidak ditemukan. Proses digagalkan`
Pastikan file ada di folder utama dengan nama persis `Piutang.xls`.

### ❌ `File 2_HCleanerAccGiroDue.py tidak ditemukan di dalam folder Dapur`
File baru ini wajib ada di `Dapur/`. Pastikan seluruh isi repositori diunduh ulang jika baru meng-clone; file ini tidak ada di versi sebelumnya.

### ❌ `Kesalahan: Pastikan ARClean_temp.xlsx, Owing_temp.xlsx, dan Avg_temp.xlsx ada`
Salah satu skrip 0–5 gagal. Jalankan manual dari `Dapur/`:
```bash
cd Dapur
python 1_CleanerAccAR.py
```

### ❌ `Terjadi kegagalan deteksi struktur kolom tabel Excel`
Kolom kunci tidak ditemukan di `Owing_temp.xlsx` atau `Avg_temp.xlsx`. Periksa nilai `ow_excel_col`, `ow_excel_sheet`, `avg_excel_key`, dan `avg_excel_sheet` di `config.conf`.

### ❌ `Kesalahan nama kolom di Google Sheets tidak ditemukan`
Nilai `ar_key_col` atau `ar_target_col` tidak ada di baris pertama sheet target. Verifikasi nama kolom di Google Sheets (termasuk kapitalisasi dan spasi).

### ❌ Kolom `Tanggal JT` selalu kosong meski ada data di Giro.xls
Kemungkinan `2_HCleanerAccGiroDue.py` menghapus semua baris karena seluruh `Tgl Cek` sudah lewat. Periksa isi `Giro.xls` — pastikan ada baris dengan tanggal cek hari ini atau yang akan datang.

### ❌ Loop berjalan tapi `Tidak ada data target kosong baru`
Semua baris sudah terisi. Normal jika semua pesanan sudah diproses. Sistem tetap menunggu baris baru.

### ❌ Nilai `Ttp Byr` tidak muncul di cell note
`ar_data_inv_pay = Ya` hanya menampilkan titip bayar jika `Nilai Faktur − Sisa Piutang > 0`. Jika faktur belum ada pembayaran sama sekali, baris ini tidak akan ditampilkan.

### ❌ Error autentikasi Google (`DefaultCredentialsError` / `invalid_grant`)
Periksa isi `credentials.json` — `private_key` harus tersalin lengkap termasuk `-----BEGIN PRIVATE KEY-----` dan `-----END PRIVATE KEY-----`.

### ❌ Loop terlalu cepat / terlalu lambat
Ubah `ar_time_interval` di `[AR]`. Minimum yang disarankan: 5 menit untuk menghindari rate limit API.

---

## 📌 Catatan Penting

- **`Piutang.xls` disalin, bukan dipindahkan** — File asli di folder utama tetap aman setelah proses.
- **`credentials.json` wajib dijaga kerahasiaannya** — Tambahkan `Dapur/credentials.json` ke `.gitignore`. Jangan pernah commit ke repositori publik.
- **Skrip 6 adalah loop tak terbatas** — Hentikan dengan `Ctrl+C`.
- **Hanya sel kosong yang diisi** — Untuk memperbarui data yang sudah terisi, kosongkan dulu secara manual di Google Sheets.
- **`2_HCleanerAccGiroDue.py` wajib ada di `Dapur/`** — Orkestrator memvalidasi keberadaannya sebelum mulai. Proses akan gagal jika file ini tidak ada.
- **Filter giro bersifat permanen per run** — Setiap kali pipeline dijalankan, `Giro_temp.xlsx` difilter ulang berdasarkan tanggal hari itu. Giro yang cair kemarin tidak akan muncul di run hari ini.
- **Kode pelanggan harus konsisten** — Gunakan awalan yang dikenali (SL, YY, MGL, dst.) agar standarisasi berjalan dengan benar.
- **File sementara di `Dapur/` dihapus saat run baru** — Jangan simpan file penting di sana.
- **Data FBACK bersifat additive** — Skrip 4 hanya menambah baris baru, tidak pernah menghapus baris yang sudah ada.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*
