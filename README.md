# 📑 Automasi AR Orderan

> **Sinkronisasi data piutang AR ke Google Sheets order management secara real-time dan berkelanjutan — dari ekspor Accurate ke cell note terstruktur siap pakai tim admin sales**

Pipeline Python tujuh langkah yang membaca ekspor AR dari Accurate (`Piutang.xls`) dan data Giro (`Giro.xls`), mengunduh data referensi dari tiga Google Sheets (Owing, AVG performa pelanggan, FallbackCash), menggabungkan semua sumber, lalu menyuntikkan **nilai total piutang** dan **ringkasan lengkap per pelanggan sebagai cell note** ke kolom target di Google Sheets order tracker — berjalan dalam loop otomatis setiap menit (bawaan per-15 menit) untuk menangkap pesanan baru secara real-time.

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
- **Cell note** → ringkasan lengkap kondisi kredit pelanggan: plafon, rata-rata bayar, riwayat, daftar faktur aktif, status OWING dan tanggal giro

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
- **Cell note terstruktur dengan 20+ flag** — Setiap komponen dalam ringkasan cell note (plafon, rata-rata bayar, nomor faktur, tanggal, umur, dll.) dapat diaktifkan atau dinonaktifkan secara individual via `config.conf`.
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
├── 📄 Jalankan Automasi.py          ← Orkestrator utama. Jalankan ini
├── 📄 Piutang.xls                   ← [INPUT] Ekspor AR dari Accurate (wajib)
├── 📄 Giro.xls                      ← [INPUT] Rekap giro/cek masuk (opsional)
├── 📄 Ekspor Data.png               ← Panduan visual cara ekspor dari Accurate
│
├── 📁 Dapur/                        ← Folder pipeline (jangan diubah)
│   ├── 📄 __init__.py
│   ├── 📄 0_DownloaderData.py       ← Unduh Owing/AVG/FallbackCash dari Google Sheets
│   ├── 📄 0_HDownloaderData.py      ← Unduh data contoh jika URL belum dikonfigurasi
│   ├── 📄 1_CleanerAccAR.py         ← Bersihkan Piutang.xls → ARClean_temp.xlsx
│   ├── 📄 2_CleanerAccGiro.py       ← Bersihkan Giro.xls → Giro_temp.xlsx (jika aktif)
│   ├── 📄 3_AddGiroToSheet.py       ← Tambahkan Tanggal JT ke ARClean dari Giro
│   ├── 📄 4_PatchFallbackCash.py    ← Tambahkan pelanggan cash dari FallbackCash
│   ├── 📄 5_AdjDateFormat.py        ← Format ulang tanggal ke format Indonesia
│   ├── 📄 6_InjectDataToSS.py       ← Loop sinkronisasi → inject ke Google Sheets
│   ├── 📄 config.conf               ← Konfigurasi utama (wajib diisi sebelum pakai)
│   └── 📄 credentials.json          ← Kredensial Google Service Account (rahasia!)
│
└── 📁 Contoh Data/                  ← Data sampel untuk mode demo
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
   │       Cek 9 file syarat di Dapur/ ada → gagal jika kurang
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
   ├─── [3] 3_AddGiroToSheet.py
   │       Cek [GIRO] giro_stats → Skip jika tidak aktif
   │       Bangun mapping {No. Faktur → "JT DD/MM/YY & ..."}
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
             │    └─ Standarisasi kode pelanggan (support multi-kode & )
             │    └─ Lookup AR data → hitung total Sisa Piutang
             │    └─ Lookup AVG data → ambil plafon, bayar, history, tier
             │    └─ Bangun cell note terstruktur (sesuai flag aktif)
             │    └─ Tambahkan flag OWING dan JT Giro per baris faktur
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

Setelah pembersihan, dua kolom spacer `SS` disisipkan (setelah `Tgl Faktur` dan setelah `Jatuh Tempo`) untuk menjaga kompatibilitas struktur kolom. Output: `ARClean_temp.xlsx`. File `Piutang.xls` di Dapur/ dihapus setelah berhasil.

---

### Skrip 2 & 3 — Giro (kondisional)

Diaktifkan hanya jika `[GIRO] giro_stats = Ya` di `config.conf`. Skrip 2 membersihkan `Giro.xls` → `Giro_temp.xlsx`. Skrip 3 membangun peta `{No. Faktur → "JT ..."}` dan menambahkan kolom `Tanggal JT` ke `ARClean_temp.xlsx`.

---

### Skrip 4 — `4_PatchFallbackCash.py`

Membandingkan kode pelanggan di `ARClean_temp.xlsx` dengan `FallbackCash_temp.xlsx`. Pelanggan yang ada di FallbackCash namun tidak memiliki piutang aktif di data AR ditambahkan sebagai baris baru dengan kolom `Kode Pelanggan`, `Nama Pelanggan`, dan `Nama Kontak` terisi, kolom faktur lainnya kosong. Ini memastikan pelanggan cash tetap muncul di ringkasan Google Sheets.

---

### Skrip 5 — `5_AdjDateFormat.py`

Mengonversi nilai datetime di kolom `Tgl Faktur` dan `Jatuh Tempo` dari format datetime objek ke teks Indonesia: `"15 Jan 2025"`, `"28 Feb 2025"`, dst.

---

### Skrip 6 — `6_InjectDataToSS.py` (Loop utama)

Inti dari seluruh sistem. Membuka koneksi ke Google Sheets, memindai baris demi baris, dan mengisi sel kosong dengan nilai + cell note AR. Detail logika:

**Standardisasi kode pelanggan:**
```
"MGL1234" → "MGL-1234"
"sl 001"  → "SL-001"
"YY1234A" → "YY-1234 A"
Awalan prefix yang dikenali: SL, YY, MKS, MGL, PW, PWT, PLU, SG, SMG, TGL, PA, KDI
```

**Multi-kode per sel:**
Sel dengan nilai `"MGL-001 & MGL-002"` akan mencari data AR untuk keduanya sekaligus dan menggabungkan hasilnya.

**Hanya mengisi sel kosong:**
Baris di mana kolom `ar_target_col` sudah berisi nilai apapun akan dilewati, bahkan jika datanya sudah berubah.

**Batch update:**
Semua perubahan dalam satu iterasi loop dikirim sekaligus via `batchUpdate` Google Sheets API, bukan satu per satu, untuk efisiensi.

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

Kosongkan `url =` untuk skip pengunduhan (gunakan data dari run sebelumnya atau mode demo).

---

### `[GIRO]` — Aktifkan/nonaktifkan fitur Giro

```ini
[GIRO]
giro_stats = Ya    ; Aktifkan: Ya | Nonaktifkan: No
```

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
avg_excel_sheet = Solo                    ; Nama sheet di Avg_temp.xlsx
avg_excel_key = NO. PELANGGAN            ; Kolom kunci untuk join dengan data AR
avg_excel_age = AVG UMUR PIUTANG         ; Kolom rata-rata umur piutang
avg_excel_val = AVG NILAI FAKTUR         ; Kolom rata-rata nilai faktur
avg_excel_inv = JUMLAH INVOICE           ; Kolom rata-rata jumlah invoice
avg_excel_plaf = PLAFON                  ; Kolom plafon kredit
avg_excel_pay = AVG BAYAR                ; Kolom rata-rata pembayaran
avg_excel_his = AVG HISTORY BAYAR (HARI) ; Kolom histori hari bayar
avg_excel_tier = TIERING                 ; Kolom tiering pelanggan
```

---

### `[AR]` — Target Google Sheets & semua flag output

```ini
[AR]
ar_url = https://docs.google.com/spreadsheets/d/ID/edit  ; URL spreadsheet order tracker
ar_time_interval = 15          ; Interval loop sinkronisasi (menit)
ar_sheet = Solo                ; Nama sheet target di spreadsheet
ar_key_col = KODE PELANGGAN   ; Nama kolom berisi kode pelanggan di sheet target
ar_prod_key_col =              ; Nama kolom produk/divisi (boleh kosong)
ar_target_col = Nominal Nota Belum Lunas  ; Kolom yang akan diisi nilai + note
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
| `ar_data_inv_val` | `Ya` | Total faktur aktif (jumlah baris) |
| `ar_data_inv_ar` | `Ya` | Sisa piutang per faktur |
| `ar_data_owing` | `Ya` | Tandai `(OWING)` di baris faktur |
| `ar_data_giro` | `Ya` | Tampilkan tanggal giro `(JT ...)` per faktur |
| `ar_data_age` | `Ya` | Hitung umur piutang (hari dari Tgl Faktur hingga hari ini) |

---

## 🔑 Setup Google Sheets API

### 1. Buat Service Account

1. Buka [Google Cloud Console](https://console.cloud.google.com/) → buat atau pilih project.
2. Aktifkan **Google Sheets API** dan **Google Drive API**.
3. Masuk ke **IAM & Admin → Service Accounts** → buat Service Account baru.
4. Di tab **Keys** → buat key baru tipe **JSON** → file terunduh otomatis.

### 2. Pasang kredensial

Ganti isi `Dapur/credentials.json` dengan file JSON yang diunduh. Format yang dibutuhkan:

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

Buka setiap Google Sheets yang digunakan (OWING, ARAVG, FBACK, dan AR target), klik **Share**, lalu tambahkan `client_email` dari `credentials.json` sebagai **Editor**.

### 4. Struktur spreadsheet target (AR order tracker)

Sheet target di Google Sheets harus memiliki:
- Baris pertama = header kolom
- Kolom dengan nama sesuai `ar_key_col` (misal: `KODE PELANGGAN`) — berisi kode pelanggan per baris pesanan
- Kolom dengan nama sesuai `ar_target_col` (misal: `Nominal Nota Belum Lunas`) — ini yang akan diisi otomatis (mulai kosong)
- Kolom pertama (A) berisi tanggal order (digunakan jika `ar_data_dt_order = Ya`)

---

## 📤 Format Output: Nilai Sel & Cell Note

Setiap sel yang diisi terdiri dari dua bagian:

### Nilai Sel

Total `Sisa Piutang` seluruh faktur aktif pelanggan dalam format IDR (titik sebagai pemisah ribuan):

```
1.234.567
```

### Cell Note (hover/klik sel)

Teks terstruktur multi-baris yang dapat dilihat dengan mengarahkan kursor ke sel atau klik ikon segitiga kecil di pojok sel:

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
100001   800.000   15 HR (OWING)
100002   234.567   45 HR (JT 01/05/25)
100003   200.000   60 HR
```

Setiap baris dalam **DAFTAR RINCIAN** dapat memiliki suffix:
- `(OWING)` — nomor faktur ini ada di daftar Owing
- `(JT DD/MM/YY)` — ada tanggal cair giro dari data `Giro.xls`

---

## 🎮 Mode Demo (Tanpa URL)

Jika belum memiliki Google Sheets yang terkonfigurasi, sistem dapat dijalankan dalam **mode demo** untuk keperluan pengujian:

1. Kosongkan `url =` di seksi `[OWING]` dan `[ARAVG]` di `config.conf`
2. Isi `ar_url`, `ar_sheet`, `ar_key_col`, dan `ar_target_col` dengan spreadsheet uji Anda
3. Jalankan seperti biasa

Skrip `0_HDownloaderData.py` akan otomatis mengunduh file contoh dari folder `Contoh Data/` di repositori GitHub ini:
- `Owing_temp.xlsx` — contoh daftar OWING
- `Avg_temp.xlsx` — contoh data performa pelanggan

> File contoh tersedia di folder `Contoh Data/` repositori ini untuk referensi struktur kolom yang dibutuhkan.

---

## 🛠️ Troubleshooting

### ❌ `File Piutang.xls tidak ditemukan. Proses digagalkan`
Pastikan file ada di folder utama (sejajar dengan `Jalankan Automasi.py`) dengan nama persis `Piutang.xls`.

### ❌ `Kesalahan: Pastikan ARClean_temp.xlsx, Owing_temp.xlsx, dan Avg_temp.xlsx ada`
Salah satu skrip 0–5 gagal menghasilkan file sementara. Jalankan skrip tersebut secara manual dari dalam folder `Dapur/` untuk melihat pesan error lengkap:
```bash
cd Dapur
python 1_CleanerAccAR.py
```

### ❌ `Terjadi kegagalan deteksi struktur kolom tabel Excel`
Kolom kunci tidak ditemukan di `Owing_temp.xlsx` atau `Avg_temp.xlsx`. Periksa nilai `ow_excel_col`, `ow_excel_sheet`, `avg_excel_key`, dan `avg_excel_sheet` di `config.conf` — harus persis sesuai nama kolom/sheet di file tersebut.

### ❌ `Kesalahan nama kolom di Google Sheets tidak ditemukan`
Nilai `ar_key_col` atau `ar_target_col` tidak ditemukan di baris pertama sheet target. Buka Google Sheets target, verifikasi nama kolom persis sama (termasuk kapitalisasi dan spasi).

### ❌ Loop berjalan tapi `Tidak ada data target kosong baru yang perlu diperbarui`
Semua sel di kolom target sudah terisi. Ini normal jika semua baris sudah diproses. Sistem akan terus memantau baris baru yang ditambahkan.

### ❌ Nilai sel terisi tapi cell note tidak muncul
Kemungkinan pelanggan tidak ditemukan di `ARClean_temp.xlsx` maupun `Avg_temp.xlsx`. Periksa apakah kode pelanggan di Google Sheets (setelah standarisasi) cocok dengan kode di data AR.

### ❌ `ar_data_fraud = No` tapi data FRAUD masih muncul
Pastikan kolom `Nama Penjual` ada di `ARClean_temp.xlsx` dan berisi kata `FRAUD` pada baris yang ingin difilter. Filter dilakukan dengan `str.contains('FRAUD', case=False)`.

### ❌ Error autentikasi Google (`DefaultCredentialsError` / `invalid_grant`)
Periksa isi `Dapur/credentials.json` — pastikan file JSON valid dan `private_key` tersalin lengkap termasuk baris `-----BEGIN PRIVATE KEY-----` dan `-----END PRIVATE KEY-----`.

### ❌ Loop terlalu cepat / terlalu lambat
Ubah nilai `ar_time_interval` di seksi `[AR]` config.conf. Nilai dalam menit. Interval minimum yang disarankan: 5 menit untuk menghindari rate limit Google Sheets API.

---

## 📌 Catatan Penting

- **`Piutang.xls` disalin, bukan dipindahkan** — Orkestrator menggunakan `shutil.copy()`, sehingga file asli di folder utama tetap aman setelah proses selesai. Berbeda dengan proyek ARVIEWER yang menggunakan `shutil.move()`.
- **`credentials.json` wajib dijaga kerahasiaannya** — Jangan pernah commit file ini ke repository publik. Tambahkan `Dapur/credentials.json` ke `.gitignore`.
- **Skrip 6 adalah loop tak terbatas** — Hentikan dengan `Ctrl+C`. Terminal akan menampilkan pesan "Looping dihentikan oleh pengguna." dan orkestrator akan selesai dengan aman.
- **Hanya sel kosong yang diisi** — Skrip 6 tidak akan menimpa sel yang sudah berisi nilai. Jika perlu memperbarui data yang sudah ada, kosongkan dulu sel tersebut secara manual di Google Sheets.
- **Kode pelanggan harus konsisten** — Pastikan kode di Google Sheets order tracker dan di `Piutang.xls` menggunakan format yang sama setelah standarisasi. Gunakan awalan yang dikenali (SL, YY, MGL, dst.) untuk memastikan normalisasi berjalan dengan benar.
- **File sementara di Dapur/ dihapus saat run baru dimulai** — Orkestrator menghapus semua `*.xls` dan `*.xlsx` di Dapur/ sebelum menyalin file baru. Jangan simpan file penting di sana.
- **Data FBACK bersifat additive** — Skrip 4 hanya menambah baris baru, tidak pernah menghapus atau menimpa baris yang sudah ada di `ARClean_temp.xlsx`.

---

## 📜 Lisensi

Proyek ini dikembangkan untuk keperluan internal perusahaan. Silakan sesuaikan dengan kebutuhan organisasi Anda.

---

*Dikembangkan oleh [ACC-TAX-REIGHTEEN](https://github.com/ACC-TAX-REIGHTEEN)*a
