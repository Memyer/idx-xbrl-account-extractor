# IDX XBRL Account Extractor

Aplikasi desktop untuk otomasi ekstraksi data keuangan emiten IDX dari dokumen Inline XBRL.

## Ringkasan

Proyek ini menyediakan pipeline end-to-end untuk:
- Membuat daftar link laporan berdasarkan tahun dan periode.
- Mengunduh berkas ZIP XBRL.
- Mengekstrak isi ZIP.
- Mengambil akun-akun keuangan tertentu dari HTML XBRL.
- Membersihkan nilai negatif pada kolom yang dipilih.
- Menormalisasi nilai akun menjadi Full Amount IDR.
- Mengekspor hasil ke CSV dan Excel.

## Fitur Utama

- GUI berbasis CustomTkinter dengan 5 tab: Pengaturan, Pilih Akun, Jalankan, Log, Tentang.
- Pemilihan akun fleksibel per metrik.
- Progress per tahap dan kontrol eksekusi (run per step atau jalankan semua).
- Export hasil terpisah:
  - RAW (sebelum konversi Full IDR).
  - FULL IDR (sesudah konversi).
- Build menjadi satu file aplikasi Windows (.exe) dengan PyInstaller.

## Alur Pipeline

1. Generate Links
2. Download XBRL ZIP
3. Ekstrak ZIP
4. Ekstrak Data XBRL
5. Bersihkan Data
6. Normalisasi Full IDR
7. Export CSV + Excel

## Output

Secara default output disimpan di folder ExtractedData_XBRL.

File utama yang dihasilkan:
- 00_SUMMARY_all_companies_raw.csv
- 00_SUMMARY_all_companies_raw.xlsx
- 00_SUMMARY_all_companies_full_idr.csv
- 00_SUMMARY_all_companies_full_idr.xlsx

## Struktur Proyek

- main.py: Entry point aplikasi.
- gui/app.py: Antarmuka desktop dan orkestrasi pipeline.
- core/link_generator.py: Pembuatan link IDX.
- core/downloader.py: Pengunduhan file XBRL.
- core/unzipper.py: Ekstraksi ZIP.
- core/xbrl_extractor.py: Parsing dan ekstraksi akun XBRL.
- core/data_cleaner.py: Pembersihan data.
- core/amount_normalizer.py: Konversi Full IDR.
- core/csv_exporter.py: Konversi CSV ke Excel.
- build.bat: Script build aplikasi Windows.

## Menjalankan dari Source

Prasyarat:
- Windows
- Python 3.13+

Langkah:

```bash
pip install -r requirements.txt
python main.py
```

## Build Menjadi Satu Aplikasi (.exe)

Cara cepat:

```bat
build.bat
```

Atau manual:

```bash
python -m PyInstaller --noconfirm --clean --onefile --windowed --name IDX_Superapp --add-data "..\akun_indonesia.txt;." --add-data "core;core" --add-data "gui;gui" --hidden-import curl_cffi --hidden-import bs4 --hidden-import pandas --hidden-import openpyxl main.py
```

Hasil build:
- dist/IDX_Superapp.exe

## Catatan

Jika build PyInstaller gagal karena paket pathlib backport, jalankan:

```bash
python -m pip uninstall -y pathlib
```

## Lisensi

Belum ditentukan.
