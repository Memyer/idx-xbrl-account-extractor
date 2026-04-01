# 📊 IDX XBRL Account Extractor

> **Otomasi Ekstraksi Data Keuangan XBRL — Bursa Efek Indonesia**

[![Python](https://img.shields.io/badge/Python-3.13%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D4?logo=windows&logoColor=white)](https://www.microsoft.com/windows/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Active-brightgreen)](https://github.com/Memyer/idx-xbrl-account-extractor)

---

## ✨ Highlight

Ekstraksi data keuangan emiten **BEI** secara otomatis dari dokumen **Inline XBRL** dengan:

- ✅ **GUI Desktop Modern** – Interface intuitif dengan CustomTkinter.
- ✅ **Pilih Akun Fleksibel** – 120+ akun keuangan terstruktur per kategori.
- ✅ **Pipeline Otomatis** – 7 tahap from download sampai export.
- ✅ **Dual Output** – RAW (asli) + FULL IDR (terkonversi) dalam CSV & Excel.
- ✅ **Anti-Bot Smart** – Session handling + header matching untuk download stabil.
- ✅ **One-Click Build** – Jadikan satu file .exe siap distribusi.

---

## 🎯 Fitur Utama

### 1. 🔧 Pengaturan Fleksibel
- Konfigurasi tahun laporan dan periode (Q1–Q4, Tahunan).
- Folder output custom.
- Kurs USD ke IDR adjustable.
- Opsi pembersihan data per kolom.

### 2. 📋 Selector Akun Cerdas
- **120+ akun** dari 4 kategori:
  - Informasi Umum
  - Neraca (Balance Sheet)
  - Laba Rugi (Income Statement)
  - Arus Kas (Cash Flow)
- Search real-time.
- Tambah/hapus pilihan per item atau batch.

### 3. 🚀 Pipeline 7 Tahap
| # | Tahap | Icon | Deskripsi |
|---|-------|------|-----------|
| 1 | Generate Links | 🔗 | Buat daftar link per emiten (ticker IDX). |
| 2 | Download XBRL | ⬇️ | Unduh berkas XBRL ZIP dengan session anti-bot. |
| 3 | Ekstrak ZIP | 📦 | Ekstrak ZIP ke folder XBRL. |
| 4 | Ekstrak Akun | 🔍 | Parse HTML, ambil akun yang dipilih → CSV. |
| 5 | Bersihkan Data | 🧹 | Ubah nilai negatif sesuai setting. |
| 6 | Normalisasi IDR | 💱 | Konversi Full Amount sesuai pembulatan + mata uang. |
| 7 | Export CSV+Excel | 📊 | Simpan 2 format: RAW + FULL IDR. |

### 4. 📁 Output Terstruktur
Hasil tersimpan otomatis di **ExtractedData_XBRL/**:

```
00_SUMMARY_all_companies_raw.csv          ← Data asli (belum konversi)
00_SUMMARY_all_companies_raw.xlsx

00_SUMMARY_all_companies_full_idr.csv     ← Data konversi Full IDR
00_SUMMARY_all_companies_full_idr.xlsx
```

---

## 🏗️ Struktur Proyek

```
idx-xbrl-account-extractor/
├── main.py                    # Entry point aplikasi
├── gui/
│   ├── app.py                 # UI desktop + orkestrasi pipeline
│   └── __init__.py
├── core/
│   ├── link_generator.py      # Pembuatan link IDX
│   ├── downloader.py          # Download XBRL ZIP (anti-bot)
│   ├── unzipper.py            # Ekstraksi ZIP
│   ├── xbrl_extractor.py      # Parse & ekstrak akun XBRL
│   ├── data_cleaner.py        # Pembersihan nilai negatif
│   ├── amount_normalizer.py   # Konversi Full IDR
│   ├── csv_exporter.py        # Export ke Excel
│   └── __init__.py
├── build.bat                  # Script build Windows
├── requirements.txt           # Dependencies
├── README.md                  # Dokumentasi (file ini)
└── companies.txt              # Daftar emiten IDX
```

---

## 📥 Instalasi & Cara Pakai

### Opsi A: Jalankan dari Source (Development)

**Prasyarat:**
- Windows 10/11
- Python 3.13+

**Langkah:**

```bash
# 1. Clone repository
git clone https://github.com/Memyer/idx-xbrl-account-extractor.git
cd idx-xbrl-account-extractor/app

# 2. Install dependencies
pip install -r requirements.txt

# 3. Jalankan aplikasi
python main.py
```

### Opsi B: Aplikasi Siap Pakai (.exe)

Unduh file **IDX_Superapp.exe** dari [Releases](https://github.com/Memyer/idx-xbrl-account-extractor/releases) dan jalankan langsung.

---

## 🔨 Build Menjadi Aplikasi Windows

Proses build otomatis menggunakan PyInstaller.

### Cara Cepat:

```bash
cd app
build.bat
```

Output: `dist/IDX_Superapp.exe`

### Manual:

```bash
python -m pip uninstall -y pathlib  # Jika perlu
python -m PyInstaller \
  --onefile --windowed \
  --name IDX_Superapp \
  --add-data "../akun_indonesia.txt;." \
  --add-data "core;core" \
  --add-data "gui;gui" \
  --hidden-import curl_cffi \
  --hidden-import bs4 \
  --hidden-import pandas \
  --hidden-import openpyxl \
  main.py
```

---

## ⚠️ Troubleshooting

### Build PyInstaller Gagal: "pathlib is incompatible"

```bash
python -m pip uninstall -y pathlib
```

Lalu ulangi build.

### Error Download: "Bot Detected"

- Pastikan **Kurs USD** terisi di tab Pengaturan.
- Cek internet koneksi.
- Tunggu beberapa menit dan coba ulang.

---

## 📊 Contoh Workflow

1. Buka aplikasi → tab **Pengaturan**.
2. Atur tahun, periode, dan kurs USD.
3. Tab **Pilih Akun** → pilih akun yang diinginkan (~20–50 akun).
4. Tab **Jalankan** → klik tombol hijau **▶ Jalankan Semua**.
5. Monitor progress di **Log**.
6. Hasil otomatis di folder **ExtractedData_XBRL/**.
7. Buka file CSV/Excel untuk analisis lanjut.

---

## 🛠️ Tech Stack

| Komponen | Library |
|----------|---------|
| Desktop UI | CustomTkinter |
| Data Processing | Pandas, NumPy |
| Web Scraping | curl_cffi, BeautifulSoup4 |
| Excel Export | openpyxl |
| Build | PyInstaller |
| Sumber Data | idx.co.id (Inline XBRL) |

---

## 📝 Lisensi

MIT License — Bebas digunakan, modifikasi, dan distribusikan.

---

## 🤝 Kontribusi

Issues dan pull requests welcome! Silakan buka issue untuk bug report atau feature request.

---

## 📞 Kontak & Info

- **Repository**: https://github.com/Memyer/idx-xbrl-account-extractor
- **Author**: Memyer
- **Last Updated**: April 1, 2026

---

**Made with ❤️ untuk komunitas data keuangan Indonesia.**

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
