# 📊 IDX Superapp — Ekstraksi Data Keuangan XBRL

> **Otomasi penuh ekstraksi, normalisasi, dan analisis data keuangan emiten Bursa Efek Indonesia**

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows-0078D4?logo=windows&logoColor=white)](https://www.microsoft.com/windows/)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)
[![Status](https://img.shields.io/badge/Status-Active-brightgreen)](https://github.com/Memyer/idx-xbrl-account-extractor)

---

## ✨ Highlight

- 🖥️ **GUI Desktop Modern** — Antarmuka intuitif, bisa dipakai tanpa pengetahuan teknis
- 📋 **738 Akun XBRL** — Semua pos keuangan dari 4 kategori (Informasi Umum, Neraca, Laba Rugi, Arus Kas)
- ⚗️ **Preset Analisis** — Altman Z-Score & IBD (Interest Bearing Debt) satu klik
- 🔄 **Pipeline 7 Langkah** — Dari download hingga Excel otomatis
- 💾 **Dual Output** — Data mentah (RAW) + nilai Rupiah penuh (Full IDR)
- 📦 **Standalone EXE** — Tidak perlu install Python di komputer lain

---

## 🎯 Fitur Utama

### ⚙️ Pengaturan
- Pilih tahun laporan dan periode (Q1 / Q2 / Q3 / Tahunan)
- Folder kerja custom — semua output tersimpan rapi
- Kurs USD ke IDR adjustable untuk emiten pelapor USD

### 📋 Pilih Akun
- **738 akun XBRL** dari `Pos_XBRL_IDX_Lengkap.xlsx`
- Filter per kategori: Informasi Umum · Neraca · Laba Rugi · Arus Kas
- Search real-time
- **Preset sekali klik:**

| Preset | Jumlah Akun | Kegunaan |
|--------|-------------|----------|
| ⚗️ Altman Z-Score | 24 akun | Prediksi risiko kebangkrutan |
| 💰 IBD | 59 akun | Total utang berbunga |

### 🚀 Pipeline 7 Langkah

| # | Langkah | Deskripsi |
|---|---------|-----------|
| 1 | 🔗 Generate Links | Ambil semua ticker IDX & buat `companies.txt` |
| 2 | ⬇ Download XBRL | Unduh file ZIP laporan keuangan dari idx.co.id |
| 3 | 📦 Ekstrak ZIP | Buka ZIP ke folder XBRL |
| 4 | 🔍 Ekstrak Data XBRL | Parse HTML, ekstrak akun terpilih → CSV |
| 5 | 💱 Normalisasi IDR | Konversi ke Rupiah penuh (Full Amount) |
| 6 | 📊 Export Excel | Simpan 2 file Excel: RAW + Full IDR |
| 7 | ⚗️ Hitung Metrik | Hitung Altman Z-Score & IBD (jika akun dipilih) |

### 📁 Output

```
ExtractedData_XBRL/
├── 00_SUMMARY_all_companies_raw.csv              ← Data mentah
├── 00_SUMMARY_all_companies_raw.xlsx
├── 00_SUMMARY_all_companies_full_idr.csv         ← Nilai Rupiah penuh
├── 00_SUMMARY_all_companies_full_idr.xlsx
└── 00_SUMMARY_all_companies_full_idr_metrics.xlsx ← + Altman Z-Score & IBD
```

---

## 📥 Instalasi & Cara Pakai

### Opsi A: Langsung Pakai EXE *(Direkomendasikan)*

> Tidak perlu install Python atau library apapun.

1. Download `IDX_Superapp.exe` dari [Releases](https://github.com/Memyer/idx-xbrl-account-extractor/releases)
2. Klik dua kali → aplikasi langsung terbuka

### Opsi B: Jalankan dari Source Code

**Prasyarat:** Windows 10/11, Python 3.10+

```bash
# 1. Clone repository
git clone https://github.com/Memyer/idx-xbrl-account-extractor.git
cd idx-xbrl-account-extractor/app

# 2. Setup environment (otomatis buat venv + install dependencies)
install.bat

# 3. Jalankan
run.bat
# atau: .venv\Scripts\python main.py
```

### Opsi C: Manual

```bash
pip install -r requirements.txt
python main.py
```

---

## 🔨 Build EXE

```bat
build.bat
```

Output: `release\IDX_Superapp.exe` — siap distribusi, semua data sudah di-bundle.

---

## 🏗️ Struktur Proyek

```
idx-xbrl-account-extractor/
├── main.py                        # Entry point
├── gui/
│   └── app.py                     # UI desktop + orkestrasi pipeline
├── core/
│   ├── link_generator.py          # Generate link IDX per ticker
│   ├── downloader.py              # Download XBRL ZIP (anti-bot)
│   ├── unzipper.py                # Ekstrak ZIP
│   ├── xbrl_extractor.py          # Parse & ekstrak akun XBRL
│   ├── amount_normalizer.py       # Konversi Full IDR
│   ├── csv_exporter.py            # Export ke Excel
│   ├── metric_calculator.py       # Hitung Altman Z-Score & IBD
│   └── data_cleaner.py            # Pembersihan nilai beban bunga
├── Pos_XBRL_IDX_Lengkap.xlsx      # Referensi 738 akun XBRL
├── Altzman _IBD.xlsx              # Referensi akun preset Altman & IBD
├── IDX_Superapp.spec              # Konfigurasi PyInstaller
├── build.bat                      # Script build EXE
├── install.bat                    # Script setup environment
└── requirements.txt               # Dependencies Python
```

---

## 🛠️ Tech Stack

| Komponen | Library |
|----------|---------|
| Desktop GUI | CustomTkinter 5.2+ |
| Web Request | curl_cffi (anti-bot) |
| HTML Parsing | BeautifulSoup4, lxml |
| Data Processing | Pandas |
| Excel I/O | openpyxl |
| Build | PyInstaller 6+ |
| Sumber Data | idx.co.id — Inline XBRL |

---

## ⚠️ Troubleshooting

**Download gagal / Bot Detected**
- Pastikan koneksi internet aktif
- Tunggu beberapa menit lalu coba ulang
- Aplikasi otomatis retry dengan session baru

**Akun Altman/IBD tidak terhitung**
- Pastikan preset akun sudah dipilih di halaman **Pilih Akun** sebelum menjalankan pipeline

**Build EXE gagal**
```bash
.venv\Scripts\pip install --upgrade pyinstaller
build.bat
```

---

## 📝 Lisensi

MIT License — Bebas digunakan, dimodifikasi, dan didistribusikan.

---

**Made with ❤️ untuk komunitas data keuangan Indonesia · Last updated: April 2026**
