"""
core/xbrl_extractor.py
Extract financial metrics from XBRL HTML files.
Supports selective extraction via selected_metrics parameter.
"""
from pathlib import Path
from bs4 import BeautifulSoup
import pandas as pd

# ═══════════════════════════════════════════════════════
# KATALOG LENGKAP — dikelompokkan per kategori
# ═══════════════════════════════════════════════════════
METRICS_CATALOG = {
    "Informasi Umum": [
        "Nama entitas",
        "Kode entitas",
        "Sektor",
        "Subsektor",
        "Industri",
        "Subindustri",
        "Periode penyampaian laporan keuangan",
        "Tanggal awal periode berjalan",
        "Tanggal akhir periode berjalan",
        "Mata uang pelaporan",
        "Pembulatan yang digunakan dalam penyajian jumlah dalam laporan keuangan",
        "Jenis laporan atas laporan keuangan",
        "Tanggal laporan audit atau hasil laporan review",
        "Nama kantor akuntan publik",
        "Nama auditor",
        "Nama mitra yang bertanggung jawab",
        "Opini auditor",
        "Jumlah karyawan",
    ],
    "Neraca (Balance Sheet)": [
        # Aset
        "Jumlah aset",
        "Jumlah aset lancar",
        "Jumlah aset tidak lancar",
        "Kas dan setara kas",
        "Investasi jangka pendek",
        "Piutang usaha",
        "Piutang lain-lain",
        "Persediaan",
        "Uang muka",
        "Pajak dibayar di muka",
        "Biaya dibayar di muka",
        "Aset keuangan lancar lainnya",
        "Aset lancar lainnya",
        "Investasi pada entitas asosiasi dan ventura bersama",
        "Investasi jangka panjang lainnya",
        "Aset tetap",
        "Properti investasi",
        "Aset tak berwujud",
        "Goodwill",
        "Aset pajak tangguhan",
        "Aset tidak lancar lainnya",
        # Liabilitas
        "Jumlah liabilitas",
        "Jumlah liabilitas jangka pendek",
        "Jumlah liabilitas jangka panjang",
        "Utang usaha",
        "Utang lain-lain",
        "Pinjaman jangka pendek",
        "Utang pajak",
        "Liabilitas jangka pendek lainnya",
        "Pinjaman jangka panjang",
        "Liabilitas imbalan pascakerja",
        "Liabilitas pajak tangguhan",
        "Liabilitas jangka panjang lainnya",
        # Ekuitas
        "Jumlah ekuitas",
        "Jumlah ekuitas yang dapat diatribusikan kepada pemilik entitas induk",
        "Modal ditempatkan dan disetor penuh",
        "Tambahan modal disetor",
        "Saldo laba yang belum ditentukan penggunaannya",
        "Saldo laba yang telah ditentukan penggunaannya",
        "Komponen ekuitas lainnya",
        "Kepentingan nonpengendali",
    ],
    "Laba Rugi (Income Statement)": [
        "Penjualan dan pendapatan usaha",
        "Beban pokok penjualan dan pendapatan",
        "Jumlah laba bruto",
        "Pendapatan operasional lainnya",
        "Beban penjualan",
        "Beban umum dan administrasi",
        "Beban operasional lainnya",
        "Jumlah laba (rugi) usaha",
        "Pendapatan bunga",
        "Beban bunga dan keuangan",
        "Beban bunga",
        "Keuntungan (kerugian) selisih kurs",
        "Bagian laba (rugi) entitas asosiasi dan ventura bersama",
        "Pendapatan (beban) lain-lain",
        "Jumlah laba (rugi) sebelum pajak penghasilan",
        "Beban pajak penghasilan",
        "Jumlah laba (rugi) dari operasi yang dilanjutkan",
        "Jumlah laba (rugi) dari operasi yang dihentikan",
        "Jumlah laba (rugi)",
        "Laba (rugi) yang dapat diatribusikan ke entitas induk",
        "Laba (rugi) yang dapat diatribusikan ke kepentingan nonpengendali",
        "Jumlah laba komprehensif",
        "Laba komprehensif yang dapat diatribusikan ke entitas induk",
        "Laba (rugi) per saham dasar",
        "Laba (rugi) per saham dilusian",
    ],
    "Arus Kas (Cash Flow)": [
        "Jumlah arus kas dari aktivitas operasi",
        "Jumlah arus kas dari aktivitas investasi",
        "Jumlah arus kas dari aktivitas pendanaan",
        "Kas dan setara kas awal periode",
        "Kas dan setara kas akhir periode",
        "Kenaikan (penurunan) bersih kas dan setara kas",
        # Operasi
        "Penerimaan dari pelanggan",
        "Pembayaran kepada pemasok",
        "Pembayaran kepada karyawan",
        "Penerimaan bunga operasi",
        "Pembayaran bunga operasi",
        "Pembayaran pajak penghasilan",
        # Investasi
        "Pengeluaran untuk perolehan aset tetap",
        "Penerimaan dari penjualan aset tetap",
        "Pengeluaran untuk investasi",
        "Penerimaan dari investasi",
        # Pendanaan
        "Penerimaan pinjaman",
        "Pembayaran pinjaman",
        "Pembayaran dividen",
        "Pembayaran bunga pendanaan",
        # Non-kas (dari rekonsiliasi)
        "Depresiasi",
        "Amortisasi",
        "Penyisihan piutang tak tertagih",
    ],
}

# Flat list semua akun (urutan: General → BS → IS → CF)
ALL_METRICS: list = []
for _cat_metrics in METRICS_CATALOG.values():
    ALL_METRICS.extend(_cat_metrics)

# Default: akun yang aktif saat install pertama (sama seperti sebelumnya)
DEFAULT_METRICS = [
    "Nama entitas",
    "Kode entitas",
    "Sektor",
    "Subsektor",
    "Industri",
    "Subindustri",
    "Periode penyampaian laporan keuangan",
    "Mata uang pelaporan",
    "Pembulatan yang digunakan dalam penyajian jumlah dalam laporan keuangan",
    "Jenis laporan atas laporan keuangan",
    "Tanggal laporan audit atau hasil laporan review",
    "Jumlah aset",
    "Jumlah aset lancar",
    "Jumlah liabilitas jangka pendek",
    "Jumlah ekuitas",
    "Saldo laba yang belum ditentukan penggunaannya",
    "Saldo laba yang telah ditentukan penggunaannya",
    "Penjualan dan pendapatan usaha",
    "Jumlah laba (rugi) sebelum pajak penghasilan",
    "Laba (rugi) yang dapat diatribusikan ke entitas induk",
    "Jumlah laba bruto",
    "Beban bunga dan keuangan",
    "Jumlah liabilitas",
    "Depresiasi",
    "Amortisasi",
    "Beban bunga",
    "Pendapatan bunga",
    "Pendapatan operasional lainnya",
]

TARGET_CONTEXTS = ["CurrentYearInstant", "CurrentYearDuration"]


def _clean_value(value_text):
    if not value_text:
        return None
    value_text = value_text.strip()
    if not value_text or value_text.lower() in ["", "nan", "none"]:
        return None
    is_negative = False
    if value_text.startswith("(") and value_text.endswith(")"):
        is_negative = True
        value_text = value_text[1:-1]
    try:
        clean_text = value_text.replace(",", "")
        result = float(clean_text) if "." in clean_text else int(clean_text)
        if is_negative:
            result = -result
        return result
    except ValueError:
        return value_text


def _extract_from_dir(company_dir: Path, metrics: list) -> dict:
    company_name = company_dir.name.split("_")[0]
    metrics_found = {}
    html_files = list(company_dir.glob("*.html"))
    if not html_files:
        return {"company": company_name, "folder": company_dir.name, "metrics": {}}

    metrics_set = set(metrics)

    for html_file in html_files:
        if not metrics_set - set(metrics_found.keys()):
            break  # semua metrik sudah ditemukan
        try:
            with open(html_file, "r", encoding="utf-8") as f:
                content = f.read()
            soup = BeautifulSoup(content, "html.parser")
            rows = soup.find_all("tr")
            for row in rows:
                header_cell = row.find("td", class_="rowHeaderLeft")
                if not header_cell:
                    continue
                header_text = header_cell.get_text(strip=True)
                if header_text not in metrics_set or header_text in metrics_found:
                    continue
                value_cells = row.find_all("td", class_="valueCell")
                for value_cell in value_cells:
                    ix_elems = value_cell.find_all(["ix:nonfraction", "ix:nonnumeric"])
                    for ix_elem in ix_elems:
                        if ix_elem.get("contextref", "") not in TARGET_CONTEXTS:
                            continue
                        if ix_elem.get("xsi:nil") == "true":
                            continue
                        value = ix_elem.get_text(strip=True)
                        sign_attr = ix_elem.get("sign", "")
                        cell_text = value_cell.get_text(strip=True)
                        has_parens = cell_text.startswith("(") and cell_text.endswith(")")
                        cleaned = _clean_value(value)
                        if cleaned is not None and isinstance(cleaned, (int, float)):
                            if has_parens or sign_attr == "-":
                                cleaned = -abs(cleaned)
                        if cleaned is not None:
                            metrics_found[header_text] = {"value": cleaned, "file": html_file.name}
                            break
                    if header_text in metrics_found:
                        break
        except Exception:
            continue

    return {"company": company_name, "folder": company_dir.name, "metrics": metrics_found}


def extract_all(xbrl_dir: str, output_dir: str,
                selected_metrics: list = None,
                progress_callback=None, log_callback=None,
                stop_flag=None) -> str:
    """
    Ekstrak data XBRL dari semua folder perusahaan.

    Args:
        selected_metrics: List nama akun yang akan diekstrak.
                          Jika None, gunakan DEFAULT_METRICS.
    Returns:
        Path ke summary CSV, atau None jika gagal.
    """
    def log(msg):
        if log_callback: log_callback(msg)
        else: print(msg)

    metrics = selected_metrics if selected_metrics else DEFAULT_METRICS
    log(f"Mengekstrak {len(metrics)} akun yang dipilih...")

    src = Path(xbrl_dir)
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    company_dirs = [d for d in src.iterdir() if d.is_dir()]
    if not company_dirs:
        log(f"Tidak ada folder perusahaan di: {xbrl_dir}")
        return None

    total = len(company_dirs)
    log(f"Ditemukan {total} perusahaan.\n")
    all_results = []

    for i, company_dir in enumerate(company_dirs, 1):
        if stop_flag and stop_flag():
            log("Dihentikan oleh user.")
            break
        log(f"[{i}/{total}] {company_dir.name}")
        result = _extract_from_dir(company_dir, metrics)
        all_results.append(result)
        found = len(result["metrics"])
        log(f"  → {found}/{len(metrics)} akun ditemukan")
        if progress_callback:
            progress_callback(i, total)

    # Build summary
    summary_data = []
    for result in all_results:
        row = {"Company": result["company"], "Folder": result["folder"]}
        for metric in metrics:
            row[metric] = result["metrics"].get(metric, {}).get("value", None)
        summary_data.append(row)

    df = pd.DataFrame(summary_data)
    out_file = out / "00_SUMMARY_all_companies.csv"
    df.to_csv(out_file, index=False, encoding="utf-8-sig")
    log(f"\n✓ Summary → {out_file}")
    log(f"✓ {len(all_results)} perusahaan, {len(metrics)} akun")
    return str(out_file)
