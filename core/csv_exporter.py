"""
core/csv_exporter.py
Convert summary CSV to Excel (XLSX) — keduanya (CSV + XLSX) tersimpan.
"""
import pandas as pd
from pathlib import Path
import shutil


def export_to_excel(csv_path: str, excel_path: str = None, log_callback=None) -> str:
    """
    Convert CSV ke Excel (.xlsx). File CSV tetap disimpan.

    Args:
        csv_path: Path ke file CSV input (sudah ada dari step extract)
        excel_path: Path output Excel (default: sama tapi .xlsx)
        log_callback: Fungsi logging

    Returns:
        Path file Excel yang dibuat, atau None jika gagal.
    """
    def log(msg):
        if log_callback: log_callback(msg)
        else: print(msg)

    csv_p = Path(csv_path)
    if not csv_p.exists():
        log(f"File CSV tidak ditemukan: {csv_path}")
        return None

    if excel_path is None:
        excel_path = str(csv_p.with_suffix(".xlsx"))

    try:
        log(f"Membaca CSV: {csv_path}")
        df = pd.read_csv(csv_path, encoding="utf-8-sig")
        log(f"  → {len(df)} baris, {len(df.columns)} kolom")

        log(f"Menyimpan ke Excel: {excel_path}")
        with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Data Keuangan")
            # Auto-fit kolom
            ws = writer.sheets["Data Keuangan"]
            for col_cells in ws.columns:
                max_len = max(
                    (len(str(cell.value)) if cell.value else 0) for cell in col_cells
                )
                ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 40)

        log(f"✓ Excel berhasil dibuat   : {excel_path}")
        log(f"✓ CSV tetap tersimpan     : {csv_path}")
        return excel_path

    except Exception as e:
        log(f"✗ Error saat export Excel: {e}")
        return None

