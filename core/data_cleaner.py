"""
core/data_cleaner.py
Bersihkan nilai negatif pada kolom beban bunga di CSV output (opsional).
"""
import pandas as pd
from pathlib import Path

ALL_CLEANABLE_COLUMNS = [
    "Beban bunga dan keuangan",
    "Beban bunga",
]


def clean_data(csv_path: str, columns_to_clean: list = None, log_callback=None) -> bool:
    """
    Ubah nilai negatif menjadi positif pada kolom yang dipilih.

    Args:
        csv_path: Path ke file CSV
        columns_to_clean: List nama kolom yang akan dibersihkan.
                          Jika None, gunakan semua default kolom beban bunga.
        log_callback: Fungsi logging

    Returns:
        True jika berhasil, False jika gagal.
    """
    def log(msg):
        if log_callback: log_callback(msg)
        else: print(msg)

    if columns_to_clean is None:
        columns_to_clean = ALL_CLEANABLE_COLUMNS

    if not columns_to_clean:
        log("  Tidak ada kolom yang dipilih untuk dibersihkan, step dilewati.")
        return True

    path = Path(csv_path)
    if not path.exists():
        log(f"File tidak ditemukan: {csv_path}")
        return False

    try:
        df = pd.read_csv(csv_path, encoding="utf-8-sig")
        changed = []
        for col in columns_to_clean:
            if col in df.columns:
                def _to_abs(v):
                    if pd.isna(v):
                        return v
                    try:
                        return abs(float(v))
                    except (ValueError, TypeError):
                        return v
                df[col] = df[col].apply(_to_abs)
                changed.append(col)
                log(f"  ✓ Kolom '{col}' diubah ke nilai positif")

        if not changed:
            log("  Tidak ada kolom yang cocok ditemukan di CSV.")
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        log(f"✓ File berhasil diperbarui: {csv_path}")
        return True

    except Exception as e:
        log(f"✗ Error saat cleaning: {e}")
        return False

