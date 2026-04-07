"""
core/data_cleaner.py
Bersihkan nilai negatif pada kolom beban bunga di CSV output.

Kolom beban bunga selalu diubah ke nilai absolut (positif) secara otomatis —
diperlukan agar kalkulasi Altman Z-Score dan IBD menghasilkan nilai yang benar.
"""
import pandas as pd
from pathlib import Path

# Kolom yang selalu dibersihkan (nilai negatif → positif)
COLUMNS_TO_CLEAN = [
    "Beban bunga dan keuangan",
    "Beban bunga",
]


def clean_data(csv_path: str, log_callback=None) -> bool:
    """
    Ubah nilai negatif menjadi positif untuk kolom beban bunga.

    Args:
        csv_path: Path ke file CSV
        log_callback: Fungsi logging

    Returns:
        True jika berhasil, False jika gagal.
    """
    def log(msg):
        if log_callback: log_callback(msg)
        else: print(msg)

    path = Path(csv_path)
    if not path.exists():
        log(f"File tidak ditemukan: {csv_path}")
        return False

    try:
        df = pd.read_csv(csv_path, encoding="utf-8-sig")
        changed = []
        for col in COLUMNS_TO_CLEAN:
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
                log(f"  ✓ '{col}' → nilai absolut")

        if not changed:
            log("  Kolom beban bunga tidak ditemukan di CSV — dilewati.")
        df.to_csv(csv_path, index=False, encoding="utf-8-sig")
        log(f"✓ File diperbarui: {csv_path}")
        return True

    except Exception as e:
        log(f"✗ Error saat cleaning: {e}")
        return False
