"""
core/amount_normalizer.py
Normalize selected financial accounts to full amount IDR.
"""
import pandas as pd
from pathlib import Path

ROUNDING_COLUMN = "Pembulatan yang digunakan dalam penyajian jumlah dalam laporan keuangan"
CURRENCY_COLUMN = "Mata uang pelaporan"


def _normalize_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip().lower()


def _unit_multiplier(rounding_value) -> int:
    text = _normalize_text(rounding_value)
    if not text:
        return 1

    if "full amount" in text or "satuan" in text or "unit" in text or "penuh" in text:
        return 1
    if "ribu" in text or "thousand" in text:
        return 1_000
    if "juta" in text or "million" in text:
        return 1_000_000
    if "miliar" in text or "milyar" in text or "billion" in text:
        return 1_000_000_000
    if "triliun" in text or "trillion" in text:
        return 1_000_000_000_000
    return 1


def _currency_multiplier(currency_value, usd_rate: float) -> float:
    text = _normalize_text(currency_value)
    if not text:
        return 1.0

    if "idr" in text or "rupiah" in text:
        return 1.0
    if "usd" in text or text == "$" or "dollar" in text:
        return float(usd_rate)
    return 1.0


def normalize_to_full_idr(
    csv_path: str,
    selected_metrics: list,
    usd_rate: float,
    log_callback=None,
) -> bool:
    """
    Normalize selected metric columns in CSV to full amount and IDR.

    Rules:
    - Unit multiplier comes from ROUNDING_COLUMN
    - Currency multiplier comes from CURRENCY_COLUMN
    - Final value = original * unit_multiplier * currency_multiplier
    """

    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    path = Path(csv_path)
    if not path.exists():
        log(f"File tidak ditemukan: {csv_path}")
        return False

    if usd_rate <= 0:
        log("Nilai kurs USD harus lebih besar dari 0.")
        return False

    try:
        df = pd.read_csv(path, encoding="utf-8-sig")
    except Exception as e:
        log(f"Gagal membaca CSV: {e}")
        return False

    missing = [c for c in [ROUNDING_COLUMN, CURRENCY_COLUMN] if c not in df.columns]
    if missing:
        log(
            "Kolom referensi normalisasi tidak ditemukan: "
            + ", ".join(missing)
            + "."
        )
        return False

    if not selected_metrics:
        log("Tidak ada akun dipilih untuk dinormalisasi.")
        return True

    skipped = {"Company", "Folder", ROUNDING_COLUMN, CURRENCY_COLUMN}
    target_cols = [c for c in selected_metrics if c in df.columns and c not in skipped]

    if not target_cols:
        log("Tidak ada kolom akun yang cocok untuk dinormalisasi di CSV.")
        return True

    unit_factor = df[ROUNDING_COLUMN].apply(_unit_multiplier)
    currency_factor = df[CURRENCY_COLUMN].apply(lambda v: _currency_multiplier(v, usd_rate))
    total_factor = unit_factor * currency_factor

    converted_cells = 0
    for col in target_cols:
        numeric_values = pd.to_numeric(df[col], errors="coerce")
        mask = numeric_values.notna()
        if mask.any():
            df.loc[mask, col] = numeric_values[mask] * total_factor[mask]
            converted_cells += int(mask.sum())

    try:
        df.to_csv(path, index=False, encoding="utf-8-sig")
    except Exception as e:
        log(f"Gagal menyimpan CSV hasil normalisasi: {e}")
        return False

    log(f"✓ Normalisasi full amount + IDR selesai ({len(target_cols)} kolom akun).")
    log(f"✓ Total sel numerik yang dikonversi: {converted_cells}")
    log(f"✓ Kurs USD yang digunakan: {usd_rate}")
    return True
