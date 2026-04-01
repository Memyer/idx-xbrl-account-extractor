"""
core/unzipper.py
Extract ZIP files from FinancialStatements to Folder_XBRL.
"""
import os
import zipfile
from pathlib import Path


def unzip_all(source_dir: str, dest_dir: str,
              progress_callback=None, log_callback=None,
              stop_flag=None) -> dict:
    """
    Ekstrak semua ZIP dari source_dir ke dest_dir.

    Returns:
        Dict: success, skipped, errors
    """
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    src = Path(source_dir)
    dst = Path(dest_dir)
    dst.mkdir(parents=True, exist_ok=True)

    stats = {"success": 0, "skipped": 0, "errors": 0}

    zip_files = list(src.glob("*.zip"))
    if not zip_files:
        log(f"Tidak ada file ZIP ditemukan di: {source_dir}")
        return stats

    total = len(zip_files)
    log(f"Ditemukan {total} file ZIP untuk diekstrak.\n")

    for i, zip_path in enumerate(zip_files, 1):
        if stop_flag and stop_flag():
            log("Ekstraksi dihentikan oleh user.")
            break

        filename = zip_path.stem           # e.g. "AALI_inlineXBRL"
        code = filename.split("_")[0]      # e.g. "AALI"
        dest_folder = dst / f"{code}_InlineXBRL"

        if dest_folder.exists() and any(dest_folder.iterdir()):
            log(f"[{i}/{total}] {code} - Sudah ada, dilewati.")
            stats["skipped"] += 1
        else:
            try:
                dest_folder.mkdir(exist_ok=True)
                with zipfile.ZipFile(zip_path, "r") as zf:
                    zf.extractall(dest_folder)
                log(f"[{i}/{total}] {code} - Berhasil diekstrak.")
                stats["success"] += 1
            except zipfile.BadZipFile:
                log(f"[{i}/{total}] {code} - Error: File ZIP tidak valid.")
                stats["errors"] += 1
            except Exception as e:
                log(f"[{i}/{total}] {code} - Error: {e}")
                stats["errors"] += 1

        if progress_callback:
            progress_callback(i, total)

    log(f"\n--- Ringkasan Ekstraksi ---")
    log(f"Berhasil : {stats['success']}")
    log(f"Dilewati : {stats['skipped']}")
    log(f"Error    : {stats['errors']}")
    return stats
