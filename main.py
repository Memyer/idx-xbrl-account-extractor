"""
main.py — Entry point for IDX Superapp EXE
"""
import sys
import os

# Pastikan direktori app ditambahkan ke path agar import berjalan saat frozen (EXE)
if getattr(sys, "frozen", False):
    # Jika dijalankan sebagai EXE via PyInstaller
    base_path = os.path.dirname(sys.executable)
else:
    base_path = os.path.dirname(os.path.abspath(__file__))

sys.path.insert(0, base_path)

from gui.app import run

if __name__ == "__main__":
    run()
