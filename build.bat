@echo off
REM ============================================================
REM  IDX Superapp - Build Script
REM  Menghasilkan IDX_Superapp.exe di folder dist/
REM ============================================================

echo ============================================================
echo  IDX Superapp - Build EXE
echo ============================================================
echo.

REM Pindah ke folder app (lokasi script ini)
cd /d "%~dp0"

echo [1/3] Menginstall dependencies...
pip install -r requirements.txt
echo.

echo [2/3] Membuild EXE dengan PyInstaller...
pyinstaller ^
  --onefile ^
  --windowed ^
  --name "IDX_Superapp" ^
  --add-data "..\akun_indonesia.txt;." ^
  --add-data "core;core" ^
  --add-data "gui;gui" ^
  --hidden-import curl_cffi ^
  --hidden-import bs4 ^
  --hidden-import pandas ^
  --hidden-import openpyxl ^
  main.py

echo.
echo [3/3] Selesai!
echo.

if exist "dist\IDX_Superapp.exe" (
    echo  SUCCESS: dist\IDX_Superapp.exe berhasil dibuat!
    echo  Ukuran file:
    dir dist\IDX_Superapp.exe | findstr IDX_Superapp
) else (
    echo  ERROR: Build gagal. Cek output PyInstaller di atas.
)

echo.
echo Tekan sembarang tombol untuk keluar...
pause > nul
