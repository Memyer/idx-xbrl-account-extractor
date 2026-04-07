@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo.
echo  ============================================================
echo   IDX Superapp ^| Setup Installer
echo   Menyiapkan environment Python untuk menjalankan aplikasi
echo  ============================================================
echo.

REM ── Cek Python ──────────────────────────────────────────────────
where python >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python tidak ditemukan!
    echo.
    echo  Silakan install Python 3.10 atau lebih baru:
    echo    https://www.python.org/downloads/
    echo.
    echo  PENTING: Centang "Add Python to PATH" saat instalasi.
    echo.
    pause & exit /b 1
)

for /f "tokens=2 delims= " %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo  [OK] Python %PYVER% ditemukan

REM Cek versi minimum 3.10
for /f "tokens=1,2 delims=." %%a in ("%PYVER%") do (
    if %%a LSS 3 (
        echo  [ERROR] Python versi minimal 3.10 diperlukan. Versi saat ini: %PYVER%
        pause & exit /b 1
    )
    if %%a EQU 3 if %%b LSS 10 (
        echo  [ERROR] Python versi minimal 3.10 diperlukan. Versi saat ini: %PYVER%
        pause & exit /b 1
    )
)

echo.
echo  ── Step 1/3: Membuat virtual environment ───────────────────
if exist ".venv" (
    echo  [INFO] Virtual environment sudah ada, dilewati.
) else (
    python -m venv .venv
    if errorlevel 1 (
        echo  [ERROR] Gagal membuat virtual environment!
        pause & exit /b 1
    )
    echo  [OK] Virtual environment dibuat.
)

echo.
echo  ── Step 2/3: Install dependencies ──────────────────────────
echo  [INFO] Menginstall paket Python, mohon tunggu...
.venv\Scripts\pip install --upgrade pip --quiet
.venv\Scripts\pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo  [ERROR] Gagal install dependencies!
    echo  Coba jalankan manual: pip install -r requirements.txt
    pause & exit /b 1
)
echo  [OK] Semua dependencies berhasil diinstall.

echo.
echo  ── Step 3/3: Buat shortcut jalankan aplikasi ───────────────
echo @echo off > run.bat
echo cd /d "%%~dp0" >> run.bat
echo .venv\Scripts\python main.py >> run.bat
echo  [OK] Shortcut 'run.bat' dibuat.

echo.
echo  ============================================================
echo   INSTALASI SELESAI!
echo  ============================================================
echo.
echo   Untuk menjalankan aplikasi:
echo     Klik dua kali file  run.bat
echo     atau jalankan:  .venv\Scripts\python main.py
echo.
pause
