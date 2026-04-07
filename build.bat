@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo.
echo  ============================================================
echo   IDX Superapp ^| Build EXE
echo  ============================================================
echo.

REM ── Cek Python ──────────────────────────────────────────────────
where python >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python tidak ditemukan. Install Python 3.10+ terlebih dahulu.
    echo          https://www.python.org/downloads/
    pause & exit /b 1
)
for /f "tokens=*" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo  [OK] %PYVER%

REM ── Gunakan venv jika ada, fallback ke Python sistem ────────────
if exist ".venv\Scripts\python.exe" (
    set PYTHON=.venv\Scripts\python.exe
    set PIP=.venv\Scripts\pip.exe
    set PYINST=.venv\Scripts\pyinstaller.exe
    echo  [OK] Menggunakan virtual environment (.venv)
) else (
    set PYTHON=python
    set PIP=pip
    set PYINST=pyinstaller
    echo  [INFO] .venv tidak ditemukan, menggunakan Python sistem
)

echo.
echo  ── Step 1/3: Install dependencies ──────────────────────────
%PIP% install -r requirements.txt --quiet
if errorlevel 1 (
    echo  [ERROR] Gagal install dependencies!
    pause & exit /b 1
)
echo  [OK] Dependencies terpasang.

echo.
echo  ── Step 2/3: Build EXE ──────────────────────────────────────
if exist "dist\IDX_Superapp.exe" del /f /q "dist\IDX_Superapp.exe"
if exist "build" rd /s /q "build"

%PYINST% IDX_Superapp.spec --noconfirm
if errorlevel 1 (
    echo.
    echo  [ERROR] Build gagal! Lihat pesan di atas.
    pause & exit /b 1
)

echo.
echo  ── Step 3/3: Siapkan folder release ─────────────────────────
if not exist "release" mkdir release

REM Salin EXE
copy /y "dist\IDX_Superapp.exe" "release\IDX_Superapp.exe" >nul
if errorlevel 1 (
    echo  [ERROR] Gagal menyalin EXE ke folder release!
    pause & exit /b 1
)

REM Semua file data sudah di-bundle ke dalam EXE (via spec datas)
REM Tidak perlu salin file xlsx terpisah

echo.
echo  ============================================================
echo   BUILD BERHASIL!
echo  ============================================================
echo.
echo   File EXE tersimpan di:
echo     release\IDX_Superapp.exe
echo.

for %%f in ("release\IDX_Superapp.exe") do (
    set /a SIZE=%%~zf / 1048576
    echo   Ukuran : !SIZE! MB
)

echo.
echo   Cara distribusi:
echo     Cukup salin seluruh folder 'release\' ke komputer tujuan.
echo     Tidak perlu install Python atau library apapun.
echo.
pause
