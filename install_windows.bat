@echo off
chcp 65001 >nul
title INSTALLER BOT CKG KEMENKES

echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║  🤖  INSTALLER BOT CKG KEMENKES              ║
echo  ║      sehatindonesiaku.kemkes.go.id           ║
echo  ╚══════════════════════════════════════════════╝
echo.

:: Pindah ke folder tempat file .bat ini berada
cd /d "%~dp0"

:: ── Cek Python ───────────────────────────────────────────
echo [1/5] Memeriksa Python...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo  ❌ Python tidak ditemukan!
    echo.
    echo  Silakan download dan instal Python dari:
    echo  https://www.python.org/downloads/
    echo.
    echo  PENTING: Saat instalasi, centang "Add Python to PATH"
    echo.
    pause
    start https://www.python.org/downloads/
    exit /b 1
)
for /f "tokens=*" %%i in ('python --version') do echo  ✅ Ditemukan: %%i

:: ── Buat Virtual Environment ──────────────────────────────
echo.
echo [2/5] Membuat virtual environment...
if not exist "venv" (
    python -m venv venv
    echo  ✅ Virtual environment berhasil dibuat
) else (
    echo  ✅ Virtual environment sudah ada, dilewati
)

:: ── Install dependensi Python ─────────────────────────────
echo.
echo [3/5] Menginstal dependensi Python...
echo      (Proses ini mungkin memakan waktu 1-3 menit, harap tunggu)
venv\Scripts\pip install --upgrade pip -q
venv\Scripts\pip install -r requirements.txt -q
if %errorlevel% neq 0 (
    echo  ❌ Gagal menginstal dependensi. Periksa koneksi internet Anda.
    pause
    exit /b 1
)
echo  ✅ Dependensi berhasil diinstal

:: ── Install Playwright + Chromium ────────────────────────
echo.
echo [4/5] Menginstal browser Chromium (~150 MB, harap tunggu)...
venv\Scripts\playwright install chromium
if %errorlevel% neq 0 (
    echo  ❌ Gagal menginstal Chromium.
    pause
    exit /b 1
)
echo  ✅ Browser Chromium berhasil diinstal

:: ── Buat konfigurasi awal ─────────────────────────────────
echo.
echo [5/5] Memeriksa file konfigurasi...
if not exist "gui_config.json" (
    (
        echo {
        echo     "email": "",
        echo     "password": "",
        echo     "default_phone": "081234567890",
        echo     "excel_path": "ckg.xlsx",
        echo     "auto_advance": true,
        echo     "headless": false
        echo }
    ) > gui_config.json
    echo  ⚠️  File gui_config.json dibuat. Isi email dan password Kemenkes Anda!
) else (
    echo  ✅ gui_config.json sudah ada
)

:: ── Buat shortcut/launcher ────────────────────────────────
echo @echo off > jalankan_bot.bat
echo cd /d "%%~dp0" >> jalankan_bot.bat
echo venv\Scripts\python gui.py >> jalankan_bot.bat
echo  ✅ File jalankan_bot.bat dibuat

:: ── Selesai ───────────────────────────────────────────────
echo.
echo  ╔══════════════════════════════════════════════╗
echo  ║  ✅  INSTALASI BERHASIL!                     ║
echo  ╚══════════════════════════════════════════════╝
echo.
echo  Langkah selanjutnya:
echo  1. Buka file gui_config.json dengan Notepad
echo     dan isi email + password login Kemenkes Anda
echo  2. Pastikan file ckg.xlsx sudah ada di folder ini
echo  3. Jalankan bot dengan klik dua kali: jalankan_bot.bat
echo.
pause
