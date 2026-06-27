@echo off
chcp 65001 >nul
title Installer Bot CKG Kemenkes

:: ============================================================
::  Pindah ke folder tempat file .bat ini berada
:: ============================================================
cd /d "%~dp0"

cls
echo.
echo  ================================================================
echo   🤖  INSTALLER BOT CKG KEMENKES  (ASIK sehatindonesiaku.go.id)
echo  ================================================================
echo.
echo  Halo! Installer ini akan menyiapkan Bot CKG di laptop Anda.
echo  Proses ini hanya perlu dilakukan SATU KALI.
echo  Pastikan laptop terhubung ke internet sebelum melanjutkan.
echo.
echo  ⏳ Estimasi waktu: 3-10 menit (tergantung kecepatan internet)
echo.
pause

:: ============================================================
::  LANGKAH 1 - Cek Python
:: ============================================================
cls
echo.
echo  ┌─────────────────────────────────────────────────────┐
echo  │  LANGKAH 1/4 — Memeriksa Python...                 │
echo  └─────────────────────────────────────────────────────┘
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    cls
    echo.
    echo  ================================================================
    echo   ❌  PYTHON BELUM TERINSTAL
    echo  ================================================================
    echo.
    echo  Python adalah program wajib agar Bot CKG bisa berjalan.
    echo.
    echo  Cara install Python (mudah, gratis):
    echo.
    echo   1. Browser akan terbuka ke halaman download Python
    echo   2. Klik tombol kuning besar "Download Python"
    echo   3. Buka file yang diunduh (installer Python)
    echo   4. ⚠️  PENTING: Centang "Add Python to PATH" di bagian BAWAH
    echo   5. Klik "Install Now" dan tunggu selesai
    echo   6. Setelah selesai, TUTUP jendela hitam ini
    echo      lalu klik 2x file install_windows.bat lagi dari awal
    echo.
    echo  ================================================================
    echo.
    pause
    start https://www.python.org/downloads/
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version 2^>^&1') do set PYVER=%%i
echo   ✅ Python ditemukan: %PYVER%
echo.
timeout /t 1 >nul

:: ============================================================
::  LANGKAH 2 - Buat Virtual Environment
:: ============================================================
cls
echo.
echo  ┌─────────────────────────────────────────────────────┐
echo  │  LANGKAH 2/4 — Menyiapkan lingkungan Python...     │
echo  └─────────────────────────────────────────────────────┘
echo.

if not exist "venv" (
    echo   Membuat virtual environment baru...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo.
        echo  ❌  Gagal membuat virtual environment.
        echo.
        echo  Coba langkah berikut:
        echo   - Pastikan Python terinstal dengan benar
        echo   - Coba instal ulang Python dari python.org
        echo   - Pastikan tidak ada antivirus yang menghalangi
        echo.
        pause
        exit /b 1
    )
    echo   ✅ Lingkungan Python berhasil dibuat.
) else (
    echo   ✅ Lingkungan Python sudah ada, dilewati.
)
echo.
timeout /t 1 >nul

:: ============================================================
::  LANGKAH 3 - Install Dependensi
:: ============================================================
cls
echo.
echo  ┌─────────────────────────────────────────────────────┐
echo  │  LANGKAH 3/4 — Mengunduh paket-paket yang           │
echo  │                diperlukan... (Harap tunggu)          │
echo  └─────────────────────────────────────────────────────┘
echo.
echo   Sedang mengunduh dan menginstal program pendukung...
echo   Jangan tutup jendela ini sampai selesai!
echo.

:: Upgrade pip diam-diam
call venv\Scripts\python.exe -m pip install --upgrade pip --quiet 2>nul

:: Install requirements -- tampilkan output agar terlihat progresnya
call venv\Scripts\pip install -r requirements.txt
if %errorlevel% neq 0 (
    cls
    echo.
    echo  ================================================================
    echo   ❌  GAGAL MENGUNDUH PAKET PROGRAM
    echo  ================================================================
    echo.
    echo  Kemungkinan penyebab dan solusi:
    echo.
    echo   1. ❌ Tidak ada internet
    echo      → Pastikan laptop Anda terhubung ke WiFi/data
    echo.
    echo   2. ❌ Antivirus memblokir unduhan
    echo      → Coba nonaktifkan sementara antivirus, lalu jalankan
    echo        install_windows.bat lagi
    echo.
    echo   3. ❌ Python versi terlalu lama (butuh 3.9 - 3.12)
    echo      → Unduh Python terbaru dari python.org dan instal ulang
    echo.
    echo  Jika masih gagal, hubungi Admin dengan mengirim foto
    echo  tampilan jendela hitam ini ke WhatsApp.
    echo.
    pause
    exit /b 1
)
echo.
echo   ✅ Paket program berhasil diinstal.
echo.
timeout /t 1 >nul

:: ============================================================
::  LANGKAH 4 - Install Browser Chromium
:: ============================================================
cls
echo.
echo  ┌─────────────────────────────────────────────────────┐
echo  │  LANGKAH 4/4 — Mengunduh browser otomatis...       │
echo  │               (Ukuran ±150 MB, harap tunggu)        │
echo  └─────────────────────────────────────────────────────┘
echo.
echo   Bot menggunakan browser tersembunyi untuk mengisi CKG.
echo   Sedang mengunduh browser (mungkin 2-5 menit)...
echo   Jangan tutup jendela ini sampai selesai!
echo.

call venv\Scripts\playwright install chromium
if %errorlevel% neq 0 (
    cls
    echo.
    echo  ================================================================
    echo   ❌  GAGAL MENGUNDUH BROWSER OTOMATIS
    echo  ================================================================
    echo.
    echo  Kemungkinan penyebab:
    echo   - Koneksi internet terputus di tengah proses
    echo   - Antivirus memblokir unduhan file besar
    echo.
    echo  Solusi: Jalankan install_windows.bat lagi dari awal.
    echo  Browser yang sudah diunduh sebagian akan dilanjutkan otomatis.
    echo.
    pause
    exit /b 1
)
echo.
echo   ✅ Browser otomatis berhasil diinstal.
echo.
timeout /t 1 >nul

:: ============================================================
::  Buat file konfigurasi jika belum ada
:: ============================================================
if not exist "gui_config.json" (
    (
        echo {
        echo     "email": "",
        echo     "password": "",
        echo     "default_phone": "",
        echo     "excel_path": "ckg.xlsx",
        echo     "auto_advance": true,
        echo     "headless": false
        echo }
    ) > gui_config.json
)

:: ============================================================
::  Buat file jalankan_bot.bat
:: ============================================================
(
    echo @echo off
    echo chcp 65001 ^>nul
    echo cd /d "%%~dp0"
    echo title Bot CKG Kemenkes
    echo echo.
    echo echo  Memulai Bot CKG... Harap tunggu sebentar.
    echo echo.
    echo venv\Scripts\python.exe gui.py
    echo if %%errorlevel%% neq 0 ^(
    echo     echo.
    echo     echo  Bot berhenti dengan error. Hubungi Admin.
    echo     pause
    echo ^)
) > jalankan_bot.bat

:: ============================================================
::  Selesai!
:: ============================================================
cls
echo.
echo  ================================================================
echo   ✅  INSTALASI BERHASIL!
echo  ================================================================
echo.
echo  Bot CKG sudah siap digunakan di laptop ini.
echo.
echo  ─────────────────────────────────────────────────────────────
echo   LANGKAH SELANJUTNYA:
echo  ─────────────────────────────────────────────────────────────
echo.
echo   🔑  AKTIVASI (hanya 1x):
echo       Klik 2x file  jalankan_bot.bat
echo       Ikuti petunjuk aktivasi di layar
echo       (Hubungi Admin untuk mendapatkan kunci aktivasi)
echo.
echo   📂  Pastikan file ckg.xlsx ada di folder ini sebelum
echo       menjalankan bot.
echo.
echo   ▶️   Setelah aktif, untuk menjalankan bot cukup:
echo       Klik 2x →  jalankan_bot.bat
echo.
echo  ================================================================
echo.
pause
