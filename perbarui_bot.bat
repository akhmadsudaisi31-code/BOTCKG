@echo off
:: =============================================================
::  SKRIP PEMBARUAN OTOMATIS BOT CKG KEMENKES (WINDOWS)
::  Jalankan file ini untuk mengunduh versi terbaru dari Git.
:: =============================================================
title Pembaruan Bot CKG Kemenkes

echo 🔄 Menarik pembaruan dari repositori Git (git pull)...
set GIT_TERMINAL_PROMPT=0
set GIT_SSH_COMMAND=ssh -o BatchMode=yes

git pull origin main
if %ERRORLEVEL% NEQ 0 (
    git pull
)

if exist venv (
    echo 📦 Memperbarui pustaka dependensi Python...
    call venv\Scripts\python.exe -m pip install --upgrade pip
    call venv\Scripts\pip.exe install -r requirements.txt
    
    echo 🌐 Memperbarui browser Chromium Playwright...
    call venv\Scripts\playwright.exe install chromium
) else (
    echo ⚠️ Virtual environment 'venv' tidak ditemukan. Silakan jalankan install_windows.bat terlebih dahulu.
)

echo ✅ Pembaruan selesai!
pause
