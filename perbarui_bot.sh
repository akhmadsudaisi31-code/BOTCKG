#!/bin/bash
# =============================================================
#  SKRIP PEMBARUAN OTOMATIS BOT CKG KEMENKES
#  Jalankan skrip ini untuk mengunduh versi terbaru:
#     bash perbarui_bot.sh
# =============================================================

set -e

echo "🔄 Menarik pembaruan dari repositori Git (git pull)..."
# Disable credential prompts to avoid hanging if credentials aren't set
export GIT_TERMINAL_PROMPT=0
export GIT_SSH_COMMAND="ssh -o BatchMode=yes"

git pull origin main || git pull

if [ -d "venv" ]; then
    echo "📦 Memperbarui pustaka dependensi Python..."
    ./venv/bin/pip install --upgrade pip
    ./venv/bin/pip install -r requirements.txt
    
    echo "🌐 Memperbarui browser Chromium Playwright..."
    ./venv/bin/playwright install chromium
else
    echo "⚠️ Virtual environment 'venv' tidak ditemukan. Silakan jalankan install.sh terlebih dahulu."
fi

echo "✅ Pembaruan selesai! Silakan jalankan bot dengan: bash start_bot.sh"
