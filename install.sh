#!/bin/bash
# =============================================================
#  INSTALLER BOT CKG KEMENKES
#  Jalankan sekali saja di laptop baru:
#     bash install.sh
# =============================================================

set -e  # hentikan jika ada error

CYAN='\033[0;36m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
BOLD='\033[1m'
NC='\033[0m'

banner() {
    echo ""
    echo -e "${CYAN}╔══════════════════════════════════════════════╗${NC}"
    echo -e "${CYAN}║  🤖  INSTALLER BOT CKG KEMENKES              ║${NC}"
    echo -e "${CYAN}║      sehatindonesiaku.kemkes.go.id           ║${NC}"
    echo -e "${CYAN}╚══════════════════════════════════════════════╝${NC}"
    echo ""
}

step() { echo -e "\n${BOLD}${CYAN}▶ $1${NC}"; }
ok()   { echo -e "${GREEN}  ✅ $1${NC}"; }
warn() { echo -e "${YELLOW}  ⚠️  $1${NC}"; }
err()  { echo -e "${RED}  ❌ $1${NC}"; exit 1; }

banner

# ── Cek sistem operasi ─────────────────────────────────────
step "Memeriksa sistem operasi..."
if [[ "$OSTYPE" != "linux-gnu"* ]]; then
    err "Installer ini hanya untuk Linux (Ubuntu/Debian). Untuk Windows gunakan install_windows.bat"
fi
ok "Sistem Linux terdeteksi"

# ── Cek Python 3 ────────────────────────────────────────────
step "Memeriksa Python 3..."
if ! command -v python3 &> /dev/null; then
    warn "Python 3 belum terinstal. Menginstal sekarang..."
    sudo apt-get update -q
    sudo apt-get install -y python3 python3-pip python3-venv python3-tk
fi
PYTHON_VER=$(python3 --version)
ok "Ditemukan: $PYTHON_VER"

# ── Cek Tkinter (GUI) ───────────────────────────────────────
step "Memeriksa Tkinter (library GUI)..."
if ! python3 -c "import tkinter" 2>/dev/null; then
    warn "Tkinter belum ada. Menginstal..."
    sudo apt-get install -y python3-tk
fi
ok "Tkinter tersedia"

# ── Buat Virtual Environment ────────────────────────────────
step "Membuat virtual environment Python..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
    ok "Virtual environment 'venv' berhasil dibuat"
else
    ok "Virtual environment 'venv' sudah ada, dilewati"
fi

# ── Install dependensi Python ───────────────────────────────
step "Menginstal dependensi Python (playwright, pandas, openpyxl, Pillow)..."
venv/bin/pip install --upgrade pip -q
venv/bin/pip install -r requirements.txt -q
ok "Semua dependensi Python berhasil diinstal"

# ── Install Browser Chromium via Playwright ─────────────────
step "Menginstal browser Chromium (hanya sekali, ~150 MB)..."
venv/bin/playwright install chromium
ok "Browser Chromium berhasil diinstal"

# ── Instal dependensi sistem untuk Playwright ───────────────
step "Menginstal dependensi sistem untuk browser..."
venv/bin/playwright install-deps chromium 2>/dev/null || \
    sudo apt-get install -y \
        libglib2.0-0 libnss3 libnspr4 libdbus-1-3 \
        libatk1.0-0 libatk-bridge2.0-0 libexpat1 \
        libx11-6 libxcomposite1 libxdamage1 libxext6 \
        libxfixes3 libxrandr2 libgbm1 libxcb1 \
        libxkbcommon0 libpango-1.0-0 libcairo2 libasound2 2>/dev/null || true
ok "Dependensi sistem selesai"

# ── Buat file konfigurasi awal jika belum ada ──────────────
step "Memeriksa file konfigurasi..."
if [ ! -f "gui_config.json" ]; then
    cat > gui_config.json << 'EOF'
{
    "email": "",
    "password": "",
    "default_phone": "081234567890",
    "excel_path": "ckg.xlsx",
    "auto_advance": true,
    "headless": false
}
EOF
    warn "File gui_config.json dibuat. Isi email dan password login Kemenkes Anda sebelum menjalankan bot!"
else
    ok "gui_config.json sudah ada"
fi

# ── Buat script launcher ─────────────────────────────────────
step "Membuat shortcut peluncur (start_bot.sh)..."
cat > start_bot.sh << 'LAUNCHER'
#!/bin/bash
cd "$(dirname "$0")"
./venv/bin/python gui.py
LAUNCHER
chmod +x start_bot.sh
ok "start_bot.sh siap digunakan"

# ── Selesai ─────────────────────────────────────────────────
echo ""
echo -e "${GREEN}╔══════════════════════════════════════════════╗${NC}"
echo -e "${GREEN}║  ✅  INSTALASI BERHASIL!                     ║${NC}"
echo -e "${GREEN}╚══════════════════════════════════════════════╝${NC}"
echo ""
echo -e "${BOLD}Langkah selanjutnya:${NC}"
echo -e "  1. Buka file ${YELLOW}gui_config.json${NC} dan isi email & password login Kemenkes"
echo -e "  2. Pastikan file ${YELLOW}ckg.xlsx${NC} sudah ada di folder ini"
echo -e "  3. Jalankan bot dengan perintah:"
echo -e "     ${CYAN}bash start_bot.sh${NC}"
echo -e "     atau klik dua kali file ${CYAN}start_bot.sh${NC}"
echo ""
