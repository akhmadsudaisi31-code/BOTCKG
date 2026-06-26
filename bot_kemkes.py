"""
Bot CKG — sehatindonesiaku.kemkes.go.id
Bot + Kontrol Manual via Terminal

Titik konfirmasi:
  1. Setelah Cek NIK     → cek data yang muncul di browser
  2. Sebelum submit form → verifikasi akhir semua isian
  3. Saat error          → pilih: coba lagi / skip / berhenti

Jalankan: venv/bin/python bot_kemkes.py
"""

import asyncio
import re
import sys
import base64
import io
from pathlib import Path
from datetime import datetime

import pandas as pd
from playwright.async_api import async_playwright, Page, Browser, BrowserContext
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from PIL import Image

# ============================================================
#  GUI INTEGRATION BRIDGE
# ============================================================
GUI_BRIDGE_ACTIVE = False
gui_bridge = None
current_page = None
_current_input_task = None  # asyncio.Task yang sedang menjalankan input_pasien

class GUIBridge:
    def __init__(self):
        import queue
        self.log_queue = queue.Queue()
        self.request_queue = queue.Queue()
        self.response_queue = queue.Queue()
        self.running = True
        self.auto_advance = True

    def request_input(self, req_type, data=None):
        self.request_queue.put((req_type, data))
        response = self.response_queue.get()
        return response

    def log(self, msg, level="INFO"):
        self.log_queue.put((msg, level))

# ============================================================
#  KONFIGURASI (Dimuat dinamis dari gui_config.json)
# ============================================================
import json

USERNAME = ""
PASSWORD = ""
BASE_URL  = "https://sehatindonesiaku.kemkes.go.id"
HEADLESS  = False
TIMEOUT   = 30_000
SCREENSHOT_DIR = Path("screenshots")
FILE_EXCEL      = Path("ckg.xlsx")
JEDA_ANTAR_DATA = 1500   # ms jeda antar pasien
MULAI_DARI      = 1      # baris ke-berapa (No.) untuk mulai, berguna jika resume

# Muat kredensial dan konfigurasi dari gui_config.json untuk menghindari hardcode kredensial
CONFIG_FILE = Path("gui_config.json")
if CONFIG_FILE.exists():
    try:
        with open(CONFIG_FILE, "r") as f:
            config_data = json.load(f)
            USERNAME = config_data.get("email", USERNAME)
            PASSWORD = config_data.get("password", PASSWORD)
            if "excel_path" in config_data:
                FILE_EXCEL = Path(config_data["excel_path"])
            if "headless" in config_data:
                HEADLESS = config_data["headless"]
    except Exception:
        pass

BULAN_ID = {
    1:"Jan", 2:"Feb", 3:"Mar", 4:"Apr", 5:"Mei", 6:"Jun",
    7:"Jul", 8:"Agt", 9:"Sep", 10:"Okt", 11:"Nov", 12:"Des"
}

# Fallback — urutan sesuai prioritas (nama paling umum duluan)
BULAN_FALLBACK = {
    1: ["Jan", "Januari"],
    2: ["Feb", "Februari"],
    3: ["Mar", "Maret"],
    4: ["Apr", "April"],
    5: ["Mei", "May"],
    6: ["Jun", "Juni"],
    7: ["Jul", "Juli"],
    8: ["Agt", "Agu", "Aug", "Agustus"],   # ← Agt sesuai screenshot
    9: ["Sep", "Sept", "September"],
    10:["Okt", "Oct", "Oktober"],
    11:["Nov", "November"],
    12:["Des", "Dec", "Desember"],
}

# ============================================================
#  TERMINAL UI
# ============================================================

# ANSI escape codes for beautiful terminal styling
RESET = "\033[0m"
BOLD = "\033[1m"
UNDERLINE = "\033[4m"

# Colors
RED = "\033[91m"
GREEN = "\033[92m"
YELLOW = "\033[93m"
BLUE = "\033[94m"
MAGENTA = "\033[95m"
CYAN = "\033[96m"
WHITE = "\033[97m"
GRAY = "\033[90m"

GARIS  = "─" * 58
GARIS2 = "═" * 58

def log(msg: str, level: str = "INFO"):
    ts_time = datetime.now().strftime("%H:%M:%S")
    ts_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ikon = {"INFO":"ℹ️ ","OK":"✅","WARN":"⚠️ ","ERR":"❌","WAIT":"⏳","BOT":"🤖"}.get(level,"  ")
    
    color = RESET
    if level == "INFO":
        color = CYAN
    elif level == "OK":
        color = GREEN
    elif level == "WARN":
        color = YELLOW
    elif level == "ERR":
        color = RED
    elif level == "BOT":
        color = BLUE
    elif level == "WAIT":
        color = MAGENTA

    # Cetak ke terminal
    print(f"{GRAY}[{ts_time}]{RESET} {color}{ikon} {msg}{RESET}")
    
    # Kirim ke GUI
    if GUI_BRIDGE_ACTIVE and gui_bridge:
        gui_bridge.log(msg, level)

    # Catat ke file log lokal (bot_run.log)
    try:
        clean_msg = re.sub(r'\033\[[0-9;]*m', '', msg)
        with open("bot_run.log", "a", encoding="utf-8") as log_file:
            log_file.write(f"[{ts_date}] [{level}] {clean_msg}\n")
    except Exception:
        pass


def print_box_line(content: str, color: str = YELLOW, text_color: str = WHITE, is_bold: bool = False, inner_width: int = 50):
    clean_content = re.sub(r'\033\[[0-9;]*m', '', content)
    pad_len = inner_width - len(clean_content)
    if pad_len < 0:
        content = content[:inner_width-3] + "..."
        pad_len = 0
    
    bold_code = BOLD if is_bold else ""
    print(f"  {color}║{RESET}  {bold_code}{text_color}{content}{" " * pad_len}{RESET}  {color}║{RESET}")


def tanya(prompt: str, pilihan: dict) -> str:
    inner_width = 50
    border = "═" * (inner_width + 4)
    print()
    print(f"  {YELLOW}╔{border}╗{RESET}")
    
    # Split prompt into lines
    lines = []
    current_line = ""
    for word in prompt.split(" "):
        if len(current_line) + len(word) + 1 > (inner_width - 5):
            lines.append(current_line)
            current_line = word
        else:
            current_line = (current_line + " " + word).strip()
    if current_line:
        lines.append(current_line)
        
    for i, line in enumerate(lines):
        icon = "⏸️  " if i == 0 else "    "
        print_box_line(f"{icon}{line}", color=YELLOW, text_color=WHITE, is_bold=True, inner_width=inner_width)
        
    print(f"  {YELLOW}╠{border}╣{RESET}")
    for kunci, label in pilihan.items():
        k = "ENTER" if kunci == "enter" else kunci.upper()
        print_box_line(f"[{k}] {label}", color=YELLOW, text_color=CYAN, is_bold=False, inner_width=inner_width)
    print(f"  {YELLOW}╚{border}╝{RESET}")

    while True:
        raw = input(f"\n  {BOLD}{YELLOW}>> Pilihan Anda:{RESET} ").strip().lower()
        if raw == "" and "enter" in pilihan:
            return "enter"
        if raw in pilihan:
            return raw
        print(f"  {RED}⚠️  Ketik salah satu: {list(pilihan.keys())} atau tekan ENTER{RESET}")


def cetak_data_pasien(nomor: int, total: int, nama: str, nik: str, jk: str, tgl: str):
    inner_width = 50
    border = "═" * (inner_width + 4)
    print()
    print(f"  {BLUE}╔{border}╗{RESET}")
    print_box_line(f"📋 PASIEN [{nomor}/{total}]", color=BLUE, text_color=WHITE, is_bold=True, inner_width=inner_width)
    print(f"  {BLUE}╠{border}╣{RESET}")
    
    print_box_line(f"Nama          : {nama}", color=BLUE, text_color=CYAN, is_bold=True, inner_width=inner_width)
    print_box_line(f"NIK           : {nik}", color=BLUE, text_color=YELLOW, is_bold=False, inner_width=inner_width)
    print_box_line(f"Jenis Kelamin : {jk}", color=BLUE, text_color=WHITE, is_bold=False, inner_width=inner_width)
    print_box_line(f"Tanggal Lahir : {tgl}", color=BLUE, text_color=GREEN, is_bold=False, inner_width=inner_width)
    
    print(f"  {BLUE}╚{border}╝{RESET}")
    print()

    # Kirim info pasien ke GUI bridge agar label Nama & NIK terupdate
    if GUI_BRIDGE_ACTIVE and gui_bridge:
        gui_bridge.log(f"__PATIENT_INFO__|{nomor}|{total}|{nama}|{nik}", "INFO")


# ============================================================
#  HELPER
# ============================================================

def image_to_binary_ascii(base64_str, width=50):
    try:
        if "base64," in base64_str:
            base64_str = base64_str.split("base64,")[1]
        img_bytes = base64.b64decode(base64_str)
        img = Image.open(io.BytesIO(img_bytes))
        img = img.convert("L")
        
        original_width, original_height = img.size
        aspect_ratio = original_height / original_width
        
        height = int(width * aspect_ratio)
        img = img.resize((width, height))
        
        pixels_list = list(img.getdata())
        min_val = min(pixels_list)
        max_val = max(pixels_list)
        threshold = min_val + (max_val - min_val) * 0.85
        
        pixels = img.load()
        lines = []
        for y in range(height):
            line = []
            for x in range(width):
                gray = pixels[x, y]
                if gray < threshold:
                    line.append("██")
                else:
                    line.append("  ")
            lines.append("".join(line))
        return "\n".join(lines)
    except Exception as e:
        return f"Error rendering CAPTCHA: {e}"


# Exception khusus agar input_pasien bisa sinyal SKIP ke loop utama
class SkipPasien(BaseException):
    pass

class StopBotException(BaseException):
    pass

def check_stop_request():
    if GUI_BRIDGE_ACTIVE and gui_bridge:
        if not gui_bridge.running:
            raise StopBotException("Bot dihentikan oleh user")
        if getattr(gui_bridge, "skip_requested", False) or getattr(gui_bridge, "skip_active", False):
            gui_bridge.skip_requested = False
            gui_bridge.skip_active = False
            raise SkipPasien("Skip diminta oleh user melalui GUI")


async def poll_gui_signals():
    """
    Background task yang memantau sinyal stop/skip dari GUI.
    Ketika skip diterima: cancel asyncio.Task input_pasien yang sedang berjalan
    agar CancelledError merambat keluar secara bersih tanpa tertahan oleh
    try-except Exception lokal di dalam input_pasien.
    """
    global current_page, _current_input_task
    while True:
        try:
            await asyncio.sleep(0.05)  # poll lebih cepat (50ms) agar lebih responsif
            if not GUI_BRIDGE_ACTIVE or not gui_bridge:
                continue

            # ── Cek request STOP ─────────────────────────────
            if not gui_bridge.running:
                if current_page and not current_page.is_closed():
                    try:
                        await current_page.close()
                    except Exception:
                        pass
                # Cancel task aktif jika ada
                if _current_input_task and not _current_input_task.done():
                    _current_input_task.cancel()
                break

            # ── Cek request SKIP ─────────────────────────────
            if getattr(gui_bridge, "skip_requested", False):
                gui_bridge.skip_requested = False
                gui_bridge.skip_active = True

                # 1. Cancel asyncio.Task input_pasien agar keluar bersih
                #    CancelledError tidak tertangkap oleh except Exception,
                #    sehingga pasti merambat ke loop utama.
                if _current_input_task and not _current_input_task.done():
                    _current_input_task.cancel()

                # 2. Navigasi browser ke form pendaftaran sekaligus
                #    agar halaman sudah siap ketika loop utama memanggil
                #    buka_form_daftar_baru() untuk pasien berikutnya.
                if current_page and not current_page.is_closed():
                    try:
                        await current_page.goto(
                            f"{BASE_URL}/ckg-pendaftaran-individu",
                            wait_until="domcontentloaded",
                            timeout=8000
                        )
                    except Exception:
                        pass  # halaman mungkin sudah pindah karena cancel

        except Exception:
            pass


async def screenshot(page: Page, nama: str):
    pass


async def force_js_click(locator) -> bool:
    """
    Memaksakan klik elemen menggunakan JavaScript injection untuk memotong Actionability checks Playwright.
    Sangat ampuh untuk mengatasi elemen yang tertutup overlay/backdrop/dialog lain.
    """
    try:
        if hasattr(locator, "first"):
            target = locator.first
        else:
            target = locator
        
        if await target.count() > 0:
            await target.evaluate("el => el.click()")
            return True
    except Exception:
        pass
    return False


async def klik_opsional(page: Page, locator, timeout: int = 4000):
    try:
        if not await force_js_click(locator):
            await locator.click(timeout=timeout)
        await page.wait_for_timeout(400)
    except Exception:
        pass


async def klik_popup_kuota(page: Page):
    """
    Deteksi popup 'Kuota Pemeriksaan Habis' dan tangani 2 tahap otomatis:
      Tahap 1 — popup kuota: klik 'Lanjut' (btn-outline-primary)
      Tahap 2 — konfirmasi warning: klik 'Lanjut' (btn-fill-warning)
    Memeriksa hingga 5 kali dan memverifikasi popup benar-benar tertutup.
    Menggunakan force_js_click untuk mencegah hambatan akibat overlay.
    """
    KUOTA_KEYWORDS = ["kuota pemeriksaan habis", "kuota pemeriksaan", "kuota habis"]

    async def popup_kuota_masih_ada() -> bool:
        """Cek apakah popup kuota masih terlihat di halaman."""
        try:
            # Cara paling andal: cek tombol 'Pilih Tanggal Lain' yang hanya ada di popup kuota
            btn_pilih_lain = page.locator("button").filter(
                has_text=re.compile(r"Pilih Tanggal Lain", re.IGNORECASE)
            ).first
            if await btn_pilih_lain.is_visible():
                return True
            # Cek lewat teks popup secara cepat menggunakan locator text regex
            popup_text_locator = page.locator("text=/kuota pemeriksaan habis|kuota pemeriksaan|kuota habis/i").first
            if await popup_text_locator.is_visible():
                return True
            return False
        except Exception:
            return False

    # Loop hingga popup kuota benar-benar tertutup (maks 5 percobaan)
    for attempt in range(1, 6):
        if not await popup_kuota_masih_ada():
            break  # popup sudah tidak ada, selesai

        log(f"  ⚠️  Popup 'Kuota Pemeriksaan Habis' terdeteksi (percobaan {attempt}/5).", "WARN")

        try:
            # Tahap 1: Klik tombol 'Lanjut' di dalam popup kuota
            pilih_lain = page.locator("button").filter(
                has_text=re.compile(r"Pilih Tanggal Lain", re.IGNORECASE)
            ).first

            if await pilih_lain.is_visible():
                # Popup kuota masih ada — cari tombol Lanjut
                lanjut_btn = page.locator(
                    "button.btn-outline-primary, div.z-1000 button, div.z-1100 button"
                ).filter(has_text=re.compile(r"\bLanjut\b", re.IGNORECASE)).first

                if await lanjut_btn.is_visible():
                    log("  🤖 Mengklik 'Lanjut' di popup kuota...", "BOT")
                    if not await force_js_click(lanjut_btn):
                        await lanjut_btn.click(timeout=2000)
                    await page.wait_for_timeout(600)
                else:
                    # Tidak ada Lanjut — pilih tanggal lain
                    log("  ⚠️  Tombol Lanjut tidak ada, mengklik 'Pilih Tanggal Lain'...", "WARN")
                    if not await force_js_click(pilih_lain):
                        await pilih_lain.click(timeout=2000)
                    await page.wait_for_timeout(500)
                    try:
                        grid = page.locator(".form-data-individu form .grid-cols-7.mt-2")
                        alt_btn = grid.locator("button").nth(13)
                        if await alt_btn.is_visible():
                            if not await force_js_click(alt_btn):
                                await alt_btn.click(timeout=2000)
                            await page.wait_for_timeout(500)
                            log("  ✓ Tanggal pemeriksaan alternatif berhasil dipilih.", "OK")
                    except Exception as e:
                        log(f"  ⚠️  Gagal memilih tanggal alternatif: {e}", "WARN")
        except Exception:
            pass

        try:
            # Tahap 2: Konfirmasi warning (btn-fill-warning atau div.z-99999)
            warning_btn = page.locator("button.btn-fill-warning").filter(
                has_text=re.compile(r"\bLanjut\b", re.IGNORECASE)
            ).first
            if not await warning_btn.is_visible():
                warning_btn = page.locator("div.z-99999 button").filter(
                    has_text=re.compile(r"\bLanjut\b", re.IGNORECASE)
                ).first
            if await warning_btn.is_visible():
                log("  🤖 Mengklik konfirmasi 'Lanjut' (warning)...", "BOT")
                if not await force_js_click(warning_btn):
                    await warning_btn.click(timeout=2000)
                await page.wait_for_timeout(600)
        except Exception:
            pass

        # Beri waktu halaman merespons sebelum cek ulang
        await page.wait_for_timeout(600)

    # Verifikasi akhir
    if await popup_kuota_masih_ada():
        log("  ⚠️  Popup kuota masih ada setelah 5 percobaan. Melanjutkan paksa...", "WARN")
    else:
        log("  ✓ Popup kuota berhasil ditutup.", "OK")


async def cek_sesi_berakhir(page: Page, username: str = None, password: str = None) -> bool:
    """
    Deteksi popup atau halaman 'Sesi Berakhir' / 'Session Expired'.
    Jika terdeteksi: login ulang otomatis dan kembalikan True.
    Jika tidak: kembalikan False.
    """
    try:
        sesi_popup = page.locator("text=/[Ss]esi.*[Bb]erakhir|[Ss]ession.*[Ee]xpired/").first
        login_redirect = "login" in page.url.lower()

        sesi_visible = False
        try:
            sesi_visible = await sesi_popup.is_visible()
        except Exception:
            pass

        if sesi_visible or login_redirect:
            log("  ⚠️  Sesi berakhir terdeteksi — login ulang...", "WARN")
            try:
                await page.get_by_role("button", name="Login").click(timeout=3000)
            except Exception:
                pass
            await page.wait_for_timeout(500)
            await login(page, username=username, password=password)
            log("  ✓ Login ulang berhasil.", "OK")
            return True
    except Exception:
        pass
    return False


async def cek_overlay_blok(page: Page):
    """
    Deteksi overlay gelap (backdrop) yang memblokir form.
    XPath: div[@class='fixed inset-0 bg-[rgba(0,0,0,0.8)] opacity-50 z-1000']
    Jika terlihat → raise SkipPasien agar baris di-skip dan diwarnai kuning.
    """
    try:
        overlay = page.locator(
            "xpath=//div[contains(@class,'fixed') and contains(@class,'inset-0') "
            "and contains(@class,'opacity-50') and contains(@class,'z-1000')]"
        ).first
        if await overlay.is_visible():
            log("  ⚠️  Overlay blok terdeteksi (NIK sudah terdaftar / validasi gagal) — pasien di-skip.", "WARN")
            # Tutup modal jika ada tombol X
            try:
                await page.locator("button.btn-transparent").first.click(timeout=2000)
            except Exception:
                pass
            raise SkipPasien("Overlay blok: pasien tidak dapat didaftarkan")
    except SkipPasien:
        raise
    except Exception:
        pass


async def is_element_really_visible(locator) -> bool:
    """
    Checks if a locator points to an element that is actually visible to the user.
    Traverses parent elements to ensure no ancestor is hidden via display, visibility, or opacity.
    """
    try:
        if await locator.count() == 0:
            return False
        if not await locator.is_visible():
            return False
        
        visible = await locator.first.evaluate("""(element) => {
            let el = element;
            while (el) {
                if (el.nodeType === Node.ELEMENT_NODE) {
                    const style = window.getComputedStyle(el);
                    if (style.display === 'none' || 
                        style.visibility === 'hidden' || 
                        parseFloat(style.opacity) === 0) {
                        return false;
                    }
                }
                el = el.parentElement;
            }
            return element.offsetWidth > 0 && element.offsetHeight > 0;
        }""")
        return bool(visible)
    except Exception:
        try:
            return await locator.first.is_visible()
        except Exception:
            return False


async def cek_popup_periksa_kembali(page: Page):
    """
    Jika muncul modal khusus (misal .w-[40%]), klik tombol 'Periksa Kembali'.
    Berguna untuk menutup notifikasi validasi NIK/Data Peserta agar bisa lanjut mengisi form.
    """
    try:
        popup = page.locator(".w-\\[40\\%\\]").first
        if await is_element_really_visible(popup):
            log("  ⚠️  Popup validasi data terdeteksi, mengklik 'Periksa Kembali'...", "WARN")
            btn = page.get_by_role("button", name=re.compile(r"Periksa Kembali", re.IGNORECASE)).first
            if await is_element_really_visible(btn):
                await btn.click()
                await page.wait_for_timeout(500)
                log("  ✓ Popup 'Periksa Kembali' ditutup.", "OK")
    except Exception:
        pass


async def cek_popup_kategori_pasien(page: Page):
    """
    Jika muncul popup kategori sasaran bermasalah, klik 'Kembali' dan minta konfirmasi user
    apakah lolos atau skip.
    """
    xpath_txt_kategori = (
        "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
        "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
        "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
        "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
        "/div[@class='p-2']/div/div[@class='w-full']"
        "/div[@class='fixed top-0 left-0 w-full h-full flex justify-center items-center backdrop-blur-5 z-1000']"
        "/div[@class='w-[40%] max-h-full relative z-1000 rounded-lg bg-white p-4 shadow-gmail lt-md:w-[80%] lt-sm:w-full overflow-auto']"
        "/div[@class='p-2']/div/div[@class='flex flex-col justify-center items-center gap-3 text-[20px] font-bold mb-1']"
        "/div[@class='pb-2 text-center']"
    )
    xpath_btn_kembali = (
        "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
        "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
        "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
        "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
        "/div[@class='p-2']/div/div[@class='w-full']"
        "/div[@class='fixed top-0 left-0 w-full h-full flex justify-center items-center backdrop-blur-5 z-1000']"
        "/div[@class='w-[40%] max-h-full relative z-1000 rounded-lg bg-white p-4 shadow-gmail lt-md:w-[80%] lt-sm:w-full overflow-auto']"
        "/div[@class='p-2']/div/div[@class='flex justify-center gap-2 font-600 text-base w-full mt-6']"
        "/div[@class='w-auto lt-sm:w-full'][1]/button[@class='w-fill btn-outline-primary h-11']"
    )
    try:
        popup_el = page.locator(f"xpath={xpath_txt_kategori}").first
        if await is_element_really_visible(popup_el):
            msg_text = await popup_el.inner_text()
            log(f"  ⚠️  Popup kategori sasaran terdeteksi: '{msg_text}'", "WARN")
            
            # Klik tombol Kembali
            log("  🚀 Mengklik tombol 'Kembali'...", "BOT")
            btn_kembali = page.locator(f"xpath={xpath_btn_kembali}").first
            await btn_kembali.click()
            await page.wait_for_timeout(800)
            
            # Tanya user untuk konfirmasi lolos/skip
            while True:
                if GUI_BRIDGE_ACTIVE and gui_bridge:
                    pilihan = gui_bridge.request_input("kategori_pasien", msg_text)
                else:
                    pilihan = input("\n  ❓ Konfirmasi status kategori pasien ini (l = Lolos / s = Skip): ").strip().lower()
                
                if pilihan in ["l", "lolos"]:
                    log("  ✓ Pasien diloloskan untuk melanjutkan pendaftaran.", "OK")
                    break
                elif pilihan in ["s", "skip"]:
                    log("  ⏭ Pasien di-skip sesuai permintaan user.", "WARN")
                    raise SkipPasien("Kategori pasien dilewati (user skip)")
                else:
                    if not GUI_BRIDGE_ACTIVE:
                        print("Input tidak valid. Masukkan 'l' atau 's'.")
    except SkipPasien:
        raise
    except Exception:
        pass


async def cek_popup_tidak_valid(page: Page) -> bool:
    """
    Deteksi popup 'data valid' atau 'data tidak valid' secara agresif dengan polling.
    Menunggu hingga 3.0 detik agar popup muncul, atau mendeteksi jika halaman sudah berpindah.
    Menggunakan pencarian teks native dan mengurangi jeda polling untuk kecepatan maksimal.
    """
    try:
        # XPath untuk tombol Selanjutnya Step 1 (jika tombol ini hilang/tidak visible, artinya halaman sudah berpindah)
        xpath_btn_selanjutnya_step1 = (
            "//div[@id='__nuxt']/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[5]/div[2]/div/div/div[6]/div[2]/div[2]/div/div/div[3]/div/button"
        )

        log("  ⏳ Menunggu deteksi popup validasi data (polling super cepat)...", "WAIT")

        # Polling setiap 100ms selama maksimal 30 kali (total 3.0 detik)
        for attempt in range(30):
            # 1. Cek jika halaman sudah berpindah (tidak ada popup, submit berhasil langsung)
            try:
                btn_sel_xpath = page.locator(f"xpath={xpath_btn_selanjutnya_step1}").first
                btn_sel_role = page.get_by_role("button", name="Selanjutnya").first
                
                # Jika tombol Selanjutnya sudah tidak visible, berarti sudah berhasil submit dan pindah halaman
                if not (await btn_sel_xpath.is_visible()) and not (await btn_sel_role.is_visible()):
                    log("  ✓ Halaman terdeteksi sudah berpindah (tidak ada popup pendaftaran yang menghalangi).", "OK")
                    return False
            except Exception:
                pass

            # 2. Cek spesifik popup "Data peserta tidak valid" via teks native (sangat cepat)
            tidak_valid_el = page.locator("text=/data peserta tidak valid|nik tidak valid|nik salah|tidak terdaftar/i").first
            if await tidak_valid_el.is_visible():
                log("  ⚠️ Popup 'Data peserta tidak valid' terdeteksi secara spesifik.", "WARN")
                popup_peserta_tidak_valid = page.locator("div.fixed.z-1000, div.fixed.z-1100, div[class*='z-1000'], div[class*='z-1100']").filter(
                    has_text=re.compile(r"Data peserta tidak valid|nik tidak valid|nik salah|tidak terdaftar", re.IGNORECASE)
                ).first
                # Cari tombol Periksa Kembali atau Tutup di dalam popup ini
                btn_periksa = popup_peserta_tidak_valid.locator("button, [role='button'], div.cursor-pointer").filter(
                    has_text=re.compile(r"(Periksa|Tutup|Kembali|OK)", re.IGNORECASE)
                ).first
                if await btn_periksa.is_visible():
                    log(f"  🤖 Menutup popup tidak valid: '{await btn_periksa.inner_text()}'...", "BOT")
                    if not await force_js_click(btn_periksa):
                        await btn_periksa.click(timeout=2000)
                else:
                    btn_any = popup_peserta_tidak_valid.locator("button, [role='button'], div.cursor-pointer").first
                    if await btn_any.count() > 0:
                        log("  🤖 Menutup popup tidak valid via tombol alternatif...", "BOT")
                        if not await force_js_click(btn_any):
                            await btn_any.click(timeout=2000)
                await page.wait_for_timeout(300) # Kurangi delay dari 800ms ke 300ms
                return True  # kembalikan True karena data tidak valid (skip)

            # 3. Cek spesifik popup "Data peserta valid" via teks native (sangat cepat)
            valid_el = page.locator("text=/data peserta valid|data anda valid/i").first
            if await valid_el.is_visible():
                log("  ✓ Popup 'Data peserta valid' terdeteksi secara spesifik.", "OK")
                popup_peserta_valid = page.locator("div.fixed.z-1000, div.fixed.z-1100, div[class*='z-1000'], div[class*='z-1100']").filter(
                    has_text=re.compile(r"Data peserta valid", re.IGNORECASE)
                ).first
                btn_lanjutkan = popup_peserta_valid.locator("button, [role='button'], div.cursor-pointer").filter(
                    has_text=re.compile(r"Lanjutkan", re.IGNORECASE)
                ).first
                if await btn_lanjutkan.is_visible():
                    log(f"  🚀 Mengklik tombol 'Lanjutkan' pada popup: '{await btn_lanjutkan.inner_text()}'...", "BOT")
                    if not await force_js_click(btn_lanjutkan):
                        await btn_lanjutkan.click(timeout=2000)
                else:
                    btn_any = popup_peserta_valid.locator("button, [role='button'], div.cursor-pointer").first
                    if await btn_any.count() > 0:
                        log("  🚀 Mengklik tombol alternatif pada popup data peserta valid...", "BOT")
                        if not await force_js_click(btn_any):
                            await btn_any.click(timeout=2000)
                await page.wait_for_timeout(300) # Kurangi delay dari 800ms ke 300ms
                return False  # kembalikan False karena data valid (lanjut)

            # 4. Scan Strategi 1: XPath spesifik user (Prioritas Utama)
            try:
                xpath_txt_popup = 'xpath=//*[@id="__nuxt"]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[5]/div[2]/div/div/div[6]/div[2]/div[2]/div/div/div[1]/div'
                xpath_btn_popup_new = 'xpath=//*[@id="__nuxt"]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[4]/div[2]/div/div/div[6]/div[2]/div[2]/div/div/div[3]/div/button'
                xpath_btn_popup_old = 'xpath=//*[@id="__nuxt"]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[5]/div[2]/div/div/div[6]/div[2]/div[2]/div/div/div[3]/div/button'
                
                btn_el = page.locator(xpath_btn_popup_new).first
                if not await btn_el.is_visible():
                    btn_el = page.locator(xpath_btn_popup_old).first

                if await btn_el.is_visible():
                    txt_el = page.locator(xpath_txt_popup).first
                    txt = await txt_el.inner_text() if await txt_el.count() > 0 else ""
                    txt_lower = txt.lower()
                    
                    if "nomor whatsapp" in txt_lower or "tanggal lahir" in txt_lower or "isi data wali" in txt_lower:
                        pass
                    elif "kuota pemeriksaan" in txt_lower or "kuota habis" in txt_lower:
                        pass
                    else:
                        log(f"  🔍 Teks popup (user XPath): '{txt.strip()[:80]}'", "INFO")
                        is_tidak_valid = any(k in txt_lower for k in ["tidak valid", "tidak terdaftar", "salah", "gagal", "tidak cocok"])
                        if is_tidak_valid or not txt:
                            log("  ⚠️ Popup NIK TIDAK VALID terkonfirmasi via XPath — menutup popup.", "WARN")
                            if not await force_js_click(btn_el):
                                await btn_el.click(timeout=2000)
                            await page.wait_for_timeout(300)
                            return True
                        is_valid = ("valid" in txt_lower and "tidak" not in txt_lower) or "peserta" in txt_lower or "lanjut" in txt_lower
                        if is_valid:
                            log("  ✓ Popup NIK VALID terkonfirmasi via XPath — melanjutkan.", "OK")
                            if not await force_js_click(btn_el):
                                await btn_el.click(timeout=2000)
                            await page.wait_for_timeout(300)
                            return False
            except Exception:
                pass

            # 5. Scan Strategi 2: Scan semua elemen fixed/modal yang visible (Menggunakan native .is_visible() untuk kecepatan)
            try:
                modal_selectors = [
                    "div.fixed.z-9000",
                    "div.fixed.z-1100",
                    "div.fixed.z-1000",
                    "div[class*='z-9000']",
                    "div[class*='z-1100']",
                    "div[class*='z-1000']",
                    "div[role='dialog']",
                    "div[role='alertdialog']",
                ]

                for sel in modal_selectors:
                    elements = page.locator(sel)
                    cnt = await elements.count()
                    for i in range(cnt):
                        el = elements.nth(i)
                        if not await el.is_visible():
                            continue
                        text = (await el.inner_text() or "").strip()
                        if not text:
                            continue
                        text_lower = text.lower()

                        if "nomor whatsapp" in text_lower or "tanggal lahir" in text_lower or "isi data wali" in text_lower:
                            continue
                        if "kuota pemeriksaan" in text_lower or "kuota habis" in text_lower:
                            continue

                        is_invalid_kw = any(k in text_lower for k in ["tidak valid", "tidak terdaftar", "data tidak", "nik tidak", "nik salah", "gagal", "tidak cocok", "tidak ditemukan", "tidak aktif"])
                        is_valid_kw = (
                            any(k in text_lower for k in ["data valid", "data anda valid", "peserta valid", "sudah terdaftar", "berhasil terverifikasi", "lanjutkan pendaftaran", "data peserta valid"])
                            or ("valid" in text_lower and "tidak" not in text_lower)
                        )

                        if not (is_valid_kw or is_invalid_kw):
                            continue

                        if is_invalid_kw:
                            log(f"  ⚠️ Popup TIDAK VALID terdeteksi: '{text[:100]}'", "WARN")
                            btn = el.locator("button, [role='button'], div.cursor-pointer").first
                            if await btn.count() > 0 and await btn.is_visible():
                                log(f"  🤖 Menutup popup tidak valid: '{await btn.inner_text()}'...", "BOT")
                                if not await force_js_click(btn):
                                    await btn.click(timeout=2000)
                                await page.wait_for_timeout(300)
                            return True

                        if is_valid_kw:
                            log(f"  ✓ Popup VALID terdeteksi: '{text[:100]}'", "OK")
                            btn_confirm = el.locator("button, [role='button'], div.cursor-pointer").filter(
                                has_text=re.compile(r"(Data Valid|Ya|Lanjut|OK|Konfirmasi)", re.IGNORECASE)
                            ).first
                            if await btn_confirm.count() > 0 and await btn_confirm.is_visible():
                                log(f"  🚀 Mengklik konfirmasi valid: '{await btn_confirm.inner_text()}'...", "BOT")
                                if not await force_js_click(btn_confirm):
                                    await btn_confirm.click(timeout=2000)
                            else:
                                btn_any = el.locator("button, [role='button'], div.cursor-pointer").first
                                if await btn_any.count() > 0:
                                    log("  🚀 Mengklik tombol alternatif di popup valid...", "BOT")
                                    if not await force_js_click(btn_any):
                                        await btn_any.click(timeout=2000)
                            await page.wait_for_timeout(300)
                            return False
            except Exception:
                pass

            # Beri jeda sebelum polling berikutnya (100ms)
            await page.wait_for_timeout(100)

        # ── Strategi 3: Fallback teks cepat ──
        try:
            invalid_text_locator = page.locator("text=/nik tidak valid|nik salah|data tidak valid|tidak terdaftar/i").first
            if await invalid_text_locator.is_visible():
                log("  ⚠️ Terdeteksi teks 'tidak valid' di halaman (fallback cepat).", "WARN")
                return True
        except Exception:
            pass

    except Exception as e:
        log(f"  ⚠️ Gagal mengecek popup valid/tidak valid: {e}", "WARN")
    
    log("  ✓ Tidak ada popup validasi yang terdeteksi. Melanjutkan...", "OK")
    return False


def update_excel_row_color(no_value, color_type):
    """
    Color cells in the row matching `no_value` (No.) in Excel.
    color_type: 'green' | 'yellow' | 'red'
    """
    try:
        if not FILE_EXCEL.exists():
            return
        wb = load_workbook(FILE_EXCEL)
        ws = wb.active
        
        target_row = None
        for row_idx in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=row_idx, column=1).value
            if cell_val is not None:
                try:
                    if int(cell_val) == int(no_value):
                        target_row = row_idx
                        break
                except ValueError:
                    if str(cell_val).strip() == str(no_value).strip():
                        target_row = row_idx
                        break
        
        if target_row:
            fill_map = {
                "green": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
                "yellow": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
                "red": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
            }
            fill = fill_map.get(color_type)
            if fill:
                for col_idx in range(1, ws.max_column + 1):
                    ws.cell(row=target_row, column=col_idx).fill = fill
                # Backup file Excel sebelum melakukan penyimpanan
                import shutil
                try:
                    backup_path = FILE_EXCEL.with_suffix(FILE_EXCEL.suffix + ".bak")
                    shutil.copy2(FILE_EXCEL, backup_path)
                except Exception as e_bak:
                    log(f"Gagal membuat backup Excel sebelum menyimpan: {e_bak}", "WARN")
                wb.save(FILE_EXCEL)
                log(f"Baris Excel No.{no_value} berhasil diwarnai {color_type}.", "OK")
    except Exception as e:
        log(f"Gagal mewarnai Excel: {e}", "WARN")


def temukan_input_terakhir():
    """
    Find the last processed row (marked with a pattern fill) in ckg.xlsx.
    Returns (last_processed_no, last_processed_name, status_str).
    """
    if not FILE_EXCEL.exists():
        return None, None, None
    try:
        wb = load_workbook(FILE_EXCEL)
        ws = wb.active
        
        last_processed_no = None
        last_processed_name = None
        status_str = None
        
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(row=r, column=1)
            # If the cell has fill and pattern fill is solid, it has been processed
            if cell.fill and cell.fill.fill_type == "solid":
                no_val = cell.value
                name_val = ws.cell(row=r, column=2).value
                if no_val is not None:
                    last_processed_no = no_val
                    last_processed_name = name_val
                    
                    # Detect color/status from cell fill color
                    # Convert rgb to string to check color parts
                    rgb = str(cell.fill.start_color.rgb).upper()
                    if "C6EFCE" in rgb:
                        status_str = "BERHASIL (Hijau)"
                    elif "FFEB9C" in rgb:
                        status_str = "SKIP (Kuning)"
                    elif "FFC7CE" in rgb:
                        status_str = "GAGAL/CANCEL (Merah)"
                    else:
                        status_str = f"DIPROSES (Color RGB: {rgb})"
                    
        return last_processed_no, last_processed_name, status_str
    except Exception as e:
        log(f"Gagal mendeteksi input terakhir di Excel: {e}", "WARN")
        return None, None, None


# ============================================================
#  LOGIN
# ============================================================

async def login(page: Page, username: str = None, password: str = None):
    # Store initial inputs, but allow updating them in the loop
    user_email = username if username else USERNAME
    user_pass  = password if password else PASSWORD

    while True:
        log("Membuka halaman login…", "INFO")
        await page.goto(f"{BASE_URL}/auth/login", wait_until="domcontentloaded", timeout=TIMEOUT)
        await page.wait_for_timeout(1500)

        await page.get_by_role("textbox", name="Email").fill(user_email)
        await page.wait_for_timeout(200)
        await page.get_by_role("textbox", name="Kata sandi").fill(user_pass)
        await page.wait_for_timeout(200)

        log("Email & password terisi.", "OK")

        # CAPTCHA — minta user untuk memasukkan secara manual di terminal
        try:
            captcha_field = page.locator("xpath=//input[@id='input-captcha']")
            if await captcha_field.count() > 0:
                # Coba tampilkan Captcha di terminal jika memungkinkan
                try:
                    captcha_img = None
                    img_locators = [
                        page.locator("img[src*='captcha']"),
                        page.locator("img[class*='captcha']"),
                        page.locator("img[id*='captcha']"),
                        page.locator("#input-captcha").locator("xpath=../preceding-sibling::*//img"),
                        page.locator("#input-captcha").locator("xpath=..//img"),
                        page.locator("form img")
                    ]
                    for loc in img_locators:
                        if await loc.count() > 0 and await is_element_really_visible(loc.first):
                            captcha_img = loc.first
                            break
                    
                    img_bytes = None
                    if captcha_img:
                        img_bytes = await captcha_img.screenshot()
                        if not GUI_BRIDGE_ACTIVE:
                            base64_str = base64.b64encode(img_bytes).decode('utf-8')
                            ascii_art = image_to_binary_ascii(base64_str)
                            print("\n" + "=" * 54)
                            print("  [ KODE CAPTCHA DI BAWAH INI ]")
                            print("=" * 54)
                            print(ascii_art)
                            print("=" * 54 + "\n")
                except Exception:
                    pass

                if GUI_BRIDGE_ACTIVE and gui_bridge:
                    captcha_val = gui_bridge.request_input("captcha", img_bytes)
                else:
                    captcha_val = input("  >> Masukkan Kode CAPTCHA yang tampil di browser: ").strip()
                await captcha_field.fill(captcha_val)
                await page.wait_for_timeout(200)
                log(f"Mengisi CAPTCHA: {captcha_val}", "OK")
            
            log("Mengklik tombol Masuk...", "BOT")
            masuk_btn = page.locator("button[type='submit']").first
            if await masuk_btn.count() == 0:
                masuk_btn = page.get_by_role("button", name="Masuk").first
            await masuk_btn.click(timeout=10000)
            await page.wait_for_timeout(3000)
        except Exception as e:
            log(f"Gagal klik Masuk: {e}", "WARN")


        # Periksa apakah sudah di dashboard secara programatis
        await page.wait_for_timeout(1000)
        current_url = page.url
        is_dashboard = "dashboard" in current_url.lower() or await page.get_by_role("button", name="CKG Umum").count() > 0
        
        if is_dashboard:
            log("Terdeteksi sudah masuk ke Dashboard.", "OK")
            break
        else:
            log(f"Browser belum masuk ke Dashboard (URL saat ini: {current_url})", "WARN")
            if GUI_BRIDGE_ACTIVE and gui_bridge:
                res = gui_bridge.request_input("login_failure", current_url)
                if isinstance(res, tuple) and len(res) == 3:
                    pilihan, email_new, pass_new = res
                else:
                    pilihan = str(res)
                    email_new = None
                    pass_new = None
                    
                if pilihan == "2":
                    user_email = email_new
                    user_pass = pass_new
                elif pilihan == "3":
                    log("Memaksa lanjut sesuai instruksi user...", "WARN")
                    break
            else:
                print("  Pilih tindakan:")
                print("  [1] Ulangi proses login (halaman akan dimuat ulang)")
                print("  [2] Masukkan Email & Kata Sandi baru")
                print("  [3] Paksa lanjut (abaikan deteksi dashboard)")
                pilihan = input("  >> Masukkan pilihan [1/2/3, default: 1]: ").strip()
                
                if pilihan == "2":
                    user_email = input("  >> Masukkan Email baru: ").strip()
                    user_pass  = input("  >> Masukkan Kata sandi baru: ").strip()
                elif pilihan == "3":
                    log("Memaksa lanjut sesuai instruksi user...", "WARN")
                    break
            # default: loop repeats, loading page again

    # Popup opsional setelah login
    await klik_opsional(page, page.locator("div:nth-child(85) > div").first)
    await klik_opsional(page, page.locator("#verify").nth(1))
    await klik_opsional(page, page.get_by_role("button", name="Setuju"))

    log("Login selesai.", "OK")
    await screenshot(page, "login_success")


# ============================================================
#  NAVIGASI
# ============================================================

async def buka_form_daftar_baru(page: Page):
    # 0. Cek jika url bukan halaman form pendaftaran, navigasikan langsung ke form untuk mempercepat proses (menghindari delay menu clicks)
    if "ckg-pendaftaran-individu" not in page.url:
        log("  🌐 Menavigasi langsung ke form pendaftaran CKG...", "INFO")
        try:
            await page.goto(f"{BASE_URL}/ckg-pendaftaran-individu", wait_until="domcontentloaded", timeout=15000)
            await page.wait_for_timeout(1000)
        except Exception as e:
            log(f"  ⚠️ Gagal navigasi langsung ke form: {e}. Mencoba menu fallback...", "WARN")

    # 1. Cek apakah pop up overlay z-20 aktif (menandakan form aktif/terbuka)
    xpath_overlay_z20 = (
        "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
        "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
        "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
        "/div[@class='fixed inset-0 bg-[rgba(0,0,0,0.8)] opacity-50 z-20']"
    )
    try:
        overlay = page.locator(f"xpath={xpath_overlay_z20}").first
        if await overlay.is_visible(timeout=1500):
            log("  ✓ Pop up overlay z-20 aktif (form terbuka), langsung mengisi data.", "OK")
            return
    except Exception:
        pass

    # 2. Cek apakah form sudah terbuka (misal input NIK sudah terlihat)
    try:
        nik_input = page.get_by_role("textbox", name="NIK *")
        if await nik_input.is_visible(timeout=1500):
            log("  ✓ Form sudah terbuka, langsung mengisi data tanpa navigasi ulang.", "OK")
            return
    except Exception:
        pass

    # 3. Cek apakah tombol 'Daftar Baru' terlihat di halaman (tanpa perlu klik menu CKG Umum)
    try:
        daftar_baru_btn = page.get_by_role("button", name="Daftar Baru")
        if await daftar_baru_btn.is_visible(timeout=1500):
            log("  🤖 Menemukan tombol 'Daftar Baru', langsung mengeklik...", "BOT")
            await daftar_baru_btn.click()
            await page.wait_for_timeout(1000)
            return
    except Exception:
        pass

    # 4. Fallback: Navigasi penuh dari menu
    log("Navigasi ke form CKG Umum → Daftar Baru…", "BOT")
    try:
        await page.get_by_role("button", name="CKG Umum").click()
        await page.wait_for_timeout(800)
        await page.locator('[id="menu_cari/daftarkan_individu"]').click()
        await page.wait_for_timeout(800)
        await page.get_by_role("button", name="Daftar Baru").click()
        await page.wait_for_timeout(1000)
    except Exception as e:
        log(f"  ⚠️ Gagal navigasi menu fallback: {e}", "ERR")
        raise e


# ============================================================
#  DATE PICKER
# ============================================================

async def pilih_tanggal(page: Page, tahun: int, bulan: int, hari: int):
    """
    Memilih tanggal pada Vue Datepicker berdasarkan rekaman Playwright.
    Alur: buka picker → klik tahun (langsung masuk Decade/Year grid) →
          navigasi decade → klik tahun → klik bulan → klik hari.
    Data diambil dari Excel kolom F (Tanggal Lahir).

    Referensi recording:
      page.get_by_role("button", name="2026").nth(1).click()   # → decade view
      page.get_by_role("cell", name="2000").click()            # → pilih tahun
      page.get_by_role("cell", name="Jun", exact=True).click() # → pilih bulan
      page.locator("div").filter(has_text=re.compile(r"^24$")).click() # → pilih hari
    """
    log(f"  🗓  Memilih tanggal lahir: {hari:02d}-{bulan:02d}-{tahun}", "BOT")

    # 1. Buka picker
    try:
        await page.locator('[id="Tanggal Lahir"] .mx-input-wrapper').click(timeout=5000)
    except Exception:
        await page.locator("div").filter(has_text=re.compile(r"^Pilih tanggal lahir$")).nth(2).click()
    await page.wait_for_timeout(500)

    # 2. Pastikan popup muncul
    popup = page.locator("xpath=//div[contains(@class, 'mx-datepicker-popup')]").first
    await popup.wait_for(state="visible", timeout=5000)

    # 3. Klik tombol tahun di header → langsung masuk Year/Decade grid (1 klik)
    #    Sesuai recording: page.get_by_role("button", name="2026").nth(1).click()
    await popup.locator("button.mx-btn-current-year").click()
    await page.wait_for_timeout(400)

    # 4. Navigasi decade hingga tahun target masuk dalam rentang yang tampil
    for _ in range(20):
        try:
            header_text = await popup.locator(".mx-calendar-header-label").text_content()
            years = [int(x) for x in re.findall(r"\d{4}", header_text)]
        except Exception:
            years = []

        if len(years) >= 2:
            decade_start, decade_end = years[0], years[-1]
        elif len(years) == 1:
            decade_start = years[0]
            decade_end   = decade_start + 9
        else:
            decade_start, decade_end = 2020, 2029

        if decade_start <= tahun <= decade_end:
            break
        elif tahun < decade_start:
            await popup.locator("button.mx-btn-icon-double-left").click()
        else:
            await popup.locator("button.mx-btn-icon-double-right").click()
        await page.wait_for_timeout(200)

    # 5. Klik tahun — sesuai recording: page.get_by_role("cell", name="2000").click()
    await popup.get_by_role("cell", name=str(tahun)).click()
    await page.wait_for_timeout(400)

    # 6. Klik bulan — sesuai recording: page.get_by_role("cell", name="Jun", exact=True).click()
    bulan_ok = False
    for nama_bulan in [BULAN_ID[bulan]] + BULAN_FALLBACK[bulan]:
        try:
            cell = popup.get_by_role("cell", name=nama_bulan, exact=True)
            if await cell.count() > 0:
                await cell.click()
                bulan_ok = True
                await page.wait_for_timeout(400)
                break
        except Exception:
            continue

    if not bulan_ok:
        log(f"  ⚠️  Nama bulan tidak ditemukan, fallback ke indeks ke-{bulan}...", "WARN")
        month_cells = popup.locator(".mx-table-month td.cell")
        if await month_cells.count() >= 12:
            await month_cells.nth(bulan - 1).click()
            await page.wait_for_timeout(400)
            bulan_ok = True

    if not bulan_ok:
        raise Exception(f"Gagal memilih bulan ke-{bulan}")

    # 7. Klik hari
    #    Prioritas 1: ISO title (mitigasi overlap hari bulan tetangga)
    iso_date_str = f"{tahun}-{bulan:02d}-{hari:02d}"
    day_cell_iso = popup.locator(f"td.cell[title='{iso_date_str}']").first

    if await day_cell_iso.count() > 0:
        await day_cell_iso.click()
        await page.wait_for_timeout(300)
        log(f"  ✓ Tanggal {hari:02d}-{bulan:02d}-{tahun} dipilih (ISO title).", "OK")
    else:
        # Prioritas 2: div teks persis — sesuai recording:
        # page.locator("div").filter(has_text=re.compile(r"^24$")).click()
        day_div = popup.locator("div").filter(has_text=re.compile(rf"^{hari}$")).first
        if await day_div.count() > 0:
            await day_div.click()
            await page.wait_for_timeout(300)
            log(f"  ✓ Tanggal {hari:02d}-{bulan:02d}-{tahun} dipilih (div text).", "OK")
        else:
            # Prioritas 3: td.cell teks hari, bukan bulan tetangga
            day_cells = popup.locator("td.cell:not(.not-current-month)")
            clicked = False
            for i in range(await day_cells.count()):
                cell = day_cells.nth(i)
                txt = (await cell.text_content()).strip()
                if txt == str(hari):
                    await cell.click()
                    clicked = True
                    await page.wait_for_timeout(300)
                    break
            if clicked:
                log(f"  ✓ Tanggal {hari:02d}-{bulan:02d}-{tahun} dipilih (td fallback).", "OK")
            else:
                raise Exception(f"Hari {hari} tidak ditemukan di grid tanggal ({iso_date_str})")

async def pilih_tanggal_wali(page: Page, tahun: int, bulan: int, hari: int):
    """
    Memilih tanggal pada Vue Datepicker untuk Wali (vue2-datepicker).
    """
    log(f"    🗓  Memilih tanggal lahir wali: {hari:02d}-{bulan:02d}-{tahun}", "BOT")

    # 1. Buka picker
    try:
        await page.locator(".mx-input").nth(2).click(timeout=5000)
    except Exception:
        try:
            await page.locator(".mx-input-wrapper").last.click(timeout=5000)
        except Exception as e:
            raise Exception(f"Gagal membuka datepicker Wali: {e}")
    await page.wait_for_timeout(600)

    # 2. Dapatkan popup yang sedang aktif/terbuka
    popup = page.locator("xpath=//div[contains(@class, 'mx-datepicker-popup')]").first
    await popup.wait_for(state="visible", timeout=5000)

    # 3. Klik button current-year
    await popup.locator("button.mx-btn-current-year").click()
    await page.wait_for_timeout(500)

    # 4. Navigasi dekade
    for _ in range(25):
        try:
            header_text = await popup.locator(".mx-calendar-header-label").text_content()
            years = [int(x) for x in re.findall(r"\d{4}", header_text)]
        except Exception:
            years = []

        if len(years) >= 2:
            decade_start, decade_end = years[0], years[-1]
        elif len(years) == 1:
            decade_start = years[0]
            decade_end   = decade_start + 9
        else:
            decade_start, decade_end = 2020, 2029

        if decade_start <= tahun <= decade_end:
            break
        elif tahun < decade_start:
            await popup.locator("button.mx-btn-icon-double-left").click()
        else:
            await popup.locator("button.mx-btn-icon-double-right").click()
        await page.wait_for_timeout(300)

    # 5. Pilih tahun
    await popup.get_by_role("cell", name=str(tahun)).click()
    await page.wait_for_timeout(500)

    # 6. Pilih bulan
    bulan_ok = False
    for nama_bulan in [BULAN_ID[bulan]] + BULAN_FALLBACK[bulan]:
        try:
            cell = popup.get_by_role("cell", name=nama_bulan, exact=True)
            if await cell.count() > 0:
                await cell.click()
                bulan_ok = True
                await page.wait_for_timeout(500)
                break
        except Exception:
            continue

    if not bulan_ok:
        month_cells = popup.locator(".mx-table-month td.cell")
        if await month_cells.count() >= 12:
            await month_cells.nth(bulan - 1).click()
            await page.wait_for_timeout(500)
            bulan_ok = True

    # 7. Pilih hari
    iso_date_str = f"{tahun}-{bulan:02d}-{hari:02d}"
    day_cell_iso = popup.locator(f"td.cell[title='{iso_date_str}']").first
    if await day_cell_iso.count() > 0:
        await day_cell_iso.click()
        await page.wait_for_timeout(400)
    else:
        day_div = popup.locator("div").filter(has_text=re.compile(rf"^{hari}$")).first
        if await day_div.count() > 0:
            await day_div.click()
            await page.wait_for_timeout(400)
        else:
            day_cells = popup.locator("td.cell:not(.not-current-month)")
            clicked = False
            for i in range(await day_cells.count()):
                cell = day_cells.nth(i)
                txt = (await cell.text_content()).strip()
                if txt == str(hari):
                    await cell.click()
                    clicked = True
                    await page.wait_for_timeout(300)
                    break
            if not clicked:
                raise Exception(f"Hari {hari} tidak ditemukan di grid tanggal ({iso_date_str})")


# ============================================================
#  INPUT SATU PASIEN — dengan konfirmasi manual
# ============================================================


async def input_pasien(page: Page, row: dict, nomor: int, total: int) -> str:
    check_stop_request()
    """
    Kembalikan: 'ok' | 'skip' | 'quit'
    """
    nama  = str(row.get("Nama Lengkap", "")).strip()
    # Parsing NIK secara aman untuk menghindari hilangnya digit ke-16 akibat floating-point precision
    nik_raw = str(row.get("NIK", "")).strip()
    if nik_raw.endswith(".0"):
        nik_raw = nik_raw[:-2]
    nik = "".join(re.findall(r"\d+", nik_raw)).zfill(16)
    jk    = str(row.get("Jenis Kelamin", "")).strip().upper()
    tgl   = pd.Timestamp(row.get("Tanggal Lahir"))
    no_hp = str(row.get("No HP", "")).strip()
    no_excel = int(row.get("No", 0))

    # Fungsi pembantu untuk memaksakan klik elemen menggunakan JavaScript Injection (Force Click)
    async def js_click(locator_or_element):
        try:
            if hasattr(locator_or_element, "first"):
                target = locator_or_element.first
            else:
                target = locator_or_element
            
            if await target.count() > 0:
                await target.evaluate("el => el.click()")
                return True
        except Exception:
            pass
        try:
            # Fallback jika locator normal tidak bekerja
            await page.evaluate("el => el.click()", locator_or_element)
            return True
        except Exception:
            pass
        return False

    # Tampilkan data di terminal
    cetak_data_pasien(nomor, total, nama, nik, jk, tgl.strftime("%d-%m-%Y"))

    xpath_txt_data_ditemukan = (
        "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
        "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
        "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
        "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
        "/div[@class='p-2']/div/div[@class='w-full']"
        "/div[@class='fixed top-0 left-0 w-full h-full flex justify-center items-center backdrop-blur-5 z-1000']"
        "/div[@class='w-[25%] lt-lg:w-[40%] max-h-full relative z-1000 rounded-lg bg-white p-4 shadow-gmail lt-md:w-[80%] lt-sm:w-full overflow-auto']"
        "/div[@class='p-2']/div[@class='text-[#62686F]']"
        "/div[@class='flex flex-col justify-center items-center gap-3 text-[20px] font-bold mb-1']"
        "/div[@class='pb-2 text-center text-black']"
    )

    xpath_btn_gunakan_data = (
        "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
        "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
        "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
        "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
        "/div[@class='p-2']/div/div[@class='w-full']"
        "/div[@class='fixed top-0 left-0 w-full h-full flex justify-center items-center backdrop-blur-5 z-1000']"
        "/div[@class='w-[25%] lt-lg:w-[40%] max-h-full relative z-1000 rounded-lg bg-white p-4 shadow-gmail lt-md:w-[80%] lt-sm:w-full overflow-auto']"
        "/div[@class='p-2']/div[@class='text-[#62686F]']"
        "/div[@class='flex justify-center gap-2 font-600 text-base w-full mt-6']"
        "/div[@class='w-auto lt-sm:w-full'][2]"
        "/button[@class='w-fill btn-fill-primary h-11']"
    )

    xpath_datepicker_tgl_pemeriksaan = (
        "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
        "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
        "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
        "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
        "/div[@class='p-2']/div/div[@class='w-full']"
        "/div[@class='form-data-individu']/form/div[@class='flex gap-5 w-full lt-sm:flex-col']"
        "/div[@class='flex flex-col gap-3 w-full'][2]/div[2]"
        "/div[@class='relative p-3 pt-2 border rounded-md bg-white shadow-gmail']"
    )

    xpath_btn_selanjutnya_step1 = (
        "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
        "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
        "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
        "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
        "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
        "/div[@class='p-2']/div/div[@class='w-full']"
        "/div[@class='form-data-individu']/form/div[@class='flex justify-end py-1']/div"
        "/button[@class='w-fill btn-fill-primary h-11']/div[@class='flex flex-row justify-center gap-2']/div[@class='tracking-wide']"
    )

    # ── Isi NIK ───────────────────────────────────────────────
    log(f"  ✏️  Mengisi NIK: {nik}", "BOT")
    await page.get_by_role("textbox", name="NIK *").click()
    await page.get_by_role("textbox", name="NIK *").fill(nik)
    await page.get_by_role("button", name="Cek NIK").click()

    # Tunggu sebentar hingga salah satu elemen respon muncul (max 2 detik)
    log("  ⏳ Menunggu respon NIK...", "WAIT")
    for _ in range(20):
        popup_periksa = page.locator(".w-\\[40\\%\\]").first
        popup_ditemukan = page.locator(f"xpath={xpath_txt_data_ditemukan}").first
        popup_ditemukan_text = page.locator("div.text-black", has_text="Data Peserta ditemukan").first
        if await popup_ditemukan_text.count() == 0:
            popup_ditemukan_text = page.get_by_text("Data Peserta ditemukan").first
        nama_input = page.get_by_role("textbox", name="Nama Lengkap *")
        if (await popup_periksa.is_visible()) or (await popup_ditemukan.is_visible()) or (await popup_ditemukan_text.is_visible()) or (await nama_input.is_editable()):
            break
        await page.wait_for_timeout(100)

    # Tangani kemungkinan munculnya popup "Periksa Kembali" setelah klik Cek NIK
    await cek_popup_periksa_kembali(page)
    await cek_popup_kategori_pasien(page)

    # ── Cek Popup "Data Ditemukan" ────────────────────────────
    data_ditemukan = False
    try:
        # Polling selama maks 3 detik untuk memberi waktu popup selesai loading/animasi
        popup_txt = None
        for _ in range(15):
            loc = page.locator("div.text-black, h2, h3, div, p", has_text="Data Peserta ditemukan").first
            if await loc.count() == 0:
                loc = page.get_by_text("Data Peserta ditemukan").first
            if await loc.count() == 0:
                loc = page.locator(f"xpath={xpath_txt_data_ditemukan}").first
                
            if await loc.count() > 0 and await is_element_really_visible(loc):
                popup_txt = loc
                break
            await page.wait_for_timeout(200)
            
        if popup_txt:
            log("  ✓ Popup 'Data Peserta ditemukan' terdeteksi. Menggunakan data terdaftar...", "OK")
            # Cek tombol Gunakan Data
            btn_gunakan = page.get_by_role("button", name="Gunakan Data").first
            if await btn_gunakan.count() == 0:
                btn_gunakan = page.locator("button", has_text="Gunakan Data").first
            if await btn_gunakan.count() == 0:
                btn_gunakan = page.locator(f"xpath={xpath_btn_gunakan_data}").first
            
            try:
                await btn_gunakan.click(timeout=3000)
            except Exception:
                await js_click(btn_gunakan)
                
            log("  ✓ Tombol 'Gunakan Data' diklik.", "OK")
            await page.wait_for_timeout(500)
            data_ditemukan = True
    except Exception as e:
        log(f"  ⚠️ Gagal mendeteksi/memproses popup Data Peserta ditemukan: {e}", "WARN")

    if data_ditemukan:
        # Jika data ditemukan, atur tanggal pemeriksaan
        log("  🗓  Memilih tanggal pemeriksaan (Hari ini) via datepicker...", "BOT")
        try:
            today_day = str(datetime.now().day)
            
            # Cari calendar container (div shadow-gmail)
            calendar_container = page.locator("div.relative.p-3.pt-2.border.rounded-md.bg-white.shadow-gmail").first
            if await calendar_container.count() == 0:
                calendar_container = page.locator(f"xpath={xpath_datepicker_tgl_pemeriksaan}").first
            
            # Jika container tidak visible, klik dulu untuk membukanya
            if not await is_element_really_visible(calendar_container):
                log("  🗓  Calendar container tersembunyi, mengklik datepicker trigger...", "BOT")
                await page.locator(f"xpath={xpath_datepicker_tgl_pemeriksaan}").click()
                await page.wait_for_timeout(500)
                
            # Sekarang cari tombol tanggal today di dalam container tersebut
            day_btn = None
            try:
                # Cari button yang memiliki indikator bulatan berwarna hijau/tosca (rgb(0, 184, 172))
                dot_btn = calendar_container.locator("button").filter(
                    has=page.locator("span[style*='0, 184, 172']")
                ).first
                if await dot_btn.count() > 0 and await is_element_really_visible(dot_btn):
                    day_btn = dot_btn
                    log("  ✓ Menemukan hari ini via indikator bulatan berwarna.", "OK")
            except Exception:
                pass
            
            # Fallback: cari berdasarkan angka tanggal hari ini
            if not day_btn or await day_btn.count() == 0:
                day_btn = calendar_container.locator("button, div.cursor-pointer").filter(
                    has_text=re.compile(rf"^\s*{today_day}\s*$", re.IGNORECASE)
                ).first
            
            if await day_btn.count() > 0:
                await day_btn.click()
                await page.wait_for_timeout(200)
                log(f"  ✓ Tanggal {today_day} berhasil dipilih.", "OK")
            else:
                log(f"  ⚠️ Tidak menemukan tombol tanggal {today_day} di calendar container. Mencoba scan grid...", "WARN")
                grid_loc = page.locator(".form-data-individu form .grid-cols-7.mt-2, div.grid-cols-7.mt-2")
                day_btn_alt = grid_loc.locator("div.cursor-pointer, button").filter(
                    has_text=re.compile(rf"^\s*{today_day}\s*$", re.IGNORECASE)
                ).first
                await day_btn_alt.click()
                await page.wait_for_timeout(200)
                log(f"  ✓ Tanggal {today_day} berhasil dipilih (scan grid).", "OK")
        except Exception as e:
            log(f"  ⚠️ Gagal memilih tanggal pemeriksaan: {e}", "WARN")
    else:
        # ── Isi Nama ──────────────────────────────────────────────
        log(f"  ✏️  Mengisi Nama: {nama}", "BOT")
        await page.get_by_role("textbox", name="Nama Lengkap *").click()
        await page.get_by_role("textbox", name="Nama Lengkap *").fill(nama)
        await page.wait_for_timeout(200)

        # ── Tanggal Lahir ─────────────────────────────────────────
        await pilih_tanggal(page, tgl.year, tgl.month, tgl.day)

        # ── Jenis Kelamin ─────────────────────────────────────────
        log(f"  ✏️  Memilih Jenis Kelamin: {jk}", "BOT")
        await page.locator("div").filter(
            has_text=re.compile(r"^Pilih jenis kelamin$")
        ).nth(1).click()
        await page.wait_for_timeout(400)
        if jk == "LAKI-LAKI":
            await page.locator("form").get_by_text("Laki-laki").click()
        else:
            await page.locator("form").get_by_text("Perempuan").click()
        await page.wait_for_timeout(200)

        # ── No. WhatsApp ──────────────────────────────────────────
        # Gunakan nomor HP dari Excel jika ada, jika tidak gunakan nomor HP default dari GUI
        whatsapp_val = no_hp
        if not whatsapp_val and GUI_BRIDGE_ACTIVE and gui_bridge:
            whatsapp_val = getattr(gui_bridge, "default_phone", "")
            
        if not whatsapp_val:
            whatsapp_val = "" # Fallback default (kosong)
            
        log(f"  ✏️  Mengisi No. WhatsApp: {whatsapp_val}", "BOT")
        try:
            whatsapp_input = page.locator("xpath=//input[@id='No Whatsapp']")
            await whatsapp_input.click(timeout=4000)
            await whatsapp_input.fill(whatsapp_val)
            await page.wait_for_timeout(300)
        except Exception as e:
            log(f"  ⚠️  Gagal mengisi No. WhatsApp via XPath: {e}. Mencoba get_by_role...", "WARN")
            try:
                await page.get_by_role("textbox", name="No. Whatsapp Aktif * +").click(timeout=3000)
                await page.get_by_role("textbox", name="No. Whatsapp Aktif * +").fill(whatsapp_val)
                await page.wait_for_timeout(300)
            except Exception as e2:
                log(f"  ⚠️  Gagal mengisi No. WhatsApp: {e2}", "WARN")

        # ── Tanggal Pemeriksaan (Auto Pick Today) ─────────────────
        today_day = str(datetime.now().day)
        log(f"  🗓  Memilih tanggal pemeriksaan (Hari ini): {today_day}", "BOT")
        try:
            grid = page.locator(".form-data-individu form .grid-cols-7.mt-2")
            day_buttons = grid.locator("button")
            clicked = False
            for i in range(await day_buttons.count()):
                btn = day_buttons.nth(i)
                span = btn.locator("span.font-bold").first
                if await span.count() > 0:
                    val = (await span.text_content()).strip()
                    if val == today_day:
                        await btn.click()
                        clicked = True
                        await page.wait_for_timeout(400)
                        # Cek popup Kuota Habis setelah klik tanggal
                        await klik_popup_kuota(page)
                        break
            if clicked:
                log(f"  ✓ Tanggal pemeriksaan (hari ini: {today_day}) berhasil dipilih.", "OK")
            else:
                raise Exception("Hari tidak ditemukan di grid")
        except SkipPasien:
            raise
        except Exception as e:
            log(f"  ⚠️  Gagal memilih tanggal pemeriksaan otomatis via grid: {e}. Mencoba get_by_role...", "WARN")
            try:
                await page.get_by_role("button", name=today_day, exact=True).click(timeout=3000)
                await page.wait_for_timeout(400)
                await klik_popup_kuota(page)
                log(f"  ✓ Tanggal pemeriksaan (hari ini: {today_day}) dipilih menggunakan get_by_role.", "OK")
            except SkipPasien:
                raise
            except Exception as e2:
                log(f"  ⚠️  Gagal memilih tanggal pemeriksaan: {e2}.", "WARN")

    # ── Cek apakah muncul section "Isi Data Wali" ──────────────────
    wali_header = page.locator("div").filter(
        has_text=re.compile(r"^Isi Data Wali$", re.IGNORECASE)
    )
    if await wali_header.count() == 0:
        wali_header = page.get_by_text("Isi Data Wali", exact=True)

    if await wali_header.count() > 0 and await is_element_really_visible(wali_header.first):
        log("  👤 Terdeteksi form 'Isi Data Wali'", "BOT")

        # Tampilkan 10 baris di atas pasien saat ini
        excel_row_current = "?"
        try:
            temp_df_list = pd.read_excel(FILE_EXCEL, usecols=[0, 1])
            temp_df_list.columns = ["No", "Nama Lengkap"]
            
            matched_current = temp_df_list[temp_df_list["No"] == no_excel]
            if not matched_current.empty:
                current_idx = matched_current.index[0]
                excel_row_current = current_idx + 2
                
                # Ambil 10 baris di atas index saat ini
                start_idx = max(0, current_idx - 10)
                above_rows = temp_df_list.iloc[start_idx:current_idx]
                if not above_rows.empty:
                    print("  📋 10 Baris Pasien di Atas:")
                    for idx, r_above in above_rows.iterrows():
                        excel_row = idx + 2
                        try:
                            no_val = int(r_above['No'])
                        except Exception:
                            no_val = r_above['No']
                        print(f"     [Baris Excel: {excel_row:>3}] No. {no_val:>2} | {str(r_above['Nama Lengkap']).strip()}")
                    print()
        except Exception as ex_list:
            log(f"    ⚠️ Gagal memuat daftar baris di atas: {ex_list}", "WARN")

        # Konfirmasi nomor baris excel untuk data wali
        if GUI_BRIDGE_ACTIVE and gui_bridge:
            above_rows_data = []
            try:
                for idx, r_above in above_rows.iterrows():
                    above_rows_data.append({
                        "excel_row": idx + 2,
                        "no": int(r_above['No']) if pd.notna(r_above['No']) else r_above['No'],
                        "nama": str(r_above['Nama Lengkap']).strip()
                    })
            except Exception:
                pass
            row_target_input = gui_bridge.request_input("wali", {
                "no_excel": no_excel,
                "nama": nama,
                "above_rows": above_rows_data
            })
        else:
            print()
            print("  ╔══════════════════════════════════════════╗")
            print("  ║               ISI DATA WALI              ║")
            print("  ╠══════════════════════════════════════════╣")
            print(f"  ║  Pasien Saat Ini: No. {no_excel} (Baris Excel: {excel_row_current})")
            print(f"  ║  Nama           : {nama}")
            print("  ╚══════════════════════════════════════════╝")
            row_target_input = input(f"  >> Masukkan No. baris pasien di Excel (kolom 'No') untuk dijadikan Wali [Default: {no_excel}]: ").strip()

        target_no = no_excel
        if row_target_input:
            try:
                target_no = int(row_target_input)
            except ValueError:
                log(f"    ⚠️ Input tidak valid, menggunakan default No. {no_excel}", "WARN")
                target_no = no_excel

        # Data wali default = data pasien itu sendiri
        wali_nik  = nik
        wali_nama = nama
        wali_jk   = jk
        wali_tgl  = tgl  # pd.Timestamp

        if target_no != no_excel:
            try:
                temp_df = pd.read_excel(FILE_EXCEL, usecols=[0, 1, 2, 3, 5], dtype={2: str})
                temp_df.columns = ["No", "Nama Lengkap", "NIK", "Jenis Kelamin", "Tanggal Lahir"]
                matched_rows = temp_df[temp_df["No"] == target_no]
                if not matched_rows.empty:
                    target_row_data = matched_rows.iloc[0]
                    wali_nama = str(target_row_data["Nama Lengkap"]).strip()
                    # Parsing Wali NIK secara aman
                    wali_nik_raw = str(target_row_data["NIK"]).strip()
                    if wali_nik_raw.endswith(".0"):
                        wali_nik_raw = wali_nik_raw[:-2]
                    wali_nik = "".join(re.findall(r"\d+", wali_nik_raw)).zfill(16)
                    wali_jk   = str(target_row_data["Jenis Kelamin"]).strip().upper()
                    wali_tgl  = pd.Timestamp(target_row_data["Tanggal Lahir"])
                    log(f"    👉 Menggunakan data No. {target_no} ({wali_nama}) sebagai Wali", "OK")
                else:
                    log(f"    ⚠️ No. {target_no} tidak ditemukan di Excel. Menggunakan data pasien saat ini.", "WARN")
            except Exception as ex:
                log(f"    ⚠️ Gagal mencari data wali di Excel: {ex}. Menggunakan data pasien saat ini.", "WARN")

        # 1. NIK Wali
        try:
            nik_wali_field = page.get_by_placeholder("Masukkan NIK Wali").first
            if await nik_wali_field.count() == 0:
                nik_wali_field = page.locator("input[id='nik wali'], input[name='NIK wali'], input[placeholder*='NIK Wali']").first
            if await nik_wali_field.count() == 0:
                nik_wali_field = page.locator("input[maxlength='16']").last
            
            await js_click(nik_wali_field)
            await nik_wali_field.fill(wali_nik)
            log(f"    ✅ NIK Wali: {wali_nik}", "OK")
        except Exception as e:
            log(f"    ⚠️ Gagal isi NIK Wali: {e}", "WARN")
        await page.wait_for_timeout(300)

        # 2. Nama Lengkap Wali
        try:
            nama_wali_field = page.get_by_role("textbox", name="Masukkan Nama Lengkap").first
            if await nama_wali_field.count() == 0:
                nama_wali_field = page.get_by_role("textbox", name="Nama Lengkap *").first
            if await nama_wali_field.count() == 0:
                nama_wali_field = page.locator("input[id='Nama Lengkap'], input[name='Nama Lengkap Wali'], input[placeholder*='Nama Lengkap']").first
            if await nama_wali_field.count() == 0:
                nama_wali_field = page.locator("input[maxlength='300']").first
                
            await js_click(nama_wali_field)
            await nama_wali_field.fill(wali_nama)
            log(f"    ✅ Nama Wali: {wali_nama}", "OK")
        except Exception as e:
            log(f"    ⚠️ Gagal isi Nama Wali: {e}", "WARN")
        await page.wait_for_timeout(300)

        # 3. Tanggal Lahir Wali
        try:
            await pilih_tanggal_wali(page, wali_tgl.year, wali_tgl.month, wali_tgl.day)
            log(f"    ✅ Tgl Lahir Wali: {wali_tgl.strftime('%d-%m-%Y')}", "OK")
        except Exception as e:
            log(f"    ⚠️ Gagal isi Tgl Lahir Wali: {e}", "WARN")
        await page.wait_for_timeout(400)

        # 4. Jenis Kelamin Wali
        try:
            mapped_jk = "Perempuan" if ("P" in wali_jk or "PEREMPUAN" in wali_jk) else "Laki-laki"
            log(f"  ✏️  Memilih Jenis Kelamin Wali: {mapped_jk}", "BOT")

            # 1. Coba klik trigger dropdown menggunakan XPath yang benar dari user
            xpath_jk_trigger = 'xpath=//*[@id="__nuxt"]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[5]/div[2]/div/div/div[6]/div/form/div[1]/div[1]/div[7]/div/div[2]/div[4]/div/div[2]/div[1]'
            clicked_trigger = False
            try:
                await page.locator(xpath_jk_trigger).click(timeout=3000)
                clicked_trigger = True
                log("    ✓ Trigger jenis kelamin wali diklik menggunakan XPath utama.", "OK")
            except Exception:
                pass

            # 2. Fallback: gunakan metode "Pilih jenis kelamin" yang disamakan dengan form utama
            if not clicked_trigger:
                try:
                    await page.locator("div").filter(
                        has_text=re.compile(r"^Pilih jenis kelamin$", re.IGNORECASE)
                    ).nth(2).click(timeout=3000)
                    clicked_trigger = True
                    log("    ✓ Trigger jenis kelamin wali diklik menggunakan filter text (index 2).", "OK")
                except Exception:
                    try:
                        await page.locator("div").filter(
                            has_text=re.compile(r"^Pilih jenis kelamin$", re.IGNORECASE)
                        ).nth(1).click(timeout=3000)
                        clicked_trigger = True
                        log("    ✓ Trigger jenis kelamin wali diklik menggunakan filter text (index 1).", "OK")
                    except Exception:
                        pass

            await page.wait_for_timeout(400)

            # Setelah diklik, gunakan metode yang sama seperti form utama untuk memilih opsi jenis kelamin
            # tetapi utamakan XPath opsi spesifik agar tidak salah pilih ke form utama di latar belakang.
            clicked_option = False
            xpath_opt_laki = 'xpath=//*[@id="__nuxt"]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[5]/div[2]/div/div/div[6]/div/form/div[1]/div[1]/div[7]/div/div[2]/div[4]/div/div[2]/div[3]/div/div[1]'
            xpath_opt_perempuan = 'xpath=//*[@id="__nuxt"]/main/div/div[1]/section[2]/div/div/div/div[2]/div/div[3]/div[5]/div[2]/div/div/div[6]/div/form/div[1]/div[1]/div[7]/div/div[2]/div[4]/div/div[2]/div[3]/div/div[2]'
            
            target_xpath_opt = xpath_opt_perempuan if mapped_jk == "Perempuan" else xpath_opt_laki
            try:
                await page.locator(target_xpath_opt).click(timeout=3000)
                clicked_option = True
                log(f"    ✓ Opsi Jenis Kelamin Wali '{mapped_jk}' berhasil dipilih via XPath.", "OK")
            except Exception:
                pass

            if not clicked_option:
                # Fallback: Gunakan metode form utama
                if mapped_jk == "Laki-laki":
                    await page.locator("form").get_by_text("Laki-laki").click()
                else:
                    await page.locator("form").get_by_text("Perempuan").click()
                log(f"    ✓ Opsi Jenis Kelamin Wali '{mapped_jk}' dipilih via text-matching fallback.", "OK")
            
            await page.wait_for_timeout(200)
            log(f"    ✅ Jenis Kelamin Wali: {mapped_jk}", "OK")
        except Exception as e:
            log(f"    ⚠️ Gagal pilih JK Wali: {e}", "WARN")
        await page.wait_for_timeout(300)

        # 5. Centang checkbox "phone-sama"
        try:
            chk_phone = page.locator("#phone-sama")
            clicked_chk = False
            for i in range(await chk_phone.count()):
                el = chk_phone.nth(i)
                if await is_element_really_visible(el):
                    await js_click(el)
                    clicked_chk = True
                    break
            if not clicked_chk:
                if await chk_phone.count() > 1:
                    await js_click(chk_phone.nth(1))
                else:
                    await js_click(chk_phone.first)
            log("    ✅ Checkbox 'phone-sama' dicentang.", "OK")
        except Exception as e:
            log(f"    ⚠️ Gagal centang phone-sama: {e}", "WARN")
        await page.wait_for_timeout(300)

        log("  ✅ Selesai mengisi Data Wali.", "OK")
    # ── End: Isi Data Wali ──────────────────────────────────────────

    # ── Submit Step 1 ─────────────────────────────────────────
    log("  🚀 Mengklik Selanjutnya (Step 1)…", "BOT")
    data_tidak_valid = False
    for _ in range(3):
        try:
            try:
                await page.locator(f"xpath={xpath_btn_selanjutnya_step1}").first.click(timeout=3000)
            except Exception:
                await page.get_by_role("button", name="Selanjutnya").click(timeout=5000)
            await page.wait_for_timeout(100)
            
            # Cek popup tidak valid
            if await cek_popup_tidak_valid(page):
                data_tidak_valid = True
                break

            # Tangani popup Kuota Pemeriksaan Habis setelah submit
            await klik_popup_kuota(page)
            await page.wait_for_timeout(200)
            
            # Cek jika tombol Selanjutnya sudah tidak terlihat (halaman berhasil berpindah)
            btn_sel_xpath = page.locator(f"xpath={xpath_btn_selanjutnya_step1}").first
            btn_sel_role = page.get_by_role("button", name="Selanjutnya").first
            
            if not (await btn_sel_xpath.is_visible()) and not (await btn_sel_role.is_visible()):
                log("  ✓ Berhasil submit Step 1 (halaman berpindah).", "OK")
                break
        except SkipPasien:
            raise
        except Exception:
            break

    # ── Isi Data Pendukung (Modal) jika data valid ────────────
    if not data_tidak_valid:
        xpath_status = (
            "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
            "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
            "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
            "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
            "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
            "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
            "/div[@class='p-2']/div/div[6]/form/div[@class='flex gap-5 w-full lt-sm:flex-col']"
            "/div[@class='flex flex-col gap-3 w-full']/div[2]/div[@class='w-full']"
            "/div[@class='relative text-black border-1 border-solid font-medium flex border-gray-3 focus-within:border-black border-rd-l-lg border-rd-r-lg']"
            "/div[@class='h-[2.9rem] w-full flex gap-2 cursor-pointer items-center justify-start overflow-hidden border-none bg-transparent pl-4 text-sm focus:outline-none text-black']"
        )

        xpath_disabilitas = (
            "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
            "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
            "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
            "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
            "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
            "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
            "/div[@class='p-2']/div/div[6]/form/div[@class='flex gap-5 w-full lt-sm:flex-col']"
            "/div[@class='flex flex-col gap-3 w-full']/div[3]/div[@class='w-full']"
            "/div[@class='relative text-black border-1 border-solid font-medium flex border-gray-3 focus-within:border-black border-rd-l-lg border-rd-r-lg']"
            "/div[@class='h-[2.9rem] w-full flex gap-2 cursor-pointer items-center justify-start overflow-hidden border-none bg-transparent pl-4 text-sm focus:outline-none text-black']"
        )

        xpath_pekerjaan = (
            "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
            "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
            "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
            "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
            "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
            "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
            "/div[@class='p-2']/div/div[6]/form/div[@class='flex gap-5 w-full lt-sm:flex-col']"
            "/div[@class='flex flex-col gap-3 w-full']/div[4]/div[@class='w-full']"
            "/div[@class='relative text-black border-1 border-solid font-medium flex border-gray-3 focus-within:border-black border-rd-l-lg border-rd-r-lg']"
            "/div[@class='h-[2.9rem] w-full flex gap-2 cursor-pointer items-center justify-start overflow-hidden border-none bg-transparent pl-4 text-sm focus:outline-none text-black']"
        )

        xpath_domisili = (
            "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
            "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
            "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
            "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
            "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
            "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
            "/div[@class='p-2']/div/div[6]/form/div[@class='flex gap-5 w-full lt-sm:flex-col']"
            "/div[@class='flex flex-col gap-3 w-full']/div[5]/div/div[@class='relative border-1 border-solid rounded-lg font-medium flex border-gray-3 focus-within:border-black']"
            "/div[@class='min-h-[2.9rem] w-full flex cursor-pointer items-center justify-start overflow-hidden border-none bg-transparent pl-4 text-sm focus:outline-none text-black']"
        )

        xpath_btn_container = (
            "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
            "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
            "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
            "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
            "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
            "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
            "/div[@class='p-2']/div/div[6]/form/div[@class='flex justify-end gap-2 items-center py-3']/div[2]/*"
        )

        try:
            log("  ⏳ Menunggu form data pendukung (modal) terbuka atau langsung ke Individu Terdaftar...", "WAIT")
            
            # Tentukan selector untuk mendeteksi apakah kita sudah melewati modal data pendukung
            xpath_btn_pilih_check = (
                "//div[@class='overflow-auto table-individu-terdaftar']"
                "//button[contains(@class, 'btn-outline-primary')]"
            )
            
            modal_terbuka = False
            card_modal = None
            for _ in range(60):  # 60 * 200ms = 12 detik max wait (lebih lama karena kuota bisa lambat)
                # 1. Pastikan popup kuota sudah benar-benar hilang sebelum memeriksa modal
                try:
                    btn_pilih_lain = page.locator("button").filter(
                        has_text=re.compile(r"Pilih Tanggal Lain", re.IGNORECASE)
                    ).first
                    if await btn_pilih_lain.is_visible():
                        # Popup kuota MASIH ada — coba tutup sekali lagi
                        await klik_popup_kuota(page)
                        await page.wait_for_timeout(500)
                        continue
                except Exception:
                    pass

                # 2. Cek jika popup 'Data peserta valid' masih terbuka, klik 'Lanjutkan'
                try:
                    popup_peserta_valid = page.locator("div.fixed.z-1000, div.fixed.z-1100, div[class*='z-1000'], div[class*='z-1100']").filter(
                        has_text=re.compile(r"Data peserta valid", re.IGNORECASE)
                    ).first
                    if await popup_peserta_valid.count() > 0 and await popup_peserta_valid.is_visible():
                        log("  ⚠️  Popup 'Data peserta valid' terdeteksi masih terbuka di loop modal. Mengeklik Lanjutkan...", "WARN")
                        btn_lanjutkan = popup_peserta_valid.locator("button, [role='button'], div.cursor-pointer").filter(
                            has_text=re.compile(r"Lanjutkan", re.IGNORECASE)
                        ).first
                        if await btn_lanjutkan.count() > 0 and await btn_lanjutkan.is_visible():
                            if not await force_js_click(btn_lanjutkan):
                                await btn_lanjutkan.click(timeout=2000)
                        else:
                            btn_any = popup_peserta_valid.locator("button, [role='button'], div.cursor-pointer").first
                            if await btn_any.count() > 0:
                                await force_js_click(btn_any)
                        await page.wait_for_timeout(500)
                        continue
                except Exception:
                    pass

                # 3. Cek jika popup 'Data peserta tidak valid' masih terbuka, klik 'Periksa Kembali' dan skip pasien
                try:
                    popup_peserta_tidak_valid = page.locator("div.fixed.z-1000, div.fixed.z-1100, div[class*='z-1000'], div[class*='z-1100']").filter(
                        has_text=re.compile(r"Data peserta tidak valid", re.IGNORECASE)
                    ).first
                    if await popup_peserta_tidak_valid.count() > 0 and await popup_peserta_tidak_valid.is_visible():
                        log("  ⚠️  Popup 'Data peserta tidak valid' terdeteksi terbuka di loop modal. Mengeklik Periksa Kembali dan men-skip pasien...", "WARN")
                        btn_periksa = popup_peserta_tidak_valid.locator("button, [role='button'], div.cursor-pointer").filter(
                            has_text=re.compile(r"(Periksa|Tutup|Kembali|OK)", re.IGNORECASE)
                        ).first
                        if await btn_periksa.count() > 0 and await btn_periksa.is_visible():
                            if not await force_js_click(btn_periksa):
                                await btn_periksa.click(timeout=2000)
                        else:
                            btn_any = popup_peserta_tidak_valid.locator("button, [role='button'], div.cursor-pointer").first
                            if await btn_any.count() > 0:
                                await force_js_click(btn_any)
                        await page.wait_for_timeout(500)
                        raise SkipPasien("Kategori pasien dilewati (data tidak valid terdeteksi di loop modal)")
                except SkipPasien:
                    raise
                except Exception:
                    pass

                # Cek jika modal data pendukung terlihat
                # Gunakan keyword yang UNIK untuk modal data pendukung:
                # 'status pernikahan' + 'pekerjaan' + 'alamat domisili' — tidak ada di popup kuota
                possible_card = page.locator("div.fixed.z-1000 div.rounded-lg.bg-white, div.fixed.z-1100 div.rounded-lg.bg-white").first
                if await is_element_really_visible(possible_card):
                    text = await possible_card.inner_text()
                    text_lower = text.lower()
                    # Harus mengandung 'status pernikahan' DAN ('pekerjaan' ATAU 'alamat')
                    # Ini sangat spesifik untuk modal data pendukung, tidak akan cocok dengan popup kuota
                    has_pernikahan = "status pernikahan" in text_lower
                    has_form_fields = "pekerjaan" in text_lower or "alamat domisili" in text_lower
                    if has_pernikahan and has_form_fields:
                        modal_terbuka = True
                        card_modal = possible_card
                        break
                
                # Cek jika tabel individu terdaftar atau tombol pilih sudah muncul langsung
                btn_pilih = page.locator(xpath_btn_pilih_check).first
                if await is_element_really_visible(btn_pilih):
                    log("  ✨ Terdeteksi langsung masuk ke tabel Individu Terdaftar (form data pendukung di-skip/sudah diisi).", "OK")
                    break
                    
                await page.wait_for_timeout(200)

            if modal_terbuka and card_modal:
                log("  ✓ Form data pendukung (modal) terbuka.", "OK")
                
                # Fungsi pembantu untuk mencari input/dropdown di dalam modal card secara dinamis
                async def get_modal_field(label_text, fallback_xpath):
                    try:
                        # Cari berdasarkan teks label
                        el = card_modal.locator("div").filter(has_text=re.compile(rf"^\s*({label_text})", re.IGNORECASE)).locator("div.cursor-pointer, div[class*='cursor-pointer'], textarea, input").first
                        if await el.count() > 0 and await is_element_really_visible(el):
                            return el
                    except Exception:
                        pass
                    try:
                        # Cari berdasarkan teks placeholder "Pilih..."
                        el_placeholder = card_modal.locator("div.cursor-pointer, div[class*='cursor-pointer']").filter(has_text=re.compile(rf"Pilih\s+({label_text})", re.IGNORECASE)).first
                        if await el_placeholder.count() > 0 and await is_element_really_visible(el_placeholder):
                            return el_placeholder
                    except Exception:
                        pass
                    return page.locator(f"xpath={fallback_xpath}").first

                # 1. Status Pernikahan (Kolom G)
                status_perkawinan = str(row.get("Status Pernikahan", "")).strip().upper()
                log(f"  ✏️  Memilih Status Pernikahan: {status_perkawinan}", "BOT")
                field_status = await get_modal_field("Status Pernikahan|Status Perkawinan", xpath_status)
                await js_click(field_status)
                await page.wait_for_timeout(600)
                
                # Gunakan pencocokan substring agar lebih fleksibel terhadap variasi data Excel
                if "BELUM" in status_perkawinan or "BK" in status_perkawinan or status_perkawinan == "":
                    mapped_status = "Belum Menikah"
                elif "KAWIN" in status_perkawinan or "MENIKAH" in status_perkawinan:
                    mapped_status = "Menikah"
                elif "CERAI HIDUP" in status_perkawinan or status_perkawinan == "CH":
                    mapped_status = "Cerai Hidup"
                elif "CERAI MATI" in status_perkawinan or status_perkawinan == "CM":
                    mapped_status = "Cerai Mati"
                else:
                    mapped_status = "Belum Menikah"  # Default ke Belum Menikah jika tidak cocok

                try:
                    # Cari opsi yang benar-benar visible pada modal overlay (menghindari mengklik element latar belakang)
                    option_perkawinan = page.locator(
                        "div.modal-content button, div.modal-content div.cursor-pointer, div.modal-content .list-group-item, div.modal-content [role='option'], "
                        "div[class*='z-9000'] button, div[class*='z-9000'] div.cursor-pointer, div[class*='z-9000'] [role='option'], div[class*='z-9000'] .list-group-item, "
                        "div[class*='z-1100'] button, div[class*='z-1100'] div.cursor-pointer, div[class*='z-1100'] [role='option'], div[class*='z-1100'] .list-group-item, "
                        "div.fixed.z-9000 button, div.fixed.z-9000 div.cursor-pointer, div.fixed.z-9000 .list-group-item, "
                        "div.fixed.z-1000 button, div.fixed.z-1000 div.cursor-pointer, div.fixed.z-1000 .list-group-item, "
                        "div.fixed button, div.fixed div.cursor-pointer, div.fixed .list-group-item"
                    ).filter(
                        has_text=re.compile(rf"{mapped_status}", re.IGNORECASE)
                    )
                    clicked = False
                    for i in range(await option_perkawinan.count()):
                        btn = option_perkawinan.nth(i)
                        if await is_element_really_visible(btn):
                            if await js_click(btn):
                                clicked = True
                                break
                    if not clicked:
                        await js_click(option_perkawinan.first)
                except Exception as e:
                    log(f"  ⚠️ Gagal memilih opsi perkawinan: {e}", "WARN")
                await page.wait_for_timeout(400)

                # 2. Penyandang Disabilitas (Default: Tidak memiliki disabilitas)
                log("  ✏️  Memilih Penyandang disabilitas: Tidak memiliki disabilitas", "BOT")
                field_disabilitas = await get_modal_field("Penyandang disabilitas", xpath_disabilitas)
                await js_click(field_disabilitas)
                await page.wait_for_timeout(600)
                try:
                    option_disabilitas = page.locator(
                        "div.modal-content button, div.modal-content div.cursor-pointer, div.modal-content .list-group-item, div.modal-content [role='option'], "
                        "div[class*='z-9000'] button, div[class*='z-9000'] div.cursor-pointer, div[class*='z-9000'] [role='option'], div[class*='z-9000'] .list-group-item, "
                        "div[class*='z-1100'] button, div[class*='z-1100'] div.cursor-pointer, div[class*='z-1100'] [role='option'], div[class*='z-1100'] .list-group-item, "
                        "div.fixed.z-9000 button, div.fixed.z-9000 div.cursor-pointer, div.fixed.z-9000 .list-group-item, "
                        "div.fixed.z-1000 button, div.fixed.z-1000 div.cursor-pointer, div.fixed.z-1000 .list-group-item, "
                        "div.fixed button, div.fixed div.cursor-pointer, div.fixed .list-group-item"
                    ).filter(
                        has_text=re.compile(r"Tidak memiliki disabilitas", re.IGNORECASE)
                    )
                    clicked = False
                    for i in range(await option_disabilitas.count()):
                        btn = option_disabilitas.nth(i)
                        if await is_element_really_visible(btn):
                            if await js_click(btn):
                                clicked = True
                                break
                    if not clicked:
                        await js_click(option_disabilitas.first)
                except Exception as e:
                    log(f"  ⚠️ Gagal memilih opsi disabilitas: {e}", "WARN")
                await page.wait_for_timeout(400)

                # 3. Pekerjaan (Kolom H)
                pekerjaan_val = str(row.get("Pekerjaan", "")).strip()
                log(f"  ✏️  Memilih Pekerjaan: {pekerjaan_val}", "BOT")
                field_pekerjaan = await get_modal_field("Pekerjaan", xpath_pekerjaan)
                await js_click(field_pekerjaan)
                await page.wait_for_timeout(600)

                mapped_pekerjaan = "Belum/Tidak Bekerja"
                if "WIRASWASTA" in pekerjaan_val.upper():
                    mapped_pekerjaan = "Wiraswasta"
                elif "MENGURUS RUMAH TANGGA" in pekerjaan_val.upper() or "RUMAH TANGGA" in pekerjaan_val.upper():
                    mapped_pekerjaan = "Mengurus Rumah Tangga"
                elif "BELUM" in pekerjaan_val.upper() or "TIDAK BEKERJA" in pekerjaan_val.upper():
                    mapped_pekerjaan = "Belum/Tidak Bekerja"
                elif "TIDAK TERBACA" in pekerjaan_val.upper():
                    mapped_pekerjaan = "Belum/Tidak Bekerja"

                option_pekerjaan = page.locator(
                    "div.modal-content button, div.modal-content div.cursor-pointer, div.modal-content .list-group-item, div.modal-content [role='option'], "
                    "div[class*='z-9000'] button, div[class*='z-9000'] div.cursor-pointer, div[class*='z-9000'] [role='option'], div[class*='z-9000'] .list-group-item, "
                    "div.z-9000 button, div.z-9000 div.cursor-pointer, div.z-9000 .list-group-item, "
                    "div.fixed.z-9000 button, div.fixed.z-9000 div.cursor-pointer, div.fixed.z-9000 .list-group-item, "
                    "div.fixed.z-1000 button, div.fixed.z-1000 div.cursor-pointer, div.fixed.z-1000 .list-group-item, "
                    "div.fixed button, div.fixed div.cursor-pointer, div.fixed .list-group-item"
                ).filter(
                    has_text=re.compile(rf"{mapped_pekerjaan}", re.IGNORECASE)
                ).first
                try:
                    if await option_pekerjaan.count() > 0 and await is_element_really_visible(option_pekerjaan):
                        await js_click(option_pekerjaan)
                    else:
                        xpath_btn_pekerjaan = (
                            "/html/body/div[@class='fixed top-0 left-0 z-9000 w-full h-full flex justify-center items-center backdrop-blur-5']"
                            "/div[@class='modal-content overflow-auto w-full sm:w-[400px] !overflow-auto !max-h-[500px] w-[40%] lt-sm:w-full slide-down']"
                            "/div[@class='mt-4 px-1']/div/div[5]"
                            "/button[@class='py-4 pl-0 w-full text-left border-none border-b border-b-solid border-b-gray-200 rounded-0 hover:bg-transparent hover:opacity-50 text-gray-600']"
                            "/div[@class='flex items-center justify-between gap-2']"
                        )
                        await js_click(page.locator(f"xpath={xpath_btn_pekerjaan}"))
                except Exception as e:
                    log(f"  ⚠️ Gagal memilih Pekerjaan, mencoba mengklik xpath alternatif: {e}", "WARN")
                    try:
                        await js_click(page.get_by_role("button", name=mapped_pekerjaan, exact=False))
                    except Exception:
                        pass
                await page.wait_for_timeout(400)

                # 4. Alamat Domisili (Kolom J) — Pengisian Otomatis Cascading (Provinsi > Kabupaten > Kecamatan > Kelurahan)
                alamat_domisili = str(row.get("Alamat Domisili", "")).strip()
                log(f"  ✏️  Alamat Domisili: {alamat_domisili}", "BOT")
                field_domisili = await get_modal_field("Alamat Domisili", xpath_domisili)
                await js_click(field_domisili)
                await page.wait_for_timeout(800)

                parts = [p.strip() for p in alamat_domisili.split(",")]
                desa_val = parts[0] if len(parts) > 0 else "-"
                kec_val  = parts[1] if len(parts) > 1 else "-"
                kab_val  = parts[2] if len(parts) > 2 else "-"
                prov_val = parts[3] if len(parts) > 3 else "-"
                
                # Coba pengisian alamat secara otomatis dengan tahapan Provinsi > Kabupaten > Kecamatan > Kelurahan
                success_auto_address = False
                try:
                    stages = [
                        {"name": "Provinsi", "val": prov_val, "index": 0},
                        {"name": "Kabupaten", "val": kab_val, "index": 1},
                        {"name": "Kecamatan", "val": kec_val, "index": 2},
                        {"name": "Kelurahan", "val": desa_val, "index": 3}
                    ]
                    
                    log("  🤖 Memulai pengisian alamat otomatis secara berjenjang...", "BOT")
                    
                    completed_stages = 0
                    for stage in stages:
                        check_stop_request()
                        name = stage["name"]
                        val = stage["val"]
                        idx = stage["index"]
                        
                        if val == "-":
                            log(f"    ⚠️ Nilai untuk {name} tidak valid, melewati otomatisasi.", "WARN")
                            break
                            
                        log(f"    ⏳ Memproses {name}: '{val}'...", "WAIT")
                        
                        # 1. Cari input pencarian untuk tahap ini
                        # KUNCI: Overlay menampilkan 1 input sekaligus secara bergantian (cascade):
                        # Provinsi → (setelah dipilih) → Kabupaten → Kecamatan → Kelurahan
                        # Jadi selalu cari input PERTAMA (first) yang visible di overlay, bukan nth(idx)
                        search_input = None
                        
                        # Tunggu input overlay muncul (hingga 3 detik)
                        for _wait in range(15):  # 15 * 200ms = 3 detik
                            locators_search = [
                                page.locator(f"xpath=/html/body/div[3]/div[2]/div[{idx + 1}]/input"),
                                page.locator("div.fixed.z-9000 input").first,
                                page.locator("div[class*='z-9000'] input").first,
                            ]
                            for loc in locators_search:
                                try:
                                    if await loc.count() > 0 and await is_element_really_visible(loc.first):
                                        search_input = loc.first
                                        break
                                except Exception:
                                    pass
                            if search_input:
                                break
                            await page.wait_for_timeout(200)
                                
                        if not search_input:
                            log(f"    ⚠️ Kolom input {name} tidak ditemukan atau belum aktif.", "WARN")
                            break
                            
                        # 2. Klik, fokus, isi input dan kirim ketikan
                        log(f"      🤖 Mengisi {name} dengan: '{val}'...", "BOT")
                        await js_click(search_input)
                        await page.wait_for_timeout(300)
                        await search_input.fill("")
                        await search_input.type(val, delay=80)
                        # 3. Cari dan klik item hasil pencarian yang cocok (dengan polling hingga 3.6 detik)
                        stage_clicked = False
                        
                        log(f"      ⏳ Menunggu opsi pencarian untuk {name} muncul...", "WAIT")
                        for _search_wait in range(18):  # 18 * 200ms = 3.6 detik max wait
                            # LAPISAN 1: Cari opsi secara langsung menggunakan locator teks di dalam overlay z-9000
                            try:
                                # Cari di elemen list/button/div di dalam overlay yang mengandung teks target
                                opt_locators = [
                                    page.locator("div.fixed.z-9000, div[class*='z-9000']").locator("li, button, [role='option'], div.cursor-pointer").get_by_text(val, exact=False),
                                    page.locator("div.fixed.z-9000, div[class*='z-9000']").locator("div").get_by_text(val, exact=False)
                                ]
                                
                                for loc in opt_locators:
                                    cnt = await loc.count()
                                    for i in range(cnt):
                                        opt = loc.nth(i)
                                        if await opt.is_visible():
                                            text = (await opt.text_content() or "").strip()
                                            # Hindari container besar (teks > 100 karakter)
                                            if not text or len(text) < 2 or len(text) > 100:
                                                continue
                                            # Filter kata pencarian fiktif
                                            if "pencarian" in text.lower() or "cari" in text.lower() or "pilih" in text.lower():
                                                continue
                                            
                                            log(f"      ✓ Menemukan opsi '{text}' untuk {name} (Lapis 1)", "OK")
                                            if not await force_js_click(opt):
                                                await opt.click(timeout=2000)
                                            stage_clicked = True
                                            break
                                    if stage_clicked:
                                        break
                            except Exception:
                                pass
                                
                            if stage_clicked:
                                break
                            await page.wait_for_timeout(200)

                        # LAPISAN 2: Scan manual semua elemen di z-9000 (jika lapis 1 gagal)
                        if not stage_clicked:
                            try:
                                overlay_opts = page.locator("div.fixed.z-9000 div, div.fixed.z-9000 li, div.fixed.z-9000 span, div.fixed.z-9000 p")
                                for i in range(await overlay_opts.count()):
                                    opt = overlay_opts.nth(i)
                                    if await opt.is_visible():
                                        is_input = await opt.evaluate("el => ['INPUT','TEXTAREA','SELECT'].includes(el.tagName)")
                                        if is_input:
                                            continue
                                            
                                        text = (await opt.text_content() or "").strip()
                                        if not text or len(text) < 2 or len(text) > 80:
                                            continue
                                        if "pencarian" in text.lower() or "cari" in text.lower() or "pilih" in text.lower():
                                            continue
                                            
                                        if val.lower() in text.lower() or text.lower() in val.lower():
                                            log(f"      ✓ Menemukan opsi '{text}' untuk {name} (Lapis 2)", "OK")
                                            if not await force_js_click(opt):
                                                await opt.click(timeout=2000)
                                            stage_clicked = True
                                            break
                            except Exception as ex_l2:
                                log(f"      ⚠️ Gagal memindai z-9000 overlay: {ex_l2}", "WARN")
                                
                        # LAPISAN 3: Keyboard Fallback (ArrowDown + Enter)
                        if not stage_clicked:
                            log(f"      ⌨️ Keyboard fallback untuk {name}...", "INFO")
                            try:
                                await search_input.press("ArrowDown")
                                await page.wait_for_timeout(300)
                                await search_input.press("Enter")
                                await page.wait_for_timeout(500)
                                stage_clicked = True
                                log(f"      ✓ Berhasil memilih {name} via keyboard.", "OK")
                            except Exception as ex_key:
                                check_stop_request()
                                log(f"      ⚠️ Gagal memilih {name} via keyboard: {ex_key}", "WARN")
                                
                        if not stage_clicked:
                            log(f"    ❌ Gagal memilih opsi untuk {name}.", "ERR")
                            break
                            
                        completed_stages += 1
                        await page.wait_for_timeout(1200)  # jeda agar overlay stage berikutnya render

                        
                    if completed_stages == 4:
                        log("  ✓ Seluruh tahapan alamat (Provinsi > Kabupaten > Kecamatan > Kelurahan) berhasil diisi otomatis.", "OK")
                        success_auto_address = True
                    else:
                        log(f"  ⚠️ Hanya berhasil mengisi {completed_stages}/4 tahapan alamat. Beralih ke mode manual.", "WARN")
                        
                except Exception as ex_addr:
                    check_stop_request()
                    log(f"  ⚠️ Gagal melakukan pengisian alamat otomatis berjenjang: {ex_addr}", "WARN")

                # Jika pengisian otomatis gagal, tampilkan petunjuk pengisian manual
                check_stop_request()
                if not success_auto_address:
                    if GUI_BRIDGE_ACTIVE and gui_bridge:
                        gui_bridge.request_input("manual_address", {
                            "provinsi": prov_val,
                            "kabupaten": kab_val,
                            "kecamatan": kec_val,
                            "kelurahan": desa_val
                        })
                    else:
                        print()
                        print("  ╔══════════════════════════════════════════╗")
                        print("  ║         ISI ALAMAT DOMISILI MANUAL       ║")
                        print("  ╠══════════════════════════════════════════╣")
                        print(f"  ║  Provinsi       : {prov_val}")
                        print(f"  ║  Kabupaten/Kota : {kab_val}")
                        print(f"  ║  Kecamatan      : {kec_val}")
                        print(f"  ║  Desa/Kelurahan : {desa_val}")
                        print("  ╚══════════════════════════════════════════╝")
                        input("  >> Setelah selesai isi alamat di browser, tekan ENTER untuk lanjut... ")

                # Konfirmasi/Simpan pilihan alamat domisili jika ada tombol konfirmasi
                log("    ⏳ Mengklik tombol konfirmasi/pilih alamat...", "WAIT")
                btn_konfirmasi_domisili = page.locator(
                    "div.modal-content button, "
                    "div[class*='z-9000'] button, "
                    "div[class*='z-1100'] button, "
                    "div.fixed.z-9000 button, "
                    "button"
                ).filter(
                    has_text=re.compile(r"^(Pilih|Simpan|Terapkan|Tutup|Konfirmasi)$", re.IGNORECASE)
                ).first
                
                if await btn_konfirmasi_domisili.count() > 0 and await btn_konfirmasi_domisili.is_visible():
                    await js_click(btn_konfirmasi_domisili)
                else:
                    btn_alt = page.locator("div.modal-content button").last
                    if await btn_alt.count() > 0:
                        await js_click(btn_alt)
                await page.wait_for_timeout(600)

                # 5. Detail Alamat Domisili (Kolom I)
                detail_domisili = str(row.get("Detail Domisili", "")).strip()
                log(f"  ✏️  Mengisi Detail Alamat Domisili: {detail_domisili}", "BOT")
                textarea_domisili = card_modal.locator("textarea#detail-domisili").first
                if await textarea_domisili.count() == 0:
                    textarea_domisili = page.locator("xpath=//textarea[@id='detail-domisili']").first
                await js_click(textarea_domisili)
                await textarea_domisili.fill(detail_domisili)
                await page.wait_for_timeout(300)

                # 6. Klik Selanjutnya (Step 2 - Modal)
                log("  🚀 Mengklik Selanjutnya (Step 2 - Modal)…", "BOT")
                try:
                    btn_selanjutnya = card_modal.locator("button").filter(has_text=re.compile(r"Selanjutnya", re.IGNORECASE)).first
                    await js_click(btn_selanjutnya)
                except Exception:
                    try:
                        await js_click(page.locator(f"xpath={xpath_btn_container}"))
                    except Exception:
                        btn_selanjutnya_modal = page.locator("div.modal-content button, div.z-1100 button, button").filter(has_text=re.compile(r"^\s*Selanjutnya\s*$", re.IGNORECASE)).first
                        await js_click(btn_selanjutnya_modal)
                await page.wait_for_timeout(1000)

                # Cek popup tidak valid setelah Step 2
                if await cek_popup_tidak_valid(page):
                    data_tidak_valid = True
            else:
                log("  ℹ️  Form data pendukung (modal) tidak aktif/sudah terlewati. Melanjutkan...", "INFO")

            # 7. Pilih Individu Terdaftar jika data valid
            if not data_tidak_valid:
                log("  ⏳ Menunggu tombol 'Pilih' (Individu Terdaftar) muncul...", "WAIT")
                xpath_btn_pilih = (
                    "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
                    "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
                    "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
                    "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
                    "/div[@class='p-2']/div/div[3]/div[@class='mt-5']"
                    "/div[@class='overflow-auto table-individu-terdaftar']/table[@class='w-full border-collapse text-sm']/tbody"
                    "/tr[@class='p-2 border-b-solid border-gray-1']/td[@class='p-4 flex justify-center']/div[@class='w-[77px] text-sm']"
                    "/button[@class='w-fill btn-outline-primary h-9']"
                )
                
                try:
                    pilih_btn = page.locator(f"xpath={xpath_btn_pilih}").first
                    if await pilih_btn.count() == 0 or not await is_element_really_visible(pilih_btn):
                        pilih_btn = page.locator("div.table-individu-terdaftar table button, button.btn-outline-primary, button").filter(has_text=re.compile(r"^\s*Pilih\s*$", re.IGNORECASE)).first
                    await pilih_btn.wait_for(state="visible", timeout=7000)
                    log("  🚀 Mengeklik tombol 'Pilih'...", "BOT")
                    await js_click(pilih_btn)
                    await page.wait_for_timeout(600)
                except Exception as e:
                    log(f"  ⚠️ Tombol 'Pilih' tidak ditemukan atau gagal diklik: {e}", "WARN")

                # 8. Klik Daftarkan
                log("  ⏳ Menunggu tombol 'Daftarkan' muncul...", "WAIT")
                xpath_btn_daftarkan = (
                    "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
                    "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
                    "/div[@class='fixed top-0 left-0 z-1000 w-full h-full flex justify-center items-center backdrop-blur-5']"
                    "/div[@class='w-[55%] max-h-full lt-md:w-full lt-lg:w-full relative z-1100 rounded-lg bg-white p-4 shadow-gmail w-[50%] lt-md:w-[80%] lt-sm:w-full overflow-auto']"
                    "/div[@class='p-2']/div/div[3]"
                    "/div[@class='flex justify-between items-center mt-5 lt-sm:flex-col lt-sm:justify-center lt-sm:items-center lt-sm:gap-2 w-full']"
                    "/div[@class='flex lt-sm:flex-col gap-2 lt-sm:w-full']/div[@class='w-auto lt-sm:w-full']"
                    "/button[@class='w-fill btn-fill-primary-v2 h-11']"
                )
                try:
                    daftarkan_btn = page.locator(f"xpath={xpath_btn_daftarkan}").first
                    if await daftarkan_btn.count() == 0 or not await is_element_really_visible(daftarkan_btn):
                        daftarkan_btn = page.locator("button").filter(has_text=re.compile(r"Daftarkan", re.IGNORECASE)).first
                    await daftarkan_btn.wait_for(state="visible", timeout=5000)
                    log("  🚀 Mengeklik tombol 'Daftarkan'...", "BOT")
                    await js_click(daftarkan_btn)
                    await page.wait_for_timeout(1000)
                except Exception as e:
                    log(f"  ⚠️ Tombol 'Daftar' tidak ditemukan atau gagal diklik: {e}", "WARN")

                # 9. Tunggu sukses modal "berhasil"
                log("  ⏳ Menunggu notifikasi 'Berhasil' muncul...", "WAIT")
                xpath_txt_berhasil = (
                    "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
                    "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
                    "/div[@class='fixed top-0 left-0 w-full h-full flex justify-center items-center backdrop-blur-5 z-1000']"
                    "/div[@class='w-[30%] max-h-full relative z-1000 rounded-lg bg-white p-4 shadow-gmail lt-md:w-[80%] lt-sm:w-full overflow-auto']"
                    "/div[@class='p-2']/div[@class='flex flex-col justify-center items-center gap-3 text-[20px] font-bold mb-7']/div[@class='pb-2']"
                )
                try:
                    berhasil_txt = page.locator(f"xpath={xpath_txt_berhasil}").first
                    if await berhasil_txt.count() == 0 or not await is_element_really_visible(berhasil_txt):
                        berhasil_txt = page.locator("div.fixed.z-1000 div.rounded-lg.bg-white, div.fixed.z-1100 div.rounded-lg.bg-white").filter(has_text=re.compile(r"Berhasil", re.IGNORECASE)).first
                    await berhasil_txt.wait_for(state="visible", timeout=15000)
                    log("  ✓ Notifikasi sukses 'Berhasil' terdeteksi.", "OK")
                except Exception as e:
                    log(f"  ⚠️ Gagal mendeteksi teks 'Berhasil' (timeout): {e}", "WARN")

                # 10. Klik Tutup sukses modal
                log("  🚀 Mengeklik tombol 'Tutup'...", "BOT")
                xpath_btn_tutup = (
                    "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
                    "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
                    "/div[@class='fixed top-0 left-0 w-full h-full flex justify-center items-center backdrop-blur-5 z-1000']"
                    "/div[@class='w-[30%] max-h-full relative z-1000 rounded-lg bg-white p-4 shadow-gmail lt-md:w-[80%] lt-sm:w-full overflow-auto']"
                    "/div[@class='p-2']/div[@class='flex justify-center gap-2 font-600 text-base lt-md:flex-col mt-4']/div[@class='w-auto']"
                    "/button[@class='w-fill btn-fill-primary h-11']"
                )
                try:
                    tutup_btn = page.locator(f"xpath={xpath_btn_tutup}").first
                    if await tutup_btn.count() == 0 or not await is_element_really_visible(tutup_btn):
                        tutup_btn = page.locator("div.fixed.z-1000 div.rounded-lg.bg-white button, div.fixed.z-1100 div.rounded-lg.bg-white button, button").filter(has_text=re.compile(r"^\s*Tutup\s*$", re.IGNORECASE)).first
                    await tutup_btn.wait_for(state="visible", timeout=5000)
                    await js_click(tutup_btn)
                    await page.wait_for_timeout(2000)
                except Exception as e:
                    log(f"  ⚠️ Gagal mengeklik tombol 'Tutup': {e}", "WARN")

                # 11. Cari Pasien by Name & Tandai Hadir
                log("  🔎 Melakukan pencarian pasien untuk Tandai Hadir...", "BOT")
                xpath_dropdown_filter = (
                    "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
                    "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
                    "/div[@class='w-full flex lt-sm:flex-col lt-lg:flex-col gap-1 pt-2 pb-4']"
                    "/div[@class='flex lg:w-[85%] lt:sm:w-full lt-lg:w-full']"
                    "/div[@class='border-rd-r-0 w-[20%] lt-lg:w-full lt-sm:w[50%]']"
                    "/div[@class='relative text-black border-1 border-solid font-medium flex border-gray-3 focus-within:border-black border-rd-l-lg']"
                    "/div[@class='h-[2.9rem] w-full flex gap-2 cursor-pointer items-center justify-start overflow-hidden border-none bg-transparent pl-4 text-sm focus:outline-none text-black']"
                )
                
                try:
                    # Click dropdown filter
                    dropdown_filter = page.locator(".icon.w-6.h-6.transform").first
                    if await dropdown_filter.count() == 0 or not await is_element_really_visible(dropdown_filter):
                        dropdown_filter = page.locator(f"xpath={xpath_dropdown_filter}").first
                    await js_click(dropdown_filter)
                    await page.wait_for_timeout(600)
                    
                    # Select "Nama" option
                    option_nama = page.get_by_text("Nama", exact=True).first
                    if await option_nama.count() == 0:
                        option_nama = page.locator("div.modal-content button, div.z-9000 button, button").filter(
                            has_text=re.compile(r"^\s*Nama\s*$", re.IGNORECASE)
                        ).first
                    await js_click(option_nama)
                    await page.wait_for_timeout(600)
                except Exception as e:
                    log(f"  ⚠️ Gagal memilih filter 'Nama': {e}", "WARN")

                # Input nama to search input
                try:
                    search_input = page.get_by_placeholder("Masukkan nama").first
                    if await search_input.count() == 0 or not await is_element_really_visible(search_input):
                        search_input = page.locator("xpath=//input[@id='searchNik']").first
                    await js_click(search_input)
                    await search_input.fill(nama)
                    await search_input.press("Enter")
                    await page.wait_for_timeout(2000)
                except Exception as e:
                    log(f"  ⚠️ Gagal mencari nama pasien: {e}", "WARN")

                # Klik konfirmasi tanggal
                log("  ⏳ Menunggu tombol 'Konfirmasi Hadir' atau 'Konfirmasi Tanggal' di tabel...", "WAIT")
                xpath_btn_konfirmasi_tgl = (
                    "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
                    "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
                    "/div[@class='relative w-full h-full overflow-hidden']"
                    "/div[@class='relative z-[10] overflow-hidden shadow-md sm:rounded-lg']"
                    "/div[@class='overflow-auto table-individu-terdaftar']/table[@class='w-full border-collapse text-sm']/tbody"
                    "/tr[@class='p-2 relative']/td[@class='p-4 flex justify-center min-w-[305px]']"
                    "/div[@class='flex w-full gap-2 lt-sm:flex-col']/div[@class='w-[50%] lt-sm:w-full'][1]"
                    "/div[@class='w-full text-sm']/button[@class='w-fill btn-fill-primary h-11']"
                )
                try:
                    konfirmasi_btn = page.get_by_role("button", name="Konfirmasi Hadir").first
                    if await konfirmasi_btn.count() == 0 or not await is_element_really_visible(konfirmasi_btn):
                        konfirmasi_btn = page.get_by_role("button", name="Konfirmasi Tanggal").first
                    if await konfirmasi_btn.count() == 0 or not await is_element_really_visible(konfirmasi_btn):
                        konfirmasi_btn = page.locator(f"xpath={xpath_btn_konfirmasi_tgl}").first
                    await konfirmasi_btn.wait_for(state="visible", timeout=10000)
                    await js_click(konfirmasi_btn)
                    await page.wait_for_timeout(1000)
                except Exception as e:
                    log(f"  ⚠️ Gagal klik tombol Konfirmasi: {e}", "WARN")

                # Pilih tanggal hari ini (di modal Konfirmasi Hadir)
                today_day = str(datetime.now().day)
                log(f"  🗓  Memilih tanggal konfirmasi (Hari ini): {today_day}", "BOT")
                try:
                    tgl_btn = page.get_by_role("button", name=today_day, exact=True).first
                    await js_click(tgl_btn)
                    await page.wait_for_timeout(400)
                except Exception as e:
                    log(f"  ⚠️ Gagal memilih tanggal {today_day}: {e}", "WARN")

                # Centang persetujuan/verify
                log("  ✏️  Mencentang persetujuan/verify...", "BOT")
                try:
                    verify_checkbox = page.locator("#verify").nth(1)
                    if await verify_checkbox.count() == 0 or not await is_element_really_visible(verify_checkbox):
                        verify_checkbox = page.locator("#verify").first
                    if await verify_checkbox.count() == 0 or not await is_element_really_visible(verify_checkbox):
                        verify_checkbox = page.locator("xpath=//div[@id='verify']").first
                    await js_click(verify_checkbox)
                    await page.wait_for_timeout(400)
                except Exception as e:
                    log(f"  ⚠️ Gagal mencentang checkbox verify: {e}", "WARN")

                # Klik Hadir
                log("  🚀 Mengeklik tombol 'Hadir'...", "BOT")
                xpath_btn_hadir = (
                    "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
                    "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
                    "/div[@class='fixed top-0 left-0 w-full h-full flex justify-center items-center backdrop-blur-5 z-1000']"
                    "/div[@class='w-[45%] max-h-full z-2000 relative z-1000 rounded-lg bg-white p-4 shadow-gmail lt-md:w-[80%] lt-sm:w-full overflow-auto']"
                    "/div[@class='p-2']/div[@class='font-600 text-[18px] flex gap-2']/div[@class='w-[50%]']"
                    "/button[@class='w-fill btn-fill-primary h-11']"
                )
                try:
                    hadir_btn = page.get_by_role("button", name="Hadir", exact=True).first
                    if await hadir_btn.count() == 0 or not await is_element_really_visible(hadir_btn):
                        hadir_btn = page.locator(f"xpath={xpath_btn_hadir}").first
                    await hadir_btn.wait_for(state="visible", timeout=5000)
                    await js_click(hadir_btn)
                    await page.wait_for_timeout(2000)
                except Exception as e:
                    log(f"  ⚠️ Gagal klik tombol 'Hadir': {e}", "WARN")

                # 12. Tunggu popup berhasil hadir dan tutup
                log("  ⏳ Menunggu notifikasi 'Berhasil Hadir'...", "WAIT")
                xpath_txt_berhasil_hadir = (
                    "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
                    "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
                    "/div[@class='fixed top-0 left-0 w-full h-full flex justify-center items-center backdrop-blur-5 z-1000']"
                    "/div[@class='w-[30%] max-h-full relative z-1000 rounded-lg bg-white p-4 shadow-gmail lt-md:w-[80%] lt-sm:w-full overflow-auto']"
                    "/div[@class='p-2']/div[@class='flex flex-col justify-center items-center gap-3 text-[20px] font-bold mb-7']/div[@class='pb-2']"
                )
                xpath_btn_tutup_hadir = (
                    "//div[@id='__nuxt']/main/div[@class='h-full min-h-screen w-full  lt-md:block']"
                    "/div[@class='h-auto overflow-hidden bg-white']/section[@class='relative ml-64 h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[@class='relative h-full min-h-screen overflow-hidden lt-md:ml-0 lt-md:w-full']"
                    "/div/div[2]/div[@class='mt-6 px-6 pb-4']/div[@class='w-full']"
                    "/div[@class='fixed top-0 left-0 w-full h-full flex justify-center items-center backdrop-blur-5 z-1000']"
                    "/div[@class='w-[30%] max-h-full relative z-1000 rounded-lg bg-white p-4 shadow-gmail lt-md:w-[80%] lt-sm:w-full overflow-auto']"
                    "/div[@class='p-2']/div[@class='flex justify-center gap-2 font-600 text-base w-full mt-4']/div[@class='w-auto']"
                    "/button[@class='w-fill btn-fill-primary h-11']"
                )
                try:
                    berhasil_hadir_txt = page.locator(f"xpath={xpath_txt_berhasil_hadir}").first
                    if await berhasil_hadir_txt.count() == 0 or not await is_element_really_visible(berhasil_hadir_txt):
                        berhasil_hadir_txt = page.locator("div.fixed.z-1000 div.rounded-lg.bg-white, div.fixed.z-1100 div.rounded-lg.bg-white").filter(has_text=re.compile(r"Berhasil", re.IGNORECASE)).first
                    await berhasil_hadir_txt.wait_for(state="visible", timeout=15000)
                    log("  ✓ Notifikasi sukses 'Berhasil Hadir' terdeteksi.", "OK")
                    
                    # Klik tutup berhasil hadir
                    log("  🚀 Mengeklik tombol 'Tutup' (Berhasil Hadir)...", "BOT")
                    tutup_hadir_btn = page.get_by_role("button", name="Tutup").first
                    if await tutup_hadir_btn.count() == 0 or not await is_element_really_visible(tutup_hadir_btn):
                        tutup_hadir_btn = page.locator(f"xpath={xpath_btn_tutup_hadir}").first
                    await tutup_hadir_btn.wait_for(state="visible", timeout=5000)
                    await js_click(tutup_hadir_btn)
                    await page.wait_for_timeout(2000)
                except Exception as e:
                    log(f"  ⚠️ Gagal mendeteksi/menutup 'Berhasil Hadir': {e}", "WARN")

        except Exception as e:
            log(f"  ⚠️ Gagal memproses konfirmasi kehadiran: {e}", "WARN")

    # ── Setelah submit ────────────────────────────────────────
    await page.wait_for_timeout(1000)
    if data_tidak_valid:
        log(f"[{nomor}/{total}] {nama} data TIDAK VALID.", "WARN")
    else:
        log(f"[{nomor}/{total}] {nama} selesai diproses formnya.", "BOT")
    await screenshot(page, f"form_done_{nomor}_{nama[:12].replace(' ','_')}")

    # Konfirmasi manual status akhir Excel
    if data_tidak_valid:
        print("\n  ⚠️⚠️⚠️ WARNING: DATA TIDAK VALID TERDETEKSI! ⚠️⚠️⚠️")
        print("  Silakan ketik [2] atau langsung tekan ENTER untuk SKIP pasien ini dan lanjut ke baris berikutnya.")

    default_pilihan = "2" if data_tidak_valid else "1"
    if GUI_BRIDGE_ACTIVE and gui_bridge:
        pilihan = gui_bridge.request_input("decision", {
            "nomor": nomor,
            "nama": nama,
            "data_tidak_valid": data_tidak_valid,
            "default_pilihan": default_pilihan
        })
    else:
        print("\n  >> Tentukan status pasien ini di Excel:")
        print("     [1] Lolos (Hijau - Berhasil)")
        print("     [2] Skip  (Kuning - Lewati)")
        print("     [3] Gagal (Merah - Error/Gagal)")
        print("     [r] Ulangi (Reload & Ulangi baris ini)")
        print("     [q] Berhenti (Quit)")
        
        pilihan = input(f"  >> Masukkan pilihan [1/2/3/r/q, default: {default_pilihan}]: ").strip().lower()
        if pilihan == "":
            pilihan = default_pilihan

    if pilihan == "2":
        return "skip"
    elif pilihan == "3":
        return "gagal"
    elif pilihan == "r":
        return "ulang"
    elif pilihan == "q":
        return "quit_success"
    else:
        return "ok"


# ============================================================
#  MAIN — satu data contoh
# ============================================================

async def main(username_arg=None, password_arg=None, mulai_baris_arg=None, excel_file_arg=None, headless_arg=None):
    global FILE_EXCEL, HEADLESS
    
    if excel_file_arg:
        FILE_EXCEL = Path(excel_file_arg)
    if headless_arg is not None:
        HEADLESS = headless_arg
        
    print()
    print(f"  {GARIS2}")
    print(f"  🤖 Arzachel Bot CKG — sehatindonesiaku.kemkes.go.id")
    print(f"  {GARIS2}")
    print()

    # ── Pilihan Login ────────────────────────────────────────
    if username_arg is not None and password_arg is not None:
        username_input = username_arg
        password_input = password_arg
    else:
        print("  === PILIHAN LOGIN ===")
        print("  [1] Gunakan username & password default")
        print("  [2] Masukkan username & password lain")
        pilihan_login = input("  >> Masukkan pilihan [1/2, default: 1]: ").strip()
        
        username_input = None
        password_input = None
        if pilihan_login == "2":
            username_input = input("  >> Masukkan Email: ").strip()
            password_input = input("  >> Masukkan Kata sandi: ").strip()

    # ── Baca Excel & Deteksi Baris Terakhir ──────────────────
    if not FILE_EXCEL.exists():
        log(f"File tidak ditemukan: {FILE_EXCEL}", "ERR")
        sys.exit(1)

    # Deteksi input terakhir sebelum membaca/memfilter data
    last_no, last_name, last_status = temukan_input_terakhir()
    if last_no is not None:
        msg = f"Input terakhir yang terdeteksi di Excel: No. {last_no} (Nama: {last_name}) -> Status: {last_status}"
        if "BERHASIL" in last_status:
            log(msg, "OK")
        elif "SKIP" in last_status:
            log(msg, "WARN")
        else:
            log(msg, "ERR")
        suggested_start = int(last_no) + 1
    else:
        log("Belum ada data yang ditandai diproses di Excel.", "INFO")
        suggested_start = 1

    if mulai_baris_arg is not None:
        mulai_baris = mulai_baris_arg
    else:
        mulai_baris_str = input(f"  >> Mulai dari No. berapa di Excel? [Default: {suggested_start}]: ").strip()
        if mulai_baris_str == "":
            mulai_baris = suggested_start
        else:
            try:
                mulai_baris = int(mulai_baris_str)
            except ValueError:
                mulai_baris = suggested_start
                log(f"Pilihan tidak valid, menggunakan default: {mulai_baris}", "WARN")

    # Baca Excel, ambil kolom No, Nama Lengkap, NIK, Jenis Kelamin, Tanggal Lahir, Status Perkawinan, Pekerjaan, Alamat, Alamat Domisili
    df = pd.read_excel(FILE_EXCEL, usecols=[0, 1, 2, 3, 5, 6, 7, 8, 9], dtype={2: str})
    df.columns = ["No", "Nama Lengkap", "NIK", "Jenis Kelamin", "Tanggal Lahir", "Status Pernikahan", "Pekerjaan", "Detail Domisili", "Alamat Domisili"]
    df = df.dropna(subset=["NIK"])          # hapus baris kosong
    df = df[df["No"] >= mulai_baris]        # resume dari baris tertentu
    df = df.reset_index(drop=True)

    total = len(df)
    if total == 0:
        log("Tidak ada data pasien yang tersisa untuk diproses.", "WARN")
        sys.exit(0)

    log(f"Total data yang akan diproses: {total} pasien (mulai dari No. {mulai_baris})", "OK")

    # Preview 3 baris pertama
    for _, r in df.head(3).iterrows():
        # Parsing NIK secara aman untuk preview
        nik_raw = str(r["NIK"]).strip()
        if nik_raw.endswith(".0"):
            nik_raw = nik_raw[:-2]
        nik = "".join(re.findall(r"\d+", nik_raw)).zfill(16)
        tgl = pd.Timestamp(r["Tanggal Lahir"]).strftime("%d-%m-%Y")
        log(f"  No.{int(r['No']):>2} | {r['Nama Lengkap']:<30} | {nik} | {r['Jenis Kelamin']} | {tgl}", "INFO")
    if total > 3:
        log(f"  ... dan {total - 3} data lainnya", "INFO")
    print()

    # ── Jalankan browser ─────────────────────────────────────
    async with async_playwright() as pw:
        browser: Browser = await pw.chromium.launch(
            headless=HEADLESS,
            args=["--start-maximized"],
        )
        context: BrowserContext = await browser.new_context(
            viewport=None,
            locale="id-ID",
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
        )
        page: Page = await context.new_page()
        page.set_default_navigation_timeout(30000)
        page.set_default_timeout(8000)
        
        global current_page
        current_page = page
        
        # Jalankan background task untuk memantau sinyal stop/skip dari GUI
        asyncio.create_task(poll_gui_signals())

        berhasil  = 0
        gagal     = 0
        dilewati  = 0
        gagal_list = []

        try:
            # ── Login sekali di awal ──────────────────────────
            await login(page, username=username_input, password=password_input)

            print(f"\n  {GARIS2}")
            print(f"  🤖 Mulai input batch — {total} pasien")
            print(f"  {GARIS2}\n")

            # ── Loop semua pasien ─────────────────────────────
            for urutan, (_, baris) in enumerate(df.iterrows(), start=1):
                check_stop_request()
                nomor = int(baris["No"])
                nama  = str(baris["Nama Lengkap"]).strip()
                status_pasien = None
                break_outer = False
                excel_sudah_ditulis = False  # flag untuk mencegah double-write ke Excel
                while True:
                    if page.is_closed():
                        log("  🌐 Membuat halaman browser baru untuk melanjutkan...", "INFO")
                        page = await context.new_page()
                        current_page = page
                        page.set_default_navigation_timeout(30000)
                        page.set_default_timeout(8000)
                        try:
                            await page.goto(BASE_URL, wait_until="domcontentloaded")
                        except Exception as eg:
                            log(f"  ⚠️ Gagal navigasi ke halaman utama: {eg}", "WARN")
                            
                    check_stop_request()
                    try:
                        # Cek sesi sebelum tiap pasien — auto relogin jika expired
                        await cek_sesi_berakhir(page, username=username_input, password=password_input)

                        # Buka form Daftar Baru setiap iterasi
                        await buka_form_daftar_baru(page)

                        # Bungkus input_pasien dalam asyncio.Task agar bisa di-cancel
                        # oleh poll_gui_signals saat user menekan tombol Skip.
                        # asyncio.CancelledError TIDAK tertangkap oleh except Exception,
                        # sehingga skip selalu berhasil terlepas dari try-except lokal
                        # di dalam input_pasien.
                        global _current_input_task
                        _current_input_task = asyncio.ensure_future(
                            input_pasien(page, baris.to_dict(), urutan, total)
                        )
                        hasil = await _current_input_task
                        _current_input_task = None

                        if hasil == "ok":
                            status_pasien = "berhasil"
                        elif hasil == "quit_success":
                            status_pasien = "berhasil_quit"
                            break_outer = True
                            break
                        elif hasil == "skip":
                            status_pasien = "dilewati"
                        elif hasil == "gagal":
                            status_pasien = "gagal"
                        elif hasil == "ulang":
                            log("  🔄 User memilih ulang baris ini dari menu status. Mereload halaman...", "WARN")
                            await page.reload()
                            await page.wait_for_timeout(3000)
                            continue
                        elif hasil == "quit":
                            status_pasien = "quit"
                            break_outer = True
                            break

                    except StopBotException:
                        status_pasien = "quit"
                        break_outer = True
                        break

                    except SkipPasien as sp:
                        status_pasien = "dilewati"
                        log(f"  ⏭  Skip [{nomor}] {nama}: {sp}", "WARN")
                        await screenshot(page, f"skip_{nomor}_{nama[:10].replace(' ','_')}")
                        break

                    except asyncio.CancelledError:
                        # Skip ditekan user via tombol GUI — ini jalur utama skip yang pasti berhasil.
                        # CancelledError tidak bisa ditangkap oleh except Exception di mana pun.
                        _current_input_task = None
                        if gui_bridge:
                            gui_bridge.skip_active = False
                        status_pasien = "dilewati"
                        dilewati += 1
                        excel_sudah_ditulis = True  # tandai agar post-loop tidak double-write
                        log(f"  ⏭  Skip [{nomor}] {nama}: Tombol Skip ditekan oleh user.", "WARN")
                        update_excel_row_color(nomor, "yellow")
                        # Buka form baru agar browser sudah siap untuk pasien berikutnya
                        try:
                            if not page.is_closed():
                                await buka_form_daftar_baru(page)
                        except Exception:
                            pass
                        break_outer = False
                        break

                    except Exception as e:
                        # Cek apakah ini akibat request Stop atau Skip dari GUI
                        if GUI_BRIDGE_ACTIVE and gui_bridge:
                            if not gui_bridge.running:
                                status_pasien = "quit"
                                break_outer = True
                                break
                            if getattr(gui_bridge, "skip_requested", False) or getattr(gui_bridge, "skip_active", False):
                                gui_bridge.skip_requested = False
                                gui_bridge.skip_active = False
                                status_pasien = "dilewati"
                                log(f"  ⏭  Skip [{nomor}] {nama} (Skip diminta oleh user melalui GUI)", "WARN")
                                break

                        if "StopBotException" in str(type(e)):
                            status_pasien = "quit"
                            break_outer = True
                            break
                            
                        log(f"  ✗ Gagal [{nomor}] {nama}: {e}", "ERR")
                        await screenshot(page, f"error_{nomor}_{nama[:10].replace(' ','_')}")
                        
                        if GUI_BRIDGE_ACTIVE and gui_bridge:
                            log("  ⚠️ Mengaktifkan Asisten Pintar: Meminta bantuan user untuk ambil kendali...", "WARN")
                            pilihan = gui_bridge.request_input("take_control", {
                                "nomor": nomor,
                                "nama": nama,
                                "error_msg": str(e)
                            })
                            if pilihan == "resume":
                                log("  🔄 User meminta untuk melanjutkan pengisian otomatis (Resume)...", "WARN")
                                # Lanjut tanpa reload halaman, langsung coba isi sisa field
                                continue
                            elif pilihan == "ulang":
                                log("  🔄 User meminta untuk mengulangi baris ini dari awal. Mereload halaman...", "WARN")
                                await page.reload()
                                await page.wait_for_timeout(3000)
                                continue
                            elif pilihan == "ok" or pilihan == "berhasil":
                                log("  ✓ User menandai pasien ini BERHASIL (Lolos) secara manual.", "OK")
                                status_pasien = "berhasil"
                                break
                            elif pilihan == "skip":
                                status_pasien = "dilewati"
                                break
                            elif pilihan == "gagal":
                                status_pasien = "gagal"
                                break
                            elif pilihan == "quit":
                                status_pasien = "quit"
                                break_outer = True
                                break
                        else:
                            break

                    # Jeda sebelum pasien berikutnya/reload
                    try:
                        if not page.is_closed():
                            await page.wait_for_timeout(JEDA_ANTAR_DATA)
                    except Exception:
                        pass

                    # Konfirmasi manual sebelum lanjut ke baris berikutnya
                    if urutan < total:
                        if GUI_BRIDGE_ACTIVE and gui_bridge:
                            if gui_bridge.auto_advance:
                                break
                            else:
                                pilihan_lanjut = gui_bridge.request_input("next_patient_pause", nomor)
                                if pilihan_lanjut == "r":
                                    log("  🔄 User memilih reload halaman. Mereload halaman...", "WARN")
                                    await page.reload()
                                    await page.wait_for_timeout(3000)
                                    continue
                                elif pilihan_lanjut == "q":
                                    log("Bot dihentikan oleh user.", "WARN")
                                    break_outer = True
                                    break
                                else:
                                    break
                        else:
                            pilihan_lanjut = input(f"\n  [Selesai Baris {nomor}] Tekan ENTER untuk lanjut ke pasien berikutnya, ketik 'r' untuk reload halaman & ulangi baris ini, atau 'q' untuk keluar: ").strip().lower()
                            if pilihan_lanjut == "r":
                                log("  🔄 User memilih reload halaman. Mereload halaman...", "WARN")
                                await page.reload()
                                await page.wait_for_timeout(3000)
                                continue
                            elif pilihan_lanjut == "q":
                                log("Bot dihentikan oleh user.", "WARN")
                                break_outer = True
                                break
                            else:
                                break
                    else:
                        break

                # Update status akhir dan excel setelah keluar dari loop retry.
                # Jalur CancelledError sudah menulis Excel dan menghitung statistik sendiri
                # (excel_sudah_ditulis = True), sehingga blok ini dilewati untuk pasien tersebut.
                if not excel_sudah_ditulis:
                    if status_pasien == "berhasil":
                        berhasil += 1
                        update_excel_row_color(nomor, "green")
                    elif status_pasien == "berhasil_quit":
                        berhasil += 1
                        update_excel_row_color(nomor, "green")
                    elif status_pasien == "dilewati":
                        dilewati += 1
                        update_excel_row_color(nomor, "yellow")
                    elif status_pasien == "gagal":
                        gagal += 1
                        update_excel_row_color(nomor, "red")
                        gagal_list.append(f"No.{nomor} {nama}: Diproses Gagal/Error")
                    elif status_pasien == "quit":
                        update_excel_row_color(nomor, "red")

                if break_outer:
                    break

        except Exception as e:
            log(f"Error fatal: {e}", "ERR")
            await screenshot(page, "error_fatal")
        finally:
            if not GUI_BRIDGE_ACTIVE:
                input("\nTekan ENTER untuk menutup browser…")
            await browser.close()

    # ── Laporan akhir ─────────────────────────────────────────
    print()
    log("═" * 55, "INFO")
    log(f"  ✅ Berhasil : {berhasil}", "OK")
    log(f"  ⏭  Dilewati : {dilewati}", "WARN")
    log(f"  ❌ Gagal    : {gagal}", "ERR" if gagal else "INFO")
    log(f"  📋 Total    : {total}", "INFO")
    if gagal_list:
        log("Data yang gagal:", "WARN")
        for g in gagal_list:
            log(f"  - {g}", "WARN")
    log("═" * 55, "INFO")


if __name__ == "__main__":
    asyncio.run(main())
